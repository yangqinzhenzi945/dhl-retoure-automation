import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import readline from "node:readline/promises";
import ExcelJS from "exceljs";
import { chromium } from "playwright";
import {
  TRACKING_HEADER,
  asText,
  headerMap,
  outputPathFor,
  parseShipmentResult,
  rowToRecord,
  validateRecord,
} from "./helpers.js";

const FORM_URL = "https://geschaeftskunden.dhl.de/content/dhl-rpi/gw/rpcustomerweb/OrderCustService.action";
const RESULT_SUCCESS = "Ihre Retourensendung wurde erfolgreich beauftragt!";
const ERROR_SHEET_NAME = "自动化错误记录";

function ensureErrorSheet(workbook) {
  const sheet = workbook.getWorksheet(ERROR_SHEET_NAME) ?? workbook.addWorksheet(ERROR_SHEET_NAME);
  if (sheet.rowCount === 0) {
    sheet.addRow(["时间", "Excel 行号", "Sendungsreferenz", "Kundenreferenz", "错误类型", "错误详情"]);
    sheet.getRow(1).font = { bold: true };
    sheet.columns = [18, 12, 22, 28, 18, 70].map((width) => ({ width }));
  }
  return sheet;
}

function addError(errorSheet, record, type, details) {
  errorSheet.addRow([
    new Date().toLocaleString("zh-CN", { hour12: false }),
    record.excelRow,
    record.shipmentReference,
    record.customerReference,
    type,
    details,
  ]);
}

function parseArguments(argv) {
  const args = { input: "", output: "", dryRun: false, overwrite: false, headless: false };
  for (let i = 0; i < argv.length; i += 1) {
    const value = argv[i];
    if (value === "--input") args.input = argv[++i] ?? "";
    else if (value === "--output") args.output = argv[++i] ?? "";
    else if (value === "--dry-run") args.dryRun = true;
    else if (value === "--overwrite") args.overwrite = true;
    else if (value === "--headless") args.headless = true;
    else if (!value.startsWith("--") && !args.input) args.input = value;
    else throw new Error(`未知参数：${value}`);
  }
  return args;
}

async function askForInput(terminal) {
  const answer = await terminal.question("请拖入 Excel 文件并按 Enter：\n> ");
  return answer.trim().replace(/^"|"$/g, "");
}

async function waitForDhlForm(page) {
  await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
  console.log("\n请在打开的 Chrome 中登录 DHL。登录后脚本会自动继续。\n");
  const deadline = Date.now() + 10 * 60 * 1000;
  while (Date.now() < deadline) {
    const frame = page.frames().find((candidate) => candidate.url().includes("OrderCustService.action"));
    if (frame && (await frame.getByText("Retoure beauftragen", { exact: true }).count())) return frame;
    await page.waitForTimeout(1000);
  }
  throw new Error("等待 DHL 登录超时（10 分钟）。");
}

async function openBlankOrder(page, frame) {
  const bodyText = await frame.locator("body").innerText();
  if (bodyText.includes(RESULT_SUCCESS)) {
    await frame.getByRole("button", { name: "Neues Retourenlabel beauftragen", exact: true }).click();
    await page.waitForTimeout(600);
    return page.frames().find((candidate) => candidate.url().includes("OrderCustService.action")) ?? frame;
  }
  return frame;
}

async function fillOrder(page, frame, record) {
  await frame.locator('input[name="shipmentReference"]').fill(record.shipmentReference);
  await frame.locator('input[name="customerReference"]').fill(record.customerReference);
  await frame.locator('input[name="address.sender.name1"]').fill(record.name1);
  await frame.locator('input[name="address.sender.name2"]').fill(record.name2);
  await frame.locator('input[name="address.sender.name3"]').fill(record.name3);
  await frame.locator('input[name="address.sender.plz"]').fill(record.postalCode);
  await frame.locator('input[name="address.sender.city"]').fill(record.city);
  await frame.locator('input[name="address.sender.street"]').fill(record.street);
  await frame.locator('input[name="address.sender.streetNumber"]').fill(record.streetNumber);
  await frame.locator('input[name="address.sender.email"]').fill(record.email);

  await frame.locator("#receiverSelect").click();
  await frame.getByText(record.receiver, { exact: true }).click();
  await page.waitForTimeout(200);

  const goGreenPlus = frame.locator("#goGreenPlusEnabled");
  if (await goGreenPlus.isChecked()) {
    await frame.locator('label[for="goGreenPlusEnabled"]').click();
    await page.waitForTimeout(100);
  }
  if (await goGreenPlus.isChecked()) throw new Error("GoGreen Plus 无法取消勾选。");

  const actualReference = await frame.locator('input[name="shipmentReference"]').inputValue();
  if (actualReference !== record.shipmentReference) throw new Error("网页发送参考号与 Excel 不一致。");
  const submit = frame.getByRole("button", { name: "Retoure beauftragen", exact: true });
  if (!(await submit.isEnabled())) throw new Error("DHL 提交按钮不可用，请检查必填字段。");
  return submit;
}

async function submitAndRead(page, frame, submit, record) {
  await submit.scrollIntoViewIfNeeded();
  await submit.click();
  await frame.getByText(RESULT_SUCCESS, { exact: false }).waitFor({ state: "visible", timeout: 30000 });
  const result = parseShipmentResult(await frame.locator("body").innerText());
  if (!result.shipmentNumber) throw new Error("结果页没有找到 Sendungsnummer。");
  if (result.customerReference !== record.customerReference) {
    throw new Error(`结果页客户参考号不一致：${result.customerReference}`);
  }
  return result.shipmentNumber;
}

async function main() {
  const args = parseArguments(process.argv.slice(2));
  const terminal = readline.createInterface({ input: process.stdin, output: process.stdout });
  let context;
  try {
    const inputPath = path.resolve(args.input || (await askForInput(terminal)));
    await fs.access(inputPath);
    const outputPath = path.resolve(args.overwrite ? inputPath : (args.output || outputPathFor(inputPath)));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputPath);
    const sheet = workbook.worksheets[0];
    if (!sheet) throw new Error("Excel 中没有工作表。");
    const columns = headerMap(sheet.getRow(1));
    const errorSheet = ensureErrorSheet(workbook);
    let trackingColumn = columns.get(TRACKING_HEADER);
    if (!trackingColumn) {
      trackingColumn = 12;
      sheet.getCell(1, trackingColumn).value = TRACKING_HEADER;
      sheet.getCell(1, trackingColumn).style = { ...sheet.getCell(1, 11).style };
      sheet.getColumn(trackingColumn).width = 18;
    }

    const pending = [];
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      const hasAnyInput = [...columns.values()].some((column) => asText(row.getCell(column).value));
      if (!hasAnyInput) continue;
      if (asText(row.getCell(trackingColumn).value)) continue;
      pending.push(rowToRecord(row, columns));
    }
    if (!pending.length) {
      console.log("没有待处理行：L 列已有运单号，或表格没有数据。");
      return;
    }

    console.log(`输入文件：${inputPath}`);
    console.log(`输出文件：${outputPath}`);
    console.log(`待创建：${pending.length} 张退货标签`);
    if (!args.dryRun) {
      const answer = await terminal.question("这会创建真实 DHL 退货委托。请输入 CREATE 继续：");
      if (answer.trim() !== "CREATE") {
        console.log("已取消，未创建任何标签。");
        return;
      }
    }

    const profileDir = path.join(os.homedir(), ".dhl-retoure-automation", "chrome-profile");
    context = await chromium.launchPersistentContext(profileDir, {
      channel: "chrome",
      headless: args.headless,
      viewport: null,
      args: ["--start-maximized"],
    });
    const pages = context.pages();
    const page = pages[0] ?? (await context.newPage());
    let frame = await waitForDhlForm(page);
    let successCount = 0;
    let errorCount = 0;

    for (let index = 0; index < pending.length; index += 1) {
      const record = pending[index];
      console.log(`[${index + 1}/${pending.length}] Excel 第 ${record.excelRow} 行：${record.customerReference || "（无客户参考号）"}`);
      const validationErrors = validateRecord(record);
      if (validationErrors.length) {
        addError(errorSheet, record, "数据校验失败", validationErrors.join("；"));
        errorCount += 1;
        await workbook.xlsx.writeFile(outputPath);
        console.error(`  ✗ 已跳过：${validationErrors.join("；")}`);
        continue;
      }
      try {
        frame = await openBlankOrder(page, frame);
        const submit = await fillOrder(page, frame, record);
        if (args.dryRun) {
          console.log("试运行完成：第一条有效记录已填好，尚未点击 Retoure beauftragen。");
          await terminal.question("检查浏览器后按 Enter 关闭脚本。");
          return;
        }
        const shipmentNumber = await submitAndRead(page, frame, submit, record);
        const cell = sheet.getCell(record.excelRow, trackingColumn);
        cell.value = Number(shipmentNumber);
        cell.numFmt = "0";
        successCount += 1;
        await workbook.xlsx.writeFile(outputPath);
        console.log(`  ✓ Sendungsnummer ${shipmentNumber}，进度已保存`);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        const type = message.includes("Sendungsnummer") ? "未取得运单号" : "网页处理失败";
        addError(errorSheet, record, type, message);
        errorCount += 1;
        await workbook.xlsx.writeFile(outputPath);
        console.error(`  ✗ ${message}（已写入错误记录，继续下一行）`);
        try {
          await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
          frame = await waitForDhlForm(page);
        } catch {
          console.error("  无法恢复 DHL 表单，停止后续处理。已完成的进度和错误记录均已保存。");
          break;
        }
      }
    }
    console.log(`\n处理结束：成功 ${successCount} 行，报错 ${errorCount} 行。`);
    console.log(`结果文件：${outputPath}`);
    if (errorCount) console.log(`请打开 Excel 的“${ERROR_SHEET_NAME}”工作表查看详情。`);
  } finally {
    terminal.close();
    if (context) await context.close();
  }
}

main().catch((error) => {
  console.error(`\n失败：${error.message}`);
  process.exitCode = 1;
});

