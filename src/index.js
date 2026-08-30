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
} from "./helpers.js";

const FORM_URL = "https://geschaeftskunden.dhl.de/content/dhl-rpi/gw/rpcustomerweb/OrderCustService.action";
const RESULT_SUCCESS = "Ihre Retourensendung wurde erfolgreich beauftragt!";

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
    else throw new Error(`鏈煡鍙傛暟锛?{value}`);
  }
  return args;
}

async function askForInput(terminal) {
  const answer = await terminal.question("璇锋嫋鍏?Excel 鏂囦欢骞舵寜 Enter锛歕n> ");
  return answer.trim().replace(/^"|"$/g, "");
}

async function waitForDhlForm(page) {
  await page.goto(FORM_URL, { waitUntil: "domcontentloaded" });
  console.log("\n璇峰湪鎵撳紑鐨?Chrome 涓櫥褰?DHL銆傜櫥褰曞悗鑴氭湰浼氳嚜鍔ㄧ户缁€俓n");
  const deadline = Date.now() + 10 * 60 * 1000;
  while (Date.now() < deadline) {
    const frame = page.frames().find((candidate) => candidate.url().includes("OrderCustService.action"));
    if (frame && (await frame.getByText("Retoure beauftragen", { exact: true }).count())) return frame;
    await page.waitForTimeout(1000);
  }
  throw new Error("绛夊緟 DHL 鐧诲綍瓒呮椂锛?0 鍒嗛挓锛夈€?);
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
  if (await goGreenPlus.isChecked()) throw new Error("GoGreen Plus 鏃犳硶鍙栨秷鍕鹃€夈€?);

  const actualReference = await frame.locator('input[name="shipmentReference"]').inputValue();
  if (actualReference !== record.shipmentReference) throw new Error("缃戦〉鍙戦€佸弬鑰冨彿涓?Excel 涓嶄竴鑷淬€?);
  const submit = frame.getByRole("button", { name: "Retoure beauftragen", exact: true });
  if (!(await submit.isEnabled())) throw new Error("DHL 鎻愪氦鎸夐挳涓嶅彲鐢紝璇锋鏌ュ繀濉瓧娈点€?);
  return submit;
}

async function submitAndRead(page, frame, submit, record) {
  await submit.scrollIntoViewIfNeeded();
  await submit.click();
  await frame.getByText(RESULT_SUCCESS, { exact: false }).waitFor({ state: "visible", timeout: 30000 });
  const result = parseShipmentResult(await frame.locator("body").innerText());
  if (!result.shipmentNumber) throw new Error("缁撴灉椤垫病鏈夋壘鍒?Sendungsnummer銆?);
  if (result.customerReference !== record.customerReference) {
    throw new Error(`缁撴灉椤靛鎴峰弬鑰冨彿涓嶄竴鑷达細${result.customerReference}`);
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
    if (!sheet) throw new Error("Excel 涓病鏈夊伐浣滆〃銆?);
    const columns = headerMap(sheet.getRow(1));
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
      const reference = asText(row.getCell(columns.get("Sendungsreferenz")).value);
      if (!reference) continue;
      if (asText(row.getCell(trackingColumn).value)) continue;
      pending.push(rowToRecord(row, columns));
    }
    if (!pending.length) {
      console.log("娌℃湁寰呭鐞嗚锛歀 鍒楀凡鏈夎繍鍗曞彿锛屾垨琛ㄦ牸娌℃湁鏁版嵁銆?);
      return;
    }

    console.log(`杈撳叆鏂囦欢锛?{inputPath}`);
    console.log(`杈撳嚭鏂囦欢锛?{outputPath}`);
    console.log(`寰呭垱寤猴細${pending.length} 寮犻€€璐ф爣绛綻);
    if (!args.dryRun) {
      const answer = await terminal.question("杩欎細鍒涘缓鐪熷疄 DHL 閫€璐у鎵樸€傝杈撳叆 CREATE 缁х画锛?);
      if (answer.trim() !== "CREATE") {
        console.log("宸插彇娑堬紝鏈垱寤轰换浣曟爣绛俱€?);
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

    for (let index = 0; index < pending.length; index += 1) {
      const record = pending[index];
      console.log(`[${index + 1}/${pending.length}] ${record.customerReference}`);
      frame = await openBlankOrder(page, frame);
      const submit = await fillOrder(page, frame, record);
      if (args.dryRun) {
        console.log("璇曡繍琛屽畬鎴愶細绗竴鏉″凡濉ソ锛屽皻鏈偣鍑?Retoure beauftragen銆?);
        await terminal.question("妫€鏌ユ祻瑙堝櫒鍚庢寜 Enter 鍏抽棴鑴氭湰銆?);
        return;
      }
      const shipmentNumber = await submitAndRead(page, frame, submit, record);
      const cell = sheet.getCell(record.excelRow, trackingColumn);
      cell.value = Number(shipmentNumber);
      cell.numFmt = "0";
      await workbook.xlsx.writeFile(outputPath);
      console.log(`  鉁?Sendungsnummer ${shipmentNumber}锛岃繘搴﹀凡淇濆瓨`);
    }
    console.log(`\n鍏ㄩ儴瀹屾垚锛?{outputPath}`);
  } finally {
    terminal.close();
    if (context) await context.close();
  }
}

main().catch((error) => {
  console.error(`\n澶辫触锛?{error.message}`);
  process.exitCode = 1;
});

