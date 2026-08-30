import test from "node:test";
import assert from "node:assert/strict";
import { normalizePostalCode, outputPathFor, parseShipmentResult } from "../src/helpers.js";

test("德国四位邮编自动补零", () => {
  assert.equal(normalizePostalCode(9212), "09212");
  assert.equal(normalizePostalCode("07745"), "07745");
});

test("解析 DHL 结果页", () => {
  const result = parseShipmentResult("Sendungsnummer: 123456789012\nKundenreferenz: test-ref-001\n");
  assert.deepEqual(result, { shipmentNumber: "123456789012", customerReference: "test-ref-001" });
});

test("默认输出文件不覆盖源文件", () => {
  assert.match(outputPathFor("C:/data/DHL退货标签.xlsx"), /DHL退货标签_已填写运单号\.xlsx$/);
});

