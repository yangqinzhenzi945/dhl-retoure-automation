import test from "node:test";
import assert from "node:assert/strict";
import { normalizePostalCode, outputPathFor, parseShipmentResult } from "../src/helpers.js";

test("寰峰浗鍥涗綅閭紪鑷姩琛ラ浂", () => {
  assert.equal(normalizePostalCode(9212), "09212");
  assert.equal(normalizePostalCode("07745"), "07745");
});

test("瑙ｆ瀽 DHL 缁撴灉椤?, () => {
  const result = parseShipmentResult("Sendungsnummer: 123456789012\nKundenreferenz: test-ref-001\n");
  assert.deepEqual(result, { shipmentNumber: "123456789012", customerReference: "test-ref-001" });
});

test("榛樿杈撳嚭鏂囦欢涓嶈鐩栨簮鏂囦欢", () => {
  assert.match(outputPathFor("C:/data/DHL閫€璐ф爣绛?xlsx"), /DHL閫€璐ф爣绛綺宸插～鍐欒繍鍗曞彿\.xlsx$/);
});

