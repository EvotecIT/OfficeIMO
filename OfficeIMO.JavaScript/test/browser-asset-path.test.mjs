import test from "node:test";
import assert from "node:assert/strict";
import { win32 } from "node:path";
import { browserAssetPathIsContained } from "../scripts/browser-asset-path.mjs";

test("browser asset path guard rejects Windows paths on another drive", () => {
  const base = "D:\\shaping-assets";
  assert.equal(browserAssetPathIsContained(base, "D:\\shaping-assets\\provider.js", win32), true);
  assert.equal(browserAssetPathIsContained(base, "D:\\private\\config.js", win32), false);
  assert.equal(browserAssetPathIsContained(base, "C:\\Users\\developer\\private\\config.js", win32), false);
});
