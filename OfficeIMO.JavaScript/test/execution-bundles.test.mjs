import test from "node:test";
import assert from "node:assert/strict";
import { MessageChannel } from "node:worker_threads";
import * as api from "../bundles/officeimo.mjs";
import * as execution from "../bundles/officeimo-execution.mjs";

test("standalone writer and execution modules preserve remote OfficeIMO error identity", { timeout: 5000 }, async () => {
  const errors = [
    new api.core.OfficeIMOError("RESOURCE_LIMIT", "Standalone writer limit"),
    new api.core.OfficeIMOError("NOT_SUPPORTED", "Unsupported operation without a feature"),
    new api.core.NotSupportedError("standalone shaped report")
  ];
  const { port1, port2 } = new MessageChannel();
  const handler = execution.handleWorkerExports(port2, ({ input, rows, sink, signal }) => {
    if (input === "writer") return api.writeCsvTo(rows, sink, { signal, columns: [{ header: "V" }], limits: { maxTextCharacters: 1 } });
    throw errors[input];
  });
  try {
    for (let input = 0; input < errors.length; input++) {
      const original = errors[input];
      await assert.rejects(execution.executeWorkerExport(port1, [], { write() {} }, { input }), error => {
        assert.equal(error.code, original.code);
        assert.equal(error.feature, original.feature);
        assert.equal(error.name, original.name);
        assert.equal(error.message, original.message);
        assert.ok(error instanceof api.core.OfficeIMOError);
        assert.equal(error instanceof api.core.NotSupportedError, original instanceof api.core.NotSupportedError);
        return true;
      });
    }
    await assert.rejects(execution.executeWorkerExport(port1, [["oversized"]], { write() {} }, { input: "writer" }),
      error => error.code === "RESOURCE_LIMIT" && error instanceof api.core.OfficeIMOError);
  } finally { handler.dispose(); port1.close(); port2.close(); }
});

test("OfficeIMO error brands preserve subclass-specific instanceof behavior", () => {
  class ScopedError extends api.core.OfficeIMOError {}
  class ScopedUnsupported extends api.core.NotSupportedError {}
  const base = new api.core.OfficeIMOError("RESOURCE_LIMIT", "base");
  const unsupportedCode = new api.core.OfficeIMOError("NOT_SUPPORTED", "base unsupported code");
  const unsupported = new api.core.NotSupportedError("feature");
  assert.ok(new ScopedError("INVALID_STATE", "scoped") instanceof ScopedError);
  assert.ok(new ScopedUnsupported("scoped feature") instanceof ScopedUnsupported);
  assert.ok(!(base instanceof ScopedError));
  assert.ok(!(unsupported instanceof ScopedError));
  assert.ok(!(unsupported instanceof ScopedUnsupported));
  assert.ok(!(base instanceof api.core.NotSupportedError));
  assert.ok(!(unsupportedCode instanceof api.core.NotSupportedError));
  assert.ok(!(new Error("ordinary") instanceof api.core.OfficeIMOError));
});
