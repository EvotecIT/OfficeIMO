// Hosts the OfficeIMO .NET engine off the page's main thread.
// Request:  { id, type, ...args }      Response: { id, ok, value | error }      Progress: { type: "progress", bytes }
// Types: prepare {kind, assemblies}, stage {slot, bytes, name}, clear, run {kind, target, action, options},
//        artifact {index}, probe {slot}, pages {source, index}, render {source, index, page, size}.
import { dotnet } from "./_framework/dotnet.js";

let runtime = null;
let engine = null;
let booting = null;
let downloaded = 0;
const loaded = new Set();
const prefetched = new Map(); // fingerprinted asset name -> Promise<Response>
let wanted = new Set();       // logical assembly names the current tool needs ("OfficeIMO.Word.wasm")
let lazyAssets = [];

const post = (message, transfer) => self.postMessage(message, transfer || []);

function counted(responsePromise) {
  return responsePromise.then(response => {
    if (!response.ok || !response.body) return response;
    const reader = response.body.getReader();
    const body = new ReadableStream({
      async pull(controller) {
        const { done, value } = await reader.read();
        if (done) { controller.close(); return; }
        downloaded += value.byteLength;
        post({ type: "progress", bytes: downloaded });
        controller.enqueue(value);
      },
      cancel(reason) { reader.cancel(reason); }
    });
    return new Response(body, { status: response.status, statusText: response.statusText, headers: response.headers });
  });
}

const fetchAsset = (uri, integrity) => counted(fetch(uri, integrity ? { integrity, credentials: "omit" } : { credentials: "omit" }));

function prefetchWanted() {
  for (const asset of lazyAssets) {
    const logical = asset.virtualPath || asset.name;
    if (!wanted.has(logical) || loaded.has(logical) || prefetched.has(asset.name)) continue;
    prefetched.set(asset.name, fetchAsset(new URL("./_framework/" + asset.name, import.meta.url).href, asset.integrity || asset.hash));
  }
}

function boot() {
  booting ??= (async () => {
    const started = performance.now();
    runtime = await dotnet
      .withModuleConfig({
        onConfigLoaded(config) {
          lazyAssets = config.resources?.lazyAssembly || [];
          prefetchWanted(); // start the tool's engine download alongside the runtime
        }
      })
      .withResourceLoader((type, name, defaultUri, integrity) => {
        if (type === "dotnetjs") return defaultUri;
        const early = prefetched.get(name);
        if (early) { prefetched.delete(name); return early; }
        return fetchAsset(defaultUri, integrity);
      })
      .create();
    const config = runtime.getConfig();
    engine = (await runtime.getAssemblyExports(config.mainAssemblyName)).OfficeIMO.Web.Converter.Engine.EngineExports;
    return { ms: Math.round(performance.now() - started), bytes: downloaded };
  })().catch(error => {
    runtime = engine = booting = null;
    loaded.clear();
    prefetched.clear();
    lazyAssets = [];
    downloaded = 0;
    throw error;
  });
  return booting;
}

async function prepare(kind, assemblies) {
  for (const name of assemblies || []) wanted.add(name);
  if (lazyAssets.length) prefetchWanted();
  await boot();
  const started = performance.now();
  for (const name of assemblies || []) {
    if (loaded.has(name)) continue;
    await runtime.INTERNAL.loadLazyAssembly(name);
    loaded.add(name);
  }
  if (kind) engine.Warmup(kind);
  return { ms: Math.round(performance.now() - started), bytes: downloaded };
}

const handlers = {
  prepare: ({ kind, assemblies }) => prepare(kind, assemblies),
  stage: async ({ slot, bytes, name }) => { await boot(); engine.Stage(slot, new Uint8Array(bytes), name); return true; },
  clear: async () => { await boot(); engine.ClearInputs(); return true; },
  run: async ({ kind, target, action, options }) => {
    await boot();
    const call = () => JSON.parse(engine.Run(kind, target || "", action, JSON.stringify(options || {})));
    let result = call();
    // Some documents need more than the tool preloads (for example Japanese or symbol fonts): load it, then run again.
    if (result.needs && result.needs.length) {
      await prepare(null, result.needs);
      result = call();
    }
    return result;
  },
  artifact: async ({ index }) => { await boot(); return engine.Artifact(index); },
  probe: async ({ slot }) => { await boot(); return JSON.parse(engine.Probe(slot)); },
  pages: async ({ source, index }) => { await boot(); return engine.PageCount(source, index); },
  render: async ({ source, index, page, size }) => {
    // PDF previews may contain unembedded Japanese, Arabic or symbol fonts even when
    // the operation itself only edits PDF structure and never requests conversion fonts.
    await prepare(null, ["OfficeIMO.Web.Fonts.wasm", "OfficeIMO.Web.Fonts.Fallback.wasm"]);
    return engine.RenderPage(source, index, page, size || 1280);
  },
  // WebAssembly memory only grows, so its size is the engine's peak footprint (used by the CI performance budget).
  memory: async () => { await boot(); return runtime.Module && runtime.Module.HEAPU8 ? runtime.Module.HEAPU8.byteLength : 0; }
};

// One call at a time: the .NET session is single-threaded and later calls depend on earlier ones.
let queue = Promise.resolve();
self.onmessage = ({ data }) => {
  const { id, type } = data;
  const handler = handlers[type];
  queue = queue.then(async () => {
    if (!handler) { post({ id, ok: false, error: `Unknown request '${type}'.` }); return; }
    try {
      const value = await handler(data);
      if (value instanceof Uint8Array) post({ id, ok: true, value: value.buffer }, [value.buffer]);
      else post({ id, ok: true, value });
    } catch (error) {
      post({ id, ok: false, error: String(error?.message || error) });
    }
  });
};
