// Offline adapter lifecycle checks; only the host registry and component-call boundary are mocked.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';

const source = await readFile(new URL('../Apps/OfficeIMO.Web.Converter/Components/ConverterWorkspace.razor.js', import.meta.url), 'utf8');

function registry() {
  const tools = new Map();
  return {
    tools,
    async registerTool(tool, options) {
      tools.set(tool.name, tool);
      options.signal.addEventListener('abort', () => {
        if (tools.get(tool.name) === tool) tools.delete(tool.name);
      }, { once: true });
    }
  };
}

function host(parentOrigin = 'https://example.test') {
  const parentRegistry = registry();
  const childRegistry = registry();
  const events = new Map();
  const document = { modelContext: childRegistry, body: { setAttribute() {} } };
  const window = {
    location: { origin: 'https://example.test' },
    parent: { location: { origin: parentOrigin }, document: { modelContext: parentRegistry } },
    frameElement: { matches: () => true },
    addEventListener: (name, callback) => events.set(name, callback)
  };
  const context = vm.createContext({ window, document, AbortController, URL, Blob });
  vm.runInContext(source.replace(/^export /gm, ''), context, { filename: 'ConverterWorkspace.razor.js' });
  return { context, parentRegistry, childRegistry, events };
}

const cached = host();
const converter = { invokeMethodAsync: async name => ({ success: true, method: name }) };
assert.equal(await cached.context.registerWebMcpTool(converter), true);
assert.equal(cached.parentRegistry.tools.size, 1, 'The canonical host exposes the embedded converter tool.');
assert.equal(cached.childRegistry.tools.size, 0);
assert.equal((await cached.parentRegistry.tools.get('convert_selected_document').execute({}, {})).success, true);
cached.events.get('pagehide')({ persisted: true });
assert.equal(cached.parentRegistry.tools.size, 0, 'Suspended tools must not remain callable.');
await cached.events.get('pageshow')({ persisted: true });
assert.equal(cached.parentRegistry.tools.size, 1, 'A restored app re-registers its existing converter.');
assert.equal((await cached.parentRegistry.tools.get('convert_selected_document').execute({}, {})).success, true);
await cached.context.unregisterWebMcpTool();
await cached.events.get('pageshow')({ persisted: true });
assert.equal(cached.parentRegistry.tools.size, 0, 'A disposed component must not be revived.');

const departed = host();
await departed.context.registerWebMcpTool(converter);
departed.events.get('pagehide')({ persisted: false });
await departed.events.get('pageshow')({ persisted: true });
assert.equal(departed.parentRegistry.tools.size, 0, 'A destroyed document releases its converter reference.');

const foreign = host('https://other.example');
await foreign.context.registerWebMcpTool(converter);
assert.equal(foreign.parentRegistry.tools.size, 0, 'Other-origin embeds must not register on their host.');
assert.equal(foreign.childRegistry.tools.size, 1);

const pending = host();
const registration = pending.context.registerWebMcpTool(converter);
pending.events.get('pagehide')({ persisted: true });
assert.equal(await registration, false, 'A registration aborted while awaiting the platform cannot become active.');
await pending.events.get('pageshow')({ persisted: true });
assert.equal(pending.parentRegistry.tools.size, 1);
console.log('Converter WebMCP lifecycle checks passed.');
