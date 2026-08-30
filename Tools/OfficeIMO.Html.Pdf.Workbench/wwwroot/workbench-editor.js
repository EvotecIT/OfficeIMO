export function bind(element, kind, dotnetReference) {
    if (element.__officeimoWorkbenchInput) return;
    const handler = () => {
        dotnetReference.invokeMethodAsync("NotifyInput", kind).catch(() => {});
    };
    element.__officeimoWorkbenchInput = handler;
    element.addEventListener("input", handler);
}

export function unbind(element) {
    const handler = element.__officeimoWorkbenchInput;
    if (!handler) return;
    element.removeEventListener("input", handler);
    delete element.__officeimoWorkbenchInput;
}

export function setValue(element, value) {
    if (element.value !== value) element.value = value;
}

export function beginRead(element) {
    const id = crypto.randomUUID();
    const value = element.value;
    snapshots.set(id, value);
    return { id, length: value.length };
}

export function readSlice(id, offset, length) {
    const value = snapshots.get(id);
    if (value === undefined) throw new Error("Editor snapshot expired.");
    return value.slice(offset, offset + length);
}

export function releaseRead(id) {
    snapshots.delete(id);
}
const snapshots = new Map();
