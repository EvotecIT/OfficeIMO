(function (start, cancel, encode, decode, report, maxBodyBytes) {
    "use strict";
    const headerState = new WeakMap(), responseState = new WeakMap(), signalState = new WeakMap(), controllerState = new WeakMap();
    const internalKey = {};
    function state(map, value) {
        const item = map.get(value);
        if (!item) throw new TypeError("Illegal invocation");
        return item;
    }
    function headerName(value) {
        value = String(value);
        if (!/^[!#$%&'*+.^_`|~0-9A-Za-z-]+$/.test(value)) throw new TypeError("Invalid header name");
        return value.toLowerCase();
    }
    function headerValue(value) {
        value = String(value).replace(/^[\t ]+|[\t ]+$/g, "");
        if (/[\r\n\0\u0100-\uFFFF]/.test(value)) throw new TypeError("Invalid header value");
        return value;
    }
    class Headers {
        constructor(init) {
            headerState.set(this, { values: new Map(), immutable: false });
            if (init == null) return;
            if (typeof init[Symbol.iterator] === "function") {
                for (const pair of init) {
                    const entry = Array.from(pair);
                    if (entry.length !== 2) throw new TypeError("A header entry must contain two values");
                    this.append(entry[0], entry[1]);
                }
            } else {
                for (const key of Object.keys(init)) this.append(key, init[key]);
            }
        }
        append(name, value) {
            const data = state(headerState, this);
            name = headerName(name); value = headerValue(value);
            if (data.immutable) throw new TypeError("Response headers are immutable");
            data.values.set(name, data.values.has(name) ? data.values.get(name) + ", " + value : value);
        }
        set(name, value) {
            const data = state(headerState, this);
            name = headerName(name); value = headerValue(value);
            if (data.immutable) throw new TypeError("Response headers are immutable");
            data.values.set(name, value);
        }
        delete(name) {
            const data = state(headerState, this); name = headerName(name);
            if (data.immutable) throw new TypeError("Response headers are immutable");
            data.values.delete(name);
        }
        get(name) { return state(headerState, this).values.get(headerName(name)) ?? null; }
        has(name) { return state(headerState, this).values.has(headerName(name)); }
        *entries() {
            // Re-sort on each step so additions during iteration have normal Headers ordering.
            let previous;
            while (true) {
                const values = state(headerState, this).values;
                const next = Array.from(values.keys()).sort().find(key => previous === undefined || key > previous);
                if (next === undefined) return;
                previous = next; yield [next, values.get(next)];
            }
        }
        *keys() { for (const pair of this.entries()) yield pair[0]; }
        *values() { for (const pair of this.entries()) yield pair[1]; }
        [Symbol.iterator]() { return this.entries(); }
        forEach(callback, receiver) { for (const pair of this.entries()) callback.call(receiver, pair[1], pair[0], this); }
    }
    function abortError() { const error = new Error("The operation was aborted"); error.name = "AbortError"; return error; }
    class AbortSignal {
        constructor(key) {
            if (key !== internalKey) throw new TypeError("Illegal constructor");
            signalState.set(this, { aborted: false, reason: undefined, listeners: [], algorithms: new Set(), handler: null });
        }
        get aborted() { return state(signalState, this).aborted; }
        get reason() { return state(signalState, this).reason; }
        get onabort() { return state(signalState, this).handler; }
        set onabort(value) { state(signalState, this).handler = typeof value === "function" ? value : null; }
        throwIfAborted() { if (this.aborted) throw this.reason; }
        addEventListener(type, callback, options) {
            const data = state(signalState, this);
            if (String(type) !== "abort" || callback == null) return;
            const capture = typeof options === "boolean" ? options : !!options?.capture;
            if (!data.listeners.some(item => item.callback === callback && item.capture === capture))
                data.listeners.push({ callback, capture, once: !!options?.once });
        }
        removeEventListener(type, callback, options) {
            const data = state(signalState, this), capture = typeof options === "boolean" ? options : !!options?.capture;
            if (String(type) === "abort") data.listeners = data.listeners.filter(item => item.callback !== callback || item.capture !== capture);
        }
        static abort(reason) { const controller = new AbortController(); controller.abort(reason); return controller.signal; }
    }
    class AbortController {
        constructor() { controllerState.set(this, new AbortSignal(internalKey)); }
        get signal() { return state(controllerState, this); }
        abort(reason) {
            const signal = this.signal, data = state(signalState, signal);
            if (data.aborted) return;
            data.aborted = true; data.reason = reason === undefined ? abortError() : reason;
            for (const algorithm of Array.from(data.algorithms)) algorithm();
            data.algorithms.clear();
            const event = { type: "abort", target: signal, currentTarget: signal, bubbles: false, cancelable: false, defaultPrevented: false };
            for (const item of data.listeners.slice()) {
                if (!data.listeners.includes(item)) continue;
                if (item.once) data.listeners.splice(data.listeners.indexOf(item), 1);
                try {
                    if (typeof item.callback === "function") item.callback.call(signal, event);
                    else if (typeof item.callback.handleEvent === "function") item.callback.handleEvent(event);
                } catch (error) { report(String(error)); }
            }
            if (data.handler) { try { data.handler.call(signal, event); } catch (error) { report(String(error)); } }
        }
    }
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    function toBase64(bytes) {
        if (bytes.byteLength > maxBodyBytes) throw new TypeError("Fetch request body exceeds its byte budget");
        const parts = [];
        for (let i = 0; i < bytes.length; i += 3) {
            const a = bytes[i], b = bytes[i + 1], c = bytes[i + 2];
            parts.push(alphabet[a >> 2] + alphabet[((a & 3) << 4) | ((b ?? 0) >> 4)] +
                (i + 1 < bytes.length ? alphabet[((b & 15) << 2) | ((c ?? 0) >> 6)] : "=") +
                (i + 2 < bytes.length ? alphabet[c & 63] : "="));
        }
        return parts.join("");
    }
    function fromBase64(encoded) {
        const padding = encoded.endsWith("==") ? 2 : encoded.endsWith("=") ? 1 : 0;
        const bytes = new Uint8Array(encoded.length / 4 * 3 - padding);
        let offset = 0;
        for (let i = 0; i < encoded.length; i += 4) {
            const value = (alphabet.indexOf(encoded[i]) << 18) | (alphabet.indexOf(encoded[i + 1]) << 12) |
                ((alphabet.indexOf(encoded[i + 2]) & 63) << 6) | (alphabet.indexOf(encoded[i + 3]) & 63);
            if (offset < bytes.length) bytes[offset++] = value >> 16;
            if (offset < bytes.length) bytes[offset++] = value >> 8;
            if (offset < bytes.length) bytes[offset++] = value;
        }
        return bytes.buffer;
    }
    function bodyBytes(body, headers) {
        if (body == null) return null;
        if (typeof body === "string") {
            if (!headers.has("content-type")) headers.set("content-type", "text/plain;charset=UTF-8");
            return encode(body);
        }
        if (typeof URLSearchParams !== "undefined" && body instanceof URLSearchParams) {
            if (!headers.has("content-type")) headers.set("content-type", "application/x-www-form-urlencoded;charset=UTF-8");
            return encode(body.toString());
        }
        if (body instanceof ArrayBuffer) return toBase64(new Uint8Array(body));
        if (ArrayBuffer.isView(body)) return toBase64(new Uint8Array(body.buffer, body.byteOffset, body.byteLength));
        throw new TypeError("Unsupported fetch body; use a string, URLSearchParams, ArrayBuffer or typed array");
    }
    function consume(response, kind) {
        return new Promise((resolve, reject) => {
            const data = state(responseState, response);
            if (data.used) throw new TypeError("The response body has already been consumed");
            if (data.signal?.aborted) { reject(data.signal.reason); return; }
            if (data.hasBody) data.used = true;
            if (kind === "buffer") resolve(fromBase64(data.body));
            else { const text = decode(data.body); resolve(kind === "json" ? JSON.parse(text) : text); }
        });
    }
    class Response {
        constructor(key, data, signal) {
            if (key !== internalKey) throw new TypeError("Constructed Response objects are not supported by this runtime profile");
            const headers = new Headers(data.headers);
            state(headerState, headers).immutable = true;
            responseState.set(this, { ...data, headers, signal, used: false });
        }
        get status() { return state(responseState, this).status; }
        get statusText() { return state(responseState, this).statusText; }
        get ok() { return this.status >= 200 && this.status < 300; }
        get url() { return state(responseState, this).url; }
        get redirected() { return state(responseState, this).redirected; }
        get type() { return state(responseState, this).type; }
        get headers() { return state(responseState, this).headers; }
        get bodyUsed() { return state(responseState, this).used; }
        get body() {
            if (!state(responseState, this).hasBody) return null;
            throw new TypeError("Response streams are not supported; use text(), json() or arrayBuffer()");
        }
        text() { return consume(this, "text"); }
        json() { return consume(this, "json"); }
        arrayBuffer() { return consume(this, "buffer"); }
        clone() {
            const data = state(responseState, this);
            if (data.used) throw new TypeError("The response body has already been consumed");
            return new Response(internalKey, data, data.signal);
        }
    }
    function fetch(input, init) {
        return new Promise((resolve, reject) => {
            init = init ?? {};
            if (typeof input !== "string" && !(typeof URL !== "undefined" && input instanceof URL)) throw new TypeError("fetch requires a URL string or URL");
            for (const option of ["cache", "integrity", "keepalive", "referrer", "referrerPolicy", "window", "duplex", "priority"]) {
                if (init[option] !== undefined) throw new TypeError("Unsupported fetch option: " + option);
            }
            const signal = init.signal;
            if (signal != null) state(signalState, signal);
            if (signal?.aborted) { reject(signal.reason); return; }
            const headers = new Headers(init.headers);
            const body = bodyBytes(init.body, headers);
            const request = {
                Url: String(input), Method: String(init.method ?? "GET").toUpperCase(), Headers: Object.fromEntries(headers), Body: body,
                Mode: String(init.mode ?? "cors"), Credentials: String(init.credentials ?? "same-origin"), Redirect: String(init.redirect ?? "follow")
            };
            let settled = false, id;
            const abort = () => { if (!settled) { settled = true; cancel(id); reject(signal.reason); } };
            id = start(JSON.stringify(request), json => {
                if (settled) return;
                settled = true;
                if (signal) state(signalState, signal).algorithms.delete(abort);
                try {
                    const data = JSON.parse(json);
                    if (data.error) { const error = new TypeError(data.message); error.name = data.error; reject(error); }
                    else resolve(new Response(internalKey, data, signal));
                } catch (error) { reject(error); }
            });
            if (signal) state(signalState, signal).algorithms.add(abort);
        });
    }
    return { fetch, Headers, Response, AbortController, AbortSignal };
})
