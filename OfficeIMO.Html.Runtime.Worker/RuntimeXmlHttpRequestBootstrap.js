(function (fetch, Headers, AbortController, report) {
    "use strict";
    const xhrState = new WeakMap();
    const supportedMethods = new Set(["GET", "HEAD", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"]);
    const supportedResponseTypes = new Set(["", "text", "json", "arraybuffer"]);

    function state(value) {
        const item = xhrState.get(value);
        if (!item) throw new TypeError("Illegal invocation");
        return item;
    }
    function domError(name, message) {
        const error = new Error(message);
        error.name = name;
        return error;
    }
    function event(type, target) {
        return { type, target, currentTarget: target, bubbles: false, cancelable: false, defaultPrevented: false,
            lengthComputable: false, loaded: 0, total: 0 };
    }
    function dispatch(target, type) {
        const data = state(target), value = event(type, target);
        for (const item of data.listeners.slice()) {
            if (item.type !== type || !data.listeners.includes(item)) continue;
            if (item.once) data.listeners.splice(data.listeners.indexOf(item), 1);
            try {
                if (typeof item.callback === "function") item.callback.call(target, value);
                else if (typeof item.callback?.handleEvent === "function") item.callback.handleEvent(value);
            } catch (error) { report(String(error)); }
        }
        let handler;
        try { handler = target["on" + type]; }
        catch (error) { report(String(error)); return; }
        if (typeof handler === "function") {
            try { handler.call(target, value); } catch (error) { report(String(error)); }
        }
    }
    function ready(target, value) {
        const data = state(target);
        if (data.readyState === value) return;
        data.readyState = value;
        dispatch(target, "readystatechange");
    }
    function resetResponse(data) {
        data.status = 0;
        data.statusText = "";
        data.responseURL = "";
        data.responseHeaders = null;
        data.responseText = "";
        data.response = null;
    }
    function finishFailure(target, generation, type) {
        const data = state(target);
        if (data.generation !== generation || !data.send) return;
        data.send = false;
        data.controller = null;
        resetResponse(data);
        ready(target, XMLHttpRequest.DONE);
        if (data.generation !== generation) return;
        dispatch(target, type);
        if (data.generation !== generation) return;
        dispatch(target, "loadend");
    }
    function finishSuccess(target, generation, value) {
        const data = state(target);
        if (data.generation !== generation || !data.send) return;
        if (data.responseType === "arraybuffer") {
            data.response = value;
        } else {
            data.responseText = value;
            if (data.responseType === "json") {
                try { data.response = value === "" ? null : JSON.parse(value); }
                catch { data.response = null; }
            } else data.response = value;
        }
        ready(target, XMLHttpRequest.LOADING);
        if (data.generation !== generation || !data.send) return;
        data.send = false;
        data.controller = null;
        ready(target, XMLHttpRequest.DONE);
        if (data.generation !== generation) return;
        dispatch(target, "load");
        if (data.generation !== generation) return;
        dispatch(target, "loadend");
    }

    class XMLHttpRequest {
        constructor() {
            xhrState.set(this, { readyState: 0, method: null, url: null, headers: new Headers(), send: false,
                controller: null, generation: 0, responseType: "", timeout: 0, listeners: [],
                status: 0, statusText: "", responseURL: "", responseHeaders: null, responseText: "", response: null });
        }
        get readyState() { return state(this).readyState; }
        get status() { return state(this).status; }
        get statusText() { return state(this).statusText; }
        get responseURL() { return state(this).responseURL; }
        get responseXML() { state(this); return null; }
        get response() {
            const data = state(this);
            if (data.responseType === "" || data.responseType === "text")
                return data.readyState >= XMLHttpRequest.LOADING ? data.responseText : "";
            return data.readyState === XMLHttpRequest.DONE ? data.response : null;
        }
        get responseText() {
            const data = state(this);
            if (data.responseType !== "" && data.responseType !== "text")
                throw domError("InvalidStateError", "responseText is available only for text responses");
            return data.readyState >= XMLHttpRequest.LOADING ? data.responseText : "";
        }
        get responseType() { return state(this).responseType; }
        set responseType(value) {
            const data = state(this);
            value = String(value);
            if (!supportedResponseTypes.has(value)) throw domError("SyntaxError", "Unsupported XMLHttpRequest responseType");
            if (data.readyState >= XMLHttpRequest.LOADING)
                throw domError("InvalidStateError", "responseType cannot change after response loading starts");
            data.responseType = value;
        }
        get timeout() { return state(this).timeout; }
        set timeout(value) {
            const data = state(this);
            value = Number(value);
            if (!Number.isFinite(value) || value < 0) throw new TypeError("Invalid XMLHttpRequest timeout");
            if (value !== 0) throw domError("NotSupportedError", "Per-request XMLHttpRequest timeouts are not supported; use the runtime resource deadline");
            data.timeout = 0;
        }
        get withCredentials() { state(this); return false; }
        set withCredentials(value) {
            state(this);
            if (value) throw domError("NotSupportedError", "Credentialed XMLHttpRequest is not supported");
        }
        get upload() { state(this); return null; }
        open(method, url, async, user, password) {
            const data = state(this);
            if (async === false) throw domError("NotSupportedError", "Synchronous XMLHttpRequest is not supported");
            if (user != null || password != null) throw domError("NotSupportedError", "Credentialed XMLHttpRequest is not supported");
            method = String(method).toUpperCase();
            if (!supportedMethods.has(method)) throw domError("NotSupportedError", "Unsupported XMLHttpRequest method");
            if (typeof url !== "string" && !(typeof URL !== "undefined" && url instanceof URL))
                throw new TypeError("XMLHttpRequest.open requires a URL string or URL");
            data.generation++;
            data.controller?.abort();
            data.method = method;
            data.url = String(url);
            data.headers = new Headers();
            data.send = false;
            data.controller = null;
            resetResponse(data);
            ready(this, XMLHttpRequest.OPENED);
        }
        setRequestHeader(name, value) {
            const data = state(this);
            if (data.readyState !== XMLHttpRequest.OPENED || data.send)
                throw domError("InvalidStateError", "Request headers require an opened, unsent XMLHttpRequest");
            data.headers.append(name, value);
        }
        getResponseHeader(name) {
            const data = state(this);
            if (data.readyState < XMLHttpRequest.HEADERS_RECEIVED || !data.responseHeaders) return null;
            return data.responseHeaders.get(name);
        }
        getAllResponseHeaders() {
            const data = state(this);
            if (data.readyState < XMLHttpRequest.HEADERS_RECEIVED || !data.responseHeaders) return "";
            const values = Array.from(data.responseHeaders, pair => pair[0] + ": " + pair[1]);
            return values.length === 0 ? "" : values.join("\r\n") + "\r\n";
        }
        overrideMimeType() { state(this); throw domError("NotSupportedError", "XMLHttpRequest MIME overrides are not supported"); }
        send(body) {
            const data = state(this);
            if (data.readyState !== XMLHttpRequest.OPENED || data.send)
                throw domError("InvalidStateError", "XMLHttpRequest.send requires an opened, unsent request");
            if (data.method === "GET" || data.method === "HEAD") body = null;
            data.send = true;
            resetResponse(data);
            const generation = ++data.generation;
            const controller = data.controller = new AbortController();
            dispatch(this, "loadstart");
            if (data.generation !== generation || !data.send || data.controller !== controller) return;
            fetch(data.url, { method: data.method, headers: data.headers, body, signal: controller.signal,
                mode: "cors", credentials: "same-origin", redirect: "follow" }).then(response => {
                if (data.generation !== generation || !data.send) return;
                data.status = response.status;
                data.statusText = response.statusText;
                data.responseURL = response.url;
                data.responseHeaders = response.headers;
                ready(this, XMLHttpRequest.HEADERS_RECEIVED);
                if (data.generation !== generation || !data.send) return;
                const pending = data.responseType === "arraybuffer" ? response.arrayBuffer() : response.text();
                pending.then(value => finishSuccess(this, generation, value), () => finishFailure(this, generation, "error"));
            }, error => finishFailure(this, generation, error?.name === "AbortError" ? "abort" : "error"));
        }
        abort() {
            const data = state(this);
            if (!data.send) {
                data.generation++;
                data.controller?.abort();
                data.controller = null;
                data.readyState = XMLHttpRequest.UNSENT;
                resetResponse(data);
                return;
            }
            const generation = data.generation;
            data.controller?.abort();
            finishFailure(this, generation, "abort");
            if (data.generation === generation && !data.send) data.readyState = XMLHttpRequest.UNSENT;
        }
        addEventListener(type, callback, options) {
            const data = state(this);
            type = String(type);
            if (callback == null) return;
            const capture = typeof options === "boolean" ? options : !!options?.capture;
            if (!data.listeners.some(item => item.type === type && item.callback === callback && item.capture === capture)) {
                if (data.listeners.length >= 128) throw domError("QuotaExceededError", "XMLHttpRequest listener limit exceeded");
                data.listeners.push({ type, callback, capture, once: !!options?.once });
            }
        }
        removeEventListener(type, callback, options) {
            const data = state(this), capture = typeof options === "boolean" ? options : !!options?.capture;
            type = String(type);
            data.listeners = data.listeners.filter(item => item.type !== type || item.callback !== callback || item.capture !== capture);
        }
    }
    for (const pair of [["UNSENT", 0], ["OPENED", 1], ["HEADERS_RECEIVED", 2], ["LOADING", 3], ["DONE", 4]]) {
        Object.defineProperty(XMLHttpRequest, pair[0], { value: pair[1] });
        Object.defineProperty(XMLHttpRequest.prototype, pair[0], { value: pair[1] });
    }
    return { XMLHttpRequest };
})
