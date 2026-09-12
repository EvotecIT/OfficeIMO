// A shared wrapper map preserves callback identity across native DOM prototypes.
(() => {
    const wrappers = new WeakMap();
    const handlerWrappers = new WeakMap();
    const originals = new WeakMap();
    const eventPaths = new WeakMap();
    function rememberPath(event, target, replace) {
        if (!replace && eventPaths.has(event)) return;
        const path = [];
        let current = target;
        while (current) {
            path.push(current);
            // Shadow-tree retargeting is outside this light-DOM profile.
            if (current.nodeType === 11 && current.host) {
                eventPaths.set(event, null);
                return;
            }
            if (current.nodeType === 9) {
                if (current.defaultView) path.push(current.defaultView);
                break;
            }
            current = current.nodeType ? current.parentNode : null;
        }
        eventPaths.set(event, path);
    }
    return function (add, remove, report, kind, normalizeWindow) {
        if (kind === 'dispatch') return function (event) {
            if (event instanceof Event && event.eventPhase === 0) rememberPath(event, normalizeWindow(this), true);
            add.call(this, event);
            return !event.defaultPrevented;
        };
        if (kind === 'path') return function () {
            if (!(this instanceof Event)) throw new TypeError("The receiver must be an Event");
            if (this.eventPhase === 0) return [];
            if (!eventPaths.has(this)) rememberPath(this, this.target, false);
            const path = eventPaths.get(this);
            if (path === null) {
                const error = new Error("Shadow-tree event paths are outside the supported profile");
                error.name = 'NotSupportedError';
                throw error;
            }
            return path.slice();
        };
        function wrap(listener, handler) {
            if (listener === null || (typeof listener !== 'function' && typeof listener !== 'object')) return listener;
            const map = handler ? handlerWrappers : wrappers;
            let wrapped = map.get(listener);
            if (!wrapped) {
                wrapped = function (event) {
                    try {
                        rememberPath(event, event.target, false);
                        const result = typeof listener === 'function' ? listener.call(normalizeWindow(this), event) : listener.handleEvent(event);
                        if (handler && result === false) event.preventDefault();
                        return result;
                    } catch (error) {
                        let message = 'Event listener failed.';
                        try { message = String(error); } catch (_) { }
                        report(message);
                        throw error;
                    }
                };
                map.set(listener, wrapped);
                originals.set(wrapped, listener);
            }
            return wrapped;
        }
        if (kind === 'handler') return {
            get: function () { const value = add.call(this); return value && originals.get(value) || value; },
            set: function (value) { return remove.call(this, typeof value === 'function' ? wrap(value, true) : value); }
        };
        return {
            add: function (type, listener, options) { return add.call(this, type, wrap(listener), options); },
            remove: function (type, listener, options) {
                const wrapped = listener && (typeof listener === 'function' || typeof listener === 'object') ? wrappers.get(listener) : null;
                return remove.call(this, type, wrapped || listener, options);
            }
        };
    };
})();
