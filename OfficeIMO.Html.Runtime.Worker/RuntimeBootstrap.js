// A shared wrapper map preserves callback identity across native DOM prototypes.
(() => {
    const wrappers = new WeakMap();
    const handlerWrappers = new WeakMap();
    const originals = new WeakMap();
    return function (add, remove, report, kind) {
        if (kind === 'dispatch') return function (event) {
            add.call(this, event);
            return !event.defaultPrevented;
        };
        function wrap(listener, handler) {
            if (listener === null || (typeof listener !== 'function' && typeof listener !== 'object')) return listener;
            const map = handler ? handlerWrappers : wrappers;
            let wrapped = map.get(listener);
            if (!wrapped) {
                wrapped = function (event) {
                    try {
                        const result = typeof listener === 'function' ? listener.call(this, event) : listener.handleEvent(event);
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
