(function (create, take, observe, report, queue) {
    "use strict";
    const states = new WeakMap();
    function get(value) {
        const observer=states.get(value);
        if(!observer) throw new TypeError("Illegal invocation");
        return observer;
    }
    class MutationObserver {
        constructor(callback) {
            if(typeof callback!=="function") throw new TypeError("A mutation observer callback is required");
            const callbackAdapter = records => {
                try { callback.call(this,records,this); }
                catch(error) { report(String(error)); }
            };
            states.set(this, { native: create(callbackAdapter), callback: callbackAdapter });
        }
        observe(target, options) {
            const state = get(this);
            if (!(target instanceof Node)) throw new TypeError("The observation target must be a Node");
            options = options == null ? {} : Object(options);
            const value = { childList: !!options.childList, subtree: !!options.subtree };
            for (const name of ['attributes', 'characterData', 'attributeOldValue', 'characterDataOldValue']) {
                const option = options[name];
                if (option !== undefined) value[name] = !!option;
            }
            const filter = options.attributeFilter;
            if (filter !== undefined) {
                if (filter === null || typeof filter[Symbol.iterator] !== 'function') throw new TypeError("attributeFilter must be iterable");
                value.attributeFilter = Array.from(filter, item => {
                    if (typeof item === 'symbol') throw new TypeError("A symbol is not an attribute name");
                    return String(item);
                });
            }
            if (value.attributes === undefined && (value.attributeOldValue !== undefined || value.attributeFilter !== undefined)) value.attributes = true;
            if (value.characterData === undefined && value.characterDataOldValue !== undefined) value.characterData = true;
            if ((!value.attributes && (value.attributeOldValue || value.attributeFilter !== undefined)) ||
                (!value.characterData && value.characterDataOldValue) || (!value.childList && !value.attributes && !value.characterData))
                throw new TypeError("Invalid mutation observation options");
            // Option getters and iterators may disconnect this observer. Resolve
            // its current registration set only after those conversions finish.
            observe(state.native, target, value);
        }
        disconnect() {
            const state = get(this);
            state.native.disconnect();
        }
        takeRecords() { return take(get(this).native); }
    }
    function queueMicrotask(callback) {
        if(typeof callback!=="function") throw new TypeError("A microtask callback is required");
        queue(callback);
    }
    return { MutationObserver, queueMicrotask };
})
