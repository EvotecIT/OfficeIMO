(function (register, add, remove, dispatch, handler, userDispatch) {
    "use strict";
    const targets = new WeakSet();
    class TransportEventTarget {
        constructor(listenerLimit) { targets.add(this); register(this, listenerLimit); }
        addEventListener(type, callback, options) { return add.call(this, type, callback, options); }
        removeEventListener(type, callback, options) { return remove.call(this, type, callback, options); }
        dispatchEvent(event) {
            if (!targets.has(this)) throw new TypeError("Illegal invocation");
            return userDispatch.call(this, event);
        }
    }
    // Preserve EventTarget identity while the native provider's base constructor
    // remains non-constructible. All methods resolve through the shared bindings.
    Object.setPrototypeOf(TransportEventTarget.prototype, EventTarget.prototype);
    class ProgressEvent extends Event {
        // The retained CLR Event constructor supplies a type-based hasInstance.
        // Its subclasses must retain JavaScript prototype identity instead.
        static [Symbol.hasInstance](value) { return Object.prototype.isPrototypeOf.call(this.prototype, value); }
        constructor(type, init = {}) {
            super(type, init);
            Object.setPrototypeOf(this, new.target.prototype);
            Object.defineProperties(this, {
                lengthComputable: { value: !!init.lengthComputable, enumerable: true },
                loaded: { value: Number(init.loaded ?? 0), enumerable: true },
                total: { value: Number(init.total ?? 0), enumerable: true }
            });
        }
    }
    function handlers(prototype, types) {
        for (const type of types) Object.defineProperty(prototype, "on" + type, {
            ...handler("on" + type), enumerable: true, configurable: true
        });
    }
    function fire(target, type, progress = false) {
        const event = progress ? new ProgressEvent(type) : new Event(type);
        return dispatch(target, event, true);
    }
    return { EventTarget: TransportEventTarget, ProgressEvent, handlers, fire };
})
