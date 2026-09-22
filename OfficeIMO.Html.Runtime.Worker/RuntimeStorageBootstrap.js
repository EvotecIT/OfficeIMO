(function (read, write) {
    "use strict";
    const areas = new WeakMap();
    function text(value) {
        if (typeof value === "symbol") throw new TypeError("Cannot convert a Symbol to a string");
        return String(value);
    }
    function area(receiver) {
        const result = areas.get(receiver);
        if (!result) throw new TypeError("Illegal invocation");
        return result;
    }
    function values(state) { return new Map(JSON.parse(read(state.local))); }
    function put(state, key, value) { write(state.local, 'set', key, value); }
    function remove(state, key) { write(state.local, 'remove', key, ''); }
    class Storage {
        constructor() { throw new TypeError("Illegal constructor"); }
        get length() { return values(area(this)).size; }
        key(index) {
            const state = area(this);
            if (!arguments.length) throw new TypeError("An index is required");
            if (typeof index === 'bigint') throw new TypeError("Cannot convert a BigInt to a number");
            index = Number(index) >>> 0;
            return Array.from(values(state).keys())[index] ?? null;
        }
        getItem(key) {
            const state = area(this);
            if (!arguments.length) throw new TypeError("A key is required");
            key = text(key);
            return values(state).get(key) ?? null;
        }
        setItem(key, value) {
            const state = area(this);
            if (arguments.length < 2) throw new TypeError("A key and value are required");
            key = text(key);
            value = text(value);
            put(state, key, value);
        }
        removeItem(key) {
            const state = area(this);
            if (!arguments.length) throw new TypeError("A key is required");
            remove(state, text(key));
        }
        clear() {
            const state = area(this);
            write(state.local, 'clear', '', '');
        }
        get [Symbol.toStringTag]() { return "Storage"; }
    }
    function create(local) {
        const target = Object.create(Storage.prototype);
        const state = { local };
        function visible(key) { return typeof key === "string" && !Reflect.has(target, key) && values(state).has(key); }
        const proxy = new Proxy(target, {
            get(target, key, receiver) { return visible(key) ? values(state).get(key) : Reflect.get(target, key, receiver); },
            has(target, key) { return Reflect.has(target, key) || visible(key); },
            set(target, key, value) {
                if (typeof key !== "string") return Reflect.set(target, key, value);
                put(state, key, text(value));
                return true;
            },
            deleteProperty(target, key) {
                if (typeof key !== "string") return Reflect.deleteProperty(target, key);
                remove(state, key);
                return true;
            },
            ownKeys(target) { return [...values(state).keys()].filter(visible).concat(Reflect.ownKeys(target)); },
            getOwnPropertyDescriptor(target, key) {
                return visible(key) ? { value: values(state).get(key), writable: true, enumerable: true, configurable: true } : Reflect.getOwnPropertyDescriptor(target, key);
            },
            defineProperty(target, key, descriptor) {
                if (typeof key !== "string") return Reflect.defineProperty(target, key, descriptor);
                if (!("value" in descriptor) || descriptor.configurable === false || descriptor.enumerable === false || descriptor.writable === false) return false;
                put(state, key, text(descriptor.value));
                return true;
            },
            preventExtensions() { return false; }
        });
        areas.set(proxy, state);
        return proxy;
    }
    return { Storage, localStorage: create(true), sessionStorage: create(false) };
})
