(function (limit, read, write) {
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
    function put(state, key, value) {
        const previous = state.values.get(key);
        const size = state.size + value.length + (previous === undefined ? key.length : -previous.length);
        if (size > limit) {
            const error = new Error("The session storage area exceeds MaxStorageCharacters.");
            error.name = "QuotaExceededError";
            throw error;
        }
        write(state.local, 'set', key, value);
        state.values.set(key, value);
        state.size = size;
    }
    function remove(state, key) {
        const value = state.values.get(key);
        if (value !== undefined) {
            write(state.local, 'remove', key, '');
            state.values.delete(key);
            state.size -= key.length + value.length;
        }
    }
    class Storage {
        constructor() { throw new TypeError("Illegal constructor"); }
        get length() { return area(this).values.size; }
        key(index) {
            const state = area(this);
            if (!arguments.length) throw new TypeError("An index is required");
            if (typeof index === 'bigint') throw new TypeError("Cannot convert a BigInt to a number");
            index = Number(index) >>> 0;
            return Array.from(state.values.keys())[index] ?? null;
        }
        getItem(key) {
            const state = area(this);
            if (!arguments.length) throw new TypeError("A key is required");
            return state.values.get(text(key)) ?? null;
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
            state.values.clear();
            state.size = 0;
        }
        get [Symbol.toStringTag]() { return "Storage"; }
    }
    function create(local) {
        const target = Object.create(Storage.prototype);
        const values = new Map(JSON.parse(read(local)));
        const state = { values, size: Array.from(values).reduce((size,[key,value])=>size+key.length+value.length,0), local };
        function visible(key) { return typeof key === "string" && !Reflect.has(target, key) && state.values.has(key); }
        const proxy = new Proxy(target, {
            get(target, key, receiver) { return visible(key) ? state.values.get(key) : Reflect.get(target, key, receiver); },
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
            ownKeys(target) { return [...state.values.keys()].filter(visible).concat(Reflect.ownKeys(target)); },
            getOwnPropertyDescriptor(target, key) {
                return visible(key) ? { value: state.values.get(key), writable: true, enumerable: true, configurable: true } : Reflect.getOwnPropertyDescriptor(target, key);
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
