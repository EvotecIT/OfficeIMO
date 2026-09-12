(() => {
    "use strict";
    const active = new Map();
    return function (nativeTimer, normalizeWindow, kind) {
        if (kind === 'clearTimeout' || kind === 'clearInterval') return function (handle) {
            handle = +handle | 0;
            const state = active.get(handle);
            if (state) { state.active = false; active.delete(handle); }
            return nativeTimer.call(this, handle);
        };
        return function (callback, delay, ...args) {
            if (typeof callback !== 'function') throw new TypeError("This runtime supports function timer callbacks");
            delay = Math.max(kind === 'setInterval' ? 1 : 0, +delay | 0);
            const state = { active: true, handle: 0 };
            state.handle = nativeTimer.call(this, function () {
                if (!state.active) return;
                if (kind === 'setTimeout') { state.active = false; active.delete(state.handle); }
                return callback.apply(normalizeWindow(this), args);
            }, delay);
            active.set(state.handle, state);
            return state.handle;
        };
    };
})()
