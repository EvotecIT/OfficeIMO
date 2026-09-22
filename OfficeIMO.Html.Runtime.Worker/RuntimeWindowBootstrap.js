(() => {
    "use strict";
    const active = new Set();
    return function (nativeTimer, normalizeWindow, kind, alive, track, release) {
        if (kind === 'clearTimeout' || kind === 'clearInterval') return function (handle) {
            handle = +handle | 0;
            const target = normalizeWindow(this);
            for (const state of active) {
                if (state.handle === handle && state.target === target) { state.active = false; active.delete(state); }
            }
            release(this, handle);
            return nativeTimer.call(this, handle);
        };
        return function (callback, delay, ...args) {
            if (typeof callback !== 'function') throw new TypeError("This runtime supports function timer callbacks");
            delay = Math.max(kind === 'setInterval' ? 1 : 0, +delay | 0);
            const receiver = this;
            const state = { active: true, handle: 0, target: normalizeWindow(this) };
            state.handle = nativeTimer.call(this, function () {
                if (!state.active || !alive()) return;
                if (kind === 'setTimeout') { state.active = false; active.delete(state); release(receiver, state.handle); }
                return callback.apply(normalizeWindow(this), args);
            }, delay);
            active.add(state);
            track(receiver, state.handle);
            return state.handle;
        };
    };
})()
