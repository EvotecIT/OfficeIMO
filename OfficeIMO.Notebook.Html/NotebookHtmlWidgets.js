(() => {
    const scope = document.currentScript.parentElement;
    const views = new Map();
    const links = new Map();
    const timers = new Map();
    const idOf = control => control.dataset.officeimoWidgetId;
    const kindOf = control => control.dataset.officeimoWidgetKind;
    const minOf = control => Number(control.dataset.officeimoWidgetMin);

    function append(map, key, value) {
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(value);
    }

    function read(control) {
        switch (kindOf(control)) {
            case "boolean": return control.checked;
            case "text": return control.value;
            case "text-value": return control.textContent;
            case "progress": return Number(control.dataset.officeimoWidgetValue);
            default: return control.value === "" || control.validity?.valid === false ? null : Number(control.value);
        }
    }

    function write(control, value) {
        switch (kindOf(control)) {
            case "boolean":
                control.checked = value;
                break;
            case "text-value":
                control.textContent = value;
                break;
            case "progress":
                control.dataset.officeimoWidgetValue = String(value);
                control.value = value - minOf(control);
                break;
            default:
                control.value = value;
                if (kindOf(control) === "play") control.nextElementSibling.textContent = control.value;
                break;
        }
    }

    function publish(source, value, initiatingView) {
        if (value === null) return;
        const seen = new Set([source]);
        const pending = [source];
        for (let index = 0; index < pending.length; index++) {
            const id = pending[index];
            for (const control of views.get(id) || []) {
                if (control !== initiatingView) write(control, value);
            }
            for (const target of links.get(id) || []) {
                if (!seen.has(target)) {
                    seen.add(target);
                    pending.push(target);
                }
            }
        }
    }

    function togglePlay(control) {
        const id = idOf(control);
        const min = minOf(control);
        const max = Number(control.dataset.officeimoWidgetMax);
        const step = Number(control.dataset.officeimoWidgetStep);
        const repeat = control.dataset.officeimoWidgetRepeat === "true";
        const setPlaying = playing => {
            for (const view of views.get(id)) {
                view.setAttribute("aria-pressed", String(playing));
                view.textContent = playing ? "Pause" : "Play";
            }
        };
        const stop = () => {
            clearInterval(timers.get(id));
            timers.delete(id);
            setPlaying(false);
        };
        if (timers.has(id)) {
            stop();
            return;
        }
        if (read(control) >= max) publish(id, min);
        setPlaying(true);
        timers.set(id, setInterval(() => {
            if (!scope.isConnected || !control.isConnected) {
                stop();
                return;
            }
            if (read(control) >= max) {
                if (repeat) publish(id, min);
                else stop();
                return;
            }
            publish(id, Math.min(max, read(control) + step));
        }, Number(control.dataset.officeimoWidgetInterval)));
    }

    for (const control of scope.querySelectorAll("[data-officeimo-widget-id]")) {
        append(views, idOf(control), control);
        if (kindOf(control) === "play") {
            control.addEventListener("click", () => togglePlay(control));
        } else {
            const event = control.hasAttribute("data-officeimo-widget-commit-only") ? "change" : "input";
            control.addEventListener(event, () => publish(idOf(control), read(control), control));
        }
    }
    for (const link of scope.querySelectorAll("[data-officeimo-widget-link-from]")) {
        const from = link.dataset.officeimoWidgetLinkFrom;
        const to = link.dataset.officeimoWidgetLinkTo;
        append(links, from, to);
        if (link.hasAttribute("data-officeimo-widget-link-reverse")) append(links, to, from);
    }
    for (const id of links.keys()) publish(id, read(views.get(id)[0]));
})();
