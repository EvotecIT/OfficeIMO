(() => {
    const scope = document.currentScript.closest("[data-officeimo-chart-export]") || document;
    for (const figure of scope.querySelectorAll("figure[data-officeimo-chart]")) mount(figure);

    function mount(figure) {
        if (figure.hasAttribute("data-chart-mounted")) return;
        figure.setAttribute("data-chart-mounted", "");
        const plot = figure.querySelector("svg");
        const tools = figure.querySelector("[data-officeimo-chart-tools]");
        const highlight = figure.querySelector("[data-officeimo-chart-highlight]");
        const value = figure.querySelector("[data-officeimo-chart-value]");
        const data = figure.querySelector("[data-officeimo-chart-marks]");
        const marks = data ? JSON.parse(data.content.textContent) : [];
        const { width, height } = plot.viewBox.baseVal;
        const zoomIn = tools.querySelector('[data-chart-action="in"]');
        const zoomOut = tools.querySelector('[data-chart-action="out"]');
        const reset = tools.querySelector('[data-chart-action="reset"]');
        let x = 0, y = 0, scale = 1, selected = -1;
        let drag = null, hover = null, frame = 0;

        tools.hidden = false;
        plot.setAttribute("tabindex", "0");
        plot.setAttribute("aria-keyshortcuts", "+ - 0 ArrowLeft ArrowRight ArrowUp ArrowDown");
        if (marks.length) {
            tools.querySelector('[data-chart-action="previous"]').hidden = false;
            tools.querySelector('[data-chart-action="next"]').hidden = false;
            value.hidden = false;
        }
        updateView();

        function updateView() {
            const w = width / scale;
            const h = height / scale;
            x = Math.min(width - w, Math.max(0, x));
            y = Math.min(height - h, Math.max(0, y));
            plot.setAttribute("viewBox", [x, y, w, h].join(" "));
            plot.toggleAttribute("data-chart-zoomed", scale > 1);
            zoomIn.setAttribute("aria-disabled", scale >= 8 ? "true" : "false");
            zoomOut.setAttribute("aria-disabled", scale <= 1 ? "true" : "false");
            reset.setAttribute("aria-disabled", scale <= 1 && selected < 0 ? "true" : "false");
        }

        function zoom(factor) {
            const previous = scale;
            scale = Math.min(8, Math.max(1, scale * factor));
            x += width * (1 / previous - 1 / scale) / 2;
            y += height * (1 / previous - 1 / scale) / 2;
            updateView();
        }

        function select(index, reveal = false) {
            if (selected !== index) {
                selected = index;
                const mark = marks[index];
                if (mark) {
                    for (const [i, attr] of ["x", "y", "width", "height"].entries()) highlight.setAttribute(attr, mark[i]);
                    highlight.removeAttribute("visibility");
                    value.textContent = mark[4];
                } else {
                    highlight.setAttribute("visibility", "hidden");
                    value.textContent = "";
                }
                reset.setAttribute("aria-disabled", scale <= 1 && selected < 0 ? "true" : "false");
            }
            if (reveal && scale > 1 && marks[index]) {
                const [left, top, w, h] = marks[index];
                x = left + (w - width / scale) / 2;
                y = top + (h - height / scale) / 2;
                updateView();
            }
        }

        function inspect(clientX, clientY) {
            const matrix = plot.getScreenCTM();
            if (!matrix) return;
            const point = new DOMPoint(clientX, clientY).matrixTransform(matrix.inverse());
            let nearest = -1, best = Infinity;
            for (let i = 0; i < marks.length; i++) {
                const [left, top, w, h] = marks[i];
                if (point.x < left || point.x > left + w || point.y < top || point.y > top + h) continue;
                const distance = (point.x - left - w / 2) ** 2 + (point.y - top - h / 2) ** 2;
                if (distance <= best) {
                    nearest = i;
                    best = distance;
                }
            }
            select(nearest);
        }

        function activate(action) {
            switch (action) {
                case "in": zoom(1.5); break;
                case "out": zoom(1 / 1.5); break;
                case "reset":
                    scale = 1;
                    select(-1);
                    updateView();
                    break;
                case "next": select((selected + 1) % marks.length, true); break;
                case "previous": select(selected <= 0 ? marks.length - 1 : selected - 1, true); break;
            }
        }

        tools.addEventListener("click", event => {
            const button = event.target.closest("button[data-chart-action]");
            if (button && button.getAttribute("aria-disabled") !== "true") activate(button.dataset.chartAction);
        });
        plot.addEventListener("keydown", event => {
            switch (event.key) {
                case "+": case "=": activate("in"); break;
                case "-": activate("out"); break;
                case "0": activate("reset"); break;
                case "ArrowLeft": if (scale <= 1) return; x -= width / scale / 8; updateView(); break;
                case "ArrowRight": if (scale <= 1) return; x += width / scale / 8; updateView(); break;
                case "ArrowUp": if (scale <= 1) return; y -= height / scale / 8; updateView(); break;
                case "ArrowDown": if (scale <= 1) return; y += height / scale / 8; updateView(); break;
                default: return;
            }
            if (drag) {
                const pointerId = drag.id;
                finishDrag({ pointerId });
                plot.releasePointerCapture(pointerId);
            }
            event.preventDefault();
        });
        plot.addEventListener("pointerdown", event => {
            if (scale <= 1 || !event.isPrimary || event.button !== 0) return;
            const matrix = plot.getScreenCTM();
            if (!matrix) return;
            drag = { id: event.pointerId, cx: event.clientX, cy: event.clientY, x, y, inverse: matrix.inverse(), moved: false };
            plot.setPointerCapture(event.pointerId);
        });
        plot.addEventListener("pointermove", event => {
            if (drag?.id === event.pointerId) {
                const dx = event.clientX - drag.cx;
                const dy = event.clientY - drag.cy;
                drag.moved ||= Math.abs(dx) + Math.abs(dy) >= 3;
                if (drag.moved) {
                    plot.setAttribute("data-chart-dragging", "");
                    x = drag.x - drag.inverse.a * dx - drag.inverse.c * dy;
                    y = drag.y - drag.inverse.b * dx - drag.inverse.d * dy;
                    updateView();
                }
                return;
            }
            if (event.pointerType !== "mouse" || marks.length === 0) return;
            hover = [event.clientX, event.clientY];
            if (frame) return;
            frame = requestAnimationFrame(() => {
                frame = 0;
                if (hover && !drag) inspect(...hover);
            });
        });
        plot.addEventListener("pointerleave", () => {
            hover = null;
            // Preserve the last inspected value when moving to the controls.
        });
        plot.addEventListener("pointerup", event => {
            if (event.isPrimary && !drag?.moved && marks.length) inspect(event.clientX, event.clientY);
            finishDrag(event);
        });
        plot.addEventListener("pointercancel", finishDrag);
        plot.addEventListener("lostpointercapture", finishDrag);

        function finishDrag(event) {
            if (drag?.id !== event.pointerId) return;
            drag = null;
            plot.removeAttribute("data-chart-dragging");
        }
    }
})();
