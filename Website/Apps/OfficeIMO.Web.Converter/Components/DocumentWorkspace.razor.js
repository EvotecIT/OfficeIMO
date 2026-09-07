let selectionHandler;
let menuKeyHandler;
let navigationHandler;

export function connect(component) {
    disconnect();
    selectionHandler = event => {
        if (event.source !== window.parent || event.origin !== window.location.origin ||
            event.data?.type !== "officeimo:restore-selection") return;
        const { workspace, route, tool } = event.data;
        if ([workspace, route, tool].some(value => value != null && typeof value !== "string")) return;
        component.invokeMethodAsync("RestoreSelection", workspace ?? null, route ?? null, tool ?? null);
    };
    window.addEventListener("message", selectionHandler);
    navigationHandler = event => {
        if (event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
        const link = event.target.closest?.('#workspace-navigation a[data-route], #workspace-navigation a[data-pdf-tool]');
        if (!link) return;
        event.preventDefault();
        event.stopImmediatePropagation();
        component.invokeMethodAsync('RestoreSelection', link.dataset.pdfTool ? 'pdf' : 'convert', link.dataset.route ?? null, link.dataset.pdfTool ?? null);
    };
    document.addEventListener('click', navigationHandler, true);
    menuKeyHandler = event => {
        const typing = event.target.closest?.('input, textarea, select, [contenteditable="true"], [role="textbox"]');
        if (!event.defaultPrevented && !event.isComposing && window.parent !== window &&
            ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k' ||
             event.key === '/' && !typing && !event.ctrlKey && !event.metaKey && !event.altKey)) {
            event.preventDefault();
            window.parent.postMessage({ type: 'officeimo:open-search' }, window.location.origin);
            return;
        }
        const menu = document.querySelector('#workspace-navigation.is-open');
        if (event.key !== 'Tab' || !menu || window.innerWidth >= 1024) return;
        const links = menu.querySelectorAll('a[href]');
        const toggle = document.querySelector('#workspace-menu-toggle');
        if (event.shiftKey && document.activeElement === links[0] || !event.shiftKey && document.activeElement === links[links.length - 1]) {
            event.preventDefault(); toggle?.focus();
        } else if (document.activeElement === toggle) {
            event.preventDefault(); links[event.shiftKey ? links.length - 1 : 0]?.focus();
        }
    };
    window.addEventListener('keydown', menuKeyHandler);
}
export function publishSelection(workspace, route, tool, replace) {
    const content = document.querySelector('.ocx-workspace-content');
    if (content) content.scrollTop = 0;
    if (window.innerWidth < 1024 && document.activeElement?.closest('#workspace-navigation')) {
        document.querySelector('#workspace-title')?.focus({ preventScroll: true });
    }
    if (window.parent !== window) {
        window.parent.postMessage({ type: "officeimo:workspace-selection", workspace, route, tool, replace }, window.location.origin);
    }
}
export function focusMenu(open) {
    requestAnimationFrame(() => {
        document.querySelector(open ? "#workspace-navigation a[aria-current]" : "#workspace-menu-toggle")?.focus();
    });
}
export function disconnect() {
    if (selectionHandler) window.removeEventListener("message", selectionHandler);
    if (menuKeyHandler) window.removeEventListener('keydown', menuKeyHandler);
    if (navigationHandler) document.removeEventListener('click', navigationHandler, true);
    selectionHandler = undefined;
    menuKeyHandler = undefined;
    navigationHandler = undefined;
}
