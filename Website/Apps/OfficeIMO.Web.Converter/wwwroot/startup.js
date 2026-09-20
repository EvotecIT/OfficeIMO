/* Startup feedback uses Blazor's measured resource progress, never a simulated timer. */
(function () {
    const status = document.getElementById('startup-status');
    const progress = document.getElementById('startup-progress');
    const retry = document.getElementById('startup-retry');
    const slow = document.getElementById('startup-slow');
    retry.addEventListener('click', () => location.reload());
    let last = -1;
    const interval = setInterval(() => {
        if (!status.isConnected) { clearInterval(interval); return; }
        const value = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--blazor-load-percentage'));
        if (!Number.isFinite(value)) return;
        const percent = Math.min(100, Math.max(0, Math.round(value)));
        progress.value = percent;
        if (percent !== last) {
            status.textContent = percent < 100 ? `Downloading workspace · ${percent}%` : 'Starting the document workspace…';
            last = percent;
        }
    }, 250);
    const delay = setTimeout(() => {
        if (!slow.isConnected) return;
        slow.hidden = false;
        retry.hidden = false;
    }, 10000);
    function failed(error) {
        clearInterval(interval);
        clearTimeout(delay);
        if (!status.isConnected) return;
        document.getElementById('startup-title').textContent = 'The workspace could not start';
        status.textContent = 'Check your connection, then reload to try again.';
        progress.hidden = true;
        retry.hidden = false;
        console.error('OfficeIMO workspace startup failed.', error);
    }
    try {
        Blazor.start().then(() => { clearInterval(interval); clearTimeout(delay); }).catch(failed);
    } catch (error) { failed(error); }
})();
