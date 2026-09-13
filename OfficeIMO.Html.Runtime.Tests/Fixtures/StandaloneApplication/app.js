import { visibleRows, renderRows, total } from './view.js';

const title = document.querySelector('#title');
const region = document.querySelector('#region');
const results = document.querySelector('#results');
const totalElement = document.querySelector('#total');
const storedTitle = localStorage.getItem('application:title');
const storedRegion = localStorage.getItem('application:region');
if (storedTitle !== null) title.value = storedTitle;
region.value = storedRegion === null ? 'all' : storedRegion;

window.applicationMutations = 0;
new MutationObserver(records => {
    applicationMutations += records.length;
}).observe(results, { childList: true, subtree: true });

const dataResponse = await fetch('./data.json');
const rows = await dataResponse.json();
const show = () => {
    const visible = visibleRows(rows, region.value);
    renderRows(results, visible);
    totalElement.textContent = `Total: ${total(visible)}`;
};

title.addEventListener('input', event => localStorage.setItem('application:title', event.currentTarget.value));
region.addEventListener('change', event => {
    localStorage.setItem('application:region', event.currentTarget.value);
    show();
});
document.querySelector('form').addEventListener('submit', () => sessionStorage.setItem('application:submitted', 'yes'));
window.addEventListener('pagehide', event => sessionStorage.setItem('application:pagehide', String(event.persisted)));
window.addEventListener('unload', () => sessionStorage.setItem('application:unload', 'yes'));

show();
const healthResponse = await fetch('./health.txt');
const health = document.querySelector('#health');
health.textContent = healthResponse.ok ? 'Service ready' : `Service degraded (${healthResponse.status})`;
health.setAttribute('data-state', healthResponse.ok ? 'ready' : 'degraded');
window.applicationReady = true;
