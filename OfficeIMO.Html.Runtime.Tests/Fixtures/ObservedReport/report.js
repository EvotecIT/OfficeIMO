await new Promise(resolve => document.addEventListener('DOMContentLoaded', resolve, {once: true}));
const data = await fetch('./items.json').then(response => response.json());
if (document.readyState !== 'complete') await new Promise(resolve => window.addEventListener('load', resolve, {once: true}));
window.reportReadyState = document.readyState;
const items = document.querySelector('#items');
const total = document.querySelector('#total');
window.reportOrder = [];
new MutationObserver(records => {
    if (!records.some(record => record.target === items || items.contains(record.target))) return;
    reportOrder.push('observer');
    const amount = [...items.querySelectorAll('[data-amount]')]
        .reduce((sum, cell) => sum + Number(cell.getAttribute('data-amount')), 0);
    total.textContent = 'Total: ' + amount;
}).observe(document, {subtree: true, childList: true, attributes: true, attributeFilter: ['data-amount']});
Promise.resolve().then(() => reportOrder.push('before'));
for (const item of data) {
    const row = document.createElement('tr');
    const name = document.createElement('td');
    const amount = document.createElement('td');
    name.textContent = item.name;
    amount.textContent = String(item.amount);
    amount.setAttribute('data-amount', item.amount);
    row.appendChild(name);
    row.appendChild(amount);
    items.appendChild(row);
}
Promise.resolve().then(() => reportOrder.push('after'));
document.querySelector('#update').onclick = () => {
    const amount = items.querySelectorAll('[data-amount]')[1];
    amount.setAttribute('data-amount', 20);
    amount.textContent = '20';
};
