export function visibleRows(rows, region) {
    return region === 'all' ? rows : rows.filter(row => row.region === region);
}

export function renderRows(target, rows) {
    target.textContent = '';
    for (const row of rows) {
        const item = document.createElement('li');
        item.textContent = `${row.region}: ${row.amount}`;
        item.setAttribute('data-amount', row.amount);
        target.appendChild(item);
    }
}

export function total(rows) {
    return rows.reduce((sum, row) => sum + row.amount, 0);
}
