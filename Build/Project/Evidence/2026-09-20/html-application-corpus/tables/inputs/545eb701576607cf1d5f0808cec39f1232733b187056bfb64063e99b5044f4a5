const body = document.querySelector('#ledger-body');
let total = 2230;
for (let index = 1; index <= 42; index++) {
  const amount = 100 + index;
  total += amount;
  const row = document.createElement('tr');
  row.innerHTML = `<th scope="row">AC-${String(index).padStart(3, '0')}</th><td>${index % 2 ? 'North' : 'South'}</td><td>Retained ledger entry ${index}</td><td>${amount}</td>`;
  body.appendChild(row);
}
document.querySelector('#total').textContent = String(total);
window.applicationCorpusReady = true;
