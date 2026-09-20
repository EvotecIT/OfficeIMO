import { formatRows } from './format.js';

window.graphBootOrder.push('module');
fetch('./data.json').then(response => response.json()).then(rows => {
  document.querySelector('#status').textContent = `Loaded ${formatRows(rows)}`;
  const button = document.querySelector('#summarize');
  button.disabled = false;
  button.addEventListener('click', async () => {
    const { summarize } = await import('./summary.js');
    document.querySelector('#summary').textContent = `Qualified graph total: ${summarize(rows)}`;
    window.applicationCorpusReady = true;
  });
  window.applicationCorpusInteractive = true;
});
