document.querySelector('#prepare').addEventListener('click', function () {
  const title = document.querySelector('#title').value;
  const region = document.querySelector('#region').value;
  const approval = document.querySelector('#approved').checked ? 'approved' : 'pending';
  const notes = document.querySelector('#notes').value;
  document.querySelector('#summary').textContent = `${title} | ${region} | ${approval} | ${notes}`;
  window.applicationCorpusReady = true;
});
