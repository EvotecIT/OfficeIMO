history.replaceState({ phase: 'review' }, '', '/report/review');
document.querySelector('#approve').addEventListener('click', function () {
  history.pushState({ phase: 'approved' }, '', '/report/approved');
  document.querySelector('#state').textContent = 'Approved route snapshot';
  window.applicationCorpusReady = true;
});
