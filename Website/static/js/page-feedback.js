// "Was this page helpful?" on docs and API pages. The answer stays in the page:
// nothing is stored or sent. "No" points the reader at the prefilled issue link.
(function () {
  'use strict';

  function issueUrl(block) {
    var heading = document.querySelector('main h1');
    var title = (heading ? heading.textContent : document.title).replace(/\s+/g, ' ').trim();
    var page = window.location.origin + window.location.pathname;
    var body = 'Page: ' + page + '\nTitle: ' + title;
    var pkg = block.getAttribute('data-feedback-package');
    if (pkg) body += '\nPackage: ' + pkg;
    body += '\n\nWhat is wrong or missing?\n\n';
    return 'https://github.com/EvotecIT/OfficeIMO/issues/new?labels=documentation&title=' +
      encodeURIComponent('Docs: ' + title) + '&body=' + encodeURIComponent(body);
  }

  function init(block) {
    // API pages are generated without page context, so their issue link is filled in here.
    var issue = block.querySelector('[data-feedback-issue][data-feedback-dynamic]');
    if (issue) issue.href = issueUrl(block);

    var buttons = block.querySelectorAll('[data-feedback-answer]');
    Array.prototype.forEach.call(buttons, function (button) {
      button.addEventListener('click', function () {
        var answer = button.getAttribute('data-feedback-answer');
        Array.prototype.forEach.call(buttons, function (other) {
          other.setAttribute('aria-pressed', other === button ? 'true' : 'false');
        });
        Array.prototype.forEach.call(block.querySelectorAll('[data-feedback-reply]'), function (reply) {
          reply.hidden = reply.getAttribute('data-feedback-reply') !== answer;
        });
        block.setAttribute('data-feedback-state', answer);
      });
    });
  }

  function start() {
    Array.prototype.forEach.call(document.querySelectorAll('[data-page-feedback]'), init);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
