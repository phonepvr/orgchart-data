// Boot diagnostics. Written in old-style (ES5) JavaScript on purpose, so it
// still runs on browsers/servers where the main app scripts fail — its whole
// job is to explain a blank page instead of leaving it blank.
(function () {
  var problems = [];

  window.addEventListener('error', function (e) {
    var t = e.target;
    if (t && t !== window && (t.src || t.href)) {
      problems.push('Failed to load: ' + (t.src || t.href));
    } else if (e.message) {
      problems.push(e.message + (e.filename ? ' (' + e.filename + ':' + e.lineno + ')' : ''));
    }
  }, true);

  window.addEventListener('unhandledrejection', function (e) {
    problems.push('Unhandled promise rejection: ' + e.reason);
  });

  document.addEventListener('securitypolicyviolation', function (e) {
    problems.push('Content-Security-Policy blocked: ' + (e.blockedURI || '(inline)') +
      ' [' + e.violatedDirective + ']');
  });

  window.addEventListener('load', function () {
    setTimeout(function () {
      if (window.__orgSenseBooted) return;
      var box = document.getElementById('boot-fallback');
      var list = document.getElementById('boot-problems');
      if (!box || !list) return;
      if (problems.length === 0) {
        problems.push('No script errors were captured, but the app never started. ' +
          'Check that the js/, js/render/ and vendor/ folders were uploaded next to index.html.');
      }
      for (var i = 0; i < problems.length; i++) {
        var li = document.createElement('li');
        li.textContent = problems[i];
        list.appendChild(li);
      }
    }, 1500);
  });
})();
