/* Applies a saved light/dark choice before the page paints, so there is no flash of the wrong theme. */
(function () {
  try {
    var t = localStorage.getItem('t2x-theme');
    if (t === 'dark' || t === 'light') document.documentElement.setAttribute('data-theme', t);
  } catch (e) { /* storage blocked: follow the system theme */ }
})();
