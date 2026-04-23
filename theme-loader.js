/* Synchronous theme loader. Must be a blocking script in <head> before
   styles.css loads so that the data-theme attribute is in place before
   the first paint. Kept minimal and defensive. */
(function () {
  'use strict';
  var stored = null;
  try { stored = localStorage.getItem('uc-leave-calc-theme'); } catch (e) { /* ignore */ }
  var resolved = (stored === 'light' || stored === 'dark')
    ? stored
    : (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  document.documentElement.setAttribute('data-theme', resolved);
})();
