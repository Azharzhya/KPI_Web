/*
 * PWA: pendaftaran service worker + tombol "Pasang" (Android / Chrome).
 * Banner pemasangan hanya muncul saat browser menyatakan aplikasi bisa dipasang,
 * tidak muncul bila sudah dipasang / dibuka dari layar utama, dan bisa ditutup (ditahan 7 hari).
 */
(function () {
  'use strict';

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('sw.js').catch(function (error) {
        console.warn('Service worker gagal didaftarkan:', error);
      });
    });
  }

  var banner = document.getElementById('installBanner');
  var installButton = document.getElementById('installButton');
  var dismissButton = document.getElementById('installDismiss');
  if (!banner || !installButton) return;

  var DISMISS_KEY = 'datastaf-install-dismissed';
  var DISMISS_MS = 7 * 24 * 60 * 60 * 1000;
  var deferredPrompt = null;

  function isInstalled() {
    return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone === true;
  }

  function recentlyDismissed() {
    try {
      var saved = Number(localStorage.getItem(DISMISS_KEY) || 0);
      return saved && Date.now() - saved < DISMISS_MS;
    } catch (error) { return false; }
  }

  function hideBanner() { banner.classList.add('hidden'); }

  window.addEventListener('beforeinstallprompt', function (event) {
    event.preventDefault();
    deferredPrompt = event;
    if (!isInstalled() && !recentlyDismissed()) banner.classList.remove('hidden');
  });

  installButton.addEventListener('click', function () {
    if (!deferredPrompt) return;
    var promptEvent = deferredPrompt;
    deferredPrompt = null;
    hideBanner();
    promptEvent.prompt();
    promptEvent.userChoice.catch(function () {});
  });

  if (dismissButton) {
    dismissButton.addEventListener('click', function () {
      hideBanner();
      try { localStorage.setItem(DISMISS_KEY, String(Date.now())); } catch (error) { /* abaikan */ }
    });
  }

  window.addEventListener('appinstalled', function () {
    deferredPrompt = null;
    hideBanner();
  });
})();
