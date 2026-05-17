(function registerActivityReportPwa() {
  if (!("serviceWorker" in navigator)) return;

  function appBaseUrl() {
    const manifestLink = document.querySelector('link[rel="manifest"]');
    if (manifestLink && manifestLink.href) {
      return new URL("./", manifestLink.href);
    }
    return new URL("./", window.location.href);
  }

  window.addEventListener("load", () => {
    const base = appBaseUrl();
    const swUrl = new URL("sw.js", base).href;
    const scope = new URL("./", base).href;

    navigator.serviceWorker
      .register(swUrl, { scope })
      .then((reg) => {
        reg.update();
      })
      .catch((err) => {
        console.warn("[pwa] Service worker registration failed:", err);
      });
  });
})();
