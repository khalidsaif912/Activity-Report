(function registerActivityReportPwa() {
  if (!("serviceWorker" in navigator)) return;

  const swUrl = new URL("/sw.js", window.location.origin).href;

  window.addEventListener("load", () => {
    navigator.serviceWorker
      .register(swUrl, { scope: "/" })
      .then((reg) => {
        reg.update();
      })
      .catch((err) => {
        console.warn("[pwa] Service worker registration failed:", err);
      });
  });
})();
