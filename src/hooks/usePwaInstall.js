import { useEffect, useRef, useState } from "react";

const SNOOZE_KEY = "mediadeck_pwa_install_snoozed_at";
const SNOOZE_DAYS = 7;

const isStandalone = () =>
  window.matchMedia?.("(display-mode: standalone)")?.matches ||
  window.navigator?.standalone === true;

const snoozeIsActive = () => {
  try {
    const v = Number(localStorage.getItem(SNOOZE_KEY) || 0);
    return v ? Date.now() - v < SNOOZE_DAYS * 24 * 60 * 60 * 1000 : false;
  } catch {
    return false;
  }
};

export function usePwaInstall() {
  const deferredPrompt = useRef(null);
  const [canInstall, setCanInstall] = useState(false);
  const [showBanner, setShowBanner] = useState(false);
  const [isInstalling, setIsInstalling] = useState(false);

  useEffect(() => {
    if (isStandalone()) return;

    const onBeforeInstall = (e) => {
      e.preventDefault();
      deferredPrompt.current = e;
      setCanInstall(true);
      if (!snoozeIsActive()) setShowBanner(true);
    };

    const onInstalled = () => {
      deferredPrompt.current = null;
      setCanInstall(false);
      setShowBanner(false);
      try { localStorage.removeItem(SNOOZE_KEY); } catch {}
    };

    window.addEventListener("beforeinstallprompt", onBeforeInstall);
    window.addEventListener("appinstalled", onInstalled);
    return () => {
      window.removeEventListener("beforeinstallprompt", onBeforeInstall);
      window.removeEventListener("appinstalled", onInstalled);
    };
  }, []);

  const install = async () => {
    const e = deferredPrompt.current;
    if (!e) { setCanInstall(false); setShowBanner(false); return; }
    setIsInstalling(true);
    try {
      await e.prompt();
      const { outcome } = await e.userChoice;
      deferredPrompt.current = null;
      setCanInstall(false);
      setShowBanner(false);
      if (outcome === "accepted") {
        // Request notification permission so the app can send alerts
        if ("Notification" in window && Notification.permission === "default") {
          await Notification.requestPermission();
        }
        // Request persistent storage so the OS won't evict the app's cached data
        if (navigator.storage?.persist) {
          await navigator.storage.persist();
        }
      } else {
        try { localStorage.setItem(SNOOZE_KEY, String(Date.now())); } catch {}
      }
    } catch {
      setCanInstall(false);
      setShowBanner(false);
    } finally {
      setIsInstalling(false);
    }
  };

  const dismissBanner = () => {
    try { localStorage.setItem(SNOOZE_KEY, String(Date.now())); } catch {}
    setShowBanner(false);
  };

  return { canInstall, showBanner, isInstalling, install, dismissBanner };
}
