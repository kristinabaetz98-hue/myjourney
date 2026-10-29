
/* MyJourney – Outlook-Verbindung */

(() => {
  const CONFIG = window.MYJOURNEY_CONFIG || {};

  const FUNCTION_URL =
    "https://appdqolpwvjktltvtqvj.supabase.co/functions/v1/outlook-connect";

  if (
    !CONFIG.url ||
    !CONFIG.publishableKey ||
    !window.supabase?.createClient
  ) {
    console.error("Outlook: Supabase ist nicht verfügbar.");
    return;
  }

  // Dieselbe Supabase-Konfiguration wie in cloud.js.
  // Keinen abweichenden storageKey verwenden.
  const client = window.supabase.createClient(
    CONFIG.url,
    CONFIG.publishableKey,
    {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true
      }
    }
  );

  let connecting = false;

  async function connectOutlook(button) {
    if (connecting) return;

    connecting = true;
    button.disabled = true;
    button.textContent = "Outlook wird verbunden ...";

    try {
      const { data, error } = await client.auth.getSession();

      if (error || !data.session?.access_token) {
        throw new Error(
          "Bitte melde dich zuerst bei MyJourney an."
        );
      }

      const response = await fetch(FUNCTION_URL + "/start", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + data.session.access_token,
          "apikey": CONFIG.publishableKey
        },
        body: JSON.stringify({})
      });

      const result = await response.json();

      if (!response.ok || !result.authorizationUrl) {
        throw new Error(
          result.error ||
          "Die Microsoft-Anmeldung konnte nicht gestartet werden."
        );
      }

      // Microsoft-Anmeldung öffnen.
      window.location.assign(result.authorizationUrl);

    } catch (error) {
      alert(
        "Outlook-Verbindung fehlgeschlagen:\n" +
        (error.message || "Unbekannter Fehler")
      );

      button.disabled = false;
      button.textContent = "✉️ Outlook verbinden";
      connecting = false;
    }
  }

  // Capture-Phase: Verhindert, dass der alte
  // Hinweis aus app.js zusätzlich erscheint.
  document.addEventListener("click", (event) => {
    const button = event.target.closest?.("#connectOutlook");

    if (!button) return;

    event.preventDefault();
    event.stopPropagation();

    void connectOutlook(button);
  }, true);

  // Ergebnis nach der Microsoft-Anmeldung anzeigen.
  const params = new URLSearchParams(window.location.search);
  const status = params.get("outlook");

  const messages = {
    connected: "🎉 Dein Outlook-Konto wurde erfolgreich verbunden!",
    cancelled: "Die Outlook-Anmeldung wurde abgebrochen.",
    expired: "Die Outlook-Anmeldung ist abgelaufen. Bitte versuche es erneut.",
    error: "Die Outlook-Verbindung konnte nicht abgeschlossen werden."
  };

  if (status && messages[status]) {
    alert(messages[status]);

    params.delete("outlook");

    const remaining = params.toString();

    window.history.replaceState(
      {},
      "",
      window.location.pathname +
      (remaining ? "?" + remaining : "") +
      window.location.hash
    );
  }
})();
