
(() => {
  const FUNCTION_URL =
    "https://appdqolpwvjktltvtqvj.supabase.co/functions/v1/outlook-connect";

  const CONFIG = window.MYJOURNEY_CONFIG;

  if (!CONFIG?.url || !CONFIG?.publishableKey || !window.supabase) {
    console.error("Outlook: Supabase-Konfiguration fehlt.");
    return;
  }

  const client = window.supabase.createClient(
    CONFIG.url,
    CONFIG.publishableKey,
    {
      auth: {
        storageKey: "myjourney-auth-token"
      }
    }
  );

  let connecting = false;

  document.addEventListener("click", async (event) => {
    const button = event.target.closest("#connectOutlook");
    if (!button) return;

    // Der bestehende app.js-Button zeigt sonst nur einen Hinweis.
    // Deshalb verhindern wir hier dessen bisherige Aktion nicht
    // durch Änderung der app.js, sondern ergänzen die Verbindung.
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
          result.error || "Microsoft-Anmeldung konnte nicht gestartet werden."
        );
      }

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
  }, true);

  const params = new URLSearchParams(window.location.search);
  const status = params.get("outlook");

  if (status) {
    const messages = {
      connected: "Dein Outlook-Konto wurde erfolgreich verbunden! 🎉",
      cancelled: "Die Outlook-Anmeldung wurde abgebrochen.",
      expired: "Die Outlook-Anmeldung ist abgelaufen. Bitte versuche es erneut.",
      error: "Die Outlook-Verbindung konnte nicht abgeschlossen werden."
    };

    if (messages[status]) {
      alert(messages[status]);
    }

    params.delete("outlook");

    const remaining = params.toString();
    const cleanUrl =
      window.location.pathname +
      (remaining ? "?" + remaining : "") +
      window.location.hash;

    window.history.replaceState({}, "", cleanUrl);
  }
})();
