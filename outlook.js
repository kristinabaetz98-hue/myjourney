
/* MyJourney – Outlook-Verbindung und Live-Status */

(() => {
  const CONFIG = window.MYJOURNEY_CONFIG || {};

  const BASE_URL =
    "https://appdqolpwvjktltvtqvj.supabase.co/functions/v1";

  if (
    !CONFIG.url ||
    !CONFIG.publishableKey ||
    !window.supabase?.createClient
  ) {
    console.error("Outlook: Supabase ist nicht verfügbar.");
    return;
  }

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
  let checking = false;

  async function getSession() {
    const { data, error } = await client.auth.getSession();

    if (error || !data.session?.access_token) {
      throw new Error("Bitte zuerst bei MyJourney anmelden.");
    }

    return data.session;
  }

  async function requestFunction(name, body = {}) {
    const session = await getSession();

    const response = await fetch(BASE_URL + "/" + name, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + session.access_token,
        "apikey": CONFIG.publishableKey
      },
      body: JSON.stringify(body)
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(
        result.error || "Die Anfrage ist fehlgeschlagen."
      );
    }

    return result;
  }

  function statusArea() {
    const button = document.querySelector("#connectOutlook");

    if (!button) return null;

    let area = document.querySelector("#outlookLiveStatus");

    if (!area) {
      area = document.createElement("div");
      area.id = "outlookLiveStatus";
      area.style.marginTop = "12px";
      button.insertAdjacentElement("afterend", area);
    }

    return area;
  }

  async function updateOutlookStatus() {
    const button = document.querySelector("#connectOutlook");
    const area = statusArea();

    if (!button || !area || checking || connecting) return;

    checking = true;
    area.textContent = "Outlook-Verbindung wird geprüft ...";

    try {
      const result = await requestFunction("outlook-status");

      if (!document.contains(button)) return;

      const card = button.closest(".card");

      if (result.connected) {
        button.textContent = "✓ Outlook verbunden";
        button.disabled = true;

        area.replaceChildren();

        const status = document.createElement("p");
        status.style.color = "#21734b";
        status.style.fontWeight = "600";
        status.textContent = "✓ Microsoft Outlook ist verbunden";

        const email = document.createElement("p");
        email.className = "muted";
        email.textContent = result.email
          ? "Konto: " + result.email
          : "Dein Outlook-Konto ist verknüpft.";

        const info = document.createElement("p");
        info.className = "muted";
        info.textContent =
          "Die Verbindung ist eingerichtet. Der automatische Import von Buchungsbestätigungen folgt als Nächstes.";

        area.append(status, email, info);

        // Alten, inzwischen falschen Hinweis ausblenden.
        if (card) {
          card.querySelectorAll("p").forEach((p) => {
            if (
              p !== status &&
              p !== email &&
              p !== info &&
              (
                p.textContent.includes("noch nicht eingerichtet") ||
                p.textContent.includes("automatische E-Mail-Abruf folgt") ||
                p.textContent.includes("automatische E-Mail-Abruf")
              )
            ) {
              p.hidden = true;
            }
          });
        }

      } else {
        button.disabled = false;
        button.textContent = "✉️ Outlook verbinden";

        area.textContent = "Outlook ist noch nicht verbunden.";
      }

    } catch (error) {
      if (document.contains(area)) {
        area.textContent =
          "Verbindungsstatus konnte nicht geladen werden: " +
          error.message;
      }
    } finally {
      checking = false;
    }
  }

  async function connectOutlook(button) {
    if (connecting) return;

    connecting = true;
    button.disabled = true;
    button.textContent = "Outlook wird verbunden ...";

    try {
      const result = await requestFunction(
        "outlook-connect/start"
      );

      if (!result.authorizationUrl) {
        throw new Error(
          "Microsoft-Anmeldung konnte nicht gestartet werden."
        );
      }

      window.location.assign(result.authorizationUrl);

    } catch (error) {
      alert(
        "Outlook-Verbindung fehlgeschlagen:\n" +
        error.message
      );

      button.disabled = false;
      button.textContent = "✉️ Outlook verbinden";
      connecting = false;
    }
  }

  // Alten Klick-Hinweis aus app.js unterdrücken.
  document.addEventListener("click", (event) => {
    const button = event.target.closest?.("#connectOutlook");

    if (!button) return;

    event.preventDefault();
    event.stopPropagation();

    if (!button.disabled) {
      void connectOutlook(button);
    }
  }, true);

  // Erkennt, wenn der Bereich "Mehr" geöffnet wird.
  const view = document.querySelector("#view");

  if (view) {
    const observer = new MutationObserver(() => {
      if (
        document.querySelector("#connectOutlook") &&
        !checking
      ) {
        void updateOutlookStatus();
      }
    });

    observer.observe(view, {
      childList: true
    });
  }

  // Microsoft-Rückmeldung nach der Anmeldung.
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
