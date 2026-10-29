
/* MyJourney – Outlook-Verbindung und Posteingangs-Scan.
   Die Buchungsübersicht wird ausschließlich von outlook-smart.js erzeugt.
*/
(() => {
  const CONFIG = window.MYJOURNEY_CONFIG || {};
  const BASE_URL =
    "https://appdqolpwvjktltvtqvj.supabase.co/functions/v1";

  if (
    !CONFIG.url ||
    !CONFIG.publishableKey ||
    !window.supabase?.createClient
  ) {
    console.error("Outlook: Supabase nicht verfügbar.");
    return;
  }

  const client = window.supabase.createClient(
    CONFIG.url,
    CONFIG.publishableKey,
    {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true,
      },
    }
  );

  let connecting = false;
  let checking = false;
  let scanning = false;
  let scanStopRequested = false;
  let lastButton = null;

  async function session() {
    const { data, error } = await client.auth.getSession();

    if (error || !data.session?.access_token) {
      throw new Error("Bitte zuerst bei MyJourney anmelden.");
    }

    return data.session;
  }

  async function callFunction(name, body = {}) {
    const currentSession = await session();

    const response = await fetch(BASE_URL + "/" + name, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + currentSession.access_token,
        apikey: CONFIG.publishableKey,
      },
      body: JSON.stringify(body),
    });

    const result = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(
        result.error || "Die Anfrage ist fehlgeschlagen."
      );
    }

    return result;
  }

  function element(tag, text) {
    const node = document.createElement(tag);

    if (text !== undefined) {
      node.textContent = text;
    }

    return node;
  }

  function createButton(text, onClick) {
    const button = element("button", text);
    button.type = "button";
    button.style.padding = "10px 14px";
    button.style.margin = "5px 5px 5px 0";
    button.style.borderRadius = "9px";
    button.style.cursor = "pointer";
    button.addEventListener("click", onClick);
    return button;
  }

  function area() {
    const button = document.querySelector("#connectOutlook");
    if (!button) return null;

    let container = document.querySelector("#outlookLiveStatus");

    if (!container) {
      container = element("div");
      container.id = "outlookLiveStatus";
      container.style.marginTop = "14px";
      button.insertAdjacentElement("afterend", container);
    }

    return { button, container };
  }

  function hideOldNotice(button) {
    const card = button.closest(".card");
    if (!card) return;

    card.querySelectorAll("p").forEach((paragraph) => {
      const text = paragraph.textContent || "";

      if (
        text.includes("noch nicht eingerichtet") ||
        text.includes("automatische E-Mail-Abruf")
      ) {
        paragraph.hidden = true;
      }
    });
  }

  function notifySmartView() {
    document.dispatchEvent(
      new CustomEvent("myjourney:outlook-updated")
    );
  }

  function setProgress(message) {
    const progress = document.querySelector(
      "#outlookScanProgress"
    );

    if (progress) {
      progress.textContent = message;
    }
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  async function connectOutlook(button) {
    if (connecting) return;

    connecting = true;
    button.disabled = true;
    button.textContent = "Outlook wird verbunden ...";

    try {
      const result = await callFunction(
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

  async function refreshScanProgress() {
    try {
      const data = await callFunction(
        "outlook-import",
        { action: "list" }
      );

      if (!scanning) {
        const progress = data.progress || {};

        setProgress(
          (progress.scanned_count || 0) +
          " E-Mails bisher geprüft. " +
          (progress.completed
            ? "Posteingang vollständig durchsucht."
            : "Der Scan kann gestartet oder fortgesetzt werden.")
        );
      }

      notifySmartView();
    } catch (error) {
      setProgress(
        "Fortschritt konnte nicht geladen werden: " +
        error.message
      );
    }
  }

  async function startScan() {
    if (scanning) return;

    const panel = document.querySelector(
      "#outlookImportPanel"
    );

    if (!panel) return;

    scanning = true;
    scanStopRequested = false;

    const scanButton = panel.querySelector(
      "#outlookStartScan"
    );

    const refreshButton = panel.querySelector(
      "#outlookRefreshScan"
    );

    if (scanButton) scanButton.disabled = true;
    if (refreshButton) refreshButton.disabled = true;

    const stopButton = createButton(
      "Scan pausieren",
      () => {
        scanStopRequested = true;
        stopButton.disabled = true;
        stopButton.textContent = "Wird pausiert ...";
      }
    );

    stopButton.id = "outlookStopScan";

    const controls = panel.querySelector(
      "#outlookScanControls"
    );

    (controls || panel).append(stopButton);

    try {
      let completed = false;

      while (!completed && !scanStopRequested) {
        const result = await callFunction(
          "outlook-import",
          { action: "scan" }
        );

        completed = !!result.completed;

        setProgress(
          (result.scanned || 0) +
          " E-Mails geprüft. " +
          (completed
            ? "Posteingang vollständig durchsucht."
            : "Suche läuft ...")
        );

        notifySmartView();

        if (result.retryAfter) {
          setProgress(
            "Microsoft bittet um eine kurze Pause. " +
            (result.scanned || 0) +
            " E-Mails bisher geprüft."
          );

          await sleep(result.retryAfter * 1000);
        } else if (!completed) {
          await sleep(250);
        }

        if (!document.contains(panel)) {
          scanStopRequested = true;
        }
      }

      if (scanStopRequested) {
        setProgress(
          "Scan pausiert. Beim nächsten Start geht es weiter."
        );
      }

      notifySmartView();
    } catch (error) {
      setProgress(
        "Scan unterbrochen: " + error.message
      );
    } finally {
      scanning = false;
      stopButton.remove();

      if (scanButton) scanButton.disabled = false;
      if (refreshButton) refreshButton.disabled = false;

      notifySmartView();
    }
  }

  async function updateOutlookStatus() {
    const ui = area();

    if (!ui || checking || connecting || scanning) return;

    checking = true;
    ui.container.textContent =
      "Outlook-Verbindung wird geprüft ...";

    try {
      const result = await callFunction("outlook-status");

      if (!document.contains(ui.button)) return;

      ui.container.replaceChildren();

      if (!result.connected) {
        ui.button.disabled = false;
        ui.button.textContent = "✉️ Outlook verbinden";

        ui.container.append(
          element("p", "Outlook ist noch nicht verbunden.")
        );

        return;
      }

      ui.button.disabled = true;
      ui.button.textContent = "✓ Outlook verbunden";

      hideOldNotice(ui.button);

      const status = element(
        "p",
        "✓ Microsoft Outlook ist verbunden"
      );

      status.style.color = "#21734b";
      status.style.fontWeight = "600";

      ui.container.append(status);

      if (result.email) {
        ui.container.append(
          element("p", "Konto: " + result.email)
        );
      }

      const panel = element("div");
      panel.id = "outlookImportPanel";
      panel.style.marginTop = "18px";
      panel.style.padding = "14px";
      panel.style.border = "1px solid #ddd";
      panel.style.borderRadius = "12px";

      panel.append(
        element("h3", "Reisebuchungen aus Outlook")
      );

      panel.append(
        element(
          "p",
          "MyJourney durchsucht deinen Posteingang " +
          "einschließlich aller Unterordner nach " +
          "Reise- und Buchungsbestätigungen."
        )
      );

      const controls = element("div");
      controls.id = "outlookScanControls";

      const scanButton = createButton(
        "🔎 Posteingang nach Buchungen durchsuchen",
        startScan
      );

      scanButton.id = "outlookStartScan";
      controls.append(scanButton);

      const refreshButton = createButton(
        "↻ Ergebnisse aktualisieren",
        refreshScanProgress
      );

      refreshButton.id = "outlookRefreshScan";
      controls.append(refreshButton);

      panel.append(controls);

      const progress = element("p");
      progress.id = "outlookScanProgress";
      progress.style.marginTop = "12px";
      panel.append(progress);

      // WICHTIG:
      // Hier wird absichtlich KEINE alte Buchungsliste
      // mit "Vormerken" oder "Ausblenden" erstellt.
      //
      // outlook-smart.js fügt die ausführlichen
      // Buchungskarten selbst in dieses Panel ein.

      ui.container.append(panel);

      await refreshScanProgress();
    } catch (error) {
      if (document.contains(ui.container)) {
        ui.container.textContent =
          "Verbindungsstatus konnte nicht geladen werden: " +
          error.message;
      }
    } finally {
      checking = false;
    }
  }

  document.addEventListener(
    "click",
    (event) => {
      const button = event.target.closest?.(
        "#connectOutlook"
      );

      if (!button) return;

      event.preventDefault();
      event.stopPropagation();

      if (!button.disabled) {
        void connectOutlook(button);
      }
    },
    true
  );

  function detectOutlookView() {
    const button = document.querySelector(
      "#connectOutlook"
    );

    if (button && button !== lastButton) {
      lastButton = button;
      void updateOutlookStatus();
    }

    if (!button) {
      lastButton = null;
    }
  }

  const observer = new MutationObserver(
    detectOutlookView
  );

  observer.observe(document.body, {
    childList: true,
    subtree: true,
  });

  detectOutlookView();

  const params = new URLSearchParams(
    window.location.search
  );

  const outlookStatus = params.get("outlook");

  const messages = {
    connected:
      "🎉 Dein Outlook-Konto wurde erfolgreich verbunden!",
    cancelled:
      "Die Outlook-Anmeldung wurde abgebrochen.",
    expired:
      "Die Outlook-Anmeldung ist abgelaufen.",
    error:
      "Die Outlook-Verbindung konnte nicht abgeschlossen werden.",
  };

  if (outlookStatus && messages[outlookStatus]) {
    alert(messages[outlookStatus]);

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
