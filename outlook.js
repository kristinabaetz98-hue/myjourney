
/* MyJourney – Outlook-Verbindung und Buchungsimport */

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
    },
  );

  let connecting = false;
  let checking = false;
  let scanning = false;
  let lastButton = null;
  let scanStopRequested = false;

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
        Authorization:
          "Bearer " + currentSession.access_token,
        apikey: CONFIG.publishableKey,
      },
      body: JSON.stringify(body),
    });

    const result = await response.json();

    if (!response.ok) {
      throw new Error(
        result.error || "Die Anfrage ist fehlgeschlagen.",
      );
    }

    return result;
  }

  function element(tag, text, className = "") {
    const node = document.createElement(tag);

    if (text !== undefined) {
      node.textContent = text;
    }

    if (className) {
      node.className = className;
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

    let container =
      document.querySelector("#outlookLiveStatus");

    if (!container) {
      container = element("div");
      container.id = "outlookLiveStatus";
      container.style.marginTop = "14px";

      button.insertAdjacentElement(
        "afterend",
        container,
      );
    }

    return { button, container };
  }

  function hideOldNotice(button) {
    const card = button.closest(".card");

    if (!card) return;

    card.querySelectorAll("p").forEach((paragraph) => {
      if (
        paragraph.textContent.includes(
          "noch nicht eingerichtet",
        ) ||
        paragraph.textContent.includes(
          "automatische E-Mail-Abruf",
        )
      ) {
        paragraph.hidden = true;
      }
    });
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
          element("p", "Outlook ist noch nicht verbunden."),
        );

        return;
      }

      ui.button.disabled = true;
      ui.button.textContent = "✓ Outlook verbunden";

      hideOldNotice(ui.button);

      const status = element(
        "p",
        "✓ Microsoft Outlook ist verbunden",
      );

      status.style.color = "#21734b";
      status.style.fontWeight = "600";

      ui.container.append(status);

      if (result.email) {
        ui.container.append(
          element("p", "Konto: " + result.email),
        );
      }

      const panel = element("div");
      panel.id = "outlookImportPanel";
      panel.style.marginTop = "18px";
      panel.style.padding = "14px";
      panel.style.border = "1px solid #ddd";
      panel.style.borderRadius = "12px";

      panel.append(
        element("h3", "Reisebuchungen aus Outlook"),
      );

      const description = element(
        "p",
        "MyJourney sucht im gesamten Postfach nach Reise- und Buchungsbestätigungen.",
      );

      panel.append(description);

      const controls = element("div");

      controls.append(
        createButton(
          "🔎 Gesamtes Postfach durchsuchen",
          startScan,
        ),
      );

      controls.append(
        createButton(
          "↻ Ergebnisse aktualisieren",
          loadCandidates,
        ),
      );

      panel.append(controls);

      const progress = element("p");
      progress.id = "outlookScanProgress";
      progress.style.marginTop = "12px";

      panel.append(progress);

      const results = element("div");
      results.id = "outlookImportResults";

      panel.append(results);

      ui.container.append(panel);

      await loadCandidates();

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

  async function connectOutlook(button) {
    if (connecting) return;

    connecting = true;
    button.disabled = true;
    button.textContent = "Outlook wird verbunden ...";

    try {
      const result = await callFunction(
        "outlook-connect/start",
      );

      if (!result.authorizationUrl) {
        throw new Error(
          "Microsoft-Anmeldung konnte nicht gestartet werden.",
        );
      }

      window.location.assign(result.authorizationUrl);
    } catch (error) {
      alert(
        "Outlook-Verbindung fehlgeschlagen:\n" +
        error.message,
      );

      button.disabled = false;
      button.textContent = "✉️ Outlook verbinden";
      connecting = false;
    }
  }

  function setProgress(message) {
    const progress =
      document.querySelector("#outlookScanProgress");

    if (progress) {
      progress.textContent = message;
    }
  }

  function sleep(ms) {
    return new Promise((resolve) => {
      setTimeout(resolve, ms);
    });
  }

  async function startScan() {
    if (scanning) return;

    scanning = true;
    scanStopRequested = false;

    const panel =
      document.querySelector("#outlookImportPanel");

    if (!panel) {
      scanning = false;
      return;
    }

    const buttons = panel.querySelectorAll("button");

    buttons.forEach((button) => {
      button.disabled = true;
    });

    const stopButton = createButton(
      "Scan pausieren",
      () => {
        scanStopRequested = true;
        stopButton.disabled = true;
        stopButton.textContent = "Wird pausiert ...";
      },
    );

    panel.append(stopButton);

    try {
      let completed = false;

      while (!completed && !scanStopRequested) {
        const result = await callFunction(
          "outlook-import",
          { action: "scan" },
        );

        completed = !!result.completed;

        setProgress(
          result.scanned +
          " E-Mails geprüft. " +
          (completed
            ? "Postfach vollständig durchsucht."
            : "Suche läuft ..."),
        );

        if (result.retryAfter) {
          setProgress(
            "Microsoft bittet um eine kurze Pause. " +
            result.scanned +
            " E-Mails bisher geprüft.",
          );

          await sleep(result.retryAfter * 1000);
        } else if (!completed) {
          await sleep(250);
        }

        if (
          !document.querySelector("#outlookImportPanel")
        ) {
          scanStopRequested = true;
        }
      }

      if (scanStopRequested) {
        setProgress(
          "Scan pausiert. Beim nächsten Start geht es weiter.",
        );
      }

      await loadCandidates();

    } catch (error) {
      setProgress("Scan unterbrochen: " + error.message);
      await loadCandidates().catch(() => {});
    } finally {
      scanning = false;
      stopButton.remove();

      const currentPanel =
        document.querySelector("#outlookImportPanel");

      if (currentPanel) {
        currentPanel.querySelectorAll("button").forEach(
          (button) => {
            button.disabled = false;
          },
        );
      }
    }
  }

  function typeName(type) {
    const names = {
      flight: "✈️ Flug",
      hotel: "🏨 Unterkunft",
      transport: "🚆 Transport",
      activity: "🎟️ Aktivität",
      other: "📩 Reisebuchung",
    };

    return names[type] || names.other;
  }

  function dateText(value) {
    if (!value) return "";

    const date = new Date(value);

    return Number.isNaN(date.getTime())
      ? ""
      : date.toLocaleDateString("de-DE");
  }

  async function reviewCandidate(id, status) {
    try {
      await callFunction("outlook-import", {
        action: "review",
        id,
        status,
      });

      await loadCandidates();
    } catch (error) {
      alert(error.message);
    }
  }

  async function loadCandidates() {
    const results =
      document.querySelector("#outlookImportResults");

    if (!results) return;

    results.textContent = "Buchungen werden geladen ...";

    try {
      const data = await callFunction(
        "outlook-import",
        { action: "list" },
      );

      if (!document.contains(results)) return;

      results.replaceChildren();

      const progress = data.progress || {};

      if (!scanning) {
        setProgress(
          (progress.scanned_count || 0) +
          " E-Mails bisher geprüft. " +
          (progress.completed
            ? "Gesamtscan abgeschlossen."
            : "Der Scan kann gestartet oder fortgesetzt werden."),
        );
      }

      const total = element(
        "h4",
        (data.total || 0) +
        " mögliche Buchungen gefunden",
      );

      results.append(total);

      if (!data.items?.length) {
        results.append(
          element(
            "p",
            "Noch keine Buchungen gefunden.",
          ),
        );

        return;
      }

      const note = element(
        "p",
        "Hier erscheinen die neuesten 200 Treffer. " +
        "„Vormerken“ markiert eine Buchung zur späteren " +
        "Übernahme in deine Reiseplanung.",
      );

      note.style.fontSize = "13px";
      results.append(note);

      for (const item of data.items) {
        const card = element("div");

        card.style.padding = "12px 0";
        card.style.borderBottom = "1px solid #ddd";

        const title = element(
          "strong",
          typeName(item.booking_type) +
          " – " +
          (item.subject || "Ohne Betreff"),
        );

        card.append(title);

        const meta = element(
          "p",
          [
            item.sender,
            dateText(item.received_at),
          ].filter(Boolean).join(" · "),
        );

        meta.style.fontSize = "12px";
        meta.style.opacity = "0.7";

        card.append(meta);

        if (item.preview) {
          const preview = element(
            "p",
            item.preview,
          );

          preview.style.fontSize = "13px";
          card.append(preview);
        }

        if (item.status === "accepted") {
          const selected = element(
            "p",
            "✓ Zur Übernahme vorgemerkt",
          );

          selected.style.color = "#21734b";
          card.append(selected);
        }

        if (item.status === "ignored") {
          card.append(
            element("p", "Ausgeblendet"),
          );
        }

        const actions = element("div");

        if (item.status !== "accepted") {
          actions.append(
            createButton(
              "✓ Vormerken",
              () => reviewCandidate(
                item.id,
                "accepted",
              ),
            ),
          );
        }

        if (item.status !== "ignored") {
          actions.append(
            createButton(
              "Ausblenden",
              () => reviewCandidate(
                item.id,
                "ignored",
              ),
            ),
          );
        }

        if (item.status !== "new") {
          actions.append(
            createButton(
              "Zurücksetzen",
              () => reviewCandidate(
                item.id,
                "new",
              ),
            ),
          );
        }

        card.append(actions);
        results.append(card);
      }

    } catch (error) {
      results.textContent =
        "Buchungen konnten nicht geladen werden: " +
        error.message;
    }
  }

  document.addEventListener(
    "click",
    (event) => {
      const button =
        event.target.closest?.("#connectOutlook");

      if (!button) return;

      event.preventDefault();
      event.stopPropagation();

      if (!button.disabled) {
        void connectOutlook(button);
      }
    },
    true,
  );

  function detectOutlookView() {
    const button =
      document.querySelector("#connectOutlook");

    if (button && button !== lastButton) {
      lastButton = button;
      void updateOutlookStatus();
    }

    if (!button) {
      lastButton = null;
    }
  }

  const observer = new MutationObserver(
    detectOutlookView,
  );

  observer.observe(document.body, {
    childList: true,
    subtree: true,
  });

  detectOutlookView();

  const params = new URLSearchParams(
    window.location.search,
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
      window.location.hash,
    );
  }
})();
