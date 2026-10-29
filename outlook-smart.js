
/* MyJourney – Outlook Smart Booking UI */
(() => {
  const CONFIG = window.MYJOURNEY_CONFIG || {};
  const URL =
    "https://appdqolpwvjktltvtqvj.supabase.co/functions/v1/outlook-import";

  if (
    !window.supabase?.createClient ||
    !CONFIG.url ||
    !CONFIG.publishableKey
  ) return;

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

  let busy = false;
  let lastPanel = null;
  let candidates = [];
  let searchTerm = "";
  let showCancelled = true;

  const el = (tag, text) => {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    return node;
  };

  const style = (node, values) => {
    Object.assign(node.style, values);
    return node;
  };

  function button(text, action, primary = false) {
    const node = el("button", text);
    node.type = "button";
    node.className = primary ? "" : "secondary";

    style(node, {
      borderRadius: "12px",
      padding: "10px 14px",
      margin: "4px 5px 4px 0",
      cursor: "pointer",
      fontSize: "13px",
      fontWeight: "600"
    });

    node.addEventListener("click", action);
    return node;
  }

  async function api(action, extra = {}) {
    const { data, error } = await client.auth.getSession();

    if (error || !data.session?.access_token) {
      throw Error("Bitte bei MyJourney anmelden.");
    }

    const response = await fetch(URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: CONFIG.publishableKey,
        Authorization:
          "Bearer " + data.session.access_token
      },
      body: JSON.stringify({ action, ...extra })
    });

    const result = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw Error(result.error || "Serverfehler");
    }

    return result;
  }

  function formatDate(value) {
    if (!value) return "Nicht erkannt";

    const match = String(value).match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );

    if (match) {
      return `${match[3]}.${match[2]}.${match[1]}`;
    }

    const date = new Date(value);

    return Number.isNaN(date.getTime())
      ? String(value)
      : date.toLocaleDateString("de-DE");
  }

  function statusOf(candidate) {
    return candidate.details?.booking_status || "unclear";
  }

  function iconOf(type) {
    return {
      flight: "✈️",
      hotel: "🏨",
      transport: "🚆",
      transfer: "🚕",
      activity: "🎟️",
      other: "✉️"
    }[type] || "✉️";
  }

  function statusLabel(status) {
    if (status === "cancelled") return "Stornierung erkannt";
    if (status === "changed") return "Änderung erkannt";
    return "Bitte prüfen";
  }

  function getTrips() {
    return Array.isArray(state.trips) ? state.trips : [];
  }

  function getItems() {
    return Array.isArray(state.items) ? state.items : [];
  }

  function tripMatch(details) {
    const destination = String(
      details.destination || ""
    ).trim().toLowerCase();

    const date = String(details.start || "").slice(0, 10);

    if (!destination || !date) return null;

    const matches = getTrips().filter(trip => {
      const text = [
        trip.destination,
        trip.name
      ].filter(Boolean).join(" ").toLowerCase();

      const locationMatches =
        text.includes(destination) ||
        destination.includes(
          String(trip.destination || "")
            .trim()
            .toLowerCase()
        ) && !!String(trip.destination || "").trim();

      const start = String(trip.start || "").slice(0, 10);
      const end = String(trip.end || "").slice(0, 10);

      return (
        locationMatches &&
        start &&
        end &&
        date >= start &&
        date <= end
      );
    });

    return matches.length === 1 ? matches[0] : null;
  }

  function makeItem(candidate, tripId) {
    const details = candidate.details || {};

    if (statusOf(candidate) === "cancelled") {
      throw Error(
        "Diese Nachricht wurde als Stornierung erkannt."
      );
    }

    if (!details.start) {
      throw Error(
        "Ein Reisedatum fehlt. Bitte diese Buchung manuell prüfen."
      );
    }

    const allowed = [
      "flight",
      "hotel",
      "transfer",
      "activity",
      "other"
    ];

    const originalType =
      candidate.booking_type === "transport"
        ? "transfer"
        : candidate.booking_type;

    const type = allowed.includes(originalType)
      ? originalType
      : "other";

    return {
      id: crypto.randomUUID(),
      tripId,
      type,
      title:
        candidate.subject ||
        details.title ||
        "Outlook-Buchung",
      start: details.start,
      end: details.end || "",
      location: details.destination || "",
      address: "",
      lat: null,
      lng: null,
      confirmation: details.confirmation || "",
      cost:
        typeof details.cost_eur === "number"
          ? details.cost_eur
          : 0,
      notes: [
        "Importiert aus Outlook",
        candidate.sender
          ? "Absender: " + candidate.sender
          : "",
        details.amount != null
          ? "Originalpreis: " +
            details.amount +
            " " +
            (details.currency || "")
          : "Preis nicht eindeutig erkannt",
        candidate.preview || ""
      ].filter(Boolean).join("\n"),
      source: "outlook",
      outlookMessageId: candidate.microsoft_message_id
    };
  }

  async function assign(candidate, trip) {
    if (
      getItems().some(
        item =>
          item.outlookMessageId ===
          candidate.microsoft_message_id
      )
    ) {
      alert("Diese Buchung wurde bereits übernommen.");
      return;
    }

    let item;

    try {
      item = makeItem(candidate, trip.id);
    } catch (error) {
      alert(error.message);
      return;
    }

    state.items.push(item);
    save();
    render();

    try {
      await api("review", {
        id: candidate.id,
        status: "accepted"
      });
    } catch (error) {
      alert(
        "Die Buchung wurde lokal gespeichert, aber der " +
        "Outlook-Status konnte nicht aktualisiert werden: " +
        error.message
      );
    }

    await show();

    alert(
      "Buchung zu „" + trip.name + "“ hinzugefügt."
    );
  }

  async function createTrip(candidate) {
    const details = candidate.details || {};

    if (statusOf(candidate) === "cancelled") {
      alert("Stornierte Buchungen können nicht übernommen werden.");
      return;
    }

    if (!details.start) {
      alert(
        "Es fehlt ein Reisedatum. Bitte die Buchung prüfen."
      );
      return;
    }

    const name = prompt(
      "Wie soll die neue Reise heißen?",
      details.destination || "Neue Reise"
    );

    if (!name?.trim()) return;

    const start = String(details.start).slice(0, 10);
    const rawEnd = String(
      details.end || details.start
    ).slice(0, 10);

    const trip = {
      id: crypto.randomUUID(),
      name: name.trim(),
      destination:
        details.destination || name.trim(),
      start,
      end: rawEnd >= start ? rawEnd : start
    };

    let item;

    try {
      item = makeItem(candidate, trip.id);
    } catch (error) {
      alert(error.message);
      return;
    }

    state.trips.push(trip);
    state.items.push(item);
    selected = trip.id;

    save();
    render();

    try {
      await api("review", {
        id: candidate.id,
        status: "accepted"
      });
    } catch (error) {
      alert(
        "Die Reise wurde erstellt, aber der Outlook-Status " +
        "konnte nicht aktualisiert werden: " +
        error.message
      );
    }

    await show();
  }

  function detailField(label, value) {
    const box = style(el("div"), {
      padding: "12px",
      background: "rgba(145,130,110,.07)",
      borderRadius: "12px",
      minWidth: "0"
    });

    box.append(
      style(el("div", label), {
        fontSize: "11px",
        opacity: ".7",
        marginBottom: "5px"
      })
    );

    box.append(
      style(el("div", value || "Nicht erkannt"), {
        fontSize: "13px",
        fontWeight: "600",
        overflowWrap: "anywhere"
      })
    );

    return box;
  }

  function makeCard(candidate) {
    const details = candidate.details || {};
    const status = statusOf(candidate);
    const cancelled = status === "cancelled";

    const card = el("details");

    style(card, {
      border: "1px solid rgba(120,110,95,.2)",
      borderRadius: "17px",
      marginBottom: "10px",
      overflow: "hidden",
      background: "var(--card, #fffdf9)",
      boxShadow: "0 3px 12px rgba(0,0,0,.035)"
    });

    const summary = el("summary");

    style(summary, {
      display: "flex",
      alignItems: "center",
      gap: "12px",
      padding: "16px",
      cursor: "pointer",
      listStyle: "none"
    });

    const icon = style(
      el("div", iconOf(candidate.booking_type)),
      {
        width: "40px",
        height: "40px",
        flexShrink: "0",
        display: "grid",
        placeItems: "center",
        borderRadius: "13px",
        background: cancelled
          ? "#fbeae6"
          : "#f1f0e9",
        fontSize: "20px"
      }
    );

    const heading = style(el("div"), {
      flex: "1",
      minWidth: "0"
    });

    heading.append(
      style(
        el(
          "div",
          candidate.subject ||
            details.title ||
            "Buchung ohne Betreff"
        ),
        {
          fontWeight: "650",
          fontSize: "14px",
          lineHeight: "1.45",
          overflowWrap: "anywhere"
        }
      )
    );

    const statusText = style(
      el("div", statusLabel(status)),
      {
        fontSize: "11px",
        marginTop: "5px",
        color: cancelled
          ? "#a34c3c"
          : status === "changed"
            ? "#a07132"
            : "#738078"
      }
    );

    heading.append(statusText);

    const chevron = style(el("span", "⌄"), {
      fontSize: "24px",
      opacity: ".55",
      transition: "transform .2s"
    });

    summary.append(icon, heading, chevron);

    card.addEventListener("toggle", () => {
      chevron.style.transform = card.open
        ? "rotate(180deg)"
        : "rotate(0deg)";
    });

    card.append(summary);

    const content = style(el("div"), {
      padding: "0 16px 18px"
    });

    const separator = style(el("div"), {
      borderTop: "1px solid rgba(120,110,95,.15)",
      marginBottom: "15px"
    });

    content.append(separator);

    const grid = style(el("div"), {
      display: "grid",
      gridTemplateColumns:
        "repeat(auto-fit,minmax(145px,1fr))",
      gap: "8px"
    });

    grid.append(
      detailField("📍 Reiseziel", details.destination),
      detailField("📅 Beginn", formatDate(details.start)),
      detailField("📅 Ende", formatDate(details.end)),
      detailField(
        "💳 Preis",
        details.amount != null
          ? details.amount +
            " " +
            (details.currency || "")
          : "Nicht erkannt"
      ),
      detailField(
        "🔖 Buchungsnummer",
        details.confirmation
      ),
      detailField("✉️ Absender", candidate.sender)
    );

    content.append(grid);

    if (!candidate.analyzed_at) {
      content.append(
        style(
          el(
            "p",
            "Diese E-Mail wurde noch nicht im Detail analysiert."
          ),
          {
            fontSize: "12px",
            opacity: ".7"
          }
        )
      );
    }

    const mail = el("details");

    style(mail, {
      marginTop: "15px",
      padding: "12px",
      borderRadius: "12px",
      border: "1px solid rgba(120,110,95,.18)"
    });

    const mailSummary = el(
      "summary",
      "📩 Vollständige Original-E-Mail anzeigen"
    );

    style(mailSummary, {
      cursor: "pointer",
      fontWeight: "600",
      fontSize: "13px"
    });

    const mailBody = style(
      el("pre", "E-Mail zum Laden öffnen."),
      {
        whiteSpace: "pre-wrap",
        overflowWrap: "anywhere",
        maxHeight: "450px",
        overflowY: "auto",
        fontFamily: "inherit",
        fontSize: "12px",
        lineHeight: "1.7",
        marginTop: "14px"
      }
    );

    mail.append(mailSummary, mailBody);

    mail.addEventListener("toggle", async () => {
      if (!mail.open || mail.dataset.loaded) return;

      mailBody.textContent =
        "Original-E-Mail wird aus Outlook geladen …";

      try {
        const message = await api("message_detail", {
          id: candidate.id
        });

        mailBody.textContent = [
          "Von: " +
            (message.sender || candidate.sender || ""),
          "Datum: " +
            formatDate(
              message.received_at || candidate.received_at
            ),
          "Betreff: " +
            (message.subject || candidate.subject || ""),
          "",
          message.body || "Kein Nachrichtentext verfügbar."
        ].join("\n");

        mail.dataset.loaded = "1";
      } catch (error) {
        mailBody.textContent =
          "E-Mail konnte nicht geladen werden: " +
          error.message;
      }
    });

    content.append(mail);

    const actions = style(el("div"), {
      marginTop: "15px"
    });

    if (cancelled) {
      actions.append(
        style(
          el(
            "p",
            "⚠️ Als Stornierung erkannt. Bitte anhand " +
            "der Original-E-Mail überprüfen."
          ),
          {
            color: "#a34c3c",
            fontSize: "12px"
          }
        )
      );
    } else {
      const match = tripMatch(details);

      if (match) {
        actions.append(
          style(
            el(
              "p",
              "💡 Mögliche passende Reise: " + match.name
            ),
            {
              fontSize: "12px",
              color: "#507866"
            }
          )
        );
      }

      const trips = getTrips();

      if (trips.length) {
        const select = el("select");

        style(select, {
          display: "block",
          width: "100%",
          maxWidth: "420px",
          padding: "11px",
          borderRadius: "10px",
          margin: "12px 0",
          border: "1px solid #cfcac0",
          background: "#fff"
        });

        const placeholder = el(
          "option",
          "Bitte Reise auswählen …"
        );

        placeholder.value = "";
        placeholder.selected = true;
        placeholder.disabled = true;
        select.append(placeholder);

        for (const trip of trips) {
          const option = el("option", trip.name);
          option.value = trip.id;
          select.append(option);
        }

        actions.append(select);

        actions.append(
          button(
            "✓ Zu dieser Reise hinzufügen",
            async () => {
              const trip = trips.find(
                item => item.id === select.value
              );

              if (!trip) {
                alert("Bitte zuerst eine Reise auswählen.");
                return;
              }

              try {
                await assign(candidate, trip);
              } catch (error) {
                alert(error.message);
              }
            },
            true
          )
        );
      }

      actions.append(
        button("＋ Neue Reise erstellen", async () => {
          try {
            await createTrip(candidate);
          } catch (error) {
            alert(error.message);
          }
        })
      );
    }

    actions.append(
      button("Ausblenden", async () => {
        try {
          await api("review", {
            id: candidate.id,
            status: "ignored"
          });

          await show();
        } catch (error) {
          alert(error.message);
        }
      })
    );

    content.append(actions);
    card.append(content);

    return card;
  }

  function renderCandidates() {
    const target = document.querySelector(
      "#outlookSmartResults"
    );

    if (!target) return;

    target.replaceChildren();

    const filtered = candidates.filter(candidate => {
      const text = [
        candidate.subject,
        candidate.sender,
        candidate.details?.destination
      ].filter(Boolean).join(" ").toLowerCase();

      const matchesSearch = text.includes(
        searchTerm.toLowerCase()
      );

      const matchesStatus =
        showCancelled ||
        statusOf(candidate) !== "cancelled";

      return matchesSearch && matchesStatus;
    });

    if (!filtered.length) {
      const empty = style(
        el(
          "div",
          candidates.length
            ? "Keine passenden Buchungen gefunden."
            : "Aktuell sind keine offenen Buchungen vorhanden."
        ),
        {
          padding: "24px",
          textAlign: "center",
          opacity: ".7"
        }
      );

      target.append(empty);
      return;
    }

    for (const candidate of filtered) {
      target.append(makeCard(candidate));
    }
  }

  async function show() {
    const target = document.querySelector(
      "#outlookSmartResults"
    );

    if (!target) return;

    try {
      const result = await api("list");

      candidates = (result.items || []).filter(
        candidate => candidate.status === "new"
      );

      const count = document.querySelector(
        "#outlookSmartCount"
      );

      if (count) {
        count.textContent =
          candidates.length + " offene Vorschläge";
      }

      renderCandidates();
    } catch (error) {
      target.replaceChildren(
        el("p", "Fehler: " + error.message)
      );
    }
  }

  async function analyze() {
    if (busy) return;

    busy = true;

    const action = document.querySelector(
      "#outlookAnalyzeButton"
    );

    if (action) action.disabled = true;

    try {
      let total = 0;
      let remaining = 0;

      for (let index = 0; index < 10; index++) {
        const result = await api("analyze");

        total += result.analyzed || 0;
        remaining = result.remaining || 0;

        const progress = document.querySelector(
          "#outlookSmartProgress"
        );

        if (progress) {
          progress.textContent =
            total +
            " E-Mails ausgewertet. " +
            (remaining === 0
              ? "Analyse abgeschlossen."
              : remaining + " weitere warten.");
        }

        if (!result.analyzed || remaining === 0) {
          break;
        }
      }

      await show();
    } catch (error) {
      alert("Analyse fehlgeschlagen: " + error.message);
    } finally {
      busy = false;

      if (action) action.disabled = false;
    }
  }

  function mount() {
    const panel = document.querySelector(
      "#outlookImportPanel"
    );

    if (!panel || panel === lastPanel) return;

    lastPanel = panel;

    const wrap = el("section");
    wrap.id = "outlookSmart";

    style(wrap, {
      marginTop: "22px",
      paddingTop: "16px",
      borderTop: "1px solid rgba(120,110,95,.18)"
    });

    const heading = el("h3", "Deine Buchungsvorschläge");

    style(heading, {
      marginBottom: "6px",
      fontSize: "19px"
    });

    wrap.append(heading);

    const count = style(
      el("p", "Vorschläge werden geladen …"),
      {
        fontSize: "12px",
        opacity: ".7",
        marginTop: "0"
      }
    );

    count.id = "outlookSmartCount";
    wrap.append(count);

    const toolbar = style(el("div"), {
      display: "flex",
      flexWrap: "wrap",
      gap: "6px",
      margin: "12px 0"
    });

    const analyzeButton = button(
      "✨ Buchungsdetails auslesen",
      () => void analyze(),
      true
    );

    analyzeButton.id = "outlookAnalyzeButton";

    toolbar.append(analyzeButton);

    toolbar.append(
      button("↻ Aktualisieren", () => void show())
    );

    wrap.append(toolbar);

    const progress = style(el("p"), {
      fontSize: "12px",
      opacity: ".75"
    });

    progress.id = "outlookSmartProgress";
    wrap.append(progress);

    const search = el("input");
    search.type = "search";
    search.placeholder = "🔎 Buchungen durchsuchen …";

    style(search, {
      display: "block",
      width: "100%",
      boxSizing: "border-box",
      padding: "13px",
      borderRadius: "12px",
      border: "1px solid rgba(120,110,95,.25)",
      margin: "12px 0",
      fontSize: "14px"
    });

    search.addEventListener("input", () => {
      searchTerm = search.value;
      renderCandidates();
    });

    wrap.append(search);

    const toggle = el("label");

    style(toggle, {
      display: "flex",
      alignItems: "center",
      gap: "8px",
      fontSize: "12px",
      marginBottom: "15px",
      cursor: "pointer"
    });

    const checkbox = el("input");
    checkbox.type = "checkbox";
    checkbox.checked = showCancelled;

    checkbox.addEventListener("change", () => {
      showCancelled = checkbox.checked;
      renderCandidates();
    });

    toggle.append(
      checkbox,
      el("span", "Mögliche Stornierungen anzeigen")
    );

    wrap.append(toggle);

    const results = el("div");
    results.id = "outlookSmartResults";
    wrap.append(results);

    const cleanup = button(
      "🧹 Offene Vorschläge zurücksetzen",
      async () => {
        if (
          !confirm(
            "Alle noch nicht übernommenen Outlook-" +
            "Vorschläge löschen und den Scan zurücksetzen?"
          )
        ) return;

        try {
          await api("clean_and_rescan");
          await show();

          alert(
            "Vorschläge zurückgesetzt. " +
            "Du kannst den Posteingang erneut scannen."
          );
        } catch (error) {
          alert(error.message);
        }
      }
    );

    style(cleanup, {
      marginTop: "18px",
      opacity: ".7",
      fontSize: "12px"
    });

    wrap.append(cleanup);
    panel.append(wrap);

    void show();
  }

  document.addEventListener(
    "myjourney:outlook-updated",
    () => void show()
  );

  new MutationObserver(mount).observe(
    document.body,
    {
      subtree: true,
      childList: true
    }
  );

  mount();
})();
