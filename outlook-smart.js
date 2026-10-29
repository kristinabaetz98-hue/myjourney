
/* MyJourney – Outlook-Reisezuordnung */
(() => {
  const CONFIG = window.MYJOURNEY_CONFIG || {};
  const URL = CONFIG.url + "/functions/v1/outlook-import";

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

  const el = (tag, text) => {
    const x = document.createElement(tag);
    if (text !== undefined) x.textContent = text;
    return x;
  };

  const btn = (text, fn, primary = false) => {
    const x = el("button", text);
    x.type = "button";
    x.className = primary ? "primary" : "secondary";
    x.addEventListener("click", fn);
    return x;
  };

  async function api(action, extra = {}) {
    const { data, error } = await client.auth.getSession();

    if (error || !data.session?.access_token) {
      throw Error("Bitte bei MyJourney anmelden.");
    }

    const r = await fetch(URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: CONFIG.publishableKey,
        Authorization: "Bearer " + data.session.access_token
      },
      body: JSON.stringify({ action, ...extra })
    });

    const j = await r.json().catch(() => ({}));

    if (!r.ok) throw Error(j.error || "Serverfehler");
    return j;
  }

  function tripMatch(d) {
    const destination = String(
      d.destination || ""
    ).toLocaleLowerCase("de").trim();

    const date = d.start?.slice(0, 10) || "";

    const scored = state.trips.map(t => {
      const words = [t.destination, t.name]
        .join(" ")
        .toLocaleLowerCase("de");

      const place = destination &&
        words.includes(destination);

      const overlap = date &&
        t.start &&
        t.end &&
        date >= t.start.slice(0, 10) &&
        date <= t.end.slice(0, 10);

      return {
        trip: t,
        score: (place ? 3 : 0) + (overlap ? 2 : 0)
      };
    });

    scored.sort((a, b) => b.score - a.score);

    return scored[0]?.score >= 2
      ? scored[0].trip
      : null;
  }

  function makeItem(c, tripId) {
    const d = c.details || {};

    if (!d.start) {
      throw Error("Es fehlt ein eindeutiges Reisedatum.");
    }

    const type = c.booking_type === "transport"
      ? "transfer"
      : c.booking_type;

    const allowed = [
      "flight",
      "hotel",
      "transfer",
      "activity",
      "other"
    ];

    return {
      id: crypto.randomUUID(),
      tripId,
      type: allowed.includes(type) ? type : "other",
      title: d.title || c.subject || "Outlook-Buchung",
      start: d.start,
      end: d.end || "",
      location: d.destination || "",
      address: "",
      lat: null,
      lng: null,
      confirmation: d.confirmation || "",
      cost: typeof d.cost_eur === "number"
        ? d.cost_eur
        : 0,
      notes: [
        "Outlook: " + (c.sender || ""),
        d.amount != null && d.currency
          ? `Originalpreis: ${d.amount} ${d.currency}`
          : "",
        d.cost_eur == null
          ? "Preis bitte manuell überprüfen."
          : ""
      ].filter(Boolean).join("\n"),
      source: "outlook",
      outlookMessageId: c.microsoft_message_id
    };
  }

  async function assign(c, trip) {
    const duplicate = state.items.some(
      i => i.outlookMessageId === c.microsoft_message_id
    );

    if (duplicate) {
      alert("Diese Buchung wurde bereits übernommen.");
      return;
    }

    let item;

    try {
      item = makeItem(c, trip.id);
    } catch (e) {
      alert(e.message);
      return;
    }

    state.items.push(item);
    save();
    render();

    try {
      await api("review", {
        id: c.id,
        status: "accepted"
      });
    } catch (e) {
      console.warn("Outlook-Status:", e);
    }

    alert(`Buchung zu „${trip.name}“ hinzugefügt.`);
    await show();
  }

  async function createTrip(c) {
    const d = c.details || {};

    if (!d.start) {
      alert("Kein eindeutiges Reisedatum erkannt.");
      return;
    }

    const name = prompt(
      "Wie soll deine neue Reise heißen?",
      d.destination || "Neue Reise"
    );

    if (!name?.trim()) return;

    const start = d.start.slice(0, 10);
    const proposedEnd = (d.end || d.start).slice(0, 10);
    const end = proposedEnd >= start ? proposedEnd : start;

    const trip = {
      id: crypto.randomUUID(),
      name: name.trim(),
      destination: d.destination || name.trim(),
      start,
      end
    };

    let item;

    try {
      item = makeItem(c, trip.id);
    } catch (e) {
      alert(e.message);
      return;
    }

    state.trips.push(trip);
    state.items.push(item);
    selected = trip.id;

    save();
    render();

    try {
      await api("review", {
        id: c.id,
        status: "accepted"
      });
    } catch (e) {
      console.warn("Outlook-Status:", e);
    }

    alert(`Reise „${trip.name}“ wurde erstellt.`);
    await show();
  }

  function addTripSelector(c, card) {
    const select = el("select");
    select.style.width = "100%";
    select.style.marginBottom = "8px";

    for (const trip of state.trips) {
      const option = el("option", trip.name);
      option.value = trip.id;
      select.append(option);
    }

    card.append(select);

    card.append(
      btn("Dieser Reise zuordnen", () => {
        const trip = state.trips.find(
          t => t.id === select.value
        );

        if (trip) {
          void assign(c, trip).catch(
            e => alert(e.message)
          );
        }
      })
    );
  }

  async function show() {
    const target = document.querySelector(
      "#outlookSmartResults"
    );

    if (!target) return;

    target.replaceChildren(
      el("p", "Reisevorschläge werden geladen …")
    );

    try {
      const data = await api("list");

      const candidates = (data.items || [])
        .filter(c => c.status === "new");

      target.replaceChildren();

      if (!candidates.length) {
        target.append(
          el("p", "Aktuell keine offenen Reisevorschläge.")
        );
        return;
      }

      target.append(
        el("h4", `${candidates.length} Buchungen prüfen`)
      );

      for (const c of candidates) {
        const d = c.details || {};
        const card = el("article");

        card.className = "card";
        card.style.margin = "12px 0";
        card.style.padding = "18px";
        card.style.borderRadius = "16px";

        const icon = c.booking_type === "flight"
          ? "✈️"
          : c.booking_type === "hotel"
            ? "🏨"
            : c.booking_type === "transport"
              ? "🚆"
              : "🎟️";

        const heading = el(
          "h3",
          `${icon} ${d.title || c.subject}`
        );

        heading.style.marginTop = "0";
        card.append(heading);

        const details = [
          d.destination,
          d.start,
          d.end && d.end !== d.start
            ? "bis " + d.end
            : ""
        ].filter(Boolean).join(" · ");

        card.append(
          el("p", details || "Reisedaten noch unklar")
        );

        let price = "Preis nicht erkannt";

        if (typeof d.cost_eur === "number") {
          price = d.cost_eur.toLocaleString(
            "de-DE",
            {
              style: "currency",
              currency: "EUR"
            }
          );
        } else if (d.amount != null) {
          price = `${d.amount} ${d.currency || ""}`;
        }

        const priceLine = el("p", price);
        priceLine.style.fontWeight = "600";
        card.append(priceLine);

        if (!c.analyzed_at) {
          const note = el(
            "p",
            "Buchungsdetails noch nicht analysiert."
          );
          note.className = "muted";
          card.append(note);
        }

        const match = tripMatch(d);

        if (match) {
          card.append(
            el(
              "p",
              `💡 Passt vermutlich zu „${match.name}“`
            )
          );

          card.append(
            btn(
              "✓ Zu dieser Reise hinzufügen",
              () => void assign(c, match)
                .catch(e => alert(e.message)),
              true
            )
          );
        }

        if (state.trips.length) {
          addTripSelector(c, card);
        }

        card.append(
          btn(
            "＋ Neue Reise erstellen",
            () => void createTrip(c)
              .catch(e => alert(e.message))
          )
        );

        card.append(
          btn("Ignorieren", async () => {
            try {
              await api("review", {
                id: c.id,
                status: "ignored"
              });

              await show();
            } catch (e) {
              alert(e.message);
            }
          })
        );

        target.append(card);
      }

    } catch (e) {
      target.replaceChildren(
        el("p", "Fehler: " + e.message)
      );
    }
  }

  async function analyze() {
    if (busy) return;

    busy = true;

    const button = document.querySelector(
      "#outlookAnalyzeButton"
    );

    if (button) button.disabled = true;

    try {
      let count = 0;

      for (let i = 0; i < 10; i++) {
        const result = await api("analyze");

        count += result.analyzed || 0;

        const progress = document.querySelector(
          "#outlookSmartProgress"
        );

        if (progress) {
          progress.textContent =
            `${count} E-Mails analysiert. ` +
            (result.remaining === 0
              ? "Fertig."
              : "Weitere folgen …");
        }

        if (!result.analyzed || result.remaining === 0) {
          break;
        }
      }

      await show();

    } catch (e) {
      alert("Analyse fehlgeschlagen: " + e.message);

    } finally {
      busy = false;
      if (button) button.disabled = false;
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
    wrap.style.marginTop = "20px";

    wrap.append(
      el("h3", "✨ Deine Reisevorschläge")
    );

    wrap.append(
      el(
        "p",
        "Erkannte Buchungen prüfen und mit einem Klick deiner Reise hinzufügen."
      )
    );

    const analyzeButton = btn(
      "🔎 Buchungsdetails auslesen",
      () => void analyze(),
      true
    );

    analyzeButton.id = "outlookAnalyzeButton";
    wrap.append(analyzeButton);

    wrap.append(
      btn(
        "↻ Vorschläge aktualisieren",
        () => void show()
      )
    );

    const progress = el("p");
    progress.id = "outlookSmartProgress";
    wrap.append(progress);

    const results = el("div");
    results.id = "outlookSmartResults";
    wrap.append(results);

    // Neue Bereinigungsfunktion
    wrap.append(
      btn(
        "🧹 Alte Vorschläge bereinigen",
        async () => {
          if (!confirm(
            "Alle noch nicht übernommenen Outlook-Vorschläge entfernen und Scan zurücksetzen?"
          )) return;

          try {
            await api("clean_and_rescan");
            await show();

            alert(
              "Vorschläge bereinigt. Starte jetzt die Postfachsuche erneut."
            );
          } catch (e) {
            alert("Bereinigung fehlgeschlagen: " + e.message);
          }
        }
      )
    );

    panel.append(wrap);
    void show();
  }

  new MutationObserver(mount).observe(
    document.body,
    {
      subtree: true,
      childList: true
    }
  );

  mount();
})();
