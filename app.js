const ANTHROPIC_API_KEY = "sk-ant-INSERISCI-QUI-LA-TUA-API-KEY";
const MODEL = "claude-sonnet-4-20250514";

const SYSTEM_PROMPT = `Sei l'assistente tecnico post-vendita di Schenker, azienda leader nella produzione di dissalatori a recupero di energia per barche da diporto.

Quando ricevi un'email di richiesta assistenza, analizzala e verifica se contiene le seguenti 3 informazioni obbligatorie:

1. SERIALE: numero seriale del dissalatore
2. ERRORE_PANNELLO: tipo di errore visualizzato sul pannello —
   - pannello BASIC: numero di lampeggi della spia
   - pannello MINITOUCH / TOUCH / DIGITAL: messaggio di errore testuale visualizzato
3. DESCRIZIONE_VIDEO: descrizione dettagliata del problema + video del manometro durante il funzionamento + video del problema quando possibile

LINGUA: rileva la lingua dell'email e rispondi nella stessa lingua del cliente. Il sommario tecnico è sempre in italiano.

Rispondi SOLO con un oggetto JSON valido, senza testo aggiuntivo né backtick:
{
  "seriale": { "presente": true/false, "valore": "valore trovato o null" },
  "errore_pannello": { "presente": true/false, "parziale": true/false, "valore": "valore trovato o null" },
  "descrizione_video": { "presente": true/false, "parziale": true/false, "valore": "breve sintesi o null" },
  "tutte_presenti": true/false,
  "lingua": "it/en/fr/de/es",
  "bozza_risposta": "Risposta nella lingua del cliente. Se mancano info, richiederle in modo professionale. Se manca il tipo di pannello, spiegare la differenza tra BASIC (lampeggi) e MINITOUCH/TOUCH/DIGITAL (messaggio testuale). Se completa, confermare presa in carico.",
  "sommario_tecnico": "Riassunto in italiano per l'ingegnere se tutte_presenti è true, altrimenti null."
}`;

let currentDraft = "";

Office.onReady(() => {});

async function analyzeEmail() {
  const statusEl = document.getElementById("status");
  const resultEl = document.getElementById("result");

  statusEl.textContent = "Lettura email in corso...";
  resultEl.style.display = "none";

  try {
    const item = Office.context.mailbox.item;
    const subject = item.subject || "(nessun oggetto)";

    item.body.getAsync(Office.CoercionType.Text, async (bodyResult) => {
      if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
        statusEl.textContent = "Errore nella lettura del corpo email.";
        return;
      }

      const emailText = `Oggetto: ${subject}\n\n${bodyResult.value}`;
      statusEl.textContent = "Analisi in corso con Claude...";

      try {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 25000);

        const response = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          signal: controller.signal,
          headers: {
            "Content-Type": "application/json",
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01"
          },
          body: JSON.stringify({
            model: MODEL,
            max_tokens: 1000,
            system: SYSTEM_PROMPT,
            messages: [{ role: "user", content: `Analizza questa email:\n\n${emailText}` }]
          })
        });

        clearTimeout(timeout);

        if (!response.ok) {
          const errData = await response.json().catch(() => ({}));
          const code = response.status;
          let hint = "";
          if (code === 401) hint = " → API key non valida.";
          else if (code === 429) hint = " → Credito esaurito.";
          else if (code === 529) hint = " → Servizio sovraccarico, riprova.";
          statusEl.textContent = `Errore API ${code}${hint} ${errData.error?.message || ""}`;
          return;
        }

        const data = await response.json();
        const text = data.content?.map(i => i.text || "").join("") || "";

        let result;
        try {
          result = JSON.parse(text.trim());
        } catch(e) {
          statusEl.textContent = `Errore parsing: ${text.substring(0, 200)}`;
          return;
        }

        renderResult(result);
        statusEl.textContent = "";

      } catch(err) {
        statusEl.textContent = err.name === "AbortError"
          ? "Timeout: nessuna risposta. Riprova."
          : `Errore: ${err.message}`;
      }
    });

  } catch(err) {
    statusEl.textContent = `Errore: ${err.message}`;
  }
}

function renderResult(r) {
  const fields = [
    { key: "seriale",           label: "Numero seriale" },
    { key: "errore_pannello",   label: "Errore pannello" },
    { key: "descrizione_video", label: "Descrizione + video" }
  ];

  let rows = "";
  fields.forEach(f => {
    const info = r[f.key];
    let badge, detail;
    if (info.presente) {
      badge = `<span class="badge-ok">Presente</span>`;
      detail = info.valore ? `<span style="color:#444;font-size:11px;margin-left:6px">${info.valore}</span>` : "";
    } else if (info.parziale) {
      badge = `<span class="badge-par">Parziale</span>`;
      detail = "";
    } else {
      badge = `<span class="badge-miss">Mancante</span>`;
      detail = "";
    }
    rows += `<div class="row"><span>${f.label}${detail}</span>${badge}</div>`;
  });

  document.getElementById("info-rows").innerHTML = rows;

  currentDraft = r.bozza_risposta || "";
  const langFlag = { it: "🇮🇹", en: "🇬🇧", fr: "🇫🇷", de: "🇩🇪", es: "🇪🇸" };
  const flag = langFlag[r.lingua] || "🌐";

  const actionSection = document.getElementById("action-section");
  if (r.tutte_presenti) {
    actionSection.innerHTML = `
      <div class="section-title" style="color:#107c10">✔ Ticket completo — pronto per analisi</div>
      <div class="draft">${r.sommario_tecnico || ""}</div>`;
  } else {
    actionSection.innerHTML = `
      <div class="section-title">Bozza risposta al cliente ${flag}</div>
      <div class="draft">${currentDraft}</div>`;
  }

  document.getElementById("result").style.display = "block";
}

function copyDraft() {
  if (!currentDraft) return;
  const el = document.createElement("textarea");
  el.value = currentDraft;
  document.body.appendChild(el);
  el.select();
  document.execCommand("copy");
  document.body.removeChild(el);
  document.getElementById("status").textContent = "Bozza copiata negli appunti.";
  setTimeout(() => { document.getElementById("status").textContent = ""; }, 2000);
}

function replyWithDraft() {
  if (!currentDraft) return;
  Office.context.mailbox.item.displayReplyForm({
    htmlBody: currentDraft.replace(/\n/g, "<br>")
  });
}
