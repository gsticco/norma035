require("dotenv").config();
const express = require("express");
const path = require("path");
const { google } = require("googleapis");

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const SHEET_ID = "1zU9eBp4FgOscu_itZOkvEF62xDabmxMtV-pvBTCHrNA";
const DRIVE_FOLDER_ID = "1lK9c6e2P3be6iNJ8ek1MHHhmL04QRXTE";
const CREDENTIALS_FILE = path.join(__dirname, "fluid-skyline-440202-j3-f1a8b510425a.json");

// ── Auth helpers ──────────────────────────────────────────────────────────────
// Supports both local file (dev) and environment variable (Railway/production)

function getAuth(scopes) {
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try {
      const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
      return new google.auth.GoogleAuth({ credentials, scopes });
    } catch (e) {
      console.error("Error al parsear GOOGLE_CREDENTIALS_JSON:", e.message);
    }
  }
  return new google.auth.GoogleAuth({ keyFile: CREDENTIALS_FILE, scopes });
}

function sheetsClient(auth) {
  return google.sheets({ version: "v4", auth });
}

function docsClient(auth) {
  return google.docs({ version: "v1", auth });
}

function driveClient(auth) {
  return google.drive({ version: "v3", auth });
}

// ── Sheets: ensure headers ────────────────────────────────────────────────────

async function ensureHeaders() {
  const auth = getAuth(["https://www.googleapis.com/auth/spreadsheets"]);
  const sheets = sheetsClient(auth);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "Hoja1!A1",
  });
  if (!res.data.values || res.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: "Hoja1!A1",
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          "Fecha",
          "Empresa",
          "Persona de contacto",
          "Email diagnóstico",
          "Email informe",
          "Tamaño empresa",
          "P1 – Diagnóstico riesgo",
          "P2 – Política prevención",
          "P3 – Señales de alerta",
          "P4 – Canal de reporte",
          "P5 – Perspectiva de género",
          "Puntaje total",
          "Semáforo",
          "Qué cumple",
          "Qué le falta",
          "Link informe Drive",
          "Diagnóstico completo",
        ]],
      },
    });
  }
}

ensureHeaders().catch(console.error);

// ── Sheets: check if empresa already diagnosed ────────────────────────────────

async function checkEmpresaExiste(nombreEmpresa) {
  try {
    const auth = getAuth(["https://www.googleapis.com/auth/spreadsheets.readonly"]);
    const sheets = sheetsClient(auth);
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Hoja1!B:B",
    });
    const rows = res.data.values || [];
    const nombre = nombreEmpresa.toLowerCase().trim();
    return rows.some((row) => row[0] && row[0].toLowerCase().trim() === nombre);
  } catch (e) {
    console.error("Error al verificar empresa:", e);
    return false;
  }
}

// ── Sheets: append row ────────────────────────────────────────────────────────

async function appendToSheet(rowData) {
  const auth = getAuth(["https://www.googleapis.com/auth/spreadsheets"]);
  const sheets = sheetsClient(auth);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "Hoja1!A1",
    valueInputOption: "RAW",
    requestBody: { values: [rowData] },
  });
}

// ── Drive: create Google Doc with diagnosis ───────────────────────────────────

async function createDriveDoc(nombreEmpresa, diagnosticoTexto, parsed) {
  const auth = getAuth([
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
  ]);
  const docs = docsClient(auth);
  const drive = driveClient(auth);

  const fecha = new Date().toLocaleDateString("es-MX", {
    timeZone: "America/Mexico_City",
    year: "numeric", month: "long", day: "numeric",
  });

  const titulo = `Diagnóstico NOM-035 — ${nombreEmpresa}`;

  // Create empty doc in the folder
  const fileRes = await drive.files.create({
    requestBody: {
      name: titulo,
      mimeType: "application/vnd.google-apps.document",
      parents: [DRIVE_FOLDER_ID],
    },
    fields: "id,webViewLink",
    supportsAllDrives: true,
  });

  const docId = fileRes.data.id;
  const docLink = fileRes.data.webViewLink;

  // Semáforo emoji
  const semaforoEmoji =
    parsed.semaforo === "VERDE" ? "🟢" :
    parsed.semaforo === "AMARILLO" ? "🟡" : "🔴";

  // Build document content via batchUpdate
  const requests = [
    // Title
    { insertText: { location: { index: 1 }, text: `${titulo}\n` } },
    { updateParagraphStyle: {
        range: { startIndex: 1, endIndex: titulo.length + 1 },
        paragraphStyle: { namedStyleType: "HEADING_1" },
        fields: "namedStyleType",
    }},

    // Metadata block
    { insertText: { location: { index: titulo.length + 2 }, text:
      `Fecha: ${fecha}\nEmpresa: ${parsed.empresa || nombreEmpresa}\nPersona de contacto: ${parsed.contacto || "—"}\nEmail: ${parsed.email || "—"}\nTamaño de la empresa: ${parsed.tamano || "—"}\n\n`
    }},
  ];

  // Compute offset after title + metadata
  const metaText = `Fecha: ${fecha}\nEmpresa: ${parsed.empresa || nombreEmpresa}\nPersona de contacto: ${parsed.contacto || "—"}\nEmail: ${parsed.email || "—"}\nTamaño de la empresa: ${parsed.tamano || "—"}\n\n`;
  let idx = titulo.length + 2 + metaText.length;

  // Semáforo heading
  const semaforoLine = `Resultado: ${semaforoEmoji} ${parsed.semaforo}\n`;
  requests.push({ insertText: { location: { index: idx }, text: semaforoLine } });
  requests.push({ updateParagraphStyle: {
    range: { startIndex: idx, endIndex: idx + semaforoLine.length },
    paragraphStyle: { namedStyleType: "HEADING_2" },
    fields: "namedStyleType",
  }});
  idx += semaforoLine.length;

  // Qué cumple
  const cumpleHeading = "Qué está cumpliendo bien\n";
  requests.push({ insertText: { location: { index: idx }, text: cumpleHeading } });
  requests.push({ updateParagraphStyle: {
    range: { startIndex: idx, endIndex: idx + cumpleHeading.length },
    paragraphStyle: { namedStyleType: "HEADING_3" },
    fields: "namedStyleType",
  }});
  idx += cumpleHeading.length;
  const cumpleBody = `${parsed.cumple || "—"}\n\n`;
  requests.push({ insertText: { location: { index: idx }, text: cumpleBody } });
  idx += cumpleBody.length;

  // Qué le falta
  const faltaHeading = "Qué le falta mejorar\n";
  requests.push({ insertText: { location: { index: idx }, text: faltaHeading } });
  requests.push({ updateParagraphStyle: {
    range: { startIndex: idx, endIndex: idx + faltaHeading.length },
    paragraphStyle: { namedStyleType: "HEADING_3" },
    fields: "namedStyleType",
  }});
  idx += faltaHeading.length;
  const faltaBody = `${parsed.falta || "—"}\n\n`;
  requests.push({ insertText: { location: { index: idx }, text: faltaBody } });
  idx += faltaBody.length;

  // Puntaje
  const puntajeHeading = "Puntaje por dimensión\n";
  requests.push({ insertText: { location: { index: idx }, text: puntajeHeading } });
  requests.push({ updateParagraphStyle: {
    range: { startIndex: idx, endIndex: idx + puntajeHeading.length },
    paragraphStyle: { namedStyleType: "HEADING_3" },
    fields: "namedStyleType",
  }});
  idx += puntajeHeading.length;
  const puntajeBody =
    `P1 – Diagnóstico de riesgo psicosocial: ${parsed.p1_nivel || "—"}/5\n` +
    `P2 – Política y prevención documentada: ${parsed.p2_nivel || "—"}/5\n` +
    `P3 – Señales de alerta organizacional: ${parsed.p3_nivel || "—"}/5\n` +
    `P4 – Canal de reporte y atención a casos: ${parsed.p4_nivel || "—"}/5\n` +
    `P5 – Perspectiva de género transversal: ${parsed.p5_nivel || "—"}/5\n` +
    `Puntaje total: ${parsed.puntaje_total || "—"}/25\n\n`;
  requests.push({ insertText: { location: { index: idx }, text: puntajeBody } });
  idx += puntajeBody.length;

  // Diagnóstico completo
  const diagHeading = "Diagnóstico completo\n";
  requests.push({ insertText: { location: { index: idx }, text: diagHeading } });
  requests.push({ updateParagraphStyle: {
    range: { startIndex: idx, endIndex: idx + diagHeading.length },
    paragraphStyle: { namedStyleType: "HEADING_2" },
    fields: "namedStyleType",
  }});
  idx += diagHeading.length;
  requests.push({ insertText: { location: { index: idx }, text: diagnosticoTexto + "\n" } });

  await docs.documents.batchUpdate({ documentId: docId, requestBody: { requests } });

  return { docId, docLink };
}

// ── System prompt ─────────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `Eres una especialista en la NOM-035-STPS-2018 (Norma 035) de México, del equipo de Grow - Género y Trabajo. Tu misión es hacer un diagnóstico inicial sobre si una empresa cumple con la norma, integrando perspectiva de género interseccional.

═══════════════════════════════════════
FLUJO COMPLETO — SEGUÍ ESTE ORDEN EXACTO
═══════════════════════════════════════

PASO 0 — BIENVENIDA Y DATOS ORGANIZACIONALES
Presentate brevemente. Explicá que antes de empezar necesitás tres datos:
1. Nombre de la organización
2. Nombre de la persona de contacto
3. Email de contacto (para seguimiento interno de Grow)

Pedí los tres en un solo mensaje. Esperá la respuesta antes de continuar.

PASO 0B — VERIFICACIÓN DE DUPLICADO
Una vez que tengas el nombre de la empresa, el sistema verificará automáticamente si ya realizó el autodiagnóstico. Si ya existe, NO hagas las preguntas. Mostrá el mensaje de empresa ya registrada (ver sección al final).

PASO 1 — TAMAÑO DE LA EMPRESA
Preguntá cuántos trabajadores tiene la empresa. Aclarár que esto determina las obligaciones bajo la norma: hasta 15 trabajadores tienen obligaciones básicas, de 16 a 50 intermedias, y más de 50 obligaciones completas.

═══════════════════════════════════════
LAS 5 PREGUNTAS — UNA A LA VEZ
═══════════════════════════════════════

PREGUNTA 1
Hacé esta pregunta exactamente así, sin título ni encabezado:
"¿De qué manera identifica tu empresa si hay situaciones de estrés, sobrecarga o malestar en el equipo? ¿Con qué herramienta o mecanismo, y qué hacen con esa información?"

⚠️ REPREGUNTA OBLIGATORIA EN P1: Después de que respondan, si NO mencionaron explícitamente las Guías de Referencia de la STPS (Guía I, II o III de la NOM-035), preguntá específicamente: "¿Han utilizado alguna de las Guías de Referencia de la STPS para la NOM-035 (Guía I, II o III) para hacer ese diagnóstico, o se basaron en otro tipo de instrumento?" Esperá la respuesta antes de continuar a P2.

PREGUNTA 2
Hacé esta pregunta exactamente así, sin título ni encabezado:
"¿Tu empresa tiene por escrito alguna política o compromiso sobre cómo prevenir el malestar, el acoso o la violencia en el trabajo? ¿Quién la conoce y cómo se comunica?"

PREGUNTA 3
Hacé esta pregunta exactamente así, sin título ni encabezado:
"¿Llevan registro de rotación, ausentismo o conflictos en el equipo? ¿Saben si esas situaciones se concentran más en algún área, tipo de puesto o grupo de personas?"

PREGUNTA 4
Hacé esta pregunta exactamente así, sin título ni encabezado:
"Cuando alguien en tu empresa vive una situación difícil — un conflicto, una situación de acoso o algo que le genera malestar — ¿tiene un lugar claro a dónde ir? ¿Cómo funciona ese proceso y qué pasa después?"

PREGUNTA 5
Hacé esta pregunta exactamente así, sin título ni encabezado:
"Pensando en todo lo que mencionaste: ¿alguna vez han analizado si las mujeres, los hombres u otros grupos del equipo viven de manera distinta el trabajo, el estrés o las oportunidades? ¿Eso influye en cómo diseñan sus prácticas de RRHH?"

═══════════════════════════════════════
REGLAS DURANTE LAS PREGUNTAS
═══════════════════════════════════════
- Una sola pregunta por mensaje.
- Después de cada respuesta, comentá brevemente (1 línea) qué implica para la norma y continuá.
- Si la persona no entiende un concepto, explicalo con palabras simples.
- Tono: profesional pero accesible.

═══════════════════════════════════════
EVALUACIÓN INTERNA — RÚBRICA POR PREGUNTA
═══════════════════════════════════════
Evaluá cada respuesta en una escala del 1 al 5. Usá esta rúbrica:

Nivel 1: No existe práctica ni intención. Riesgo inmediato de incumplimiento.
Nivel 2: Hay algo informal o esporádico, sin documentación ni proceso.
Nivel 3: Existe una práctica básica pero genérica, sin análisis ni enfoque diferenciado.
Nivel 4: Hay proceso documentado pero sin perspectiva de género ni análisis por grupos.
Nivel 5: Proceso documentado, con análisis diferenciado por género y grupos, con evidencia.

Para P1 específicamente: si usan las Guías de Referencia STPS → suma +1 al nivel. Si no las usan pero tienen otro instrumento válido → nivel 3 como máximo.

Puntaje total = suma de los 5 niveles (mínimo 5, máximo 25).

═══════════════════════════════════════
DIAGNÓSTICO FINAL
═══════════════════════════════════════

SEMÁFORO:
🔴 ROJO (5–10 pts): Cumplimiento bajo — riesgos significativos, lejos de cumplir la norma.
🟡 AMARILLO (11–17 pts): Cumplimiento medio — hay avances pero faltan elementos clave.
🟢 VERDE (18–25 pts): Cumplimiento alto — cumple los aspectos principales, foco en mejora continua.

El diagnóstico debe incluir:
1. Semáforo con explicación de 2–3 líneas
2. "Qué está cumpliendo bien:" — listado concreto
3. "Qué le falta mejorar:" — listado concreto con justificación citando lo que dijeron
4. Puntaje por pregunta (P1: X/5, P2: X/5, etc.) con una línea de justificación por cada una
5. "Acción prioritaria:" — la más urgente

Cerrá con:
"Para continuar podés participar de un encuentro online con el equipo especializado de Grow - Género y Trabajo para resolver dudas, o contactarte a info@generoytrabajo.com para solicitar un presupuesto por la asesoría técnica. Para un diagnóstico completo se recomienda aplicar las guías de referencia de la STPS."

═══════════════════════════════════════
PREGUNTA FINAL — EMAIL PARA INFORME
═══════════════════════════════════════
Después del diagnóstico, preguntá:
"¿Te enviamos el informe completo a [email que dieron al inicio] o preferís indicar otro email?"

VALIDACIÓN DE EMAIL OBLIGATORIA: Antes de confirmar, verificá que el email tenga formato válido:
- Debe contener exactamente un @
- Debe tener al menos un punto (.) después del @
- El dominio debe tener al menos 2 caracteres después del último punto (por ejemplo: .com, .mx, .org)
- Ejemplos inválidos: pepa@com, usuario@, @dominio.com, texto sin arroba
- Si el email no cumple estas condiciones, decí: "Ese email no parece tener un formato válido (por ejemplo, falta el dominio como .com o .mx). ¿Podés verificarlo e indicármelo de nuevo?"
- No avances hasta tener un email con formato válido.

Una vez confirmado el email válido, respondé: "Perfecto, el informe llegará a [email confirmado] en las próximas horas. ¡Gracias por completar el diagnóstico!"

═══════════════════════════════════════
EMPRESA YA REGISTRADA
═══════════════════════════════════════
Si el sistema indica que la empresa ya realizó el diagnóstico, NO cierres la conversación. En cambio, preguntá:
"Veo que [nombre empresa] ya realizó el autodiagnóstico con Grow anteriormente. ¿Se trata de una sucursal o sede diferente? Si es así, por favor aclará el nombre completo con la sucursal (por ejemplo: 'Babsa — Sucursal Monterrey') y continuamos con el diagnóstico. Si es la misma sede, podés escribirnos a info@generoytrabajo.com para revisar los resultados anteriores."

Si confirman que es una sucursal diferente, tomá el nombre completo con sucursal como nombre de empresa y continuá el flujo normalmente desde PASO 1.
Si confirman que es la misma sede, cerrá con: "Perfecto, escribinos a info@generoytrabajo.com y te compartimos el diagnóstico anterior."

═══════════════════════════════════════
BLOQUE DE DATOS — AL TERMINAR TODO
═══════════════════════════════════════
Al final de tu último mensaje (después de confirmar el email del informe), incluí este bloque invisible para el usuario:

[DATOS]{"empresa":"nombre","contacto":"nombre persona contacto","email":"email diagnóstico","email_informe":"email para enviar informe","tamano":"número o rango de trabajadores tal como lo dijo la empresa","p1":"respuesta textual p1","p2":"respuesta textual p2","p3":"respuesta textual p3","p4":"respuesta textual p4","p5":"respuesta textual p5","p1_nivel":0,"p2_nivel":0,"p3_nivel":0,"p4_nivel":0,"p5_nivel":0,"puntaje_total":0,"semaforo":"ROJO/AMARILLO/VERDE","cumple":"resumen de lo que cumple","falta":"resumen de lo que le falta"}[/DATOS]`;

// ── Check empresa endpoint ────────────────────────────────────────────────────

app.post("/api/check-empresa", async (req, res) => {
  const { empresa } = req.body;
  if (!empresa) return res.json({ existe: false });
  try {
    const existe = await checkEmpresaExiste(empresa);
    res.json({ existe });
  } catch (e) {
    console.error(e);
    res.json({ existe: false });
  }
});

// ── Chat endpoint ─────────────────────────────────────────────────────────────

app.post("/api/chat", async (req, res) => {
  const { messages } = req.body;
  const msgs = (!messages || !Array.isArray(messages) || messages.length === 0)
    ? [{ role: "user", content: "Inicia el diagnóstico." }]
    : messages;

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: "API key no configurada." });

  // ── Server-side empresa duplicate check ─────────────────────────────────────
  // On the second user message (first real response with org data), check Sheets
  const userMessages = msgs.filter(m => m.role === "user");
  if (userMessages.length === 2) {
    // User just provided org name/contact/email — try to extract empresa name
    const lastUserMsg = userMessages[userMessages.length - 1].content || "";
    const lines = lastUserMsg.split(/\n/).map(l => l.trim()).filter(Boolean);
    let empresaNombre = "";
    for (const line of lines) {
      const m = line.match(/(?:empresa|organización|organizacion|compañía|compania)[:\s]+(.+)/i);
      if (m) { empresaNombre = m[1].trim(); break; }
    }
    if (!empresaNombre && lines[0] && lines[0].split(" ").length <= 8) {
      empresaNombre = lines[0];
    }
    if (empresaNombre) {
      const yaExiste = await checkEmpresaExiste(empresaNombre);
      if (yaExiste) {
        msgs.push({
          role: "assistant",
          content: `[SISTEMA INTERNO — NO MOSTRAR AL USUARIO: La empresa "${empresaNombre}" ya tiene un diagnóstico registrado en la base de datos. Aplicá el flujo de EMPRESA YA REGISTRADA: preguntá si es una sucursal diferente antes de continuar.]`
        });
        // Add a dummy user turn to keep alternating roles valid
        msgs.push({ role: "user", content: "(continuar)" });
      }
    }
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-5",
        max_tokens: 1800,
        system: SYSTEM_PROMPT,
        messages: msgs,
      }),
    });

    if (!response.ok) {
      const err = await response.json();
      return res.status(response.status).json({ error: err.error?.message || "Error de la API." });
    }

    const data = await response.json();
    let reply = data.content.map((b) => b.text || "").join("");

    // ── Extract [DATOS] block and save ───────────────────────────────────────
    const dataMatch = reply.match(/\[DATOS\](.*?)\[\/DATOS\]/s);
    if (dataMatch) {
      try {
        const parsed = JSON.parse(dataMatch[1]);
        const fecha = new Date().toLocaleString("es-MX", { timeZone: "America/Mexico_City" });
        const diagnosticoLimpio = reply.replace(/\[DATOS\].*?\[\/DATOS\]/s, "").trim();

        // Create Google Doc
        let docLink = "";
        try {
          const { docLink: link } = await createDriveDoc(
            parsed.empresa || "Empresa",
            diagnosticoLimpio,
            parsed
          );
          docLink = link;
        } catch (driveErr) {
          console.error("Error al crear Doc en Drive:", driveErr);
        }

        // Save to Sheets
        await appendToSheet([
          fecha,
          parsed.empresa || "",
          parsed.contacto || "",
          parsed.email || "",
          parsed.email_informe || "",
          parsed.tamano || "",
          parsed.p1 || "",
          parsed.p2 || "",
          parsed.p3 || "",
          parsed.p4 || "",
          parsed.p5 || "",
          parsed.puntaje_total || "",
          parsed.semaforo || "",
          parsed.cumple || "",
          parsed.falta || "",
          docLink,
          diagnosticoLimpio,
        ]);
      } catch (e) {
        console.error("Error al guardar datos:", e);
      }

      reply = reply.replace(/\[DATOS\].*?\[\/DATOS\]/s, "").trim();
    }

    res.json({ reply });
  } catch (error) {
    console.error("Error al llamar a la API:", error);
    res.status(500).json({ error: "Error interno del servidor." });
  }
});

// ── Start ─────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en http://localhost:${PORT}`));
