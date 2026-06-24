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

function sheetsClient(auth) { return google.sheets({ version: "v4", auth }); }
function docsClient(auth)   { return google.docs({ version: "v1", auth }); }
function driveClient(auth)  { return google.drive({ version: "v3", auth }); }

async function ensureHeaders() {
  const auth = getAuth(["https://www.googleapis.com/auth/spreadsheets"]);
  const sheets = sheetsClient(auth);
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "Hoja1!A1" });
  if (!res.data.values || res.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: "Hoja1!A1", valueInputOption: "RAW",
      requestBody: { values: [[
        "Fecha","Empresa","Persona de contacto","Email diagnostico","Email informe","Tamano empresa","Pais",
        "P1 Diagnostico psicosocial","P2 Acciones post-diagnostico","P3 Politica prevencion",
        "P4 Eventos traumaticos","P5 Examenes medicos","P6 Registros documentales",
        "P7 Senales alerta canal","P8 Liderazgos capacitaciones","P9 Perspectiva genero",
        "Puntaje NOM-035 (max40)","Semaforo NOM-035","Puntaje Genero (max7)","Semaforo Genero",
        "Link informe Drive",
      ]] },
    });
    console.log("Headers de Sheets creados OK");
  }
}
ensureHeaders().catch(e => console.error("Error en ensureHeaders:", e));

async function checkEmpresaExiste(nombreEmpresa) {
  try {
    const auth = getAuth(["https://www.googleapis.com/auth/spreadsheets.readonly"]);
    const sheets = sheetsClient(auth);
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "Hoja1!B:B" });
    const rows = res.data.values || [];
    const nombre = nombreEmpresa.toLowerCase().trim();
    return rows.some((row) => row[0] && row[0].toLowerCase().trim() === nombre);
  } catch (e) { console.error("Error al verificar empresa:", e); return false; }
}

async function appendToSheet(rowData) {
  console.log("Iniciando appendToSheet con", rowData.length, "columnas...");
  const auth = getAuth(["https://www.googleapis.com/auth/spreadsheets"]);
  const sheets = sheetsClient(auth);
  const result = await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: "Hoja1!A1", valueInputOption: "RAW",
    requestBody: { values: [rowData] },
  });
  console.log("appendToSheet OK — updatedRange:", result.data.updates && result.data.updates.updatedRange);
  return result;
}


async function createDriveDoc(nombreEmpresa, diagnosticoTexto, detalleTexto, parsed) {
  const auth = getAuth(["https://www.googleapis.com/auth/documents","https://www.googleapis.com/auth/drive"]);
  const docs = docsClient(auth);
  const drive = driveClient(auth);
  const fecha = new Date().toLocaleDateString("es-MX", { timeZone: "America/Mexico_City", year: "numeric", month: "long", day: "numeric" });
  const pais = parsed.pais || "";
  const titulo = "Diagnostico NOM-035 " + nombreEmpresa + (pais ? " – " + pais : "");
  const fileRes = await drive.files.create({
    requestBody: { name: titulo, mimeType: "application/vnd.google-apps.document", parents: [DRIVE_FOLDER_ID] },
    fields: "id,webViewLink", supportsAllDrives: true,
  });
  const docId = fileRes.data.id;
  const docLink = fileRes.data.webViewLink;
  const lbl = (v) => v === "VERDE" ? "VERDE" : v === "AMARILLO" ? "AMARILLO" : "ROJO";

  const requests = [
    { insertText: { location: { index: 1 }, text: titulo + "\n" } },
    { updateParagraphStyle: { range: { startIndex: 1, endIndex: titulo.length + 1 }, paragraphStyle: { namedStyleType: "HEADING_1" }, fields: "namedStyleType" }},
  ];

  const metaText = "Fecha: " + fecha +
    "\nEmpresa: " + (parsed.empresa || nombreEmpresa) +
    "\nContacto: " + (parsed.contacto || "-") +
    "\nEmail: " + (parsed.email || "-") +
    "\nTamano: " + (parsed.tamano || "-") +
    "\nPais: " + (parsed.pais || "-") + "\n\n";
  requests.push({ insertText: { location: { index: titulo.length + 2 }, text: metaText } });
  let idx = titulo.length + 2 + metaText.length;

  // --- Semáforos ---
  const sem1 = "Semaforo NOM-035: " + lbl(parsed.semaforo_norma) + " (" + (parsed.puntaje_norma || "-") + "/40)\n";
  requests.push({ insertText: { location: { index: idx }, text: sem1 } });
  requests.push({ updateParagraphStyle: { range: { startIndex: idx, endIndex: idx + sem1.length }, paragraphStyle: { namedStyleType: "HEADING_2" }, fields: "namedStyleType" }});
  idx += sem1.length;

  const sem2 = "Semaforo Perspectiva de Genero: " + lbl(parsed.semaforo_genero) + " (" + (parsed.puntaje_genero || "-") + "/7)\n";
  requests.push({ insertText: { location: { index: idx }, text: sem2 } });
  requests.push({ updateParagraphStyle: { range: { startIndex: idx, endIndex: idx + sem2.length }, paragraphStyle: { namedStyleType: "HEADING_2" }, fields: "namedStyleType" }});
  idx += sem2.length;
  requests.push({ insertText: { location: { index: idx }, text: "\n" } });
  idx += 1;

  // --- Evaluación detallada por dimensión ---
  if (detalleTexto && detalleTexto.trim().length > 10) {
    const detH = "Evaluacion detallada por dimension\n";
    requests.push({ insertText: { location: { index: idx }, text: detH } });
    requests.push({ updateParagraphStyle: { range: { startIndex: idx, endIndex: idx + detH.length }, paragraphStyle: { namedStyleType: "HEADING_2" }, fields: "namedStyleType" }});
    idx += detH.length;
    const detB = detalleTexto.trim() + "\n\n";
    requests.push({ insertText: { location: { index: idx }, text: detB } });
    idx += detB.length;
  }

  // --- Resumen cumple/falta ---
  for (const s of [
    { heading: "Que esta cumpliendo bien\n", body: (parsed.cumple || "-") + "\n\n" },
    { heading: "Que le falta mejorar\n",     body: (parsed.falta  || "-") + "\n\n" }
  ]) {
    requests.push({ insertText: { location: { index: idx }, text: s.heading } });
    requests.push({ updateParagraphStyle: { range: { startIndex: idx, endIndex: idx + s.heading.length }, paragraphStyle: { namedStyleType: "HEADING_3" }, fields: "namedStyleType" }});
    idx += s.heading.length;
    requests.push({ insertText: { location: { index: idx }, text: s.body } });
    idx += s.body.length;
  }

  // --- Puntaje por dimensión ---
  const pH = "Puntaje por dimension\n";
  requests.push({ insertText: { location: { index: idx }, text: pH } });
  requests.push({ updateParagraphStyle: { range: { startIndex: idx, endIndex: idx + pH.length }, paragraphStyle: { namedStyleType: "HEADING_3" }, fields: "namedStyleType" }});
  idx += pH.length;
  const pB = [
    "P1 Diagnostico psicosocial: " + (parsed.p1_nivel||"-") + "/5",
    "P2 Acciones post-diagnostico: " + (parsed.p2_nivel||"-") + "/5",
    "P3 Politica de prevencion: " + (parsed.p3_nivel||"-") + "/5",
    "P4 Eventos traumaticos: " + (parsed.p4_nivel||"-") + "/5",
    "P5 Examenes medicos: " + (parsed.p5_nivel||"-") + "/5",
    "P6 Registros documentales: " + (parsed.p6_nivel||"-") + "/5",
    "P7 Senales y canal: " + (parsed.p7_nivel||"-") + "/5",
    "P8 Liderazgos: " + (parsed.p8_nivel||"-") + "/5",
    "PUNTAJE NOM-035: " + (parsed.puntaje_norma||"-") + "/40",
    "",
    "P9 Perspectiva de genero: " + (parsed.p9_nivel||"-") + "/5",
    "PUNTAJE GENERO: " + (parsed.puntaje_genero||"-") + "/7",
    "",
  ].join("\n");
  requests.push({ insertText: { location: { index: idx }, text: pB } });
  idx += pB.length;

  await docs.documents.batchUpdate({ documentId: docId, requestBody: { requests } });
  return { docId, docLink };
}


const SYSTEM_PROMPT = [
  "Eres una especialista en NOM-035-STPS-2018 de Mexico, del equipo de Grow - Genero y Trabajo. Haces diagnosticos preliminares de cumplimiento de la norma con perspectiva de genero interseccional.",
  "",
  "FLUJO OBLIGATORIO — SEGUILO EN ORDEN",
  "",
  "PASO 0 — BIENVENIDA",
  "Presentate con este mensaje exacto (podes adaptarlo levemente):",
  "",
  '"Hola, soy parte del equipo de Grow - Genero y Trabajo. Esta herramienta hace un diagnostico PRELIMINAR del cumplimiento de la NOM-035-STPS-2018, y va un paso mas alla: evalua tambien si las practicas incorporan perspectiva de genero interseccional, aunque la norma no lo requiera. Por eso los resultados incluyen dos semaforos: uno para cumplimiento de la norma y otro para perspectiva de genero.',
  "",
  "Al finalizar te enviamos por email el reporte completo con el detalle por dimension, lo que estan haciendo bien y los pasos recomendados. Para eso necesito cuatro datos:",
  "1. Nombre de la organizacion",
  "2. Pais donde opera",
  "3. Nombre de la persona de contacto",
  "4. Email de contacto",
  "",
  'Esta informacion es solo para uso interno de Grow - Genero y Trabajo. No la compartimos con terceros."',
  "",
  "Espera los tres datos en un solo mensaje antes de continuar.",
  "",
  "PASO 0B — VERIFICACION DE DUPLICADO",
  "Verifica si la empresa ya hizo el diagnostico. Si ya existe, no hagas las preguntas — usa el mensaje de EMPRESA YA REGISTRADA.",
  "",
  "PASO 1 — TAMANO DE LA EMPRESA",
  "Pregunta cuantos trabajadores tiene. Aclara: hasta 15 = obligaciones basicas, 16-50 = intermedias, mas de 50 = completas.",
  "",
  "LAS 9 PREGUNTAS — UNA A LA VEZ",
  "",
  "PREGUNTA 1",
  '"De que manera identifica tu empresa si hay situaciones de estres, sobrecarga o malestar en el equipo? Con que herramienta o mecanismo lo hacen, con que frecuencia, y que hacen con esa informacion?"',
  "REPREGUNTA GUIAS (si no las menciono): Han utilizado alguna de las Guias de Referencia de la STPS para la NOM-035 (Guia I, II o III), o se basaron en otro instrumento?",
  "REPREGUNTA PERIODICIDAD (si no la menciono): Cuando fue la ultima vez que hicieron esa evaluacion?",
  "",
  "PREGUNTA 2",
  '"Despues de hacer esa evaluacion: que hicieron con los resultados? Los analizaron por area, puesto o grupo? Implementaron algun programa de intervencion? Compartieron los resultados con el personal?"',
  "",
  "PREGUNTA 3",
  '"Tu empresa tiene por escrito una politica que aborde especificamente: (1) prevencion de riesgos psicosociales, (2) prevencion de violencia laboral, y (3) promocion de un entorno organizacional favorable? Quien la conoce y como se comunica?"',
  "",
  "PREGUNTA 4",
  '"La empresa tiene algun proceso para identificar a personas que hayan sufrido o presenciado eventos traumaticos graves — accidentes, asaltos, actos de violencia? Como las canalizan a atencion medica o psicologica?"',
  "",
  "PREGUNTA 5",
  '"Cuando alguien muestra signos de afectacion emocional o psicologica severa, o cuando hay situaciones de violencia: realizan evaluaciones clinicas o los derivan a atencion medica o psicologica? Como funciona ese proceso?"',
  "",
  "PREGUNTA 6",
  '"Llevan un registro formal de los resultados de las evaluaciones, las medidas adoptadas y los nombres de las personas evaluadas clinicamente? Donde se documenta y quien tiene acceso?"',
  "",
  "PREGUNTA 7",
  '"Llevan registro de indicadores como rotacion, ausentismo o conflictos? Identifican si se concentran en algun area o grupo? Y cuando alguien vive una situacion dificil — conflicto, acoso, malestar — tiene un lugar claro a donde ir? Que pasa despues?"',
  "",
  "PREGUNTA 8",
  '"Las personas en roles de liderazgo recibieron capacitacion especifica en gestion de conflictos, prevencion de violencia o liderazgo saludable? Los equipos tienen algun mecanismo para evaluar a sus lideres?"',
  "",
  "PREGUNTA 9",
  '"Alguna vez analizaron si las mujeres, los hombres u otras identidades de genero del equipo viven de manera distinta el trabajo, el estres o las oportunidades? Eso influye en como disenan sus practicas de RRHH o interpretan los resultados?"',
  "",
  "REGLAS DURANTE LAS PREGUNTAS",
  "- Una sola pregunta principal por mensaje.",
  "- Si una pregunta tiene dos o mas aspectos y la persona solo responde uno, hace al menos una repregunta sobre lo que falto antes de pasar a la siguiente.",
  "- Despues de cada respuesta, una linea breve sobre que implica para la norma.",
  "- Tono profesional pero accesible.",
  "",
  "EVALUACION INTERNA (no mostrar al usuario)",
  "Nivel 1: No existe practica ni intencion.",
  "Nivel 2: Algo informal o esporadico, sin documentacion.",
  "Nivel 3: Practica basica generica, sin analisis diferenciado.",
  "Nivel 4: Proceso documentado, sin perspectiva de genero.",
  "Nivel 5: Proceso documentado, analisis diferenciado por genero, con evidencia.",
  "",
  "Criterios: P1 sin Guias STPS = maximo 2. P2 sin difusion = maximo 3, sin programa = maximo 2. P3 sin 3 componentes = maximo 3. P4 sin proceso = nivel 1. P5 sin derivacion clinica = maximo 2.",
  "PUNTAJE NOM-035 = suma P1 a P8 (min 8, max 40). ROJO=8-16, AMARILLO=17-28, VERDE=29-40.",
  "PUNTAJE GENERO = nivel P9 mas hasta +2 si hubo analisis de genero espontaneo en P1-P8 (max 7). ROJO=1-2, AMARILLO=3-4, VERDE=5-7.",
  "",
  "DIAGNOSTICO FINAL — LO QUE MOSTRAS EN EL CHAT",
  "Mostra SOLO esto, nada mas:",
  "",
  '"Diagnostico preliminar:',
  "",
  "Semaforo NOM-035: [ROJO/AMARILLO/VERDE]",
  "[2-3 lineas sobre que implica para la empresa]",
  "",
  "Semaforo Perspectiva de Genero: [ROJO/AMARILLO/VERDE]",
  "[2-3 lineas sobre que implica]",
  "",
  'Te enviamos el reporte completo al email que nos indicaste en las proximas horas. Para cualquier duda escribinos a info@generoytrabajo.com."',
  "",
  "NO agregues nada mas. No menciones componentes de la norma, no hagas listas, no invites a encuentros, no expliques que involucra el cumplimiento real. Solo los dos semaforos y ese cierre.",
  "",
  "PREGUNTA FINAL — EMAIL",
  "Antes de mostrar el diagnostico, pregunta: Te enviamos el informe completo a [email del inicio] o preferis indicar otro email?",
  "Valida el email: debe tener @, punto despues del @, dominio de 2+ chars. Si no cumple, pide que lo verifique.",
  "Una vez confirmado el email, muestra el diagnostico y cerra con el texto indicado arriba.",
  "Despues de ese cierre, el ULTIMO mensaje debe terminar con el bloque [DETALLE] y luego el bloque [DATOS].",
  "",
  "EMPRESA YA REGISTRADA",
  "Si ya existe: pregunta si es sucursal diferente. Si si: continua desde PASO 1 con nombre completo. Si no: Escribinos a info@generoytrabajo.com y te compartimos el diagnostico anterior.",
  "",
  "BLOQUE [DETALLE] — GENERARLO JUNTO AL BLOQUE [DATOS], INVISIBLE PARA EL USUARIO",
  "Este bloque va al final del ultimo mensaje, ANTES del [DATOS]. No se muestra en el chat. Contiene la evaluacion detallada de cada pregunta para el informe.",
  "Formato exacto:",
  "[DETALLE]",
  "=== P1 – DIAGNOSTICO PSICOSOCIAL (X/5) ===",
  "Lo que dijeron: [resumen de lo que dijo la empresa]",
  "Que cumple: [aspecto concreto que ya esta bien]",
  "Que falta o debe mejorar: [aspecto concreto que falta]",
  "Referencia NOM-035: [articulo o apartado relevante]",
  "",
  "=== P2 – ACCIONES POST-DIAGNOSTICO (X/5) ===",
  "Lo que dijeron: ...",
  "Que cumple: ...",
  "Que falta o debe mejorar: ...",
  "Referencia NOM-035: ...",
  "",
  "[repetir para P3 a P9 con sus titulos: P3 POLITICA DE PREVENCION, P4 EVENTOS TRAUMATICOS, P5 EXAMENES MEDICOS, P6 REGISTROS DOCUMENTALES, P7 SENALES Y CANAL, P8 LIDERAZGOS Y CAPACITACIONES, P9 PERSPECTIVA DE GENERO]",
  "",
  "=== ACCION PRIORITARIA ===",
  "[La medida mas urgente y concreta que debe tomar la empresa]",
  "[/DETALLE]",
  "",
  "BLOQUE DE DATOS — AL FINAL DEL ULTIMO MENSAJE (despues del [DETALLE])",
  '[DATOS]{"empresa":"x","contacto":"x","email":"x","email_informe":"x","tamano":"x","pais":"x","p1":"x","p2":"x","p3":"x","p4":"x","p5":"x","p6":"x","p7":"x","p8":"x","p9":"x","p1_nivel":0,"p2_nivel":0,"p3_nivel":0,"p4_nivel":0,"p5_nivel":0,"p6_nivel":0,"p7_nivel":0,"p8_nivel":0,"p9_nivel":0,"puntaje_norma":0,"puntaje_genero":0,"semaforo_norma":"ROJO","semaforo_genero":"ROJO","cumple":"x","falta":"x"}[/DATOS]',
].join("\n");


app.post("/api/check-empresa", async (req, res) => {
  const { empresa } = req.body;
  if (!empresa) return res.json({ existe: false });
  try { res.json({ existe: await checkEmpresaExiste(empresa) }); }
  catch (e) { console.error(e); res.json({ existe: false }); }
});

app.post("/api/chat", async (req, res) => {
  const { messages } = req.body;
  const msgs = (!messages || !Array.isArray(messages) || messages.length === 0)
    ? [{ role: "user", content: "Inicia el diagnostico." }]
    : messages;

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: "API key no configurada." });

  const userMessages = msgs.filter(m => m.role === "user");
  if (userMessages.length === 2) {
    const lastUserMsg = userMessages[userMessages.length - 1].content || "";
    const lines = lastUserMsg.split(/\n/).map(l => l.trim()).filter(Boolean);
    let empresaNombre = "";
    for (const line of lines) {
      const m = line.match(/(?:empresa|organizaci[oó]n|compa[nñ][ií]a)[:\s]+(.+)/i);
      if (m) { empresaNombre = m[1].trim(); break; }
    }
    if (!empresaNombre && lines[0] && lines[0].split(" ").length <= 8) empresaNombre = lines[0];
    if (empresaNombre) {
      const yaExiste = await checkEmpresaExiste(empresaNombre);
      if (yaExiste) {
        msgs.push({ role: "assistant", content: "[SISTEMA INTERNO: La empresa " + empresaNombre + " ya tiene diagnostico registrado. Aplica flujo EMPRESA YA REGISTRADA.]" });
        msgs.push({ role: "user", content: "(continuar)" });
      }
    }
  }

  const callAnthropic = async (retries, delayMs) => {
    if (retries === undefined) retries = 3;
    if (delayMs === undefined) delayMs = 2000;
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({ model: "claude-sonnet-4-5", max_tokens: 4000, system: SYSTEM_PROMPT, messages: msgs }),
    });
    if (response.status === 529 && retries > 0) {
      await new Promise(r => setTimeout(r, delayMs));
      return callAnthropic(retries - 1, delayMs * 2);
    }
    return response;
  };

  try {
    const response = await callAnthropic();
    if (!response.ok) {
      const err = await response.json();
      const status = response.status;
      const msg = status === 529
        ? "El servicio esta temporalmente saturado. Intenta de nuevo en unos segundos."
        : (err.error && err.error.message ? err.error.message : "Error de la API.");
      return res.status(status).json({ error: msg });
    }

    const data = await response.json();
    let reply = data.content.map(function(b) { return b.text || ""; }).join("");

    // Extraer [DETALLE] (va al Drive doc, no al chat)
    let detalleTexto = "";
    const detalleMatch = reply.match(/\[DETALLE\]([\s\S]*?)\[\/DETALLE\]/);
    if (detalleMatch) {
      detalleTexto = detalleMatch[1].trim();
      reply = reply.replace(/\[DETALLE\][\s\S]*?\[\/DETALLE\]/, "").trim();
      console.log("Bloque [DETALLE] extraido, longitud:", detalleTexto.length);
    }

    // Extraer [DATOS]
    const dataMatch = reply.match(/\[DATOS\]([\s\S]*?)\[\/DATOS\]/);
    console.log("Bloque [DATOS] encontrado:", !!dataMatch);
    if (!dataMatch) {
      console.log("=== REPLY SIN [DATOS] (primeros 500 chars) ===");
      console.log(reply.substring(0, 500));
      console.log("=== FIN ===");
    }

    if (dataMatch) {
      let parsed;
      try {
        parsed = JSON.parse(dataMatch[1]);
      } catch (jsonErr) {
        console.error("Error al parsear [DATOS] JSON:", jsonErr.message);
        console.error("JSON recibido:", dataMatch[1].substring(0, 300));
      }

      if (parsed) {
        const fecha = new Date().toLocaleString("es-MX", { timeZone: "America/Mexico_City" });
        const diagnosticoLimpio = reply.replace(/\[DATOS\][\s\S]*?\[\/DATOS\]/, "").trim();

        let docLink = "";
        try {
          const result = await createDriveDoc(parsed.empresa || "Empresa", diagnosticoLimpio, detalleTexto, parsed);
          docLink = result.docLink;
          console.log("Doc creado en Drive:", docLink);
        } catch (driveErr) {
          console.error("Error al crear Doc en Drive:", driveErr.message);
        }

        try {
          await appendToSheet([
            fecha,
            parsed.empresa       || "",
            parsed.contacto      || "",
            parsed.email         || "",
            parsed.email_informe || "",
            parsed.tamano        || "",
            parsed.pais          || "",
            parsed.p1_nivel      || "",
            parsed.p2_nivel      || "",
            parsed.p3_nivel      || "",
            parsed.p4_nivel      || "",
            parsed.p5_nivel      || "",
            parsed.p6_nivel      || "",
            parsed.p7_nivel      || "",
            parsed.p8_nivel      || "",
            parsed.p9_nivel      || "",
            parsed.puntaje_norma  || "",
            parsed.semaforo_norma || "",
            parsed.puntaje_genero || "",
            parsed.semaforo_genero|| "",
            docLink,
          ]);
          console.log("Datos guardados en Sheets OK");
        } catch (sheetsErr) {
          console.error("Error al guardar en Sheets — mensaje:", sheetsErr.message);
          console.error("Error al guardar en Sheets — codigo:", sheetsErr.code);
          console.error("Error al guardar en Sheets — status:", sheetsErr.status);
          if (sheetsErr.errors) console.error("Errores detallados:", JSON.stringify(sheetsErr.errors));
        }
      }
      reply = reply.replace(/\[DATOS\][\s\S]*?\[\/DATOS\]/, "").trim();
    }

    res.json({ reply });
  } catch (error) {
    console.error("Error al llamar a la API:", error.message);
    res.status(500).json({ error: "Error interno del servidor." });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, function() { console.log("Servidor corriendo en http://localhost:" + PORT); });
