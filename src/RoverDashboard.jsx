import { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { db, auth } from "./firebase";
import {
  doc, onSnapshot, setDoc, updateDoc, deleteField, getDoc
} from "firebase/firestore";
import {
  signInWithEmailAndPassword, signOut, onAuthStateChanged
} from "firebase/auth";

// ─── ETAPAS COMPLETAS (Cartilla 2.0 – solo 3 etapas) ─────────────────────────
const STAGES = [
  {
    id: 1, numeral: "01", color: "#D63031", bg: "#2A0A0A",
    name: "Aspirante Rover", subtitle: "Etapa de Encuentro",
    investidura: "ENTREGA DE PAÑOLETA",
    description: "Primer acercamiento al clan rover. Espacio para conocer, observar y decidir con libertad sin compromisos formales.",
    areas: ["Corporalidad","Creatividad","Carácter","Afectividad","Sociabilidad","Espiritualidad"],
    areaDesc: [
      "Participación activa en reuniones y actividades iniciales",
      "Comprender estructura y funcionamiento del Clan",
      "Mostrar compromiso inicial y disciplina",
      "Integrarse sanamente al grupo",
      "Convivir y colaborar con el Clan",
      "Reflexionar sobre el lema Rover y propósito personal"
    ],
    sections: [
      { name: "Vida Rover", area: "Sociabilidad", items: [
        "Participar en actividades de encuentro mínimo 1 mes",
        "Conocer el Clan y el Grupo",
        "Divisa scout y su historia"
      ]},
      { name: "Administrativo", area: "Carácter", items: [
        "Estar inscrito en la Asociación de Scouts de Guatemala",
        "Renas vigente",
        "Safe From Harm (SFH)"
      ]},
      { name: "Espíritu Scout", area: "Espiritualidad", items: [
        "Promesa scout",
        "Ley scout",
        "Lema rover",
        "Oración rover",
        "Principios y virtudes",
        "Saludo scout y su historia",
        "La obligación diaria (buena acción)",
        "Historia del grupo",
        "Partes de la flor de lis nacional",
        "Partes de la flor de lis mundial"
      ]}
    ]
  },
  {
    id: 2, numeral: "02", color: "#E67E22", bg: "#261800",
    name: "Heraldo Rover", subtitle: "Etapa de Reflexión",
    investidura: "INVESTIDURA DE BORLAS Y HERALDO ROVER",
    description: "Momento de mirar hacia adentro y entender qué significa ser scout. Construir identidad y criterio propio sin presión.",
    areas: ["Corporalidad","Creatividad","Carácter","Afectividad","Sociabilidad","Espiritualidad"],
    areaDesc: [
      "Participación constante en actividades básicas del Clan",
      "Comprender historia, símbolos y fundamentos del escultismo",
      "Vivir la Ley, Promesa y principios Scout",
      "Desarrollar empatía mediante servicio inicial",
      "Fortalecer trabajo en equipo y disciplina grupal",
      "Interiorizar mística Rover y obligación diaria"
    ],
    sections: [
      { name: "Vida Rover", area: "Sociabilidad", items: [
        "¿Qué es el movimiento scout?",
        "Misión y visión del movimiento scout",
        "Obligación diaria y su aplicación",
        "Historia del grupo y del Clan"
      ]},
      { name: "Escultismo", area: "Creatividad", items: [
        "Historia del Escultismo Internacional",
        "Uniforme y sus partes",
        "Biografía de Baden Powell",
        "Estructura del Clan Rover"
      ]},
      { name: "Civismo", area: "Carácter", items: [
        "Declamar Jura a la Bandera",
        "Entonar Himno Nacional correctamente"
      ]},
      { name: "Lectura Formativa", area: "Espiritualidad", items: [
        "Rover: lo que son y lo que hacen",
        "Escultismo para muchachos – fogatas 1 al 3",
        "Iniciar lectura de Roverismo hacia el éxito"
      ]}
    ]
  },
  {
    id: 3, numeral: "03", color: "#27AE60", bg: "#031409",
    name: "Escudero Rover", subtitle: "Etapa de Compromiso",
    investidura: "INVESTIDURA DE ESCUDERO Y LISTONES",
    description: "Asumir responsabilidades reales. Desarrollar liderazgo, trabajo en equipo y servicio comunitario activo.",
    areas: ["Corporalidad","Creatividad","Carácter","Afectividad","Sociabilidad","Espiritualidad"],
    areaDesc: [
      "Dominio sólido de técnicas de aire libre, seguridad y auto cuidado",
      "Elaboración de heráldica, CV, proyectos y soluciones innovadoras",
      "Liderazgo, responsabilidad, disciplina, toma de decisiones",
      "Madurez emocional, empatía, mentoría y relaciones sanas",
      "Servicio activo, ciudadanía responsable y participación comunitaria",
      "Propósito, identidad Rover, coherencia con la Ley y Promesa"
    ],
    sections: [
      { name: "Vida Rover", area: "Creatividad", items: [
        "Elaboración y explicación del escudo Rover (heráldica)",
        "Diseño y significado de la horquilla",
        "Elaboración de CV Scout y CV personal",
        "Participar en una excursión de tu clan"
      ]},
      { name: "Liderazgo y Servicio", area: "Sociabilidad", items: [
        "Formar a un miembro nuevo (unidad, caminante o rover)",
        "2 reuniones de cubil (con bitácora)",
        "2 reuniones de manada o un acantonamiento (con bitácora)",
        "2 reuniones de unidad o un campamento (con bitácora)",
        "2 reuniones de caminantes o un campamento (con bitácora)"
      ]},
      { name: "Civismo", area: "Carácter", items: [
        "Historia del himno nacional",
        "Historia de jura de la bandera",
        "Conocimiento de los símbolos patrios",
        "Correcto izado, arriado y doblado de banderas",
        "Geografía nacional: volcanes, ríos, lagos y lagunas principales",
        "Conocimiento de pueblos indígenas e idiomas"
      ]},
      { name: "Escultismo", area: "Creatividad", items: [
        "Jerarquía y organización del Grupo",
        "Ejecución correcta del saludo y seña en cada circunstancia"
      ]},
      { name: "Nudos", area: "Corporalidad", items: [
        "Margarita","Ballestrinque","As de guía","Rizo","Pescador",
        "Vuelta de escota","Media llave","Corredizo","Vuelta de braza",
        "Presilla de alondra","Fugitivo","Arnés de hombre",
        "As de guía doble","Ballestrinque doble","Zarpa de gato"
      ]},
      { name: "Amarres", area: "Corporalidad", items: [
        "Diagonal","Cuadrado","Redondo","Ocho"
      ]},
      { name: "Observación y Señalización", area: "Corporalidad", items: [
        "Claves básicas","Pistas","Pruebas de Kim",
        "Señalización con silbato","Formaciones"
      ]},
      { name: "Campamento – Instalaciones", area: "Corporalidad", items: [
        "Mesa","Portada","Cercado","Asta",
        "Cocineta","Basurero","Fogata","Puente de 1-2-3 hilos",
        "Traspaso en cable","Croquis del campamento"
      ]},
      { name: "Campamento – Habilidades", area: "Corporalidad", items: [
        "Partes, cuidado y manejo de navaja, cuchillo, machete y hacha",
        "Corte adecuado de leña",
        "Mochila básica de campamento",
        "Cuidados, levantado y reparaciones de carpa de campaña",
        "Armado, uso y prevenciones de 6 tipos de fogata",
        "Cocina individual – Menú básico (pollo, arroz, carne, sopa, postre y bebida)",
        "¿Qué es? y elaboración de un croquis"
      ]},
      { name: "Seguridad y Primeros Auxilios", area: "Afectividad", items: [
        "Tipos de botiquín y contenido",
        "Armado o re abastecimiento del botiquín del Clan",
        "Manejo de emergencias y traslado de heridos",
        "Seguridad vial y desastres naturales en Guatemala",
        "Números de emergencia nacional y regional",
        "Ejercicios B.P.",
        "Vendajes scout"
      ]},
      { name: "Medio Ambiente", area: "Espiritualidad", items: [
        "Flora y fauna nacional",
        "Reservas y áreas protegidas del país",
        "Uso y conservación de la naturaleza",
        "Visitar dos parques ecológicos y documentar propósito"
      ]},
      { name: "Competencias Digitales", area: "Creatividad", items: [
        "Hardware y software",
        "Microsoft Office básico",
        "Perfiles Gmail y Outlook",
        "Seguridad en línea y huella digital",
        "TIC aplicadas a la realidad profesional"
      ]},
      { name: "Cultura y Patrimonio", area: "Espiritualidad", items: [
        "Folklore: comidas, bailes, leyendas y tradiciones",
        "Centros arqueológicos de Guatemala",
        "Regiones culturales del país"
      ]},
      { name: "Caminata Rover", area: "Corporalidad", items: [
        "Caminata mínimo 10 km con croquis, claves, fotos y bitácora"
      ]}
    ]
  }
];

// ─── KEYS ─────────────────────────────────────────────────────────────────────
const pKey = (sid, stId, sec, idx) => `${sid}|${stId}|${sec}|${idx}`;
const dKey = (sid, stId, sec, idx) => `d:${sid}|${stId}|${sec}|${idx}`;
const invKey = (sid, stId) => `inv:${sid}|${stId}`;
const attKey = (sid, date) => `att:${sid}|${date}`;

// ── Stable keys — use stableIds, never positional index ──────────────────────
// Format: s:{scoutId}|{stageStableId}|{sectionStableId}|{objectiveStableId}
const pKeyS = (sid, stageStId, secStId, objStId) =>
  `s:${sid}|${stageStId}|${secStId}|${objStId}`;
const dKeyS = (sid, stageStId, secStId, objStId) =>
  `ds:${sid}|${stageStId}|${secStId}|${objStId}`;

// ── Read progress for ONE objective: stable key first, legacy fallback ────────
// item  = the item object from stagesConfig (has stableId, legacyIdx, legacySecName)
// stage = the stage object (has stableId, id)
// sec   = the section object (has stableId, name)
// prog/dates = the full maps from Firebase
const getItemProg = (scoutId, stage, sec, item, prog) => {
  // 1. Try stable key
  if (item && item.stableId && stage.stableId && sec.stableId) {
    const sk = pKeyS(scoutId, stage.stableId, sec.stableId, item.stableId);
    if (prog[sk] !== undefined) return !!prog[sk];
  }
  // 2. Legacy fallback using legacyIdx (original position, never changes)
  const legacyIdx = item && item.legacyIdx !== undefined ? item.legacyIdx : null;
  const legacySecName = (item && item.legacySecName) || sec.name;
  if (legacyIdx !== null) {
    const lk = pKey(scoutId, stage.id, legacySecName, legacyIdx);
    if (prog[lk] !== undefined) return !!prog[lk];
  }
  return false;
};

const getItemDate = (scoutId, stage, sec, item, dates) => {
  if (item && item.stableId && stage.stableId && sec.stableId) {
    const sk = dKeyS(scoutId, stage.stableId, sec.stableId, item.stableId);
    if (dates[sk]) return dates[sk];
  }
  const legacyIdx = item && item.legacyIdx !== undefined ? item.legacyIdx : null;
  const legacySecName = (item && item.legacySecName) || sec.name;
  if (legacyIdx !== null) {
    const lk = dKey(scoutId, stage.id, legacySecName, legacyIdx);
    if (dates[lk]) return dates[lk];
  }
  return "";
};

// ─── ASISTENCIA: ESTADOS ──────────────────────────────────────────────────────
const ATT_STATES = {
  presente: { label: "Presente", short: "P", color: "#27AE60" },
  ausente:  { label: "Ausente",  short: "A", color: "#D63031" },
  excusa:   { label: "Excusa",   short: "E", color: "#E67E22" }
};
const ATT_CYCLE = ["presente", "ausente", "excusa"];

// ─── DESIGN SYSTEM TOKENS (Fase 6) ──────────────────────────────────────────
const DS = {
  bg:{ root:"#0D0E10", surface:"#17191D", card:"#17191D", cardAlt:"#1E2025",
    overlay:"rgba(8,9,11,0.88)", input:"rgba(255,255,255,0.05)" },
  brand:{ red:"#6F1725", redDeep:"#4A0F18", redLight:"#8F2235", gold:"#C5A253",
    goldMuted:"#7A6235", goldFaint:"#2A2218", black:"#0D0E10" },
  state:{ done:"#27AE60", doneDeep:"#1E8449", inprogress:"#E67E22", pending:"rgba(255,255,255,0.18)",
    info:"#5B8FA8", infoLight:"#7FB3D5", error:"#C0392B", errorLight:"#E07A6E", warning:"#E67E22" },
  text:{ primary:"#E6E7EA", secondary:"#9099AA", tertiary:"#555F6E", muted:"#353C47",
    inverse:"#0D0E10" },
  border:{ faint:"rgba(255,255,255,0.04)", subtle:"rgba(255,255,255,0.07)",
    default:"rgba(255,255,255,0.1)", strong:"rgba(255,255,255,0.18)",
    gold:"rgba(197,162,83,0.15)", goldStrong:"rgba(197,162,83,0.3)" },
  radius:{ sm:"6px", md:"9px", lg:"12px", xl:"14px", full:"9999px" },
  space:{ xs:"4px", sm:"8px", md:"12px", lg:"16px", xl:"24px" },
  shadow:{ sm:"0 2px 8px rgba(0,0,0,0.4)", md:"0 4px 16px rgba(0,0,0,0.5)",
    lg:"0 8px 28px rgba(0,0,0,0.6)", gold:"0 0 16px rgba(197,162,83,0.1)" },
  transition:{ fast:"0.12s ease", normal:"0.18s ease", slow:"0.25s ease",
    prog:"width 0.25s ease, background 0.25s ease" },
  font:{ display:"'Cinzel', serif", body:"'Source Sans 3', sans-serif",
    szXs:"9px", szSm:"10px", szMd:"12px", szBase:"13px", szLg:"15px",
    szXl:"18px", szH1:"22px", szH2:"17px" },
};

// ─── EXCEL EXPORT (RF-04) ────────────────────────────────────────────────────
// ─── REAL XLSX via SheetJS npm package (OOXML — real ZIP-based .xlsx) ──────
// Uses xlsx installed as npm dependency — always generates real OOXML, no fallback
function exportXLSX(scouts, prog, dates, attendance, activeStagesList, priorityIds) {
  const today = new Date().toISOString().split("T")[0];
  const stages = (activeStagesList && activeStagesList.length > 0) ? activeStagesList : STAGES;

  // ── Sheet 1: Resumen ──────────────────────────────────────────────────────
  const summary = [
    ["Rover","Clan / Grupo","Etapa Actual","% Etapa Actual","% Total",
     "Objetivos Completados","Objetivos Totales","Última Actividad","Prioritario","Completado"],
  ];
  scouts.forEach(sc => {
    const cs   = stages.find(s => scoutStageStats(sc, s, prog).pct < 100 &&
                               scoutStageStats(sc, s, prog).d > 0) || stages[stages.length-1];
    const csSt = scoutStageStats(sc, cs, prog);
    let totD = 0, totT = 0;
    stages.forEach(st => { const s = scoutStageStats(sc, st, prog); totD += s.d; totT += s.t; });
    const totPct = totT ? Math.round(totD / totT * 100) : 0;
    const allDates = Object.entries(dates)
      .filter(([k]) => k.startsWith(`d:${sc.id}|`))
      .map(([, v]) => v).filter(Boolean).sort().reverse();
    summary.push([
      sc.name, sc.group||"", cs.name,
      csSt.pct, totPct, totD, totT,
      allDates[0]||"", priorityIds.has(sc.id)?"Sí":"No", totPct===100?"Sí":"No"
    ]);
  });

  // ── Sheet 2: Progresión ───────────────────────────────────────────────────
  const detail = [
    ["Rover","Clan","Etapa","Sección","Área","Objetivo","Completado","Fecha Cumplimiento"],
  ];
  scouts.forEach(sc => {
    stages.forEach(st => {
      (st.sections||[]).forEach(sec => {
        // Fix F4: filter active:false before iterating
        const secItems = (sec.items||[]).filter(it => it.active !== false);
        secItems.forEach((rawItem) => {
          // Fix F1: keep original item object for stable key lookup
          const itemObj = typeof rawItem === "object" ? rawItem : null;
          const itemText = itemObj ? (itemObj.text || String(rawItem)) : String(rawItem);
          // Stable key first (pKeyS), legacy fallback via legacyIdx
          let done = false, date = "";
          if (itemObj && itemObj.stableId && st.stableId && sec.stableId) {
            const sk = pKeyS(sc.id, st.stableId, sec.stableId, itemObj.stableId);
            const li = itemObj.legacyIdx !== undefined ? itemObj.legacyIdx : null;
            const ln = itemObj.legacySecName || sec.name;
            done = !!(prog[sk] !== undefined ? prog[sk]
              : (li !== null ? prog[pKey(sc.id, st.id, ln, li)] : false));
            date = dates[dKeyS(sc.id, st.stableId, sec.stableId, itemObj.stableId)]
              || (li !== null ? dates[dKey(sc.id, st.id, ln, li)] : "") || "";
          } else {
            // Pure legacy (string items or missing stableId)
            const idx = sec.items.indexOf(rawItem);
            done = !!prog[pKey(sc.id, st.id, sec.name, idx)];
            date = dates[dKey(sc.id, st.id, sec.name, idx)] || "";
          }
          detail.push([sc.name, sc.group||"", st.name, sec.name, sec.area,
                        itemText, done ? "Sí" : "No", date]);
        });
      });
    });
  });

  // ── Sheet 3: Asistencia ───────────────────────────────────────────────────
  const att = [["Rover","Clan","Fecha","Estado","Observaciones"]];
  scouts.forEach(sc => {
    Object.entries(attendance||{})
      .filter(([k]) => k.startsWith(`att:${sc.id}|`))
      .map(([k, v]) => ({ date: k.split("|")[1], status: v.status, note: v.note||"" }))
      .sort((a, b) => a.date.localeCompare(b.date))
      .forEach(r => att.push([sc.name, sc.group||"", r.date, r.status, r.note]));
  });

  // ── Generate real OOXML .xlsx (ZIP-based binary) ─────────────────────────
  const wb = XLSX.utils.book_new();

  const ws1 = XLSX.utils.aoa_to_sheet(summary);
  ws1["!cols"] = [18,16,18,12,10,14,12,16,10,10].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws1, "Resumen");

  const ws2 = XLSX.utils.aoa_to_sheet(detail);
  ws2["!cols"] = [18,14,18,18,14,40,10,16].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, "Progresión");

  const ws3 = XLSX.utils.aoa_to_sheet(att);
  ws3["!cols"] = [18,14,12,12,24].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws3, "Asistencia");

  // writeFile uses FileSaver internally for browser download
  XLSX.writeFile(wb, `Progresion_Clan_${today}.xlsx`, { bookType: "xlsx" });
}


// ─── PDF REAL — jsPDF from CDN ──────────────────────────────────────────────
// Loads jsPDF at runtime from CDN — no npm needed
async function loadJsPDF() {
  if (window._jsPDF) return window._jsPDF;
  const mod = await import("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js")
    .catch(() => null);
  if (mod && (mod.jsPDF || (mod.default && mod.default.jsPDF))) {
    window._jsPDF = mod.jsPDF || mod.default.jsPDF;
    return window._jsPDF;
  }
  // Fallback: script tag injection
  return new Promise((resolve, reject) => {
    if (window.jspdf) { resolve(window.jspdf.jsPDF); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    s.onload = () => { window._jsPDF = window.jspdf?.jsPDF; resolve(window._jsPDF); };
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// RF-09: PDF Individual — ficha directa sin dialog
async function exportPDFFicha(sc, prog, dates, attendance, stages, onProgress) {
  if (onProgress) onProgress("Cargando generador PDF…");
  const today = new Date().toLocaleDateString("es-GT", { day:"2-digit", month:"long", year:"numeric" });
  const todayISO = new Date().toISOString().split("T")[0];

  try {
    const JsPDF = await loadJsPDF();
    const doc = new JsPDF({ orientation:"portrait", unit:"mm", format:"a4" });
    const W = 210, margin = 14;
    let y = margin;

    // ── Header ──
    doc.setFillColor(107, 15, 15);
    doc.rect(0, 0, W, 28, "F");
    doc.setTextColor(201, 162, 39);
    doc.setFontSize(7);
    doc.text("GRUPO SCOUT No. 12 · SAN PEDRO APÓSTOL · GUATEMALA", margin, 8);
    doc.setTextColor(250, 246, 240);
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.text("Ficha Oficial de Progresión Rover", margin, 18);
    doc.setFontSize(7);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(180, 160, 130);
    doc.text(`Emitida el ${today}`, W - margin - 30, 18);
    y = 36;

    // ── Scout name + stats ──
    doc.setTextColor(36, 16, 8);
    doc.setFontSize(13); doc.setFont("helvetica","bold");
    doc.text(sc.name||"(Sin nombre)", margin, y); y += 6;
    if (sc.group) { doc.setFontSize(9); doc.setFont("helvetica","normal"); doc.text(sc.group, margin, y); y += 5; }

    // Total progress
    let totD = 0, totT = 0;
    stages.forEach(st => { const s = scoutStageStats(sc, st, prog); totD += s.d; totT += s.t; });
    const totPct = totT ? Math.round(totD / totT * 100) : 0;
    doc.setFontSize(10); doc.setFont("helvetica","bold");
    doc.setTextColor(139, 26, 26);
    doc.text(`Progreso total: ${totPct}% (${totD}/${totT} objetivos)`, margin, y); y += 8;

    // ── Stage bars ──
    stages.forEach(st => {
      const s = scoutStageStats(sc, st, prog);
      if (y > 260) { doc.addPage(); y = margin; }

      doc.setFillColor(230, 220, 210);
      doc.rect(margin, y, W - margin*2, 4, "F");
      const barW = Math.round((W - margin*2) * s.pct / 100);
      doc.setFillColor(s.pct===100 ? 39 : 139, s.pct===100 ? 174 : 26, s.pct===100 ? 96 : 26);
      if (barW > 0) doc.rect(margin, y, barW, 4, "F");
      y += 7;

      doc.setFontSize(9); doc.setFont("helvetica","bold");
      doc.setTextColor(36, 16, 8);
      doc.text(`${st.numeral} ${st.name}  ${s.pct}%  (${s.d}/${s.t})`, margin, y); y += 5;

      // Objectives
      (st.sections||[]).forEach(sec => {
        // Fix F4: filter active:false; Fix F3: stable key lookup
        const secItemsPDF = (sec.items||[]).filter(it => it.active !== false);
        secItemsPDF.forEach((rawItem) => {
          const itemObj = typeof rawItem === "object" ? rawItem : null;
          const item = itemObj ? (itemObj.text || String(rawItem)) : String(rawItem);
          if (y > 272) { doc.addPage(); y = margin; }
          let done = false, date = "";
          if (itemObj && itemObj.stableId && st.stableId && sec.stableId) {
            const sk = pKeyS(sc.id, st.stableId, sec.stableId, itemObj.stableId);
            const li = itemObj.legacyIdx !== undefined ? itemObj.legacyIdx : null;
            const ln = itemObj.legacySecName || sec.name;
            done = !!(prog[sk] !== undefined ? prog[sk]
              : (li !== null ? prog[pKey(sc.id, st.id, ln, li)] : false));
            date = dates[dKeyS(sc.id, st.stableId, sec.stableId, itemObj.stableId)]
              || (li !== null ? dates[dKey(sc.id, st.id, ln, li)] : "") || "";
          } else {
            const idx = sec.items.indexOf(rawItem);
            done = !!prog[pKey(sc.id, st.id, sec.name, idx)];
            date = dates[dKey(sc.id, st.id, sec.name, idx)] || "";
          }
          doc.setFontSize(7.5); doc.setFont("helvetica", done?"normal":"normal");
          doc.setTextColor(done ? 39 : 100, done ? 120 : 80, done ? 60 : 70);
          const mark = done ? "✓" : "○";
          const label = `${mark} ${item.length>80?item.slice(0,80)+"…":item}`;
          doc.text(label, margin + 2, y);
          if (done && date) { doc.setTextColor(150,130,110); doc.text(date, W-margin-18, y); }
          y += 4.5;
        });
      });
      y += 3;
    });

    // ── Signature area ──
    if (y > 250) { doc.addPage(); y = margin; }
    y += 6;
    doc.setDrawColor(36,16,8); doc.setLineWidth(0.3);
    const sigY = y + 8;
    doc.line(margin, sigY, margin+55, sigY);
    doc.line(margin+65, sigY, margin+120, sigY);
    doc.line(margin+130, sigY, margin+180, sigY);
    doc.setFontSize(7); doc.setTextColor(80,60,40);
    doc.text("Firma del Rover", margin, sigY+4);
    doc.text("Firma del Jefe de Clan", margin+65, sigY+4);
    doc.text("Fecha", margin+130, sigY+4);

    // ── Footer ──
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      doc.setFontSize(6.5); doc.setTextColor(180,160,130);
      doc.text(`Grupo Scout No. 12 · San Pedro Apóstol · Guatemala · Cartilla 2.0`, margin, 292);
      doc.text(`${i}/${pageCount}`, W-margin, 292, { align:"right" });
    }

    doc.save(`Ficha_${sc.name.replace(/\s+/g,"_")}_${todayISO}.pdf`);
    if (onProgress) onProgress(null);
  } catch (err) {
    console.error("PDF error:", err);
    if (onProgress) onProgress(null);
    throw err;
  }
}

// RF-09: PDF Reporte General
async function exportPDFGeneral(scouts, prog, dates, attendance, stages, priorityIds, onProgress) {
  if (onProgress) onProgress("Cargando generador PDF…");
  const today = new Date().toLocaleDateString("es-GT", { day:"2-digit", month:"long", year:"numeric" });
  const todayISO = new Date().toISOString().split("T")[0];
  const stageList = (stages && stages.length > 0) ? stages : STAGES;

  try {
    const JsPDF = await loadJsPDF();
    const doc = new JsPDF({ orientation:"portrait", unit:"mm", format:"a4" });
    const W = 210, margin = 14;
    let y = margin;

    const addHeader = (pageTitle, sub) => {
      doc.setFillColor(107,15,15);
      doc.rect(0, 0, W, 26, "F");
      doc.setTextColor(201,162,39); doc.setFontSize(7);
      doc.text("GRUPO SCOUT No. 12 · SAN PEDRO APÓSTOL · GUATEMALA", margin, 8);
      doc.setTextColor(250,246,240); doc.setFontSize(13); doc.setFont("helvetica","bold");
      doc.text(pageTitle, margin, 17);
      if (sub) { doc.setFontSize(8); doc.setFont("helvetica","normal"); doc.setTextColor(200,180,150); doc.text(sub, margin, 23); }
      doc.setFontSize(7); doc.setTextColor(180,160,130);
      doc.text(`Emitido: ${today}`, W-margin-28, 17);
    };

    addHeader("Reporte General del Clan", `${scouts.length} Rover${scouts.length===1?"":"s"}`);
    y = 34;

    // ── Summary statistics ──
    let clanD = 0, clanT = 0;
    scouts.forEach(sc => stageList.forEach(st => { const s = scoutStageStats(sc,st,prog); clanD+=s.d; clanT+=s.t; }));
    const avgPct = scouts.length ? Math.round(scouts.reduce((a,sc) => {
      let d=0,t=0; stageList.forEach(st=>{const s=scoutStageStats(sc,st,prog);d+=s.d;t+=s.t;});
      return a + (t?Math.round(d/t*100):0);
    }, 0) / scouts.length) : 0;

    doc.setFillColor(245,240,235);
    doc.rect(margin, y, W-margin*2, 18, "F");
    doc.setFontSize(8); doc.setFont("helvetica","bold"); doc.setTextColor(139,26,26);
    const statsLine = `Rovers: ${scouts.length}  |  Promedio: ${avgPct}%  |  Completados: ${clanD}/${clanT} objetivos`;
    doc.text(statsLine, margin+2, y+7); y += 22;

    // ── Rover list ──
    scouts.forEach((sc, ri) => {
      let totD=0, totT=0;
      stageList.forEach(st => { const s=scoutStageStats(sc,st,prog); totD+=s.d; totT+=s.t; });
      const totPct = totT ? Math.round(totD/totT*100) : 0;
      const cs = stageList.find(s=>scoutStageStats(sc,s,prog).pct<100 && scoutStageStats(sc,s,prog).d>0) || stageList[stageList.length-1];
      const allDates = Object.entries(dates).filter(([k])=>k.startsWith(`d:${sc.id}|`)).map(([,v])=>v).filter(Boolean).sort().reverse();

      if (y > 265) {
        doc.addPage();
        addHeader("Reporte General del Clan (continuación)", "");
        y = 34;
      }

      // Row background alternating
      doc.setFillColor(ri%2===0 ? 252 : 246, ri%2===0 ? 248 : 243, ri%2===0 ? 245 : 238);
      doc.rect(margin, y-4, W-margin*2, 14, "F");

      // Name + stage
      doc.setFontSize(9); doc.setFont("helvetica","bold"); doc.setTextColor(36,16,8);
      doc.text(sc.name, margin+1, y+1);
      doc.setFontSize(7.5); doc.setFont("helvetica","normal"); doc.setTextColor(139,26,26);
      doc.text(cs.name, margin+1, y+6);

      // Progress bar
      const barX = margin+60, barW2 = 70, barH = 4;
      doc.setFillColor(220,210,200); doc.rect(barX, y, barW2, barH, "F");
      doc.setFillColor(totPct===100?39:139, totPct===100?174:26, totPct===100?96:26);
      if (totPct > 0) doc.rect(barX, y, Math.round(barW2*totPct/100), barH, "F");
      doc.setFontSize(8); doc.setFont("helvetica","bold"); doc.setTextColor(36,16,8);
      doc.text(`${totPct}%`, barX+barW2+2, y+4);

      // Last activity + priority
      doc.setFontSize(6.5); doc.setFont("helvetica","normal"); doc.setTextColor(140,120,100);
      doc.text(`${allDates[0]||"—"}`, W-margin-28, y+1);
      if (priorityIds && priorityIds.has(sc.id)) {
        doc.setTextColor(201,162,39); doc.text("★ Prioritario", W-margin-28, y+5);
      }
      y += 16;
    });

    // ── Footer on all pages ──
    const pageCount = doc.internal.getNumberOfPages();
    for (let i=1; i<=pageCount; i++) {
      doc.setPage(i);
      doc.setFontSize(6.5); doc.setTextColor(180,160,130);
      doc.text("Grupo Scout No. 12 · San Pedro Apóstol · Guatemala · Cartilla 2.0", margin, 292);
      doc.text(`${i}/${pageCount}`, W-margin, 292, { align:"right" });
    }

    doc.save(`ReporteGeneral_Clan_${todayISO}.pdf`);
    if (onProgress) onProgress(null);
  } catch (err) {
    console.error("PDF general error:", err);
    if (onProgress) onProgress(null);
    throw err;
  }
}

// ─── BACKUP GENERATOR (RF-08) ────────────────────────────────────────────────
function generateBackup(data) {
  const backup = {
    version: "2.0",
    format: "rover-dashboard-backup",
    generatedAt: new Date().toISOString(),
    schema: {
      scouts: "array of {id, name, group}",
      prog: "map of pKey→true",
      dates: "map of dKey→YYYY-MM-DD",
      invDates: "map of invKey→YYYY-MM-DD",
      attendance: "map of attKey→{status,note}",
      stagesConfig: "Firebase stages_config document",
      sysConfig: "institution config"
    },
    data: {
      scouts:      data.scouts      || [],
      prog:        data.prog        || {},
      dates:       data.dates       || {},
      invDates:    data.invDates    || {},
      attendance:  data.attendance  || {},
      stagesConfig:data.stagesConfig|| {},
      sysConfig:   data.sysConfig   || {},
    }
  };
  const blob = new Blob([JSON.stringify(backup, null, 2)],
    { type:"application/json" });
  const a = document.createElement("a");
  const today = new Date().toISOString().split("T")[0];
  a.href = URL.createObjectURL(blob);
  a.download = `Backup_RoverDashboard_${today}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  return backup;
}

// ─── BACKUP VALIDATOR (RF-09) ────────────────────────────────────────────────
function validateBackup(parsed) {
  if (!parsed || typeof parsed !== "object") return { valid:false, error:"Formato inválido" };
  if (parsed.format !== "rover-dashboard-backup") return { valid:false, error:"No es un backup de este sistema" };
  if (!parsed.version) return { valid:false, error:"Versión no encontrada" };
  if (!parsed.data) return { valid:false, error:"Sección de datos ausente" };
  const d = parsed.data;
  // Validate structure integrity
  if (!Array.isArray(d.scouts)) return { valid:false, error:"Estructura de Rovers inválida" };
  if (typeof (d.prog||{}) !== "object") return { valid:false, error:"Estructura de progresión inválida" };
  // Check for corrupted scouts
  const corruptScouts = (d.scouts||[]).filter(s => !s.id || !s.name);
  if (corruptScouts.length > 0) return {
    valid:false, error:`${corruptScouts.length} Rover(s) sin ID o nombre válido`
  };
  return {
    valid: true,
    version: parsed.version,
    generatedAt: parsed.generatedAt,
    summary: {
      scouts:    (d.scouts||[]).length,
      stages:    Object.keys(d.stagesConfig||{}).length,
      progKeys:  Object.keys(d.prog||{}).length,
      dateKeys:  Object.keys(d.dates||{}).length,
      attKeys:   Object.keys(d.attendance||{}).length,
    }
  };
}

// ── Deep 3-level merge of stagesConfig by stableId ───────────────────────────
// priority = "cur"  → current wins scalar conflicts
// priority = "bk"   → backup wins scalar conflicts
// Rules:
//   Stage/Section/Item only in one source → always kept
//   Stage/Section/Item in both → winner scalars overwrite; loser-only children added at end
//   order → rebuilt as array position (mirrors reorderSection / reorderObjective)
//   legacyIdx, legacySecName, createdAt on items → NEVER overwritten (pKey identity)
function mergeStagesConfig(curStages, bkStages, priority) {
  const allSids = new Set([...Object.keys(curStages), ...Object.keys(bkStages)]);
  const result = {};
  allSids.forEach(sid => {
    const cur = curStages[sid];
    const bk  = bkStages[sid];
    if (cur && !bk) { result[sid] = cur; return; }
    if (bk && !cur) { result[sid] = bk;  return; }
    // Both exist — merge
    const winner = priority === "cur" ? cur : bk;
    const loser  = priority === "cur" ? bk  : cur;
    result[sid] = {
      ...loser,
      ...winner,
      sections: _mergeSections(cur.sections || [], bk.sections || [], priority),
    };
  });
  return result;
}
function _mergeSections(curSecs, bkSecs, priority) {
  const curMap = Object.fromEntries(curSecs.filter(s => s.stableId).map(s => [s.stableId, s]));
  const bkMap  = Object.fromEntries(bkSecs.filter(s => s.stableId).map(s => [s.stableId, s]));
  const seen = new Set();
  const ordered = [];
  // Iterate winner side first to preserve its order; append loser-only at end
  const winnerSecs = priority === "cur" ? curSecs : bkSecs;
  const loserSecs  = priority === "cur" ? bkSecs  : curSecs;
  winnerSecs.forEach(ws => {
    if (!ws.stableId) return;
    seen.add(ws.stableId);
    const curSec = curMap[ws.stableId];
    const bkSec  = bkMap[ws.stableId];
    if (!curSec || !bkSec) { ordered.push(ws); }
    else {
      // ws is always the winning side; use explicit curSec/bkSec to avoid aliasing errors
      ordered.push({
        ...(priority === "cur" ? bkSec  : curSec),  // loser first
        ...(priority === "cur" ? curSec : bkSec),   // winner overwrites scalars
        items: _mergeItems(curSec.items || [], bkSec.items || [], priority),
      });
    }
  });
  // Append loser-only sections at end
  loserSecs.forEach(ls => {
    if (ls.stableId && !seen.has(ls.stableId)) ordered.push(ls);
  });
  // Recalculate order as array position (mirrors reorderSection)
  return ordered.map((s, i) => ({ ...s, order: i }));
}
function _mergeItems(curItems, bkItems, priority) {
  const curMap = Object.fromEntries(
    curItems.filter(it => it && it.stableId).map(it => [it.stableId, it])
  );
  const bkMap  = Object.fromEntries(
    bkItems.filter(it => it && it.stableId).map(it => [it.stableId, it])
  );
  const seen = new Set();
  const ordered = [];
  // Iterate winner side first to preserve its order; append loser-only at end
  const winnerItems = priority === "cur" ? curItems : bkItems;
  const loserItems  = priority === "cur" ? bkItems  : curItems;
  winnerItems.forEach(wi => {
    if (!wi || !wi.stableId) return;
    seen.add(wi.stableId);
    const curItem = curMap[wi.stableId];
    const bkItem  = bkMap[wi.stableId];
    if (!curItem || !bkItem) { ordered.push(wi); }
    else {
      // Use explicit curItem/bkItem to avoid aliasing errors (wi = winner side, not always cur)
      const merged = {
        ...(priority === "cur" ? bkItem  : curItem),  // loser first
        ...(priority === "cur" ? curItem : bkItem),   // winner overwrites scalars
      };
      // legacyIdx, legacySecName, createdAt: always from current (pKey identity — never change)
      for (const f of ["legacyIdx", "legacySecName", "createdAt"]) {
        if (curItem[f] != null) merged[f] = curItem[f];
      }
      ordered.push(merged);
    }
  });
  // Append loser-only items at end
  loserItems.forEach(li => {
    if (li && li.stableId && !seen.has(li.stableId)) ordered.push(li);
  });
  // Recalculate order as array position (mirrors reorderObjective)
  return ordered.map((it, i) => ({ ...it, order: i }));
}

// ── Full conflict analysis function (pure, no Firebase writes) ────────────────
function analyzeConflicts(parsed, current, currentStagesConfig) {
  const bk = parsed.data;
  const conflicts = [];   // real conflicts — require admin decision
  const autoMerge = [];   // safe to merge automatically
  const identical = [];   // duplicates — skip

  // ── ROVERS ──────────────────────────────────────────────────────────────────
  const curScoutMap = Object.fromEntries((current.scouts||[]).map(s => [s.id, s]));
  (bk.scouts||[]).forEach(bkSc => {
    const cur = curScoutMap[bkSc.id];
    if (!cur) {
      autoMerge.push({ type:"rover", category:"new", elementId:bkSc.id,
        label:`Rover nuevo: ${bkSc.name}`, impact:"Se agregará" });
    } else if (JSON.stringify(cur) === JSON.stringify(bkSc)) {
      identical.push({ type:"rover", elementId:bkSc.id, label:`Rover idéntico: ${bkSc.name}` });
    } else {
      conflicts.push({ type:"rover", category:"data_diff", elementId:bkSc.id,
        label:`Rover con datos diferentes: ${bkSc.name}`,
        current: JSON.stringify(cur), backup: JSON.stringify(bkSc),
        impact:"Sobrescribiría datos del Rover actual" });
    }
  });

  // ── PROGRESIÓN ──────────────────────────────────────────────────────────────
  const curProg = current.prog || {};
  let newProg = 0, identProg = 0, conflProg = 0;
  Object.entries(bk.prog||{}).forEach(([k, bkVal]) => {
    const curVal = curProg[k];
    if (curVal === undefined) newProg++;
    else if (curVal === bkVal) identProg++;
    else {
      conflProg++;
      if (conflProg <= 5) { // show up to 5 examples
        conflicts.push({ type:"prog", category:"value_diff", elementId:k,
          label:`Progresión diferente: clave ${k.slice(0,40)}`,
          current:String(curVal), backup:String(bkVal),
          impact:"Cambiaría estado completado/pendiente de un objetivo" });
      }
    }
  });
  if (newProg > 0) autoMerge.push({ type:"prog", category:"new",
    label:`${newProg} registros de progresión nuevos`, impact:"Se agregarán" });
  if (identProg > 0) identical.push({ type:"prog",
    label:`${identProg} registros de progresión idénticos` });
  if (conflProg > 5) conflicts.push({ type:"prog", category:"value_diff",
    label:`...y ${conflProg-5} conflictos de progresión adicionales`,
    impact:"Requieren decisión" });

  // ── FECHAS ──────────────────────────────────────────────────────────────────
  const curDates = current.dates || {};
  let newDates = 0, conflDates = 0;
  Object.entries(bk.dates||{}).forEach(([k, bkVal]) => {
    const curVal = curDates[k];
    if (!curVal) newDates++;
    else if (curVal !== bkVal) { conflDates++;
      if (conflDates <= 3) conflicts.push({ type:"dates", category:"date_diff", elementId:k,
        label:`Fecha diferente: ${k.slice(0,40)}`,
        current:curVal, backup:bkVal, impact:"Cambiaría fecha de cumplimiento" });
    }
  });
  if (newDates > 0) autoMerge.push({ type:"dates", category:"new",
    label:`${newDates} fechas de cumplimiento nuevas`, impact:"Se agregarán" });
  if (conflDates > 3) conflicts.push({ type:"dates", category:"date_diff",
    label:`...y ${conflDates-3} conflictos de fechas adicionales`, impact:"Requieren decisión" });

  // ── ASISTENCIA ──────────────────────────────────────────────────────────────
  const curAtt = current.attendance || {};
  let newAtt = 0, conflAtt = 0;
  Object.entries(bk.attendance||{}).forEach(([k, bkVal]) => {
    const curVal = curAtt[k];
    if (!curVal) newAtt++;
    else if (JSON.stringify(curVal) !== JSON.stringify(bkVal)) {
      conflAtt++;
      if (conflAtt <= 3) conflicts.push({ type:"attendance", category:"att_diff", elementId:k,
        label:`Asistencia diferente: ${k.slice(0,40)}`,
        current:JSON.stringify(curVal), backup:JSON.stringify(bkVal),
        impact:"Cambiaría estado de asistencia" });
    }
  });
  if (newAtt > 0) autoMerge.push({ type:"attendance", category:"new",
    label:`${newAtt} registros de asistencia nuevos`, impact:"Se agregarán" });

  // ── ETAPAS / SECCIONES / OBJETIVOS ─────────────────────────────────────────
  const curStages = currentStagesConfig || {};
  Object.entries(bk.stagesConfig||{}).forEach(([sid, bkStage]) => {
    const curStage = curStages[sid];
    if (!curStage) {
      autoMerge.push({ type:"stage", category:"new", elementId:sid,
        label:`Etapa nueva: ${bkStage.name||sid}`, impact:"Se agregarán la etapa y sus secciones" });
      return;
    }
    // Compare stage name/subtitle
    if (curStage.name !== bkStage.name || curStage.subtitle !== bkStage.subtitle) {
      conflicts.push({ type:"stage", category:"content_diff", elementId:sid,
        label:`Etapa con nombre diferente: ${sid}`,
        current:curStage.name, backup:bkStage.name||sid, impact:"Cambiaría nombre visible de la etapa" });
    }
    // Compare sections
    const curSecMap = Object.fromEntries((curStage.sections||[]).map(s=>[s.stableId||s.name, s]));
    (bkStage.sections||[]).forEach(bkSec => {
      const secKey = bkSec.stableId || bkSec.name;
      const curSec = curSecMap[secKey];
      if (!curSec) {
        autoMerge.push({ type:"section", category:"new", elementId:secKey,
          label:`Sección nueva: ${bkSec.name} (${sid})`, impact:"Se agregará la sección" });
        return;
      }
      if (curSec.name !== bkSec.name) {
        conflicts.push({ type:"section", category:"content_diff", elementId:secKey,
          label:`Sección con nombre diferente: ${bkSec.name}`,
          current:curSec.name, backup:bkSec.name, impact:"Cambiaría nombre de la sección" });
      }
      // Compare objectives
      const curItemMap = Object.fromEntries((curSec.items||[])
        .filter(i=>typeof i==="object" && i.stableId)
        .map(i=>[i.stableId, i]));
      (bkSec.items||[]).forEach(bkItem => {
        if (typeof bkItem === "string") return; // legacy — skip
        const curItem = curItemMap[bkItem.stableId];
        if (!curItem) {
          autoMerge.push({ type:"objective", category:"new", elementId:bkItem.stableId,
            label:`Objetivo nuevo: "${(bkItem.text||"").slice(0,40)}"`, impact:"Se agregará" });
        } else if (curItem.text !== bkItem.text) {
          conflicts.push({ type:"objective", category:"content_diff", elementId:bkItem.stableId,
            label:`Objetivo con texto diferente`,
            current:(curItem.text||"").slice(0,50), backup:(bkItem.text||"").slice(0,50),
            impact:"Cambiaría texto del objetivo" });
        }
      });
    });
  });

  return { conflicts, autoMerge, identical };
}

// ─── NORMALIZE stableIds (RF-02: migration for data without stableId) ───────────
// Called after loading stagesConfig from Firebase.
// Guarantees every stage, section, and objective has a permanent stableId.
// SAFE: never deletes data, never resets progress, never changes pKey/dKey keys.
function normalizeStagesConfig(cfg) {
  if (!cfg || typeof cfg !== "object") return cfg;
  let modified = false;

  const makeId = (prefix) => {
    const rnd = Math.random().toString(36).slice(2, 9);
    const ts  = Date.now().toString(36).slice(-4);
    return `${prefix}_${ts}${rnd}`;
  };

  const normalized = {};
  Object.entries(cfg).forEach(([key, stage]) => {
    if (!stage || typeof stage !== "object") { normalized[key] = stage; return; }

    let stageChanged = false;
    const newStage = { ...stage };

    // Ensure stage stableId
    if (!newStage.stableId) {
      newStage.stableId = key.startsWith("stage_") ? key : makeId("stage");
      stageChanged = true;
    }

    // Ensure sections have stableId
    if (Array.isArray(newStage.sections)) {
      newStage.sections = newStage.sections.map(sec => {
        const newSec = { ...sec };
        if (!newSec.stableId) {
          newSec.stableId = makeId("sec");
          stageChanged = true;
        }
        // Ensure items have stableId
        if (Array.isArray(newSec.items)) {
          newSec.items = newSec.items.map((it, idx) => {
            if (typeof it === "string") {
              // Legacy string item — upgrade to object with stableId
              stageChanged = true;
              return {
                stableId: makeId("item"),
                legacySecName: newSec.name,
                legacyIdx: idx,
                text: it,
                active: true,
                order: idx,
                createdAt: new Date().toISOString(),
                updatedAt: new Date().toISOString(),
              };
            }
            if (!it.stableId) {
              stageChanged = true;
              return { ...it, stableId: makeId("item") };
            }
            return it;
          });
        }
        return newSec;
      });
    }

    if (stageChanged) modified = true;
    normalized[newStage.stableId] = newStage;
    // If key differs from stableId, remove old key (handled by saving normalized)
  });

  return { normalized, modified };
}

// ─── ESTILOS GLOBALES ────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@500;600;700&family=Source+Sans+3:wght@400;500;600;700&display=swap');
:root {
  --bg-root:#0D0E10; --bg-surface:#17191D; --bg-card:#17191D; --bg-input:rgba(255,255,255,0.05);
  --brand-red:#6F1725; --brand-red-deep:#4A0F18; --brand-gold:#C5A253;
  --brand-gold-mut:rgba(197,162,83,0.15); --state-done:#27AE60; --state-prog:#E67E22;
  --state-info:#5B8FA8; --state-error:#C0392B;
  --text-primary:#E6E7EA; --text-secondary:#9099AA; --text-tertiary:#555F6E; --text-muted:#353C47;
  --border-faint:rgba(255,255,255,0.04); --border-subtle:rgba(255,255,255,0.07);
  --border-default:rgba(255,255,255,0.1); --border-gold:rgba(197,162,83,0.15);
  --r-sm:6px; --r-md:9px; --r-lg:12px; --r-xl:14px;
  --t-fast:0.12s ease; --t-base:0.18s ease; --t-slow:0.25s ease;
}
* { box-sizing:border-box; }
body { margin:0; background:var(--bg-root); }
.rv-root { font-family:'Source Sans 3',sans-serif; color:var(--text-primary); }
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-thumb { background:rgba(197,162,83,0.2); border-radius:3px; }
::-webkit-scrollbar-track { background:transparent; }
.fade-in { animation:fadeIn 0.22s ease; }
@keyframes fadeIn { from{opacity:0;transform:translateY(4px);} to{opacity:1;transform:translateY(0);} }
@keyframes slideDown { from{opacity:0;transform:translateY(-8px);} to{opacity:1;transform:translateY(0);} }
@keyframes toastIn { from{opacity:0;transform:translateX(20px);} to{opacity:1;transform:translateX(0);} }
@keyframes checkPop { 0%{transform:scale(1)} 40%{transform:scale(1.22)} 70%{transform:scale(0.9)} 100%{transform:scale(1)} }
@keyframes shimmer { 0%{background-position:200% 0} 100%{background-position:-200% 0} }
@keyframes syncPulse { 0%,100%{opacity:.5} 50%{opacity:1} }
@keyframes stageSlide { from{opacity:0;transform:translateX(6px)} to{opacity:1;transform:translateX(0)} }
input,textarea,select,button { font-family:inherit; }
.ni { background:var(--bg-input); border:1px solid var(--border-default); border-radius:var(--r-sm);
  color:var(--text-primary); outline:none; font-size:12px; transition:border-color var(--t-base); }
.ni:focus { border-color:rgba(197,162,83,0.45); box-shadow:0 0 0 2px rgba(197,162,83,0.08); }
.ni:disabled { opacity:0.4; cursor:not-allowed; }
.di { background:var(--bg-input); border:1px solid var(--border-default); border-radius:var(--r-sm);
  color:var(--text-primary); outline:none; font-size:12px; padding:4px 7px; color-scheme:dark;
  transition:border-color var(--t-base); }
.di:focus { border-color:rgba(197,162,83,0.45); }
.di:disabled { opacity:0.35; cursor:not-allowed; }
.btn-primary { padding:7px 16px; border-radius:var(--r-sm); border:none; cursor:pointer;
  background:var(--brand-red); color:#fff; font-size:12px; font-weight:700;
  font-family:'Source Sans 3',sans-serif; letter-spacing:0.3px;
  transition:background var(--t-fast),transform var(--t-fast),box-shadow var(--t-fast);
  display:inline-flex; align-items:center; gap:5px; }
.btn-primary:hover { background:#5A1020; box-shadow:0 3px 12px rgba(111,23,37,0.45); }
.btn-primary:active { transform:scale(0.97); }
.btn-primary:disabled { opacity:0.45; cursor:not-allowed; transform:none; }
.btn-secondary { padding:6px 14px; border-radius:var(--r-sm); cursor:pointer;
  background:rgba(255,255,255,0.04); border:1px solid var(--border-default);
  color:var(--text-secondary); font-size:12px; font-weight:600;
  font-family:'Source Sans 3',sans-serif;
  transition:background var(--t-fast),border-color var(--t-fast),color var(--t-fast);
  display:inline-flex; align-items:center; gap:5px; }
.btn-secondary:hover { background:rgba(255,255,255,0.07); border-color:rgba(255,255,255,0.18); color:var(--text-primary); }
.btn-secondary:active { transform:scale(0.97); }
.btn-secondary:disabled { opacity:0.4; cursor:not-allowed; }
.btn-action { padding:5px 11px; border-radius:var(--r-sm); cursor:pointer;
  background:rgba(197,162,83,0.08); border:1px solid rgba(197,162,83,0.18);
  color:rgba(197,162,83,0.8); font-size:11px; font-weight:600;
  font-family:'Source Sans 3',sans-serif; transition:all var(--t-fast);
  display:inline-flex; align-items:center; gap:4px; }
.btn-action:hover { background:rgba(197,162,83,0.14); border-color:rgba(197,162,83,0.35); color:#C5A253; }
.btn-action:active { transform:scale(0.97); }
.btn-danger { padding:6px 14px; border-radius:var(--r-sm); cursor:pointer;
  background:rgba(192,57,43,0.12); border:1px solid rgba(192,57,43,0.3);
  color:#E07A6E; font-size:12px; font-weight:600;
  font-family:'Source Sans 3',sans-serif; transition:all var(--t-fast); }
.btn-danger:hover { background:rgba(192,57,43,0.22); color:#E04040; }
.btn-danger:active { transform:scale(0.97); }
.rv-toast-stack { position:fixed; bottom:18px; right:18px; display:flex; flex-direction:column;
  gap:8px; z-index:9999; pointer-events:none; }
.rv-toast { display:flex; align-items:center; gap:10px; padding:10px 14px; will-change:transform,opacity;
  border-radius:var(--r-md); background:#1E2025; border:1px solid rgba(255,255,255,0.08);
  box-shadow:0 6px 24px rgba(0,0,0,0.6); font-size:12.5px;
  font-family:'Source Sans 3',sans-serif; color:var(--text-primary);
  pointer-events:all; animation:toastIn 0.2s ease; min-width:220px; max-width:320px; }
.rv-toast.success { border-color:rgba(39,174,96,0.35); }
.rv-toast.error   { border-color:rgba(192,57,43,0.4); }
.rv-toast.info    { border-color:rgba(91,143,168,0.35); }
.rv-toast.warning { border-color:rgba(230,126,34,0.35); }
.rv-toast-icon { font-size:14px; flex-shrink:0; line-height:1; }
.rv-toast-dismiss { margin-left:auto; background:none; border:none; cursor:pointer;
  color:var(--text-muted,#353C47); font-size:12px; padding:0 2px; flex-shrink:0; transition:color var(--t-fast); }
.rv-toast-dismiss:hover { color:var(--text-secondary); }
.rv-badge { display:inline-flex; align-items:center; gap:4px; padding:2px 7px;
  border-radius:9999px; font-size:9.5px; font-weight:700; letter-spacing:0.3px;
  font-family:'Cinzel',serif; line-height:1.6; border:1px solid transparent; }
.rv-badge.done     { background:rgba(39,174,96,0.12);  color:#27AE60; border-color:rgba(39,174,96,0.28); }
.rv-badge.inprog   { background:rgba(230,126,34,0.12); color:#E67E22; border-color:rgba(230,126,34,0.28); }
.rv-badge.pending  { background:rgba(255,255,255,0.04); color:#6B7585; border-color:rgba(255,255,255,0.07); }
.rv-badge.active   { background:rgba(197,162,83,0.1);  color:#C5A253; border-color:rgba(197,162,83,0.3); }
.rv-badge.complete { background:rgba(39,174,96,0.1); color:#27AE60; border-color:rgba(39,174,96,0.28);
  font-family:'Source Sans 3',sans-serif; font-size:9px; }
.rv-badge.info     { background:rgba(91,143,168,0.12); color:#7FB3D5; border-color:rgba(91,143,168,0.28); }
.rv-skeleton { background:linear-gradient(90deg,rgba(255,255,255,0.04) 25%,rgba(255,255,255,0.08) 50%,rgba(255,255,255,0.04) 75%);
  background-size:200% 100%; animation:shimmer 1.4s ease infinite; border-radius:6px; }
.rv-skel-line { height:10px; border-radius:5px; margin-bottom:8px; }
.rv-skel-bar  { height:6px;  border-radius:3px; margin-top:6px; }
.rv-skel-card { background:rgba(23,25,29,0.97); border:1px solid rgba(255,255,255,0.04);
  border-radius:12px; padding:14px 16px; margin-bottom:10px; }
.rv-sync-dot { width:6px; height:6px; border-radius:50%; background:rgba(197,162,83,0.7);
  animation:syncPulse 1s ease infinite; display:inline-block; }
.rv-confirm-overlay { position:fixed; inset:0; z-index:800; background:rgba(8,9,11,0.78);
  display:flex; align-items:center; justify-content:center; animation:fadeIn 0.15s ease; }
.rv-confirm-box { background:#1E2025; border:1px solid rgba(111,23,37,0.4);
  border-radius:14px; padding:22px 24px; max-width:320px; width:90%;
  box-shadow:0 12px 40px rgba(0,0,0,0.65); animation:slideDown 0.18s ease; }
.rv-chk { width:20px; height:20px; border-radius:6px; border:1.5px solid rgba(255,255,255,0.15);
  background:transparent; display:flex; align-items:center; justify-content:center;
  flex-shrink:0; font-size:11px; color:#fff; transition:all 0.18s ease; }
.rv-chk.popping { animation:checkPop 0.22s ease; }
.rv-item-row { display:flex; align-items:center; gap:9px; padding:6px 8px;
  border-radius:7px; border:1px solid transparent;
  transition:background 0.12s,border-color 0.12s,opacity 0.18s; }
.rv-item-row:hover { background:rgba(255,255,255,0.025); }
.rv-item-row.expanded { border-color:rgba(255,255,255,0.08); background:rgba(255,255,255,0.02); }
.rv-item-row.done { opacity:0.45; }
.rv-item-row.done:hover { opacity:0.65; }
.rv-item-detail { margin:0 8px 6px 37px; padding:8px 11px; background:rgba(13,14,16,0.9);
  border-radius:7px; display:flex; align-items:center; gap:9px;
  border:1px solid rgba(255,255,255,0.05); }
.rv-sec-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(320px,1fr));
  gap:10px; padding:12px 14px; }
.rv-filter-btn { padding:8px 11px; border-radius:6px; border:1px solid rgba(255,255,255,0.07);
  background:rgba(255,255,255,0.02); cursor:pointer; font-size:11px; font-weight:600;
  font-family:'Source Sans 3',sans-serif; transition:all 0.12s; color:#6B7585;
  min-height:36px; display:inline-flex; align-items:center; }
.rv-filter-btn:hover { border-color:rgba(255,255,255,0.15); color:#9099AA; }
.rv-prog-track { background:rgba(255,255,255,0.06); border-radius:4px; overflow:hidden; }
.rv-prog-fill  { height:100%; border-radius:4px; transition:width 0.25s ease,background 0.25s ease; will-change:width; }
.rv-prog-fill.done { background:linear-gradient(90deg,#1E8449,#27AE60)!important; }
.rv-card { background:rgba(23,25,29,0.97); border:1px solid rgba(255,255,255,0.06);
  border-radius:12px; overflow:hidden;
  transition:border-color 0.18s ease,box-shadow 0.18s ease,transform 0.12s ease; }
.rv-card-interactive { cursor:pointer; }
.rv-card-interactive:hover { border-color:rgba(255,255,255,0.1); box-shadow:0 4px 20px rgba(0,0,0,0.3); }
.rv-card-interactive:active { transform:scale(0.995); }
.rv-card-hdr { display:flex; align-items:center; justify-content:space-between;
  padding:11px 14px; cursor:pointer; user-select:none; }
.rv-card-hdr:hover { background:rgba(255,255,255,0.02); }
.rv-card-body { padding:10px 14px 12px; }
.rv-card-inv { display:flex; align-items:center; gap:10px; padding:10px 14px;
  border-radius:10px; border:1px solid rgba(255,255,255,0.06); background:rgba(23,25,29,0.97);
  transition:border-color 0.18s,background 0.18s,box-shadow 0.18s; }
.rv-card-inv:hover { border-color:rgba(255,255,255,0.1); box-shadow:0 2px 12px rgba(0,0,0,0.25); }
.rv-card-inv:active { transform:scale(0.995); }
.rv-sec-actions { display:flex; gap:4px; opacity:0; transition:opacity 0.15s; }
.rv-card-hdr:hover .rv-sec-actions { opacity:1; }
.rv-qbtn { background:rgba(255,255,255,0.05); border:1px solid rgba(255,255,255,0.07);
  border-radius:4px; color:#6B7585; font-size:9px; cursor:pointer; padding:8px 10px;
  font-family:'Source Sans 3',sans-serif; font-weight:600; transition:all 0.12s;
  min-width:36px; min-height:36px; display:inline-flex; align-items:center; justify-content:center; }
.rv-qbtn:hover { background:rgba(255,255,255,0.1); color:#9099AA; }
.rv-search { width:100%; padding:6px 9px; background:rgba(255,255,255,0.04);
  border:1px solid rgba(255,255,255,0.07); border-radius:7px; color:var(--text-primary);
  font-size:12px; outline:none; transition:border-color var(--t-base);
  font-family:'Source Sans 3',sans-serif; }
.rv-search::placeholder { color:rgba(255,255,255,0.2); }
.rv-search:focus { border-color:rgba(197,162,83,0.4); }
.rv-inline-edit { background:rgba(255,255,255,0.07); border:1px solid rgba(197,162,83,0.4);
  border-radius:5px; color:var(--text-primary); font-size:13px; font-weight:600;
  padding:2px 6px; outline:none; width:100%; font-family:'Source Sans 3',sans-serif; }
.rv-field-err { font-size:10px; color:#E07A6E; margin-top:3px; font-family:'Source Sans 3',sans-serif; }
.rv-field-req { color:#8F2235; font-size:10px; margin-left:2px; }
.rv-stage-content { animation:stageSlide 0.18s ease; will-change:transform,opacity; }
.rv-clan-table { width:100%; border-collapse:collapse; font-size:12px; }
.rv-clan-table th { padding:7px 10px; text-align:left; font-family:'Cinzel',serif;
  font-size:8.5px; letter-spacing:1px; color:#6B7585; font-weight:700;
  border-bottom:1px solid rgba(255,255,255,0.05); cursor:pointer; user-select:none; white-space:nowrap; }
.rv-clan-table th:hover { color:#C5A253; }
.rv-clan-table td { padding:7px 10px; border-bottom:1px solid rgba(255,255,255,0.03);
  vertical-align:middle; font-family:'Source Sans 3',sans-serif; }
.rv-clan-row { transition:background 0.12s; cursor:pointer; }
.rv-clan-row:hover { background:rgba(255,255,255,0.025); }
.rv-clan-row:hover .rv-quick-view { opacity:1; pointer-events:all; }
.rv-stat-card { background:rgba(23,25,29,0.97); border:1px solid rgba(255,255,255,0.06);
  border-radius:12px; padding:14px 16px; flex:1; min-width:110px; }
.rv-stat-card:hover { border-color:rgba(255,255,255,0.1); }
.rv-quick-view { opacity:0; pointer-events:none; transition:opacity 0.15s;
  position:absolute; z-index:50; right:8px; top:50%; transform:translateY(-50%);
  background:#1E2025; border:1px solid rgba(197,162,83,0.25);
  border-radius:9px; padding:10px 13px; min-width:180px;
  box-shadow:0 6px 20px rgba(0,0,0,0.5); }
.rv-priority-star { font-size:11px; cursor:pointer; transition:color 0.12s;
  padding:2px 4px; border-radius:3px; line-height:1; }
.rv-priority-star:hover { background:rgba(197,162,83,0.15); }
.rv-filter-chip { padding:8px 9px; border-radius:20px; font-size:10.5px; font-weight:600;
  border:1px solid rgba(255,255,255,0.07); background:rgba(255,255,255,0.02);
  cursor:pointer; font-family:'Source Sans 3',sans-serif; color:#6B7585;
  transition:all 0.12s; white-space:nowrap;
  min-height:36px; display:inline-flex; align-items:center; }
.rv-filter-chip:hover { border-color:rgba(255,255,255,0.14); color:#9099AA; }
.rv-filter-chip.on { background:rgba(197,162,83,0.12); border-color:rgba(197,162,83,0.35); color:#C5A253; }
.rv-report-card { background:rgba(23,25,29,0.97); border:1px solid rgba(255,255,255,0.06);
  border-radius:12px; padding:16px 18px; transition:border-color 0.15s,box-shadow 0.15s; }
.rv-report-card:hover { border-color:rgba(197,162,83,0.25); box-shadow:0 4px 16px rgba(0,0,0,0.4); }
.rv-preview-overlay { position:fixed; inset:0; background:rgba(8,9,11,0.9);
  z-index:900; display:flex; flex-direction:column; align-items:center; animation:fadeIn 0.18s; }
.rv-preview-toolbar { width:100%; max-width:960px; display:flex; align-items:center;
  gap:10px; padding:10px 16px; background:rgba(13,14,16,0.99);
  border-bottom:1px solid rgba(255,255,255,0.06); flex-shrink:0; }
.rv-preview-frame { flex:1; width:100%; max-width:960px; border:none; background:#FAF6F0; }
.add-btn:hover { background:rgba(111,23,37,0.22)!important; }
.del-btn:hover { color:#8F2235!important; }
.scout-row:hover { background:rgba(23,25,29,0.9)!important; }
.sec-hdr:hover { background:rgba(255,255,255,0.02); }
.check-row:hover { background:rgba(255,255,255,0.02); border-color:rgba(255,255,255,0.05)!important; }
.expbtn:hover { color:#C5A253!important; background:rgba(197,162,83,0.08)!important; border-radius:4px; }
.stab:hover { opacity:0.82; }
.ebtn:hover { background:rgba(197,162,83,0.15)!important; }
.att-nav:hover { background:rgba(91,143,168,0.18)!important; }
.att-cell:hover { transform:scale(1.15); }
.move-btn:hover { background:rgba(197,162,83,0.15)!important; }
.login-input:focus { border-color:rgba(197,162,83,0.5)!important; box-shadow:0 0 0 2px rgba(197,162,83,0.08)!important; }
.login-btn:hover { background:#5A1020!important; }
/* .viewer-badge reserved */
.hamburger:hover { background:rgba(255,255,255,0.06)!important; }
@media (max-width:700px) {
  /* ── Sidebar: drawer lateral en móvil ── */
  .rv-sidebar {
    position: fixed !important;
    top: 0; left: 0; bottom: 0;
    z-index: 500;
    transform: translateX(-100%);
    transition: transform 0.25s ease;
    width: 260px !important;
    box-shadow: 4px 0 24px rgba(0,0,0,0.6);
  }
  .rv-sidebar.open-mobile {
    transform: translateX(0) !important;
  }
  /* ── Overlay: visible cuando sidebar está abierto ── */
  .rv-overlay {
    display: block !important;
    position: fixed !important;
    inset: 0;
    background: rgba(0,0,0,0.6);
    z-index: 499;
  }
  /* ── Hamburger: visible en móvil ── */
  .hamburger {
    display: flex !important;
    align-items: center;
    justify-content: center;
  }
  /* ── Header: reducir padding ── */
  .rv-header { padding: 8px 12px !important; }
  /* ── rv-body: main ocupa todo el ancho ── */
  .rv-body { }
  /* ── Content area: columna en móvil ── */
  .rv-content-area { flex-direction: column !important; }
  /* ── Stage nav: barra horizontal superior ── */
  .rv-stage-nav {
    flex-direction: row !important;
    overflow-x: auto !important; overflow-y: hidden !important;
    width: 100% !important; max-height: none !important;
    border-right: none !important;
    border-bottom: 1px solid rgba(197,162,83,0.1) !important;
    padding: 6px 8px !important;
    flex-shrink: 0 !important; gap: 5px !important;
  }
  .rv-stage-nav-item {
    min-width: 80px !important; flex-direction: column !important;
    align-items: flex-start !important; padding: 6px 8px !important;
  }
  /* ── Stage content: scroll propio ── */
  .rv-stage-content { min-height: 0 !important; }
  /* ── Secciones: una columna ── */
  .rv-sec-grid { grid-template-columns:1fr!important; padding:10px!important; }
  .rv-card-hdr { padding:10px 12px; }
  .rv-card-body { padding:8px 12px 10px; }
  .rv-card-inv  { padding:10px 12px; }
  .rv-stage-nav { flex-direction:row!important; overflow-x:auto!important; overflow-y:hidden!important;
    width:100%!important; max-height:none!important; border-right:none!important;
    border-bottom:1px solid rgba(197,162,83,0.1)!important; padding:6px 10px!important;
    flex-shrink:0!important; gap:5px!important; }
  .rv-stage-nav-item { min-width:90px!important; flex-direction:row!important;
    align-items:center!important; padding:6px 10px!important; gap:6px!important; }
  .rv-content-area { flex-direction:column!important; overflow:auto!important; }
  .rv-clan-table th:nth-child(4), .rv-clan-table td:nth-child(4) { display:none; }
}
@media (min-width:1200px) {
  .rv-sec-grid { grid-template-columns:repeat(auto-fill,minmax(340px,1fr)); }
}

/* ── RF-02/RF-03: Foco visible consistente — Fase 13 ── */
:focus-visible {
  outline: 2px solid rgba(197,162,83,0.8);
  outline-offset: 2px;
  border-radius: 4px;
}
button:focus-visible, a:focus-visible, input:focus-visible, select:focus-visible {
  outline: 2px solid rgba(197,162,83,0.8);
  outline-offset: 2px;
  box-shadow: 0 0 0 4px rgba(197,162,83,0.12);
}
/* RF-04: touch targets mínimo 44px */
/* RF-09 / CORRECCIÓN 3: Touch targets ~44×44px en controles principales */
/* Desktop: icon + action buttons at 44×44 */
.rv-chk, .rv-priority-star, .hamburger {
  min-width: 44px; min-height: 44px;
  display: inline-flex; align-items: center; justify-content: center;
}
/* Expand hit area on small action buttons */
.expbtn, .del-btn, .move-btn {
  min-width: 40px; min-height: 40px;
  padding: 8px !important;
  display: inline-flex; align-items: center; justify-content: center;
}
/* att-nav: comfortable tap area */
.att-nav {
  min-height: 36px; min-width: 36px;
  display: inline-flex; align-items: center; justify-content: center;
}
/* Mobile: all interactive controls at 44px minimum */
@media (max-width: 700px) {
  .rv-chk, .rv-priority-star, .expbtn, .del-btn, .move-btn,
  .rv-qbtn, .rv-filter-btn, .rv-filter-chip, .att-nav, .hamburger {
    min-width: 44px; min-height: 44px;
  }
  .rv-filter-chip { display: inline-flex; flex-shrink: 1; }
}
/* RF-10: preferencias de accesibilidad */
.rv-large-text { font-size: 115% !important; }
.rv-large-text .rv-clan-table { font-size: 13.5px !important; }
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: 0.01ms !important;
    transition-duration: 0.01ms !important;
  }
}
.rv-reduce-motion *, .rv-reduce-motion *::before, .rv-reduce-motion *::after {
  animation-duration: 0.01ms !important;
  transition-duration: 0.01ms !important;
  will-change: auto !important;
}
/* RF-12: high contrast improvements */
.rv-high-contrast { --text-primary: #FFFFFF; --text-secondary: #E0E0E0; }
.rv-high-contrast .rv-item-row { border-color: rgba(255,255,255,0.15) !important; }
.rv-high-contrast .rv-card { border-color: rgba(255,255,255,0.2) !important; }
.rv-high-contrast .rv-prog-track { background: rgba(255,255,255,0.2) !important; }
.rv-high-contrast .ni, .rv-high-contrast .di {
  border-color: rgba(255,255,255,0.4) !important; color: #fff !important;
}
/* RF-12: large text — preserve layout */
.rv-large-text { font-size: 113%; line-height: 1.5; }
.rv-large-text .rv-clan-table { font-size: 13px !important; }
.rv-large-text .rv-stage-nav { font-size: 112%; }
/* RF-11: focus trap visual indicator */
[role="alertdialog"]:focus-within { outline: 2px solid rgba(197,162,83,0.6); }
/* RF-10: alto contraste */
.rv-high-contrast .rv-card { border-color: rgba(255,255,255,0.2) !important; }
.rv-high-contrast .rv-item-row { border-color: rgba(255,255,255,0.1) !important; }
.rv-high-contrast .rv-clan-table td { border-bottom-color: rgba(255,255,255,0.08) !important; }
.rv-high-contrast .rv-prog-track { background: rgba(255,255,255,0.15) !important; }
/* RF-07: responsive fixes */
@media (max-width:480px) {
  .rv-header-title { font-size: 12px !important; }
  .rv-preview-toolbar { flex-wrap: wrap; }
  .rv-stat-card { min-width: 90px; }
  /* Header: ocultar textos largos en botones */
  .rv-header .btn-secondary span,
  .rv-header .rv-header-sub { font-size: 9px !important; }
  /* Nav buttons en header: solo icono */
  .rv-nav-label { display: none !important; }
}
/* RF-09: keyboard shortcut hint */
.rv-kbd { display:inline-block; padding:1px 5px; border-radius:4px;
  background:rgba(255,255,255,0.07); border:1px solid rgba(255,255,255,0.15);
  font-size:9px; font-family:'Source Sans 3',sans-serif; color:#6B7585;
  line-height:1.5; }
`;

// ─── RADAR CHART ─────────────────────────────────────────────────────────────
function RadarChart({ stage, scoutId, prog, size = 200 }) {
  const cx = size / 2, cy = size / 2 + 6, r = size * 0.36;
  const areas = stage.areas;
  const n = areas.length;
  const scores = areas.map(area => {
    let t = 0, d = 0;
    stage.sections.forEach(sec => {
      if (sec.area === area) sec.items.forEach((_, i) => {
        t++; if (prog[pKey(scoutId, stage.id, sec.name, i)]) d++;
      });
    });
    return t ? d / t : 0;
  });
  const angle = i => (Math.PI * 2 * i / n) - Math.PI / 2;
  const pt = (i, radius) => [cx + radius * Math.cos(angle(i)), cy + radius * Math.sin(angle(i))];
  const polyStr = pts => pts.map(p => p.join(",")).join(" ");
  const gridLvs = [0.25, 0.5, 0.75, 1.0];
  return (
    <svg viewBox={`0 0 ${size} ${size + 12}`} width="100%" style={{ maxWidth: size + "px", display: "block", margin: "0 auto" }}>
      {gridLvs.map((lv, gi) => (
        <polygon key={gi} points={polyStr(areas.map((_, i) => pt(i, r * lv)))}
          fill="none" stroke={gi === 3 ? "rgba(201,162,39,0.2)" : "rgba(201,162,39,0.1)"} strokeWidth={gi === 3 ? 1.2 : 0.8} />
      ))}
      {areas.map((_, i) => {
        const [x, y] = pt(i, r);
        return <line key={i} x1={cx} y1={cy} x2={x} y2={y} stroke="rgba(201,162,39,0.15)" strokeWidth="0.8" />;
      })}
      <polygon points={polyStr(areas.map((_, i) => pt(i, r * scores[i])))}
        fill={`${stage.color}28`} stroke={stage.color} strokeWidth="2" strokeLinejoin="round" />
      {areas.map((_, i) => {
        const [x, y] = pt(i, r * scores[i]);
        return <circle key={i} cx={x} cy={y} r="4" fill={stage.color} stroke="#0A0604" strokeWidth="1.5" />;
      })}
      {areas.map((a, i) => {
        const lr = r + 20;
        const [lx, ly] = pt(i, lr);
        return (
          <g key={i}>
            <text x={lx} y={ly - 2} textAnchor="middle" fill="#EDE8E1" fontSize="8"
              fontFamily="'Cinzel',serif" fontWeight="600">{a}</text>
            <text x={lx} y={ly + 9} textAnchor="middle" fill={stage.color} fontSize="8.5"
              fontFamily="'Source Sans 3',sans-serif" fontWeight="700">{Math.round(scores[i]*100)}%</text>
          </g>
        );
      })}
    </svg>
  );
}

// ─── HELPERS DE CÁLCULO (compartidos por export y UI) ──────────────────────────
function scoutStageStats(scout, stage, prog) {
  let t = 0, d = 0;
  (stage.sections||[]).forEach(sec => {
    const secItems = (sec.items||[]).filter(it => it.active !== false);
    secItems.forEach((rawItem, idx) => {
      t++;
      const itemObj = typeof rawItem === "object" ? rawItem : null;
      let checked = false;
      if (itemObj && itemObj.stableId && stage.stableId && sec.stableId) {
        const sk = pKeyS(scout.id, stage.stableId, sec.stableId, itemObj.stableId);
        if (prog[sk] !== undefined) { checked = !!prog[sk]; }
        else {
          const li = itemObj.legacyIdx !== undefined ? itemObj.legacyIdx : idx;
          const ln = itemObj.legacySecName || sec.name;
          checked = !!prog[pKey(scout.id, stage.id, ln, li)];
        }
      } else {
        checked = !!prog[pKey(scout.id, stage.id, sec.name, idx)];
      }
      if (checked) d++;
    });
  });
  return { t, d, pct: t ? Math.round(d / t * 100) : 0 };
}
function scoutTotalStats(scout, prog) {
  let t = 0, d = 0;
  STAGES.forEach(st => { const s = scoutStageStats(scout, st, prog); t += s.t; d += s.d; });
  return { t, d, pct: t ? Math.round(d / t * 100) : 0 };
}
function scoutAttendanceStats(scout, attendance) {
  const list = Object.entries(attendance)
    .filter(([k]) => k.startsWith(`att:${scout.id}|`))
    .map(([k, v]) => ({ date: k.split("|")[1], status: v.status, note: v.note || "" }));
  const t = list.length;
  const p = list.filter(r => r.status === "presente").length;
  const a = list.filter(r => r.status === "ausente").length;
  const e = list.filter(r => r.status === "excusa").length;
  return { t, p, a, e, pct: t ? Math.round(p / t * 100) : 0, list: list.sort((x,y)=>x.date.localeCompare(y.date)) };
}
function scoutCurrentStage(scout, prog) {
  for (let i = STAGES.length - 1; i >= 0; i--) {
    const s = scoutStageStats(scout, STAGES[i], prog);
    if (s.pct === 100 && i < STAGES.length - 1) return STAGES[i + 1];
    if (s.d > 0) return STAGES[i];
  }
  return STAGES[0];
}

// ─── EXPORT: HTML/PDF — TODOS LOS ROVERS, DISEÑO DINÁMICO ─────────────────────
// ─── RF-10: QR code generator (API pública, sin dependencias) ────────────────
function qrUrl(text) {
  return `https://api.qrserver.com/v1/create-qr-code/?size=90x90&data=${encodeURIComponent(text)}`;
}

// ─── RF-11: Encabezado institucional para todos los documentos ────────────────
function institutionalHeader(title, subtitle, today) {
  return `
  <div style="background:linear-gradient(135deg,#6B0F0F,#8B1A1A 60%,#A52020);
    padding:18px 24px;border-radius:0;margin-bottom:0;
    display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap">
    <div style="display:flex;align-items:center;gap:14px">
      <div style="font-size:36px;color:#C9A227;line-height:1">⚜</div>
      <div>
        <div style="font-size:9px;letter-spacing:3px;color:rgba(201,162,39,0.8);font-weight:700;margin-bottom:2px">
          GRUPO SCOUT No. 12 · SAN PEDRO APÓSTOL · GUATEMALA
        </div>
        <div style="font-size:18px;font-weight:700;color:#FAF6F0;font-family:Georgia,serif">${title}</div>
        ${subtitle ? `<div style="font-size:11px;color:rgba(255,255,255,0.6);margin-top:2px">${subtitle}</div>` : ""}
      </div>
    </div>
    <div style="text-align:right">
      <div style="font-size:9px;color:rgba(201,162,39,0.7);letter-spacing:1px">EMITIDO</div>
      <div style="font-size:11px;color:#FAF6F0;font-weight:600">${today}</div>
    </div>
  </div>`;
}

// ─── RF-08: Bloque de estadísticas para reportes generales ───────────────────
function statsBlock(scouts, prog, attendance) {
  const today = new Date().toLocaleDateString("es-GT",{day:"2-digit",month:"long",year:"numeric"});
  const clanD   = scouts.reduce((a, sc) => a + scoutTotalStats(sc, prog).d, 0);
  const clanT   = scouts.reduce((a, sc) => a + scoutTotalStats(sc, prog).t, 0);
  const clanPct = clanT ? Math.round(clanD / clanT * 100) : 0;
  const avgPct  = scouts.length
    ? Math.round(scouts.reduce((a,sc) => a + scoutTotalStats(sc,prog).pct, 0) / scouts.length)
    : 0;
  const completed = scouts.filter(sc => scoutTotalStats(sc,prog).pct === 100).length;
  const items = [
    { label:"Total Rovers", value: scouts.length, color:"#8B1A1A" },
    { label:"Promedio de avance", value: avgPct+"%", color:"#C0392B" },
    { label:"Etapas al 100%", value: completed, color:"#27AE60" },
    { label:"Objetivos completados", value: clanD, color:"#2E86C1" },
    { label:"Objetivos pendientes", value: clanT-clanD, color:"#B9770E" },
    { label:"Progreso total del Clan", value: clanPct+"%", color:"#8B1A1A" },
  ];
  return `
  <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px">
    ${items.map(c => `
      <div style="flex:1;min-width:110px;background:#fff;border:1px solid #EAE0D2;
        border-radius:10px;padding:12px 14px;border-top:3px solid ${c.color}">
        <div style="font-size:22px;font-weight:700;color:${c.color};font-family:Georgia,serif">${c.value}</div>
        <div style="font-size:10px;color:#9A7A60;margin-top:2px">${c.label}</div>
      </div>`).join("")}
  </div>`;
}

// ─── RF-09: Ficha Individual del Rover ────────────────────────────────────────
function buildRoverFicha(sc, prog, dates, attendance, qrBase) {
  const today  = new Date().toLocaleDateString("es-GT",{day:"2-digit",month:"long",year:"numeric"});
  const cs     = scoutCurrentStage(sc, prog);
  const tot    = scoutTotalStats(sc, prog);
  const att    = scoutAttendanceStats(sc, attendance);
  const qrLink = qrBase ? qrBase + encodeURIComponent(sc.name) : "";

  // Última actividad
  const allDates = Object.entries(dates)
    .filter(([k]) => k.startsWith(`d:${sc.id}|`))
    .map(([,v]) => v).filter(Boolean).sort().reverse();
  const lastActivity = allDates[0] || "Sin registro";

  // Barras de etapa
  const stageBars = STAGES.map(st => {
    const s = scoutStageStats(sc, st, prog);
    const items_done = [];
    const items_pend = [];
    (st.sections||[]).forEach(sec => {
      const secItems = (sec.items||[]).filter(it => it.active !== false);
      secItems.forEach((rawItem, idx) => {
        const itemObj = typeof rawItem === "object" ? rawItem : null;
        const itemText = itemObj ? (itemObj.text || String(rawItem)) : String(rawItem);
        let done = false, dateVal = "";
        if (itemObj && itemObj.stableId && st.stableId && sec.stableId) {
          const sk = pKeyS(sc.id, st.stableId, sec.stableId, itemObj.stableId);
          const li = itemObj.legacyIdx !== undefined ? itemObj.legacyIdx : idx;
          const ln = itemObj.legacySecName || sec.name;
          done = !!(prog[sk] !== undefined ? prog[sk] : prog[pKey(sc.id, st.id, ln, li)]);
          dateVal = dates[dKeyS(sc.id, st.stableId, sec.stableId, itemObj.stableId)] ||
                    dates[dKey(sc.id, st.id, ln, li)] || "";
        } else {
          done = !!prog[pKey(sc.id, st.id, sec.name, idx)];
          dateVal = dates[dKey(sc.id, st.id, sec.name, idx)] || "";
        }
        if (done) items_done.push({ item: itemText, date: dateVal });
        else items_pend.push({ item: itemText });
      });
    });
    return `
    <div style="margin-bottom:18px;page-break-inside:avoid">
      <div style="display:flex;justify-content:space-between;align-items:center;
        padding:8px 12px;background:${s.pct===100?"#F0FAF4":"#FBF8F4"};
        border-left:4px solid ${st.color};border-radius:0 8px 8px 0;margin-bottom:8px">
        <div>
          <span style="font-family:Georgia,serif;font-size:13px;font-weight:700;color:${st.color}">${st.numeral} </span>
          <span style="font-size:13px;font-weight:700;color:#241008">${st.name}</span>
          <span style="font-size:10px;color:#9A7A60;margin-left:8px">${st.subtitle}</span>
        </div>
        <div style="text-align:right">
          <div style="font-size:18px;font-weight:700;color:${s.pct===100?"#27AE60":st.color}">${s.pct}%</div>
          <div style="font-size:9px;color:#9A7A60">${s.d}/${s.t} objetivos</div>
        </div>
      </div>
      <div style="height:6px;background:#EFE8DD;border-radius:3px;overflow:hidden;margin-bottom:10px">
        <div style="height:100%;width:${s.pct}%;background:${s.pct===100?"#27AE60":st.color};border-radius:3px"></div>
      </div>
      ${items_done.length ? `
      <div style="margin-bottom:6px">
        <div style="font-size:9px;font-weight:700;color:#27AE60;letter-spacing:1px;margin-bottom:4px">COMPLETADOS (${items_done.length})</div>
        ${items_done.map(e => `
          <div style="display:flex;gap:6px;align-items:baseline;font-size:10.5px;color:#3D5A3E;margin-bottom:2px;padding:2px 0;border-bottom:1px dotted #E8F5E9">
            <span style="color:#27AE60;flex-shrink:0">✓</span>
            <span style="flex:1">${e.item}</span>
            ${e.date ? `<span style="font-size:9.5px;color:#9A7A60;white-space:nowrap">${e.date}</span>` : ""}
          </div>`).join("")}
      </div>` : ""}
      ${items_pend.length ? `
      <div>
        <div style="font-size:9px;font-weight:700;color:#B9770E;letter-spacing:1px;margin-bottom:4px">PENDIENTES (${items_pend.length})</div>
        ${items_pend.map(e => `
          <div style="display:flex;gap:6px;align-items:baseline;font-size:10.5px;color:#7A5A40;margin-bottom:2px">
            <span style="color:#DEB887;flex-shrink:0">○</span>
            <span>${e.item}</span>
          </div>`).join("")}
      </div>` : ""}
    </div>`;
  }).join("");

  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<title>Ficha — ${sc.name}</title>
<style>
  *{box-sizing:border-box}
  body{font-family:'Helvetica Neue',Arial,sans-serif;max-width:820px;margin:0 auto;
    padding:0 0 40px;color:#241008;background:#FAF6F0;font-size:12px}
  @media print{
    body{background:#fff;padding:0}
    .noprint{display:none!important}
    @page{margin:1.5cm;size:A4}
    .page-break{page-break-before:always}
  }
  @media(max-width:600px){body{padding:10px}}
</style></head><body>

${institutionalHeader("Ficha Oficial de Progresión Rover", `Emitida el ${today}`, today)}

<div style="padding:16px 20px">
  <!-- Info general -->
  <div style="display:flex;justify-content:space-between;align-items:flex-start;
    margin-bottom:16px;flex-wrap:wrap;gap:12px">
    <div style="display:flex;gap:14px;align-items:center">
      <div style="width:52px;height:52px;border-radius:50%;background:${cs.color};
        color:#fff;display:flex;align-items:center;justify-content:center;
        font-family:Georgia,serif;font-size:22px;font-weight:700;flex-shrink:0">
        ${(sc.name||"?")[0].toUpperCase()}
      </div>
      <div>
        <div style="font-size:19px;font-weight:700;color:#241008;margin-bottom:2px">${sc.name}</div>
        ${sc.group ? `<div style="font-size:11px;color:#9A7A60;margin-bottom:3px">Clan: ${sc.group}</div>` : ""}
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <span style="background:${cs.color}18;border:1px solid ${cs.color}44;
            border-radius:20px;padding:2px 9px;font-size:10.5px;font-weight:700;color:${cs.color}">
            ${cs.numeral} ${cs.name}
          </span>
          ${tot.pct===100 ? `<span style="background:#E8F5E9;border:1px solid #27AE6044;
            border-radius:20px;padding:2px 9px;font-size:10.5px;font-weight:700;color:#27AE60">
            ✓ Progresión Completa</span>` : ""}
        </div>
      </div>
    </div>
    <div style="display:flex;gap:12px;align-items:flex-start">
      <div style="text-align:right">
        <div style="font-size:36px;font-weight:700;color:#8B1A1A;font-family:Georgia,serif;line-height:1">${tot.pct}%</div>
        <div style="font-size:10px;color:#9A7A60">${tot.d}/${tot.t} objetivos totales</div>
        <div style="font-size:10px;color:#9A7A60;margin-top:3px">Últ. actividad: ${lastActivity}</div>
        ${att.t ? `<div style="font-size:10px;color:#2E86C1;margin-top:3px">Asistencia: ${att.pct}%</div>` : ""}
      </div>
      ${qrLink ? `<div style="text-align:center">
        <img src="${qrUrl(qrLink)}" width="80" height="80" style="border-radius:6px;border:1px solid #EAE0D2" alt="QR"/>
        <div style="font-size:8px;color:#9A7A60;margin-top:2px">Escanear para<br>ver en línea</div>
      </div>` : ""}
    </div>
  </div>

  <!-- Resumen de barras -->
  <div style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap">
    ${STAGES.map(st => {
      const s = scoutStageStats(sc, st, prog);
      return `<div style="flex:1;min-width:140px;background:#fff;border:1px solid #EAE0D2;
        border-radius:10px;padding:10px 12px;border-top:3px solid ${st.color}">
        <div style="display:flex;justify-content:space-between;margin-bottom:5px">
          <span style="font-family:Georgia,serif;font-size:11px;font-weight:700;color:${st.color}">${st.numeral} ${st.name}</span>
          <span style="font-size:13px;font-weight:700;color:${s.pct===100?"#27AE60":st.color}">${s.pct}%</span>
        </div>
        <div style="height:5px;background:#EFE8DD;border-radius:3px;overflow:hidden">
          <div style="height:100%;width:${s.pct}%;background:${s.pct===100?"#27AE60":st.color};border-radius:3px"></div>
        </div>
        <div style="font-size:9.5px;color:#9A7A60;margin-top:4px">${s.d}/${s.t} objetivos</div>
      </div>`;
    }).join("")}
  </div>

  <button class="noprint" onclick="window.print()"
    style="display:block;margin:0 auto 18px;padding:9px 22px;background:#8B1A1A;color:#fff;
    border:none;border-radius:7px;cursor:pointer;font-size:13px;font-weight:600">
    🖨️ Imprimir / Guardar como PDF
  </button>

  <!-- Detalle por etapa -->
  <div style="margin-top:6px">
    <div style="font-size:9px;letter-spacing:2px;color:#B9770E;font-weight:700;margin-bottom:12px">
      DETALLE POR ETAPA
    </div>
    ${stageBars}
  </div>

  <!-- Espacio para observaciones y firmas -->
  <div style="margin-top:16px;page-break-inside:avoid">
    <div style="font-size:9px;letter-spacing:2px;color:#B9770E;font-weight:700;margin-bottom:8px">
      OBSERVACIONES DEL JEFE DE CLAN
    </div>
    <div style="border:1px solid #EAE0D2;border-radius:8px;min-height:70px;padding:10px;
      background:#fff;font-size:11px;color:#9A7A60">
      &nbsp;
    </div>
    <div style="display:flex;gap:24px;margin-top:28px;flex-wrap:wrap">
      <div style="flex:1;min-width:160px;text-align:center">
        <div style="border-top:1px solid #241008;padding-top:6px;font-size:10px;color:#241008">
          Firma del Rover
        </div>
      </div>
      <div style="flex:1;min-width:160px;text-align:center">
        <div style="border-top:1px solid #241008;padding-top:6px;font-size:10px;color:#241008">
          Firma del Jefe de Clan
        </div>
      </div>
      <div style="flex:1;min-width:160px;text-align:center">
        <div style="border-top:1px solid #241008;padding-top:6px;font-size:10px;color:#241008">
          Fecha
        </div>
      </div>
    </div>
  </div>

  <!-- Pie de página -->
  <div style="margin-top:28px;padding-top:10px;border-top:1px solid #EAE0D2;
    display:flex;justify-content:space-between;font-size:9px;color:#B9A080;flex-wrap:wrap;gap:4px">
    <span>Grupo Scout No. 12 · San Pedro Apóstol · Guatemala</span>
    <span>Cartilla 2.0 · Progresión Rover</span>
    <span>Emitida el ${today}</span>
  </div>
</div>
</body></html>`;
}

// ─── RF-01/RF-02: Reporte General del Clan (mejorado) ────────────────────────
function exportAllHTML(scouts, prog, dates, attendance, filterIds, stagesList) {
  const today = new Date().toLocaleDateString("es-GT",{day:"2-digit",month:"long",year:"numeric"});
  const filtered = filterIds ? scouts.filter(s => filterIds.includes(s.id)) : scouts;
  const qrBase = window.location.origin + window.location.pathname + "?rover=";

  const cards = filtered.map(sc => {
    const tot = scoutTotalStats(sc, prog);
    const cs  = scoutCurrentStage(sc, prog);
    const att = scoutAttendanceStats(sc, attendance);
    const allDates = Object.entries(dates)
      .filter(([k]) => k.startsWith(`d:${sc.id}|`))
      .map(([,v]) => v).filter(Boolean).sort().reverse();
    const lastDate = allDates[0] || "—";

    const stages = (stagesList && stagesList.length > 0) ? stagesList : STAGES;
    const stageBars = stages.map(st => {
      const s = scoutStageStats(sc, st, prog);
      return `
      <div style="margin-bottom:8px">
        <div style="display:flex;justify-content:space-between;font-size:10.5px;color:#6B5644;margin-bottom:2px">
          <span style="font-weight:700;color:${st.color}">${st.numeral} · ${st.name}</span>
          <span>${s.d}/${s.t} · ${s.pct}%</span>
        </div>
        <div style="height:5px;background:#EFE8DD;border-radius:3px;overflow:hidden">
          <div style="height:100%;width:${s.pct}%;background:${s.pct===100?"#27AE60":st.color};border-radius:3px"></div>
        </div>
      </div>`;
    }).join("");

    return `
    <div style="background:#fff;border:1px solid #EAE0D2;border-radius:12px;padding:16px 18px;
      margin-bottom:14px;page-break-inside:avoid;box-shadow:0 1px 3px rgba(0,0,0,0.04)">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;
        margin-bottom:12px;flex-wrap:wrap;gap:10px">
        <div style="display:flex;align-items:center;gap:10px">
          <div style="width:40px;height:40px;border-radius:50%;background:${cs.color};
            color:#fff;display:flex;align-items:center;justify-content:center;
            font-family:Georgia,serif;font-size:16px;font-weight:700;flex-shrink:0">
            ${(sc.name||"?")[0].toUpperCase()}
          </div>
          <div>
            <div style="font-size:14px;font-weight:700;color:#241008">${sc.name}</div>
            <div style="font-size:10.5px;color:${cs.color};font-weight:600">${cs.numeral} ${cs.name}</div>
            ${sc.group ? `<div style="font-size:10px;color:#9A7A60">${sc.group}</div>` : ""}
          </div>
        </div>
        <div style="display:flex;align-items:flex-start;gap:10px">
          <div style="text-align:right">
            <div style="font-size:24px;font-weight:700;color:#8B1A1A;font-family:Georgia,serif;line-height:1">${tot.pct}%</div>
            <div style="font-size:9.5px;color:#9A7A60">${tot.d}/${tot.t} objetivos</div>
            <div style="font-size:9.5px;color:#9A7A60;margin-top:2px">Últ. act.: ${lastDate}</div>
            ${att.t ? `<div style="font-size:9.5px;color:#2E86C1">Asistencia: ${att.pct}%</div>` : ""}
          </div>
          <img src="${qrUrl(qrBase + encodeURIComponent(sc.name))}" width="60" height="60"
            style="border-radius:5px;border:1px solid #EAE0D2;flex-shrink:0" alt="QR"/>
        </div>
      </div>
      ${stageBars}
    </div>`;
  }).join("");

  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<title>Reporte General — Clan Scout No. 12</title>
<style>
  *{box-sizing:border-box}
  body{font-family:'Helvetica Neue',Arial,sans-serif;max-width:900px;margin:0 auto;
    padding:0 0 50px;color:#241008;background:#FAF6F0;font-size:12px}
  @media print{
    body{background:#fff;padding:0}
    .noprint{display:none!important}
    @page{margin:1.5cm;size:A4}
  }
</style></head><body>

${institutionalHeader("Reporte General del Clan", `${filtered.length} Rover${filtered.length===1?"":"s"} · ${today}`, today)}

<div style="padding:16px 20px">
  ${statsBlock(filtered, prog, attendance)}

  <button class="noprint" onclick="window.print()"
    style="display:block;margin:0 auto 18px;padding:9px 22px;background:#8B1A1A;color:#fff;
    border:none;border-radius:7px;cursor:pointer;font-size:13px;font-weight:600">
    🖨️ Imprimir / Guardar como PDF
  </button>

  ${cards}

  <div style="margin-top:20px;padding-top:10px;border-top:1px solid #EAE0D2;
    display:flex;justify-content:space-between;font-size:9px;color:#B9A080;flex-wrap:wrap;gap:4px">
    <span>Grupo Scout No. 12 · San Pedro Apóstol · Guatemala</span>
    <span>Cartilla 2.0 · Progresión Rover</span>
    <span>Emitida el ${today}</span>
  </div>
</div>
</body></html>`;
}

// ─── RF-04: CSV mejorado con estadísticas ─────────────────────────────────────
function exportAllCSV(scouts, prog, dates, stagesList) {
  // Use provided stagesList (activeStagesList) or fall back to STAGES
  const stages = (stagesList && stagesList.length > 0) ? stagesList : STAGES;
  const today = new Date().toISOString().split("T")[0];
  const rows = [
    ["PROGRESIÓN ROVER — CLAN SCOUT NO. 12"],
    [`Generado el ${today}`, "", "", "", `Total Rovers: ${scouts.length}`],
    [],
    ["Rover","Grupo/Clan","Etapa Actual","% Etapa Actual","% Total","Etapa","Sección","Área","Requisito","Completado","Fecha"],
  ];
  scouts.forEach(scout => {
    const cs  = scoutCurrentStage(scout, prog);
    const cs_s = scoutStageStats(scout, cs, prog);
    const tot = scoutTotalStats(scout, prog);
    stages.forEach(st => (st.sections||[]).forEach(sec => {
      // Fix F4: filter active:false; Fix F2: keep itemObj for stable key
      const secItems = (sec.items||[]).filter(it => it.active !== false);
      secItems.forEach((rawItem) => {
        const itemObj = typeof rawItem === "object" ? rawItem : null;
        const itemText = itemObj ? (itemObj.text || String(rawItem)) : String(rawItem);
        let done = false, dateVal = "";
        if (itemObj && itemObj.stableId && st.stableId && sec.stableId) {
          const sk = pKeyS(scout.id, st.stableId, sec.stableId, itemObj.stableId);
          const li = itemObj.legacyIdx !== undefined ? itemObj.legacyIdx : null;
          const ln = itemObj.legacySecName || sec.name;
          done = !!(prog[sk] !== undefined ? prog[sk]
            : (li !== null ? prog[pKey(scout.id, st.id, ln, li)] : false));
          dateVal = dates[dKeyS(scout.id, st.stableId, sec.stableId, itemObj.stableId)]
            || (li !== null ? dates[dKey(scout.id, st.id, ln, li)] : "") || "";
        } else {
          const idx = sec.items.indexOf(rawItem);
          done = !!prog[pKey(scout.id, st.id, sec.name, idx)];
          dateVal = dates[dKey(scout.id, st.id, sec.name, idx)] || "";
        }
        rows.push([
          scout.name, scout.group||"",
          cs.name, cs_s.pct+"%", tot.pct+"%",
          st.name, sec.name, sec.area, itemText,
          done ? "Sí" : "No", dateVal
        ]);
      });
    }));
  });
  const csv = rows.map(r => r.map(c => `"${String(c||"").replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Progresion_Clan_${today}.csv`;
  a.click();
}

// ─── EXPORT: CSV — ASISTENCIA ─────────────────────────────────────────────────
function exportAllAttendanceCSV(scouts, attendance) {
  const today = new Date().toISOString().split("T")[0];
  const rows = [["Rover","Grupo/Clan","Fecha","Estado","Observaciones"]];
  scouts.forEach(scout => {
    Object.entries(attendance||{})
      .filter(([k]) => k.startsWith(`att:${scout.id}|`))
      .map(([k,v]) => ({date:k.split("|")[1],status:v.status,note:v.note||""}))
      .sort((a,b) => a.date.localeCompare(b.date))
      .forEach(r => rows.push([
        scout.name, scout.group||"", r.date,
        (ATT_STATES[r.status]||{}).label||r.status, r.note
      ]));
  });
  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Asistencia_Clan_${today}.csv`;
  a.click();
}

// ─── PANTALLA DE LOGIN (solo para administrador) ──────────────────────────────
function LoginScreen({ onClose }) {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    setErr(""); setBusy(true);
    try {
      await signInWithEmailAndPassword(auth, email.trim(), pass);
      onClose();
    } catch (ex) {
      setErr("Correo o contraseña incorrectos.");
    }
    setBusy(false);
  };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(5,2,1,0.85)",
      display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000,
      padding:"16px" }} onClick={onClose}>
      <form onClick={e => e.stopPropagation()} onSubmit={handleLogin}
        style={{ background:DS.bg.card, border:"1px solid rgba(197,162,83,0.25)",
          borderRadius:"14px", padding:"28px 26px", width:"100%", maxWidth:"340px",
          display:"flex", flexDirection:"column", gap:"12px" }}>
        <div style={{ textAlign:"center", marginBottom:"6px" }}>
          <div style={{ fontSize:"30px", color:DS.brand.red }}>⚜</div>
          <div style={{ fontFamily:"'Cinzel',serif", fontSize:"15px", fontWeight:"700",
            color:DS.text.primary, marginTop:"4px" }}>Acceso de Dirigente</div>
          <div style={{ fontSize:"11.5px", color:DS.text.secondary, marginTop:"3px" }}>
            Inicia sesión para editar la progresión
          </div>
        </div>
        <input className="login-input" type="email" placeholder="Correo electrónico"
          value={email} onChange={e => setEmail(e.target.value)} required autoFocus
          style={{ background:"rgba(255,255,255,0.05)", border:"1px solid rgba(255,255,255,0.12)",
            borderRadius:"7px", padding:"10px 12px", color:DS.text.primary, fontSize:"13px", outline:"none" }} />
        <input className="login-input" type="password" placeholder="Contraseña"
          value={pass} onChange={e => setPass(e.target.value)} required
          style={{ background:"rgba(255,255,255,0.05)", border:"1px solid rgba(255,255,255,0.12)",
            borderRadius:"7px", padding:"10px 12px", color:DS.text.primary, fontSize:"13px", outline:"none" }} />
        {err && <div style={{ fontSize:"12px", color:"#E74C3C" }}>{err}</div>}
        <button className="login-btn btn-primary" type="submit" disabled={busy}
          style={{ marginTop:"4px", padding:"10px", fontSize:"14px",
            fontFamily:"'Cinzel',serif", fontWeight:"700",
            justifyContent:"center", letterSpacing:"0.5px" }}>
          {busy ? "Ingresando…" : "Ingresar"}
        </button>
        <button type="button" onClick={onClose}
          className="btn-secondary"
          style={{ fontSize:"12px", marginTop:"2px", justifyContent:"center" }}>
          Cancelar — solo ver
        </button>
      </form>
    </div>
  );
}

// ─── DOCUMENTOS FIRESTORE ─────────────────────────────────────────────────────
const DOC_REF         = doc(db, "rover_app", "data");
const STAGES_DOC_REF  = doc(db, "rover_app", "stages_config");
const AUDIT_DOC_REF   = doc(db, "rover_app", "audit_log");
const TRASH_DOC_REF   = doc(db, "rover_app", "trash");
const EMPTY_DATA = { scouts: [], prog: {}, dates: {}, invDates: {}, attendance: {}, order: [] };

// ─── Genera configuración inicial de etapas desde STAGES hardcoded ────────────
// Cada objetivo obtiene un ID estable basado en stageId + secName + idx
// Esto garantiza 100% compatibilidad con claves Firebase existentes
function buildDefaultStagesConfig() {
  const cfg = {};
  STAGES.forEach(st => {
    const stageCfg = {
      id: st.id,
      stableId: `stage_${st.id}`,
      numeral: st.numeral,
      name: st.name,
      subtitle: st.subtitle,
      description: st.description,
      investidura: st.investidura,
      color: st.color,
      bg: st.bg,
      areas: st.areas,
      areaDesc: st.areaDesc,
      active: true,
      order: st.id,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      sections: st.sections.map((sec, si) => ({
        stableId: `sec_${st.id}_${si}`,
        name: sec.name,
        area: sec.area,
        active: true,
        order: si,
        items: sec.items.map((item, idx) => ({
          // stableId preserves firebase key compatibility: stageId|secName|idx
          stableId: `item_${st.id}_${si}_${idx}`,
          // legacyKey matches existing Firebase keys exactly
          legacySecName: sec.name,
          legacyIdx: idx,
          text: item,
          active: true,
          order: idx,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        }))
      }))
    };
    cfg[`stage_${st.id}`] = stageCfg;
  });
  return cfg;
}

// Toast constants at module level (must be defined before JSX render)
const TOAST_ICONS  = { success:"✓", error:"⚠", info:"●", warning:"◐" };
const TOAST_COLORS_MAP = {
  success:"#27AE60", error:"#C0392B", info:"#5B8FA8", warning:"#E67E22"
};

export default function RoverDashboard() {
  // ── Auth ──
  const [user, setUser] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [showLogin, setShowLogin] = useState(false);
  const isAdmin = !!user;

  // ── Datos (sincronizados con Firestore) ──
  const [scouts, setScouts] = useState([]);
  const [stagesConfig, setStagesConfig] = useState(null);
  const [stagesConfigSynced, setStagesConfigSynced] = useState(false);
  const [prog, setProg]     = useState({});
  const [dates, setDates]   = useState({});
  const [invDates, setInvDates] = useState({});
  const [attendance, setAttendance] = useState({});
  const [loading, setLoading] = useState(true);
  const [saveErr, setSaveErr] = useState(false);

  // ── UI local ──
  const [selId, setSelId]       = useState(null);
  const [view, setView]         = useState("main");
  const [stageId, setStageId]   = useState(1);
  const [adding, setAdding]     = useState(false);
  const [newName, setNewName]   = useState("");
  const [newGroup, setNewGroup] = useState("");
  const [nameErr, setNameErr]   = useState("");        // RF-03: validación en tiempo real
  const [scoutSearch, setScoutSearch] = useState(""); // RF-08: búsqueda por nombre
  // Fase 9: panel de clan
  const [clanSearch, setClanSearch]       = useState("");
  const [clanSort, setClanSort]           = useState("name"); // "name"|"stage"|"pct"|"date"
  const [clanSortDir, setClanSortDir]     = useState(1);      // 1=asc -1=desc
  const [clanFilter, setClanFilter]       = useState("all");  // "all"|stageId|"initiated"|"done"|"priority"|"inactive"
  const [priorityIds, setPriorityIds]     = useState(new Set());
  // Fase 13: accesibilidad
  const [a11yPrefs, setA11yPrefs] = useState(() => {
    try { return JSON.parse(localStorage.getItem("rv_a11y")||"{}"); } catch { return {}; }
  });
  const saveA11y = (upd) => {
    const next = { ...a11yPrefs, ...upd };
    setA11yPrefs(next);
    try { localStorage.setItem("rv_a11y", JSON.stringify(next)); } catch {}
  };

  // RF-07: Guardar cambios de etapa/objetivo en Firebase
  const saveStagesConfig = (newCfg) => {
    setStagesConfig(newCfg);
    setDoc(STAGES_DOC_REF, newCfg)
      .then(() => { showSync(); addToast("Configuración guardada", "success", 2000); })
      .catch(() => addToast("Error al guardar", "error", 3000));
  };
  const updateStage = (stableId, changes) => {
    const newCfg = {
      ...stagesConfig,
      [stableId]: {
        ...stagesConfig[stableId],
        ...changes,
        updatedAt: new Date().toISOString()
      }
    };
    saveStagesConfig(newCfg);
    addAudit(`Etapa actualizada: ${changes.name || stableId}`, { type:"stage", elementId: stableId });
  };
  // RF-02/RF-03: updateObjective always uses objectiveStableId — no positional fallback
  // After normalization, all objectives have stableId. secIdx only used to scope the search.
  const updateObjective = (stageStableId, sectionStableId, objectiveStableId, changes) => {
    const stage = stagesConfig[stageStableId];
    if (!stage) return;
    let found = false;
    const newSections = stage.sections.map(sec => {
      // Match section by stableId, fall back to name for legacy sections
      const secMatch = sec.stableId === sectionStableId;
      if (!secMatch) return sec;
      const newItems = sec.items.map(it => {
        if (typeof it !== "object" || it.stableId !== objectiveStableId) return it;
        found = true;
        const { _stableId, ...cleanChanges } = changes;
        return { ...it, ...cleanChanges, updatedAt: new Date().toISOString() };
      });
      return { ...sec, items: newItems };
    });
    if (!found) {
      console.warn(`updateObjective: objective ${objectiveStableId} not found in stage ${stageStableId}`);
      return;
    }
    updateStage(stageStableId, { sections: newSections });
    addAudit(`Objetivo actualizado en etapa ${stageStableId}`, { type:"objective", elementId: objectiveStableId });
  };
  const addObjective = (stageStableId, sectionStableId, text) => {
    const stage = stagesConfig[stageStableId];
    if (!stage || !text.trim()) return;
    // Find section by stableId (or name as legacy fallback)
    const sec = stage.sections.find(s => s.stableId === sectionStableId);
    if (!sec) return;
    const newItem = {
      stableId: makeStableId("item"),  // RF-04: robust permanent ID
      legacySecName: sec.name,
      legacyIdx: sec.items.length,
      text: text.trim(),
      active: true,
      order: sec.items.length,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    const newSections = stage.sections.map(s =>
      s.stableId === sectionStableId
        ? { ...s, items: [...s.items, newItem] }
        : s
    );
    updateStage(stageStableId, { sections: newSections });
    addAudit(`Objetivo agregado: "${text.trim()}"`, { type:"objective", elementId: newItem.stableId });
  };

  // RF-03: Reorder objectives within a section (drag-free: move up/down)
  const reorderObjective = (stageStableId, sectionStableId, fromIdx, toIdx) => {
    const stage = stagesConfig[stageStableId];
    if (!stage) return;
    const newSections = stage.sections.map(sec => {
      if (sec.stableId !== sectionStableId) return sec;
      const items = [...sec.items];
      const [moved] = items.splice(fromIdx, 1);
      items.splice(toIdx, 0, moved);
      return { ...sec, items: items.map((it, i) =>
        typeof it === "object" ? { ...it, order: i } : it
      )};
    });
    updateStage(stageStableId, { sections: newSections });
    addAudit(`Objetivo reordenado en etapa ${stageStableId}`, { type:"objective", elementId: stageStableId });
  };

  // RF-03: Reorder stages
  const reorderStage = (fromId, direction) => {
    const list = Object.values(stagesConfig).sort((a,b) => (a.order||a.id)-(b.order||b.id));
    const fromIdx = list.findIndex(s => s.stableId === fromId);
    const toIdx = fromIdx + direction;
    if (toIdx < 0 || toIdx >= list.length) return;
    const newCfg = { ...stagesConfig };
    newCfg[list[fromIdx].stableId] = { ...newCfg[list[fromIdx].stableId], order: toIdx };
    newCfg[list[toIdx].stableId]   = { ...newCfg[list[toIdx].stableId],   order: fromIdx };
    saveStagesConfig(newCfg);
    addAudit(`Etapa reordenada: ${list[fromIdx].name}`, { type:"stage" });
  };

  // RF-03: Add a new objective by typing its text (already exists as addObjective)
  // createStage: creates a new stage in Firebase config
  // RF-04: Robust stableId — not dependent on time, position, or order
  const makeStableId = (prefix) => {
    const rnd = Math.random().toString(36).slice(2, 9);
    const ts  = Date.now().toString(36).slice(-4);
    return `${prefix}_${ts}${rnd}`;
  };

  const createStage = (name) => {
    if (!name.trim() || !stagesConfig) return;
    const stableId = makeStableId("stage");
    // id is a numeric alias for legacy key compatibility — stableId is the canonical key
    const numericId = Date.now();
    const order = Math.max(0, ...Object.values(stagesConfig).map(s => s.order ?? 0)) + 1;
    const newStage = {
      id: numericId,         // legacy key alias only — NOT used as identity
      stableId,              // canonical permanent identity
      numeral: `${order+1}`.padStart(2,"0"),
      name: name.trim(), subtitle: "", description: "",
      investidura: `Investidura ${name.trim()}`,
      color: DS.brand.gold, bg: "rgba(197,162,83,0.06)",
      areas: [], active: true, order,
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
      sections: [],
    };
    const newCfg = { ...stagesConfig, [stableId]: newStage };
    saveStagesConfig(newCfg);
    addAudit(`Nueva etapa creada: ${name.trim()}`, { type:"stage", elementId: stableId });
    addToast(`Etapa "${name.trim()}" creada`, "success", 2500);
  };

  // RF-02/P2: Update section name/area by sectionStableId
  const updateSection = (stageStableId, sectionStableId, changes) => {
    const stage = stagesConfig[stageStableId];
    if (!stage) return;
    const newSections = stage.sections.map(sec =>
      sec.stableId === sectionStableId
        ? { ...sec, ...changes, updatedAt: new Date().toISOString() }
        : sec
    );
    updateStage(stageStableId, { sections: newSections });
    addAudit(`Sección actualizada en etapa ${stageStableId}`, { type:"stage", elementId: sectionStableId });
  };

  // RF-02/P2: Reorder section by sectionStableId — stableId never changes
  const reorderSection = (stageStableId, sectionStableId, direction) => {
    const stage = stagesConfig[stageStableId];
    if (!stage) return;
    const sections = [...(stage.sections || [])];
    const idx = sections.findIndex(s => s.stableId === sectionStableId);
    if (idx < 0) return;
    const newIdx = idx + direction;
    if (newIdx < 0 || newIdx >= sections.length) return;
    // Swap — stableIds stay with their sections, only position changes
    [sections[idx], sections[newIdx]] = [sections[newIdx], sections[idx]];
    // Update order field for all sections
    const ordered = sections.map((s, i) => ({ ...s, order: i }));
    updateStage(stageStableId, { sections: ordered });
    addAudit(`Sección reordenada en etapa ${stageStableId}`, { type:"stage", elementId: sectionStableId });
  };

  // RF-03: Add a new section to a stage
  const addSection = (stageStableId, sectionName, area) => {
    const stage = stagesConfig[stageStableId];
    if (!stage || !sectionName.trim()) return;
    const newSection = {
      stableId: makeStableId("sec"),
      name: sectionName.trim(),
      area: area.trim() || sectionName.trim(),
      active: true, order: (stage.sections||[]).length,
      items: [],
    };
    updateStage(stageStableId, { sections: [...(stage.sections||[]), newSection] });
    addAudit(`Sección agregada: "${sectionName.trim()}"`, { type:"stage", elementId: stageStableId });
  };
  const sidebarSearchRef = useRef(null);
  const scrollRef = useRef(null);

  // Fase 11: módulo de configuración
  const [settingsTab, setSettingsTab] = useState("institution"); // institution|system|help|audit|trash
  const [sysConfig, setSysConfig]     = useState(() => {
    try { return JSON.parse(localStorage.getItem("rv_config")||"{}"); } catch { return {}; }
  });
  const saveSysConfig = (upd) => {
    const next = { ...sysConfig, ...upd };
    setSysConfig(next);
    try { localStorage.setItem("rv_config", JSON.stringify(next)); } catch {}
  };
  // ── RF-01/RF-05: Audit log + Trash — Firebase persistent (Bloque 2) ──────
  const [auditLog, setAuditLog] = useState([]);
  const [auditLoaded, setAuditLoaded] = useState(false);
  const [trash, setTrash]       = useState([]);
  const [trashLoaded, setTrashLoaded] = useState(false);

  // RF-04/RF-11: lazy-load audit when settings tab opens
  const loadAudit = () => {
    if (auditLoaded) return;
    perfRef.current.firebaseReads++;
    import("firebase/firestore").then(({ getDoc, setDoc: _setDoc }) => {
      getDoc(AUDIT_DOC_REF).then(snap => {
        // Step 1: Load existing Firebase entries
        const firebaseData = snap.exists() ? snap.data() : {};
        const firebaseIds = new Set(Object.keys(firebaseData));
        const entries = Object.values(firebaseData)
          .sort((a,b) => (b.ts||"").localeCompare(a.ts||""))
          .slice(0, 100);
        setAuditLog(entries);
        setAuditLoaded(true);

        // Step 2: Check for local records to migrate
        let local = [];
        try { local = JSON.parse(localStorage.getItem("rv_audit")||"[]"); } catch {}
        if (!local.length) return;

        // Step 3: Validate local records
        const valid = local.filter(e => e && e.ts && e.action);

        // Step 4: Generate stable ID for legacy records (based on ts+action hash)
        const toMigrate = {};
        valid.forEach(e => {
          // Use existing id if present, otherwise derive from ts+action
          const stableKey = e.id || `legacy_${(e.ts||"").replace(/\W/g,"")}_${
            (e.action||"").slice(0,10).replace(/\W/g,"")
          }`;
          // Step 5: Skip duplicates already in Firebase
          if (firebaseIds.has(stableKey)) return;
          toMigrate[stableKey] = { ...e, id: stableKey, migrated: true };
        });

        if (!Object.keys(toMigrate).length) {
          // All records already migrated — safe to clean localStorage
          try { localStorage.removeItem("rv_audit"); } catch {}
          return;
        }

        // Step 6+7: Write only missing records to Firebase
        const merged = { ...firebaseData, ...toMigrate };
        updateDoc(AUDIT_DOC_REF, toMigrate)
          .catch(() => _setDoc(AUDIT_DOC_REF, merged, { merge: true }))
          .then(() => {
            // Step 8: Confirmed — mark migration complete
            // Step 9: Only NOW remove from localStorage
            try { localStorage.removeItem("rv_audit"); } catch {}
            // Update local state with migrated entries
            const allEntries = Object.values(merged)
              .sort((a,b) => (b.ts||"").localeCompare(a.ts||""))
              .slice(0, 100);
            setAuditLog(allEntries);
          })
          .catch(() => {
            // Firebase failed — DO NOT remove localStorage
            // Migration will be retried next time loadAudit() is called
            console.warn("Audit migration to Firebase failed — rv_audit preserved in localStorage");
          });
      }).catch(() => {
        // Firebase read failed — mark as loaded with empty, preserve localStorage
        setAuditLoaded(true);
      });
    });
  };
  const loadTrash = () => {
    if (trashLoaded) return;
    perfRef.current.firebaseReads++;
    import("firebase/firestore").then(({ getDoc }) => {
      getDoc(TRASH_DOC_REF).then(snap => {
        if (snap.exists()) {
          const data = snap.data();
          const items = Object.values(data)
            .filter(v => v && v.id)
            .sort((a,b) => (b.deletedAt||"").localeCompare(a.deletedAt||""));
          setTrash(items);
        }
        setTrashLoaded(true);
        // RF-10: one-time migration from localStorage
        try {
          const local = JSON.parse(localStorage.getItem("rv_trash")||"[]");
          if (local.length > 0 && !snap.exists()) {
            // migrate to Firebase
            const migData = {};
            local.forEach(item => { migData[item.id] = item; });
            setDoc(TRASH_DOC_REF, migData).catch(()=>{});
            setTrash(local);
          }
          localStorage.removeItem("rv_trash");
        } catch {}
      });
    });
  };

  // RF-01: Write audit entry to Firebase
  const addAudit = (action, meta = {}) => {
    const entry = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2,5),
      ts: new Date().toISOString(),
      action,
      user: auth.currentUser?.email || "Admin",
      ...meta,
    };
    // Optimistic update local state
    setAuditLog(prev => [entry, ...prev].slice(0, 100));
    // Persist to Firebase (updateDoc to merge, not overwrite)
    updateDoc(AUDIT_DOC_REF, { [entry.id]: entry })
      .catch(() => setDoc(AUDIT_DOC_REF, { [entry.id]: entry }, { merge: true }));
  };

  // RF-05: Move to Firebase trash
  const moveToTrash = (item) => {
    const entry = {
      ...item,
      deletedAt: new Date().toISOString(),
      deletedBy: auth.currentUser?.email || "Admin",
    };
    setTrash(prev => [entry, ...prev]);
    updateDoc(TRASH_DOC_REF, { [item.id]: entry })
      .catch(() => setDoc(TRASH_DOC_REF, { [item.id]: entry }, { merge: true }));
    addAudit(`Rover enviado a papelera: ${item.name}`, { type:"trash", elementId: item.id });
  };

  // RF-06: Restore from Firebase trash
  const restoreFromTrash = (id) => {
    const item = trash.find(t => t.id === id);
    if (!item) return;
    // RF-06: check for conflict
    const conflict = dataRef.current.scouts.find(s => s.id === id);
    if (conflict) {
      addToast(`Ya existe un Rover con ese ID activo`, "error", 3000);
      return;
    }
    const { deletedAt, deletedBy, ...sc } = item;
    const upd = [...dataRef.current.scouts, sc];
    saveScouts(upd);
    // Remove from trash in Firebase
    updateDoc(TRASH_DOC_REF, { [id]: deleteField() })
      .catch(() => {});
    setTrash(prev => prev.filter(t => t.id !== id));
    addToast(`${sc.name} restaurado`, "success", 2200);
    addAudit(`Rover restaurado desde papelera: ${sc.name}`, { type:"restore", elementId: id });
  };
  // Fase 10: panel de reportes
  const [reportPreview, setReportPreview] = useState(null); // {html, title}
  const [reportFilter, setReportFilter]   = useState("all");
  // RF-08/RF-09: Backup & Restore state
  const [backupProcessing, setBackupProcessing] = useState(false);
  const [restoreState, setRestoreState]   = useState(null);
  const [pdfProgress, setPdfProgress]     = useState(null); // null or string message // null|{step,parsed,validation,conflicts}
  const INACTIVE_DAYS_R = 21;
  const [quickViewId, setQuickViewId]     = useState(null);
  const togglePriority = (id) => setPriorityIds(p => { const n = new Set(p); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const [scoutStageFilter, setScoutStageFilter] = useState("all"); // RF-08: filtro por etapa
  const [editingScoutId, setEditingScoutId] = useState(null); // RF-05: edición inline
  const [editingName, setEditingName] = useState("");
  const [expItem, setExpItem]   = useState(null);
  const [exportMenu, setExportMenu] = useState(false);
  const [collapsedSecs, setCollapsedSecs] = useState({});
  const [collapsedStages, setCollapsedStages] = useState({});
  const [detailFilter, setDetailFilter] = useState("all");
  const [poppingItems, setPoppingItems] = useState(new Set()); // RF-11: checkPop keys
  const [attMonth, setAttMonth] = useState(() => {
    const n = new Date();
    return { y: n.getFullYear(), m: n.getMonth() };
  });
  const [sidebarOpen, setSidebarOpen] = useState(false); // móvil
  // RF-05: sistema de toasts
  const [toasts, setToasts] = useState([]);
  // RF-05: confirm dialog (acciones críticas)
  const [confirmDlg, setConfirmDlg] = useState(null); // {message, onConfirm, danger}
  const showConfirm = (message, onConfirm, danger = true) =>
    setConfirmDlg({ message, onConfirm, danger });
  const closeConfirm = () => setConfirmDlg(null);
  // RF-03: sync indicator
  const [syncing, setSyncing] = useState(false);
  // Fase 13 RF-09: global keyboard shortcuts
  useEffect(() => {
    const handler = (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      switch(e.key) {
        case "k": case "K":
          e.preventDefault();
          if (sidebarSearchRef.current) sidebarSearchRef.current.focus();
          break;
        case "1": e.preventDefault(); setView("main"); break;
        case "2": e.preventDefault(); if (isAdmin) setView("clan"); break;
        case "3": e.preventDefault(); if (isAdmin) setView("reports"); break;
        case "4": e.preventDefault(); if (isAdmin) setView("attendance"); break;
        case "5": e.preventDefault(); if (isAdmin) setView("settings"); break;
        default: break;
      }
    };
    const escHandler = (e) => {
      if (e.key !== "Escape") return;
      if (reportPreview) { setReportPreview(null); return; }
      if (confirmDlg) { closeConfirm(); return; }
      if (adding) setAdding(false);
    };
    window.addEventListener("keydown", handler);
    window.addEventListener("keydown", escHandler);
    return () => {
      window.removeEventListener("keydown", handler);
      window.removeEventListener("keydown", escHandler);
    };
  }, [isAdmin, reportPreview, confirmDlg, adding]); // eslint-disable-line

  useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = 0;
  }, [stageId]);

  const syncTimer = useRef(null);
  // RF-11: Performance monitoring (Fase 12)
  const perfRef = useRef({
    loadStart: performance.now(),
    loadEnd: null,
    syncTimes: [],
    renderCount: 0,
    firebaseReads: 0,
    cacheHits: 0,
  });

  const showSync = () => {
    const t0 = performance.now();
    setSyncing(true);
    clearTimeout(syncTimer.current);
    syncTimer.current = setTimeout(() => {
      setSyncing(false);
      // RF-11: record sync time
      const ms = Math.round(performance.now() - t0);
      perfRef.current.syncTimes = [...perfRef.current.syncTimes.slice(-19), ms];
    }, 1200);
  };
  const toastId = useRef(0);
  const addToast = (message, type = "info", duration = 2800) => {
    const id = ++toastId.current;
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), duration);
  };
  const dismissToast = (id) => setToasts(prev => prev.filter(t => t.id !== id));
  // TOAST_ICONS and TOAST_COLORS moved to module level (fix: used before declaration)

  const dataRef = useRef(EMPTY_DATA);

  // ── Auth listener ──
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => { setUser(u); setAuthReady(true); });
    return unsub;
  }, []);

  // Subscribe to stages config in Firebase
  useEffect(() => {
    const unsubStages = onSnapshot(STAGES_DOC_REF, (snap) => {
      if (snap.exists()) {
        const data = snap.data();
        if (data && Object.keys(data).length > 0) {
          // RF-02: normalize stableIds on every load (safe, idempotent migration)
          const { normalized, modified } = normalizeStagesConfig(data);
          setStagesConfig(normalized);
          setStagesConfigSynced(true);
          // Persist normalization silently if anything was missing
          if (modified) {
            setDoc(STAGES_DOC_REF, normalized).catch(() => {});
          }
          return;
        }
      }
      // No config in Firebase yet — initialize from hardcoded STAGES
      const defaultCfg = buildDefaultStagesConfig();
      // Also normalize the default config before saving
      const { normalized: normDefault } = normalizeStagesConfig(defaultCfg);
      setDoc(STAGES_DOC_REF, normDefault)
        .then(() => { setStagesConfig(normDefault); setStagesConfigSynced(true); });
    });
    return unsubStages;
  }, []);

  // ── Firestore: escucha en tiempo real ──
  const initialSelDone = useRef(false);
  const pendingWrites = useRef(0);
  useEffect(() => {
    const unsub = onSnapshot(DOC_REF,
      (snap) => {
        const d = snap.exists() ? snap.data() : EMPTY_DATA;
        const merged = { ...EMPTY_DATA, ...d };
        setLoading(false);
        setSaveErr(false);
        if (!initialSelDone.current && (merged.scouts || []).length) {
          initialSelDone.current = true;
          setSelId(merged.scouts[0].id);
        }
        // Si hay escrituras locales aún sin confirmar, no pisamos dataRef/estado
        // con este snapshot (podría reflejar un punto del servidor anterior a
        // un clic que el usuario ya hizo y que está en cola para guardarse).
        if (pendingWrites.current > 0) return;
        dataRef.current = merged;
        setScouts(merged.scouts || []);
        setProg(merged.prog || {});
        setDates(merged.dates || {});
        setInvDates(merged.invDates || {});
        setAttendance(merged.attendance || {});
      },
      (error) => { setLoading(false); setSaveErr(true); }
    );
    // RF-11: mark load complete (first snapshot = real load time)
    if (!perfRef.current.loadEnd) {
      perfRef.current.loadEnd = performance.now();
    }
    perfRef.current.firebaseReads++;
    return unsub;
  }, []);

  // ── Guardado: persiste en Firestore, en cola secuencial ──
  // (evita que clics rápidos disparen escrituras en paralelo que se pisen entre sí)
  const writeQueue = useRef(Promise.resolve());

  // persist({ campo: valor }) -> reemplaza el campo completo (para listas/objetos ya armados)
  const persist = (patch) => {
    const next = { ...dataRef.current, ...patch };
    dataRef.current = next;
    pendingWrites.current += 1;
    writeQueue.current = writeQueue.current
      .catch(() => {})
      .then(() => setDoc(DOC_REF, patch, { merge: true }))
      .then(() => { setSaveErr(false); showSync(); })
      .catch(() => { setSaveErr(true); addToast("Error al guardar. Verifica tu conexión.", "error"); })
      .finally(() => { pendingWrites.current = Math.max(0, pendingWrites.current - 1); });
    return writeQueue.current;
  };

  // persistField('mapa', 'clave', valor | null) -> escribe o BORRA una sola clave
  // dentro de un mapa anidado, usando notación de punto + deleteField().
  // Esto es necesario porque setDoc(...,{merge:true}) con un objeto completo
  // NUNCA borra claves que falten en el objeto enviado — solo las que pasamos
  // explícitamente con deleteField() se eliminan de verdad en el servidor.
  const persistField = (mapName, key, value) => {
    const nextMap = { ...dataRef.current[mapName] };
    if (value === null) delete nextMap[key]; else nextMap[key] = value;
    dataRef.current = { ...dataRef.current, [mapName]: nextMap };
    pendingWrites.current += 1;
    const fieldPath = `${mapName}.${key}`;
    writeQueue.current = writeQueue.current
      .catch(() => {})
      .then(() => updateDoc(DOC_REF, { [fieldPath]: value === null ? deleteField() : value }))
      .catch((err) => {
        // El documento aún no existe (raro, pero posible la primera vez):
        // si estamos ESCRIBIENDO, creamos el doc completo como red de seguridad.
        // Si estamos BORRANDO, no hay nada que borrar, lo ignoramos.
        if (value !== null) {
          return setDoc(DOC_REF, { [mapName]: nextMap }, { merge: true });
        }
      })
      .then(() => setSaveErr(false))
      .catch(() => setSaveErr(true))
      .finally(() => { pendingWrites.current = Math.max(0, pendingWrites.current - 1); });
    return writeQueue.current;
  };

  const saveScouts = (v) => { setScouts(v); persist({ scouts: v }); };
  const saveProg   = (v) => { setProg(v);   persist({ prog: v }); };
  const saveDates  = (v) => { setDates(v);  persist({ dates: v }); };
  const saveInv    = (v) => { setInvDates(v); persist({ invDates: v }); };
  const saveAtt    = (v) => { setAttendance(v); persist({ attendance: v }); };

  // ── Scouts CRUD (solo admin) ──
  const addScout = () => {
    if (!isAdmin) return;
    if (!newName.trim()) { setNameErr("El nombre es obligatorio"); return; }
    if (dataRef.current.scouts.some(s => s.name.trim().toLowerCase() === newName.trim().toLowerCase()))
      { setNameErr("Ya existe un Rover con ese nombre"); return; }
    const sc = { id: Date.now().toString(36) + Math.random().toString(36).slice(2,6), name: newName.trim(), group: newGroup.trim() };
    const upd = [...dataRef.current.scouts, sc];
    saveScouts(upd);
    setSelId(sc.id); setNewName(""); setNewGroup(""); setAdding(false); setNameErr("");
    addToast(`Rover ${sc.name} agregado`, "success", 2000);
    addAudit(`Rover agregado: ${sc.name}`, { type:"scout", elementId: sc.id });
  };
  const saveInlineName = (id) => {
    if (!editingName.trim()) { setEditingScoutId(null); return; }
    const upd = dataRef.current.scouts.map(s => s.id === id ? { ...s, name: editingName.trim() } : s);
    saveScouts(upd);
    setEditingScoutId(null);
    addToast("Nombre actualizado", "success", 1800);
    addAudit(`Nombre actualizado: ${editingName.trim()}`, { type:"scout" });
  };
  const deleteScout = (id) => {
    if (!isAdmin) return;
    const cur = dataRef.current;
    const upd = cur.scouts.filter(s => s.id !== id);
    saveScouts(upd); if (selId === id) setSelId(upd[0]?.id || null);

    const na = { ...cur.attendance };
    const keysToDelete = Object.keys(na).filter(k => k.startsWith(`att:${id}|`));
    if (keysToDelete.length) {
      keysToDelete.forEach(k => delete na[k]);
      dataRef.current = { ...dataRef.current, attendance: na };
      setAttendance(na);
      pendingWrites.current += 1;
      const updatePatch = {};
      keysToDelete.forEach(k => { updatePatch[`attendance.${k}`] = deleteField(); });
      writeQueue.current = writeQueue.current
        .catch(() => {})
        .then(() => updateDoc(DOC_REF, updatePatch))
        .catch(() => {}) // documento inexistente -> nada que borrar, se ignora
        .then(() => setSaveErr(false))
        .finally(() => { pendingWrites.current = Math.max(0, pendingWrites.current - 1); });
    }
  };
  const moveScout = (id, dir) => {
    if (!isAdmin) return;
    const list = dataRef.current.scouts;
    const i = list.findIndex(s => s.id === id);
    const j = i + dir;
    if (i < 0 || j < 0 || j >= list.length) return;
    const upd = [...list];
    [upd[i], upd[j]] = [upd[j], upd[i]];
    saveScouts(upd);
  };

  // ── Progresión ──
  const toggleItem = useCallback((stId, sec, idx, itemObj, stageObj, secObj) => {
    if (!isAdmin || !selId) return;
    const cur = dataRef.current;

    // Determine the key to use: stable if we have stableIds, legacy otherwise
    let k, dk;
    if (itemObj?.stableId && stageObj?.stableId && secObj?.stableId) {
      // Stable key (new writes always go here)
      k  = pKeyS(selId, stageObj.stableId, secObj.stableId, itemObj.stableId);
      dk = dKeyS(selId, stageObj.stableId, secObj.stableId, itemObj.stableId);
    } else {
      // Legacy fallback (items without stableId)
      k  = pKey(selId, stId, sec, idx);
      dk = dKey(selId, stId, sec, idx);
    }

    // Check current state — read from stable first, legacy fallback
    const wasChecked = itemObj && stageObj && secObj
      ? getItemProg(selId, stageObj, secObj, itemObj, cur.prog)
      : !!cur.prog[k];
    const now = !wasChecked;

    const np = { ...cur.prog, [k]: now };
    const ndt = { ...cur.dates };
    if (now && !ndt[dk]) ndt[dk] = new Date().toISOString().split("T")[0];
    setProg(np); setDates(ndt);
    persist({ prog: np, dates: ndt });
    addToast(wasChecked ? "Objetivo desmarcado" : "Objetivo completado ✓",
      wasChecked ? "info" : "success", 1800);
    setPoppingItems(p => { const n = new Set(p); n.add(k); return n; });
    setTimeout(() => setPoppingItems(p => { const n = new Set(p); n.delete(k); return n; }), 300);
  }, [isAdmin, selId]); // eslint-disable-line
  const setItemDate = (stId, sec, idx, val, itemObj, stageObj, secObj) => {
    if (!isAdmin || !selId) return;
    let dk;
    if (itemObj?.stableId && stageObj?.stableId && secObj?.stableId) {
      dk = dKeyS(selId, stageObj.stableId, secObj.stableId, itemObj.stableId);
    } else {
      dk = dKey(selId, stId, sec, idx);
    }
    const nd = { ...dataRef.current.dates, [dk]: val };
    saveDates(nd);
    showSync(); // RF-09: confirma autoguardado visualmente
  };
  const toggleInv = (stId) => {
    if (!isAdmin || !selId) return;
    const cur = dataRef.current;
    const k = invKey(selId, stId);
    const wasSet = !!cur.invDates[k];
    const ni = { ...cur.invDates, [k]: wasSet ? "" : new Date().toISOString().split("T")[0] };
    saveInv(ni);
    // RF-05: toast de confirmación
    addToast(wasSet ? "Investidura desmarcada" : "🎖️ Investidura registrada", wasSet ? "info" : "success", 2200);
  };
  const toggleSec = useCallback((key) => setCollapsedSecs(p => ({ ...p, [key]: !p[key] })), []);

  // ── Asistencia (solo admin marca; todos ven) ──
  const cycleAttStatus = (scoutId, date) => {
    if (!isAdmin) return;
    const cur = dataRef.current.attendance;
    const k = attKey(scoutId, date);
    const curr = cur[k];
    if (!curr) {
      const val = { status: "presente", note: "" };
      setAttendance({ ...cur, [k]: val });
      persistField("attendance", k, val);
      return;
    }
    const idx = ATT_CYCLE.indexOf(curr.status);
    if (idx === ATT_CYCLE.length - 1) {
      // último estado del ciclo -> borrar la clave de verdad (deleteField)
      const na = { ...cur };
      delete na[k];
      setAttendance(na);
      persistField("attendance", k, null);
    } else {
      const val = { ...curr, status: ATT_CYCLE[idx + 1] };
      setAttendance({ ...cur, [k]: val });
      persistField("attendance", k, val);
    }
  };

  // ── Stats helpers (atados a este componente) ──
  const stageStats = (scoutId, stage) => {
    let t = 0, d = 0;
    stage.sections.forEach(sec => {
      const secItems = (sec.items||[]).filter(it => it.active !== false);
      secItems.forEach(item => {
        t++;
        if (getItemProg(scoutId, stage, sec, item, prog)) d++;
      });
    });
    return { t, d, pct: t ? Math.round(d / t * 100) : 0 };
  };
  const totalStats = (scoutId) => {
    let t = 0, d = 0;
    activeStagesList.forEach(st => { const s = stageStats(scoutId, st); t += s.t; d += s.d; });
    return { t, d, pct: t ? Math.round(d / t * 100) : 0 };
  };
  const getCurrStage = (scoutId) => {
    const list = activeStagesList;
    for (let i = list.length - 1; i >= 0; i--) {
      const s = stageStats(scoutId, list[i]);
      if (s.pct === 100 && i < list.length - 1) return list[i + 1];
      if (s.d > 0) return list[i];
    }
    return list[0];
  };
  const areaScore = (scoutId, stage, area) => {
    let t = 0, d = 0;
    stage.sections.forEach(sec => {
      if (sec.area === area) sec.items.forEach((_, i) => {
        t++; if (prog[pKey(scoutId, stage.id, sec.name, i)]) d++;
      });
    });
    return t ? Math.round(d / t * 100) : 0;
  };
  const attendanceList = (scoutId) => Object.entries(attendance)
    .filter(([k]) => k.startsWith(`att:${scoutId}|`))
    .map(([k, v]) => ({ date: k.split("|")[1], status: v.status, note: v.note || "" }))
    .sort((a, b) => b.date.localeCompare(a.date));
  const attendanceStats = (scoutId) => {
    const list = attendanceList(scoutId);
    const t = list.length;
    const p = list.filter(r => r.status === "presente").length;
    const a = list.filter(r => r.status === "ausente").length;
    const e = list.filter(r => r.status === "excusa").length;
    return { t, p, a, e, pct: t ? Math.round(p / t * 100) : 0 };
  };

  // Derive active stages from Firebase config (or fallback to hardcoded STAGES)
  // CRITICAL: pKey/dKey/invKey still use stageId (1,2,3) and secName — never changes
  const activeStagesList = useMemo(() => {
    if (!stagesConfig) return STAGES; // use hardcoded while loading
    return Object.values(stagesConfig)
      .filter(s => s.active !== false)
      .sort((a, b) => (a.order||a.id) - (b.order||b.id))
      .map(cfg => {
        // Merge Firebase config with hardcoded visual properties
        // stageId (1,2,3) must remain to preserve Firebase key compatibility
        const fallback = STAGES.find(s => s.id === cfg.id) || STAGES[0];
        return {
          ...fallback,
          name: cfg.name || fallback.name,
          subtitle: cfg.subtitle || fallback.subtitle,
          description: cfg.description || fallback.description,
          investidura: cfg.investidura || fallback.investidura,
          active: cfg.active !== false,
          // sections rebuilt from config preserving legacySecName/legacyIdx
          sections: (cfg.sections || fallback.sections).map(sec => ({
            ...sec,
            name: sec.name,
            area: sec.area,
            items: (sec.items || [])
              .filter(it => it.active !== false)
              .sort((a, b) => (a.order||a.legacyIdx||0) - (b.order||b.legacyIdx||0))
              .map(it => typeof it === 'string' ? it : it.text),
          }))
        };
      });
  }, [stagesConfig]);

  const selScout = scouts.find(s => s.id === selId) || null;
  const activeSt = activeStagesList.find(s => s.id === stageId) || activeStagesList[0];
  const totStats = selScout ? totalStats(selId) : { t:0, d:0, pct:0 };

  // RF-04/RF-09: Memoize clan-wide metrics — recompute only when prog/dates/scouts change
  const clanMetrics = useMemo(() => {
    perfRef.current.cacheHits++;
    return scouts.map(sc => {
      const cs  = getCurrStage(sc.id);
      const ts  = totalStats(sc.id);
      const css = stageStats(sc.id, cs);
      const allDates = Object.entries(dates)
        .filter(([k]) => k.startsWith(`d:${sc.id}|`))
        .map(([,v]) => v).filter(Boolean).sort().reverse();
      const lastDate  = allDates[0] || null;
      const daysSince = lastDate
        ? Math.floor((Date.now() - new Date(lastDate)) / 86400000)
        : 9999;
      return { sc, cs, ts, css, lastDate, daysSince,
        initiated: ts.d > 0, completed: ts.pct === 100 };
    });
  }, [scouts, prog, dates]); // eslint-disable-line

  // RF-09: Memoize selected scout's stage stats
  const selectedStageStats = useMemo(() => {
    if (!selScout) return {};
    return activeStagesList.reduce((acc, st) => {
      acc[st.id] = stageStats(selId, st);
      return acc;
    }, {});
  }, [selId, prog, activeStagesList]); // eslint-disable-line

  // RF-04: sidebar list is filtered inline using clanMetrics; no extra state needed

  if (loading || !authReady) {
    // RF-09: Skeleton loader — simula la estructura real del dashboard
    return (
      <div className="rv-root" style={{ height:"100vh", display:"flex",
        flexDirection:"column", background:DS.bg.root, overflow:"hidden" }}>
        <style>{CSS}</style>
        {/* Header skeleton */}
        <div style={{ height:"52px", borderBottom:"1px solid rgba(255,255,255,0.04)",
          padding:"12px 18px", display:"flex", alignItems:"center", gap:"12px",
          background:"rgba(13,14,16,0.98)", flexShrink:0 }}>
          <div style={{ fontSize:"20px", color:DS.brand.red, opacity:.3 }}>⚜</div>
          <div className="rv-skeleton rv-skel-line" style={{ width:"140px", height:"12px" }} />
          <div style={{ flex:1 }} />
          <div className="rv-skeleton" style={{ width:"72px", height:"26px", borderRadius:"7px" }} />
        </div>
        {/* Body skeleton */}
        <div style={{ flex:1, display:"flex", overflow:"hidden" }}>
          {/* Sidebar skeleton */}
          <div style={{ width:"218px", flexShrink:0, background:DS.bg.surface,
            borderRight:"1px solid rgba(255,255,255,0.04)", padding:"12px 10px" }}>
            <div className="rv-skeleton rv-skel-line" style={{ width:"60px", height:"9px", marginBottom:"12px" }} />
            {[1,2,3].map(i => (
              <div key={i} style={{ padding:"10px 8px", marginBottom:"6px",
                background:"rgba(255,255,255,0.02)", borderRadius:"8px" }}>
                <div className="rv-skeleton rv-skel-line" style={{ width:"80%", height:"11px" }} />
                <div className="rv-skeleton rv-skel-line" style={{ width:"50%", height:"9px" }} />
                <div className="rv-skeleton rv-skel-bar" style={{ width:"100%", height:"3px" }} />
              </div>
            ))}
          </div>
          {/* Main skeleton */}
          <div style={{ flex:1, padding:"14px 16px", display:"flex", gap:"12px", overflow:"hidden" }}>
            {/* Stage nav skeleton */}
            <div style={{ width:"160px", flexShrink:0, display:"flex", flexDirection:"column", gap:"6px" }}>
              {[1,2,3].map(i => (
                <div key={i} style={{ padding:"10px 8px", borderRadius:"9px",
                  background:"rgba(255,255,255,0.02)" }}>
                  <div className="rv-skeleton rv-skel-line" style={{ width:"100%", height:"11px" }} />
                  <div className="rv-skeleton rv-skel-bar" style={{ width:"100%", height:"3px" }} />
                </div>
              ))}
            </div>
            {/* Content skeleton */}
            <div style={{ flex:1, display:"flex", flexDirection:"column", gap:"10px" }}>
              <div style={{ padding:"14px 16px", borderRadius:"12px",
                background:"rgba(14,7,2,0.97)", border:"1px solid rgba(255,255,255,0.04)" }}>
                <div className="rv-skeleton rv-skel-line" style={{ width:"30%", height:"9px", marginBottom:"6px" }} />
                <div className="rv-skeleton rv-skel-line" style={{ width:"55%", height:"15px", marginBottom:"4px" }} />
                <div className="rv-skeleton rv-skel-bar" style={{ width:"100%", height:"8px", borderRadius:"4px" }} />
                <div style={{ display:"flex", gap:"10px", marginTop:"8px" }}>
                  <div className="rv-skeleton rv-skel-line" style={{ width:"80px", height:"9px" }} />
                  <div className="rv-skeleton rv-skel-line" style={{ width:"70px", height:"9px" }} />
                </div>
              </div>
              {[1,2].map(i => (
                <div key={i} className="rv-skel-card">
                  <div className="rv-skeleton rv-skel-line" style={{ width:"40%", height:"10px" }} />
                  <div className="rv-skeleton rv-skel-line" style={{ width:"65%", height:"9px" }} />
                  <div className="rv-skeleton rv-skel-bar" style={{ width:"100%", height:"3px" }} />
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={`rv-root${a11yPrefs.reduceMotion?" rv-reduce-motion":""}${a11yPrefs.highContrast?" rv-high-contrast":""}${a11yPrefs.largeText?" rv-large-text":""}`}
      role="application" aria-label="Progresión Rover — Grupo Scout No. 12"
      style={{ height:"100vh", display:"flex", flexDirection:"column",
      background:DS.bg.root, color:DS.text.primary, overflow:"hidden" }}>
      <style>{CSS}</style>
      {showLogin && <LoginScreen onClose={() => setShowLogin(false)} />}

      {/* ── HEADER ─────────────────────────────────────────────────────────── */}
      <header className="rv-header" role="banner" aria-label="Encabezado principal"
        style={{ display:"flex", alignItems:"center",
        justifyContent:"space-between", padding:"10px 18px",
        borderBottom:"1px solid rgba(197,162,83,0.15)", flexShrink:0, gap:"8px",
        background:"rgba(13,14,16,0.98)" }}>
        <div style={{ display:"flex", alignItems:"center", gap:"11px", minWidth:0, flex:1 }}>
          <button className="hamburger" onClick={() => setSidebarOpen(o => !o)}
            aria-label="Abrir menú de Rovers"
            style={{ display:"none", background:"none", border:"1px solid rgba(255,255,255,0.1)",
              borderRadius:"6px", color:DS.brand.gold, width:"32px", height:"32px",
              cursor:"pointer", fontSize:"16px", flexShrink:0,
              alignItems:"center", justifyContent:"center" }}>☰</button>
          <span style={{ fontSize:"22px", color:DS.brand.red, lineHeight:1, flexShrink:0 }}>⚜</span>

          {/* RF-02: Encabezado principal — nombre del Rover + etapa cuando hay selección */}
          {selScout ? (
            <div style={{ minWidth:0, flex:1 }}>
              <div style={{ display:"flex", alignItems:"baseline", gap:"10px", flexWrap:"wrap" }}>
                <div className="rv-header-title" style={{ fontFamily:"'Cinzel',serif", fontSize:"17px", fontWeight:"700",
                  letterSpacing:"1px", color:DS.text.primary, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>
                  {selScout.name}
                </div>
                {(() => {
                  const cs = getCurrStage(selId);
                  return (
                    <div style={{ display:"flex", alignItems:"center", gap:"6px", flexShrink:0 }}>
                      <span style={{ fontSize:"10px", color:"rgba(255,255,255,0.2)" }}>—</span>
                      <span style={{ fontFamily:"'Cinzel',serif", fontSize:"11px", fontWeight:"600",
                        color:cs.color, letterSpacing:"0.5px", whiteSpace:"nowrap" }}>
                        {cs.numeral} · {cs.name}
                      </span>
                    </div>
                  );
                })()}
              </div>
              <div className="rv-header-sub" style={{ fontSize:"10px", color:DS.text.tertiary, letterSpacing:"1px",
                fontFamily:"'Source Sans 3',sans-serif", marginTop:"1px" }}>
                PROGRESIÓN ROVER · CARTILLA 2.0
              </div>
            </div>
          ) : (
            <div style={{ minWidth:0 }}>
              <div className="rv-header-title" style={{ fontFamily:"'Cinzel',serif", fontSize:"14px", fontWeight:"700",
                letterSpacing:"2.5px", color:DS.text.primary, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>PROGRESIÓN ROVER</div>
              <div className="rv-header-sub" style={{ fontSize:"11px", color:DS.text.tertiary, letterSpacing:"1.5px",
                fontFamily:"'Source Sans 3',sans-serif" }}>CARTILLA 2.0 · GUATEMALA</div>
            </div>
          )}

          <button onClick={() => setView("attendance")}
            aria-label="Cuadro de asistencia (Ctrl+4)"
            aria-current={view==="attendance"?"page":undefined}
            title="Asistencia (Ctrl+4)"
            style={{ marginLeft:"6px", padding:"6px 12px",
              background: view === "attendance" ? "rgba(91,143,168,0.22)" : "rgba(91,143,168,0.1)",
              border:"1px solid rgba(91,143,168,0.35)", borderRadius:"7px",
              color:DS.state.infoLight, fontSize:"12px", fontWeight:"600", cursor:"pointer",
              fontFamily:"'Source Sans 3',sans-serif", flexShrink:0 }}>
            📋
          </button>
          {isAdmin && (
            <>
              <button onClick={() => setView("clan")}
                aria-label="Panel del Clan (Ctrl+2)"
                aria-current={view==="clan"?"page":undefined}
                title="Panel Clan (Ctrl+2)"
                style={{ padding:"6px 12px",
                  background: view === "clan" ? "rgba(197,162,83,0.18)" : "rgba(197,162,83,0.07)",
                  border:"1px solid rgba(197,162,83,0.3)", borderRadius:"7px",
                  color: view === "clan" ? DS.brand.gold : DS.brand.goldMuted, fontSize:"12px",
                  fontWeight:"600", cursor:"pointer",
                  fontFamily:"'Source Sans 3',sans-serif", flexShrink:0 }}>
                ⚑ <span className="rv-nav-label">Clan</span>
              </button>
              <button onClick={() => setView("settings")}
                aria-label="Configuración (Ctrl+5)"
                aria-current={view==="settings"?"page":undefined}
                title="Configuración (Ctrl+5)"
                style={{ padding:"6px 10px",
                  background: view === "settings" ? "rgba(91,143,168,0.2)" : "rgba(91,143,168,0.07)",
                  border:"1px solid rgba(91,143,168,0.3)", borderRadius:"7px",
                  color: DS.state.infoLight, fontSize:"12px", fontWeight:"600",
                  cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif", flexShrink:0 }}>
                ⚙
              </button>
              <button onClick={() => setView("reports")}
                aria-label="Reportes y exportación (Ctrl+3)"
                aria-current={view==="reports"?"page":undefined}
                title="Reportes (Ctrl+3)"
                style={{ padding:"6px 12px",
                  background: view === "reports" ? "rgba(139,26,26,0.25)" : "rgba(139,26,26,0.08)",
                  border:"1px solid rgba(139,26,26,0.35)", borderRadius:"7px",
                  color: view === "reports" ? DS.brand.redLight || "#E04040" : DS.brand.red,
                  fontSize:"12px", fontWeight:"600", cursor:"pointer",
                  fontFamily:"'Source Sans 3',sans-serif", flexShrink:0 }}>
                📄 <span className="rv-nav-label">Reportes</span>
              </button>
            </>
          )}
        </div>

        <div style={{ display:"flex", alignItems:"center", gap:"8px", flexShrink:0 }}>
          {/* RF-03: sync indicator — aria-live already on toast stack */}
          {syncing && !saveErr && (
            <span aria-live="polite" aria-label="Guardando cambios"
              style={{ display:"flex", alignItems:"center", gap:"5px",
              fontSize:"10px", color:"rgba(197,162,83,0.6)",
              fontFamily:"'Source Sans 3',sans-serif" }}>
              <span className="rv-sync-dot" aria-hidden="true" />
              Guardando
            </span>
          )}
          {saveErr && (
            <span className="rv-badge error" style={{ fontSize:"10px" }}>⚠ Sin conexión</span>
          )}
          {isAdmin ? (
            <>
              <span className="rv-edit-badge rv-badge done" style={{ display:"inline-flex", gap:"5px" }}>
                <span style={{ width:"5px", height:"5px", borderRadius:"50%", background:DS.state.done, flexShrink:0, alignSelf:"center" }} />
                Editando
              </span>
              <button onClick={() => signOut(auth)}
                className="btn-secondary"
                aria-label="Cerrar sesión"
                title="Cerrar sesión"
                style={{ padding:"6px 12px",
                  fontFamily:"'Source Sans 3',sans-serif" }}>Salir</button>
            </>
          ) : (
            <button onClick={() => setShowLogin(true)}
              className="btn-primary"
              style={{ padding:"6px 14px", fontSize:"12px", whiteSpace:"nowrap",
                background:"rgba(139,26,26,0.18)", border:"1px solid rgba(201,48,44,0.35)" }}>
              🔒 Dirigente
            </button>
          )}
        </div>
      </header>

      <div className="rv-body" style={{ display:"flex", flex:1, overflow:"hidden", minHeight:0 }}>

        {/* Overlay móvil (cierra el sidebar al tocar fuera) */}
        {sidebarOpen && (
          <div className="rv-overlay" onClick={() => setSidebarOpen(false)} />
        )}

        {/* ── SIDEBAR ─────────────────────────────────────────────────────── */}
        <aside className={`rv-sidebar${sidebarOpen ? " open-mobile" : ""}`}
          style={{ width:"218px", flexShrink:0, background:DS.bg.surface,
          borderRight:"1px solid rgba(197,162,83,0.08)", padding:"10px 8px",
          overflowY:"auto", display:"flex", flexDirection:"column", gap:"5px",
          transition:"transform 0.25s ease" }}>
          {/* RF-08: búsqueda y filtro de Rovers */}
          <div style={{ paddingBottom:"8px", borderBottom:"1px solid rgba(255,255,255,0.04)", marginBottom:"6px" }}>
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"7px" }}>
              <span style={{ fontFamily:"'Cinzel',serif", fontSize:"10px", color:DS.text.tertiary,
                letterSpacing:"1.5px" }}>ROVERS</span>
            {isAdmin && (
              <div style={{ display:"flex", gap:"5px" }}>
                <button className="add-btn" title="Ordenar por % de progreso"
                  onClick={() => {
                    const upd = [...dataRef.current.scouts].sort((a, b) => totalStats(b.id).pct - totalStats(a.id).pct);
                    saveScouts(upd);
                  }}
                  style={{ width:"22px", height:"22px", background:"rgba(197,162,83,0.08)",
                    border:"1px solid rgba(197,162,83,0.25)", borderRadius:"50%", color:DS.brand.gold,
                    fontSize:"11px", lineHeight:1, cursor:"pointer",
                    display:"flex", alignItems:"center", justifyContent:"center",
                    transition:"background 0.15s" }}>%</button>
                <button className="add-btn" onClick={() => setAdding(a => !a)}
                  style={{ width:"22px", height:"22px", background:"rgba(139,26,26,0.1)",
                    border:"1px solid rgba(139,26,26,0.3)", borderRadius:"50%", color:DS.brand.red,
                    fontSize:"16px", lineHeight:1, cursor:"pointer",
                    display:"flex", alignItems:"center", justifyContent:"center",
                    transition:"background 0.15s" }}>+</button>
              </div>
            )}
            </div>
            {/* RF-08: búsqueda por nombre */}
            <input className="rv-search"
              ref={sidebarSearchRef}
              placeholder="🔍 Buscar Rover… (Ctrl+K)"
              aria-label="Buscar Rover por nombre"
              value={scoutSearch} onChange={e => setScoutSearch(e.target.value)} />
            {/* RF-08: filtro por etapa */}
            {scoutSearch === "" && (
              <div style={{ display:"flex", gap:"3px", marginTop:"5px", flexWrap:"wrap" }}>
                {[{ id:"all", label:"Todos" }, ...activeStagesList.map(s => ({ id:s.id, label:s.numeral }))].map(f => (
                  <button key={f.id}
                    className={`rv-qbtn${scoutStageFilter===f.id ? " active" : ""}`}
                    onClick={() => setScoutStageFilter(f.id)}
                    style={{ color: scoutStageFilter===f.id ? DS.brand.gold : undefined,
                      borderColor: scoutStageFilter===f.id ? "rgba(197,162,83,0.35)" : undefined,
                      background: scoutStageFilter===f.id ? "rgba(197,162,83,0.1)" : undefined }}>
                    {f.label}
                  </button>
                ))}
              </div>
            )}
          </div>

          {isAdmin && adding && (
            <div className="fade-in" style={{ background:"rgba(23,25,29,0.96)",
              border:`1px solid ${nameErr ? "rgba(192,57,43,0.4)" : "rgba(197,162,83,0.18)"}`,
              borderRadius:"9px", padding:"11px", marginBottom:"4px" }}>
              {/* RF-02: label + campo obligatorio */}
              <div style={{ fontSize:"9px", fontFamily:"'Cinzel',serif",
                color:DS.text.tertiary, letterSpacing:"1px", marginBottom:"5px" }}>
                NUEVO ROVER
              </div>
              <label style={{ fontSize:"10px", color:DS.text.tertiary,
                fontFamily:"'Source Sans 3',sans-serif", display:"block", marginBottom:"3px" }}>
                Nombre completo <span className="rv-field-req">*</span>
              </label>
              <input
                style={{ width:"100%", background:"rgba(255,255,255,0.05)",
                  border:`1px solid ${nameErr ? "rgba(192,57,43,0.5)" : "rgba(255,255,255,0.1)"}`,
                  borderRadius:"6px", padding:"7px 9px", color:DS.text.primary,
                  fontSize:"13px", outline:"none", marginBottom:"2px",
                  fontFamily:"'Source Sans 3',sans-serif", transition:"border-color 0.15s" }}
                placeholder="Nombre completo del Rover"
                value={newName}
                onChange={e => {
                  setNewName(e.target.value);
                  // RF-03: validación en tiempo real
                  if (!e.target.value.trim()) setNameErr("El nombre es obligatorio");
                  else if (dataRef.current.scouts.some(s => s.name.trim().toLowerCase() === e.target.value.trim().toLowerCase()))
                    setNameErr("Ya existe un Rover con ese nombre");
                  else setNameErr("");
                }}
                onKeyDown={e => e.key === "Enter" && addScout()} autoFocus />
              {/* RF-03: mensaje de error inmediato */}
              {nameErr && <div className="rv-field-err">{nameErr}</div>}
              <label style={{ fontSize:"10px", color:DS.text.tertiary,
                fontFamily:"'Source Sans 3',sans-serif", display:"block",
                marginTop:"8px", marginBottom:"3px" }}>
                Grupo / Clan <span style={{ color:DS.text.muted, fontSize:"9px" }}>(opcional)</span>
              </label>
              <input
                style={{ width:"100%", background:"rgba(255,255,255,0.05)",
                  border:"1px solid rgba(255,255,255,0.1)", borderRadius:"6px",
                  padding:"6px 9px", color:DS.text.primary, fontSize:"12px", outline:"none",
                  marginBottom:"10px", fontFamily:"'Source Sans 3',sans-serif" }}
                placeholder="Ej. Clan Takeda"
                value={newGroup} onChange={e => setNewGroup(e.target.value)} />
              <div style={{ display:"flex", gap:"5px" }}>
                <button className="btn-primary"
                  disabled={!!nameErr || !newName.trim()}
                  style={{ flex:1, padding:"6px", fontSize:"12px", justifyContent:"center" }}
                  onClick={addScout}>Agregar</button>
                <button className="btn-secondary"
                  style={{ padding:"6px 10px", fontSize:"12px" }}
                  onClick={() => { setAdding(false); setNameErr(""); setNewName(""); setNewGroup(""); }}>✕</button>
              </div>
            </div>
          )}

          {scouts.length === 0 && !adding && (
            <div style={{ textAlign:"center", padding:"28px 8px" }}>
              <div style={{ fontSize:"32px", opacity:.15, marginBottom:"8px" }}>⚜</div>
              <div style={{ fontSize:"12px", color:DS.text.tertiary, lineHeight:"1.6",
                fontFamily:"'Source Sans 3',sans-serif" }}>
                {isAdmin ? "Agrega tu primer Rover para comenzar" : "Aún no hay Rovers registrados"}
              </div>
            </div>
          )}

          {/* RF-08: filtrar la lista de Rovers */}
          {scouts
            .filter(sc => {
              const matchSearch = scoutSearch === "" ||
                sc.name.toLowerCase().includes(scoutSearch.toLowerCase());
              const matchStage = scoutStageFilter === "all" ||
                getCurrStage(sc.id).id === scoutStageFilter;
              return matchSearch && matchStage;
            })
            .map((sc, idx, arr) => {
            const cs = getCurrStage(sc.id);
            const ts = totalStats(sc.id);
            const isSel = sc.id === selId;
            const isEditing = editingScoutId === sc.id;
            const origIdx = scouts.indexOf(sc);
            return (
              <div key={sc.id} className="scout-row"
                style={{ display:"flex", alignItems:"center", gap:"6px", padding:"8px 7px",
                  background: isSel ? "rgba(28,13,5,0.99)" : "rgba(16,8,3,0.5)",
                  border:`1px solid ${isSel ? cs.color+"44" : "transparent"}`,
                  borderRadius:"8px", cursor:"pointer", transition:"all 0.15s" }}
                onClick={() => {
                  if (isEditing) return;
                  setSelId(sc.id); setView("main");
                  setStageId(getCurrStage(sc.id).id); setDetailFilter("all");
                  setSidebarOpen(false);
                }}>
                <div style={{ width:"30px", height:"30px", borderRadius:"50%",
                  background:cs.color, display:"flex", alignItems:"center", justifyContent:"center",
                  fontFamily:"'Cinzel',serif", fontSize:"13px", fontWeight:"700",
                  color:"#fff", flexShrink:0 }}>
                  {sc.name[0].toUpperCase()}
                </div>
                <div style={{ flex:1, minWidth:0 }}>
                  {/* RF-05: edición inline con doble clic — solo admin */}
                  {isAdmin && isEditing ? (
                    <input className="rv-inline-edit"
                      value={editingName}
                      onChange={e => setEditingName(e.target.value)}
                      onKeyDown={e => {
                        if (e.key === "Enter") saveInlineName(sc.id);
                        if (e.key === "Escape") setEditingScoutId(null);
                        e.stopPropagation();
                      }}
                      onBlur={() => saveInlineName(sc.id)}
                      onClick={e => e.stopPropagation()}
                      autoFocus />
                  ) : (
                    <div
                      style={{ fontSize:"13px", fontWeight:"600", color:DS.text.primary,
                        overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}
                      onDoubleClick={e => {
                        if (!isAdmin) return;
                        e.stopPropagation();
                        setEditingScoutId(sc.id);
                        setEditingName(sc.name);
                      }}
                      title={isAdmin ? "Doble clic para editar" : undefined}>
                      {sc.name}
                    </div>
                  )}
                  {sc.group && <div style={{ fontSize:"10px", color:DS.text.tertiary,
                    overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{sc.group}</div>}
                  <div style={{ fontSize:"10px", color:cs.color, marginTop:"1px" }}>{cs.name}</div>
                  <div className="rv-prog-track" style={{ height:"3px", marginTop:"4px" }}>
                    <div className={`rv-prog-fill${ts.pct===100?" done":""}`}
                      style={{ width:`${ts.pct}%`, background: ts.pct===100 ? undefined : cs.color }} />
                  </div>
                </div>
                <div style={{ display:"flex", flexDirection:"column", alignItems:"flex-end", gap:"3px" }}>
                  <span style={{ fontSize:"11px", fontWeight:"700", color:cs.color }}>{ts.pct}%</span>
                  {isAdmin && (
                    <div style={{ display:"flex", gap:"2px" }}>
                      <button className="move-btn" disabled={origIdx===0}
                        aria-label={`Subir a ${sc.name} en la lista`}
                        title="Subir"
                        style={{ background:"none", border:"none",
                          color: origIdx===0 ? "#241208" : "#5A4030", fontSize:"10px",
                          cursor: origIdx===0 ? "default" : "pointer", padding:"0 2px" }}
                        onClick={e => { e.stopPropagation(); moveScout(sc.id, -1); }}>▲</button>
                      <button className="move-btn" disabled={origIdx===scouts.length-1}
                        aria-label={`Bajar a ${sc.name} en la lista`}
                        title="Bajar"
                        style={{ background:"none", border:"none",
                          color: origIdx===scouts.length-1 ? "#241208" : "#5A4030", fontSize:"10px",
                          cursor: origIdx===scouts.length-1 ? "default" : "pointer", padding:"0 2px" }}
                        onClick={e => { e.stopPropagation(); moveScout(sc.id, 1); }}>▼</button>
                      <button className="del-btn"
                        aria-label={`Eliminar a ${sc.name}`}
                        title="Enviar a papelera"
                        style={{ background:"none", border:"none",
                        color:"#3A2010", fontSize:"10px", cursor:"pointer", padding:"0 2px",
                        transition:"color 0.12s" }}
                        onClick={e => {
                          e.stopPropagation();
                          showConfirm(
                            `¿Mover a ${sc.name} a la papelera? Podrás restaurarlo desde Configuración.`,
                            () => { moveToTrash(sc); deleteScout(sc.id); addToast("Rover movido a papelera", "info", 2200); }
                          );
                        }}>✕</button>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
          {/* RF-08: sin resultados */}
          {scouts.length > 0 && scouts.filter(sc =>
            (scoutSearch === "" || sc.name.toLowerCase().includes(scoutSearch.toLowerCase())) &&
            (scoutStageFilter === "all" || getCurrStage(sc.id).id === scoutStageFilter)
          ).length === 0 && (
            <div style={{ textAlign:"center", padding:"16px 8px", fontSize:"11px",
              color:DS.text.tertiary, fontFamily:"'Source Sans 3',sans-serif" }}>
              Sin resultados
            </div>
          )}
        </aside>

        {/* ── MAIN ──────────────────────────────────────────────────────────── */}
        <main className="rv-main" role="main" aria-label="Contenido principal" style={{ flex:1, overflow:"hidden", minHeight:0, background:DS.bg.root,
          display:"flex", flexDirection:"column" }}
          onClick={() => exportMenu && setExportMenu(false)}>

          {/* ── VISTA: PANEL DE CLAN (Fase 9, solo admin) ──────────────────── */}
          {view === "clan" && isAdmin && (() => {
            // Computar métricas del clan usando helpers existentes
            const INACTIVE_DAYS = 21;
            const today = new Date();
            // RF-04/RF-09: reuse memoized clanMetrics instead of recomputing
            const scoutMeta = clanMetrics.map(m => ({ ...m, stageSt: m.css }));

            // RF-01: KPIs del clan
            const totalRovers    = scouts.length;
            const withActivity   = scoutMeta.filter(m => m.initiated).length;
            const noActivity     = scoutMeta.filter(m => !m.initiated).length;
            const completedCount = scoutMeta.filter(m => m.completed).length;
            const totalDone      = scoutMeta.reduce((a, m) => a + m.ts.d, 0);
            const totalItems     = scoutMeta.reduce((a, m) => a + m.ts.t, 0);
            const inactiveList   = scoutMeta.filter(m => m.daysSince >= INACTIVE_DAYS);

            // RF-02/RF-03/RF-04: filtrar y ordenar
            const filtered = scoutMeta
              .filter(m => {
                const q = clanSearch.toLowerCase();
                if (q && !m.sc.name.toLowerCase().includes(q)) return false;
                if (clanFilter === "all") return true;
                if (clanFilter === "priority") return priorityIds.has(m.sc.id);
                if (clanFilter === "initiated") return m.initiated;
                if (clanFilter === "done") return m.completed;
                if (clanFilter === "inactive") return m.daysSince >= INACTIVE_DAYS;
                return m.cs.id === clanFilter; // stage filter
              })
              .sort((a, b) => {
                let va, vb;
                if (clanSort === "name")  { va = a.sc.name; vb = b.sc.name; return clanSortDir * va.localeCompare(vb); }
                if (clanSort === "stage") { va = activeStagesList.indexOf(a.cs); vb = activeStagesList.indexOf(b.cs); }
                else if (clanSort === "pct")  { va = a.stageSt.pct; vb = b.stageSt.pct; }
                else if (clanSort === "date") { va = a.lastDate||""; vb = b.lastDate||""; }
                return clanSortDir * (va > vb ? 1 : va < vb ? -1 : 0);
              });

            const sortBtn = (col, label) => (
              <th onClick={() => { if (clanSort===col) setClanSortDir(d=>-d); else { setClanSort(col); setClanSortDir(1); } }}>
                {label} {clanSort===col ? (clanSortDir===1 ? "▲" : "▼") : <span style={{opacity:.3}}>⇅</span>}
              </th>
            );

            // RF-06: actividad reciente — últimas 8 fechas únicas
            const recentActivity = Object.entries(dates)
              .filter(([, v]) => !!v)
              .map(([k, v]) => {
                const parts = k.replace("d:","").split("|");
                const scId = parts[0];
                const sc = scouts.find(s => s.id === scId);
                return sc ? { name: sc.name, date: v, key: k } : null;
              })
              .filter(Boolean)
              .sort((a, b) => b.date.localeCompare(a.date))
              .slice(0, 8);

            return (
              <div className="fade-in" style={{ flex:1, overflow:"auto", padding:"16px 18px",
                display:"flex", flexDirection:"column", gap:"14px" }}>

                {/* RF-09: Tarjetas de resumen */}
                <div style={{ display:"flex", gap:"10px", flexWrap:"wrap" }}>
                  {[
                    { label:"Rovers", value:totalRovers, color:DS.brand.gold, sub:"registrados" },
                    { label:"Con avance", value:withActivity, color:DS.state.done, sub:"con progresión iniciada" },
                    { label:"Sin inicio", value:noActivity, color:DS.text.tertiary, sub:"sin objetivos marcados" },
                    { label:"Completados", value:completedCount, color:DS.state.done, sub:"100% total" },
                    { label:"Objetivos", value:`${totalDone}/${totalItems}`, color:DS.state.infoLight, sub:"del clan completos" },
                    { label:"Sin actividad", value:inactiveList.length, color:inactiveList.length>0 ? DS.state.inprogress : DS.text.muted, sub:`+${INACTIVE_DAYS} días sin registro` },
                  ].map((card, i) => (
                    <div key={i} className="rv-stat-card"
                      onClick={() => card.label==="Sin actividad" && setClanFilter(clanFilter==="inactive"?"all":"inactive")}
                      style={{ cursor: card.label==="Sin actividad" ? "pointer" : "default",
                        borderColor: card.label==="Sin actividad" && clanFilter==="inactive" ? `${DS.state.inprogress}44` : undefined }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"22px",
                        fontWeight:"700", color:card.color, lineHeight:1 }}>{card.value}</div>
                      <div style={{ fontSize:"11px", fontWeight:"700", color:DS.text.primary,
                        fontFamily:DS.font.body, marginTop:"4px" }}>{card.label}</div>
                      <div style={{ fontSize:"9.5px", color:DS.text.muted,
                        fontFamily:DS.font.body, marginTop:"1px" }}>{card.sub}</div>
                    </div>
                  ))}
                </div>

                <div style={{ display:"flex", gap:"14px", flexWrap:"wrap", alignItems:"flex-start" }}>

                  {/* ── Tabla principal (RF-02/RF-03/RF-04/RF-05/RF-08) ── */}
                  <div style={{ flex:1, minWidth:"320px" }}>
                    {/* Search + filtros */}
                    <div style={{ display:"flex", gap:"7px", marginBottom:"10px",
                      flexWrap:"wrap", alignItems:"center" }}>
                      <input className="rv-search" placeholder="🔍 Buscar Rover…"
                        value={clanSearch} onChange={e => setClanSearch(e.target.value)}
                        style={{ flex:1, minWidth:"140px", maxWidth:"220px" }} />
                      <div style={{ display:"flex", gap:"4px", flexWrap:"wrap" }}>
                        {[
                          { id:"all", label:"Todos" },
                          ...activeStagesList.map(s => ({ id:s.id, label:s.numeral+" "+s.name })),
                          { id:"initiated", label:"Con avance" },
                          { id:"inactive", label:"⚠ Inactivos" },
                          { id:"priority", label:"★ Prioritarios" },
                        ].map(f => (
                          <button key={f.id}
                            className={`rv-filter-chip${clanFilter===f.id?" on":""}`}
                            onClick={() => setClanFilter(f.id)}>
                            {f.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* RF-08: tabla con overflow scroll */}
                    <div style={{ background:"rgba(13,14,16,0.98)", border:"1px solid rgba(255,255,255,0.05)",
                      borderRadius:"12px", overflow:"hidden" }}>
                      <div style={{ overflowX:"auto" }}>
                        <table className="rv-clan-table">
                          <thead>
                            <tr>
                              <th style={{ width:"20px", cursor:"default" }}></th>
                              {sortBtn("name",  "Rover")}
                              {sortBtn("stage", "Etapa")}
                              {sortBtn("pct",   "Avance")}
                              {sortBtn("date",  "Última act.")}
                              <th>Estado</th>
                              <th style={{ width:"60px", cursor:"default" }}>Acc.</th>
                            </tr>
                          </thead>
                          <tbody>
                            {filtered.length === 0 && (
                              <tr><td colSpan={7} style={{ textAlign:"center", padding:"24px",
                                color:DS.text.muted, fontFamily:DS.font.body }}>Sin resultados</td></tr>
                            )}
                            {/* RF-06: paginate rows for large clans */}
                            {(filtered.slice(0, 50)).map(({ sc, cs, ts, stageSt, lastDate, daysSince, initiated, completed }) => {
                              const isPriority = priorityIds.has(sc.id);
                              const isInactive = daysSince >= INACTIVE_DAYS && initiated;
                              return (
                                <tr key={sc.id} className="rv-clan-row"
                                  style={{ position:"relative" }}
                                  onClick={() => { setSelId(sc.id); setView("main");
                                    setStageId(getCurrStage(sc.id).id); setDetailFilter("all"); }}>
                                  {/* RF-11: estrella de prioridad */}
                                  <td onClick={e => { e.stopPropagation(); togglePriority(sc.id); }}>
                                    <span className="rv-priority-star"
                                      style={{ color: isPriority ? DS.brand.gold : DS.text.muted }}
                                      title={isPriority ? "Quitar seguimiento" : "Marcar seguimiento prioritario"}>
                                      {isPriority ? "★" : "☆"}
                                    </span>
                                  </td>
                                  {/* RF-02: columnas */}
                                  <td>
                                    <div style={{ display:"flex", alignItems:"center", gap:"7px" }}>
                                      <div style={{ width:"24px", height:"24px", borderRadius:"50%",
                                        background:cs.color, display:"flex", alignItems:"center",
                                        justifyContent:"center", fontSize:"11px", fontWeight:"700",
                                        color:"#fff", flexShrink:0 }}>{sc.name[0]}</div>
                                      <div>
                                        <div style={{ fontWeight:"600", color:DS.text.primary,
                                          fontSize:"12.5px" }}>
                                          {sc.name}
                                          {isPriority && <span style={{ color:DS.brand.gold,
                                            fontSize:"9px", marginLeft:"5px" }}>★</span>}
                                        </div>
                                        {sc.group && <div style={{ fontSize:"9.5px",
                                          color:DS.text.muted }}>{sc.group}</div>}
                                      </div>
                                    </div>
                                  </td>
                                  <td>
                                    <span style={{ fontFamily:DS.font.display, fontSize:"10px",
                                      color:cs.color, fontWeight:"700" }}>{cs.numeral}</span>
                                    <span style={{ fontSize:"10.5px", color:DS.text.tertiary,
                                      marginLeft:"5px" }}>{cs.name}</span>
                                  </td>
                                  <td>
                                    <div style={{ display:"flex", alignItems:"center", gap:"7px" }}>
                                      <div className="rv-prog-track" style={{ width:"60px", height:"4px" }}>
                                        <div className={`rv-prog-fill${stageSt.pct===100?" done":""}`}
                                          style={{ width:`${stageSt.pct}%`, background:cs.color }} />
                                      </div>
                                      <span style={{ fontSize:"11px", fontWeight:"700",
                                        color: stageSt.pct===100 ? DS.state.done : cs.color,
                                        minWidth:"30px" }}>{stageSt.pct}%</span>
                                    </div>
                                  </td>
                                  <td style={{ fontSize:"11px",
                                    color: isInactive ? DS.state.inprogress : DS.text.tertiary }}>
                                    {lastDate || <span style={{ color:DS.text.muted }}>—</span>}
                                    {isInactive && <span style={{ fontSize:"9px",
                                      marginLeft:"4px" }}>⚠</span>}
                                  </td>
                                  <td>
                                    {completed
                                      ? <span className="rv-badge complete">✓ Completo</span>
                                      : !initiated
                                        ? <span className="rv-badge pending">Sin inicio</span>
                                        : isInactive
                                          ? <span className="rv-badge warning" style={{
                                              background:"rgba(230,126,34,0.1)",
                                              color:DS.state.inprogress,
                                              border:"1px solid rgba(230,126,34,0.28)",
                                              borderRadius:"9999px", padding:"2px 7px",
                                              fontSize:"9.5px", fontWeight:"700",
                                              fontFamily:DS.font.display }}>Inactivo</span>
                                          : <span className="rv-badge inprog">En progreso</span>
                                    }
                                  </td>
                                  {/* RF-05: acceso al expediente */}
                                  <td onClick={e => e.stopPropagation()}>
                                    <button className="btn-action"
                                      style={{ padding:"4px 9px", fontSize:"10px" }}
                                      onClick={() => { setSelId(sc.id); setView("main");
                                        setStageId(getCurrStage(sc.id).id); setDetailFilter("all"); }}>
                                      Ver →
                                    </button>
                                  </td>
                                  {/* RF-10: vista rápida en hover */}
                                  <td style={{ padding:0, position:"relative", width:0, overflow:"visible" }}>
                                    <div className="rv-quick-view">
                                      <div style={{ fontFamily:DS.font.display, fontSize:"11px",
                                        fontWeight:"700", color:cs.color, marginBottom:"6px" }}>
                                        {cs.numeral} {cs.name}
                                      </div>
                                      <div style={{ fontSize:"12px", fontWeight:"700",
                                        color:DS.text.primary, marginBottom:"2px" }}>{sc.name}</div>
                                      {activeStagesList.map(st => {
                                        const ss = stageStats(sc.id, st);
                                        return (
                                          <div key={st.id} style={{ display:"flex", alignItems:"center",
                                            gap:"7px", marginTop:"4px" }}>
                                            <span style={{ fontSize:"9px", color:st.color,
                                              fontFamily:DS.font.display, width:"14px" }}>{st.numeral}</span>
                                            <div className="rv-prog-track" style={{ flex:1, height:"3px" }}>
                                              <div className={`rv-prog-fill${ss.pct===100?" done":""}`}
                                                style={{ width:`${ss.pct}%`, background:st.color }} />
                                            </div>
                                            <span style={{ fontSize:"9.5px", fontWeight:"700",
                                              color:st.color, width:"28px", textAlign:"right" }}>{ss.pct}%</span>
                                          </div>
                                        );
                                      })}
                                      {lastDate && <div style={{ fontSize:"9px", color:DS.text.muted,
                                        marginTop:"7px" }}>Últ. actividad: {lastDate}</div>}
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                      <div style={{ padding:"7px 12px", fontSize:"9.5px", color:DS.text.muted,
                        fontFamily:DS.font.body, borderTop:"1px solid rgba(255,255,255,0.03)",
                        display:"flex", alignItems:"center", justifyContent:"space-between",
                        flexWrap:"wrap", gap:"6px" }}>
                        <span>{filtered.length} de {scouts.length} Rovers</span>
                        {/* RF-06: page indicator if clipped */}
                        {filtered.length > 50 && (
                          <span style={{ fontSize:"9px", color:DS.text.muted }}>
                            Mostrando primeros 50 — usa filtros para acotar
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* ── Panel lateral: actividad reciente + inactivos ── */}
                  <div style={{ width:"220px", flexShrink:0, display:"flex",
                    flexDirection:"column", gap:"10px" }}>

                    {/* RF-06: Actividad reciente */}
                    <div style={{ background:"rgba(14,7,2,0.97)",
                      border:"1px solid rgba(255,255,255,0.05)",
                      borderRadius:"12px", padding:"12px 14px" }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                        color:DS.brand.goldMuted, letterSpacing:"1.5px",
                        fontWeight:"700", marginBottom:"10px" }}>ACTIVIDAD RECIENTE</div>
                      {recentActivity.length === 0
                        ? <div style={{ fontSize:"11px", color:DS.text.muted,
                            fontFamily:DS.font.body }}>Sin registros aún</div>
                        : recentActivity.map((a, i) => (
                          <div key={i} style={{ display:"flex", alignItems:"flex-start",
                            gap:"7px", marginBottom:"7px" }}>
                            <span style={{ color:DS.state.done, fontSize:"10px",
                              flexShrink:0, marginTop:"1px" }}>✓</span>
                            <div>
                              <div style={{ fontSize:"11px", fontWeight:"600",
                                color:DS.text.primary, lineHeight:"1.3" }}>{a.name}</div>
                              <div style={{ fontSize:"9.5px", color:DS.text.muted }}>{a.date}</div>
                            </div>
                          </div>
                        ))
                      }
                    </div>

                    {/* RF-07: Rovers sin actividad */}
                    {inactiveList.length > 0 && (
                      <div style={{ background:"rgba(14,7,2,0.97)",
                        border:"1px solid rgba(230,126,34,0.2)",
                        borderRadius:"12px", padding:"12px 14px" }}>
                        <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                          color:DS.state.inprogress, letterSpacing:"1.5px",
                          fontWeight:"700", marginBottom:"8px" }}>
                          ⚠ SIN ACTIVIDAD +{INACTIVE_DAYS}d
                        </div>
                        {inactiveList.map(({ sc, cs, daysSince }) => (
                          <div key={sc.id} style={{ display:"flex", alignItems:"center",
                            gap:"7px", marginBottom:"6px", cursor:"pointer" }}
                            onClick={() => { setSelId(sc.id); setView("main");
                              setStageId(getCurrStage(sc.id).id); setDetailFilter("all"); }}>
                            <div style={{ width:"20px", height:"20px", borderRadius:"50%",
                              background:cs.color, display:"flex", alignItems:"center",
                              justifyContent:"center", fontSize:"9px", fontWeight:"700",
                              color:"#fff", flexShrink:0 }}>{sc.name[0]}</div>
                            <div style={{ flex:1, minWidth:0 }}>
                              <div style={{ fontSize:"11px", fontWeight:"600",
                                color:DS.text.primary, overflow:"hidden",
                                textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{sc.name}</div>
                              <div style={{ fontSize:"9.5px", color:DS.state.inprogress }}>
                                {daysSince >= 9000 ? "Sin inicio" : `${daysSince}d sin actividad`}
                              </div>
                            </div>
                            <span onClick={e => { e.stopPropagation(); togglePriority(sc.id); }}
                              className="rv-priority-star"
                              style={{ color: priorityIds.has(sc.id) ? DS.brand.gold : DS.text.muted }}>
                              {priorityIds.has(sc.id) ? "★" : "☆"}
                            </span>
                          </div>
                        ))}
                      </div>
                    )}

                    {/* RF-11: Seguimiento prioritario */}
                    {priorityIds.size > 0 && (
                      <div style={{ background:"rgba(14,7,2,0.97)",
                        border:`1px solid ${DS.brand.gold}33`,
                        borderRadius:"12px", padding:"12px 14px" }}>
                        <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                          color:DS.brand.gold, letterSpacing:"1.5px",
                          fontWeight:"700", marginBottom:"8px" }}>★ SEGUIMIENTO PRIORITARIO</div>
                        {scouts.filter(s => priorityIds.has(s.id)).map(sc => {
                          const cs = getCurrStage(sc.id);
                          const ts = totalStats(sc.id);
                          return (
                            <div key={sc.id} style={{ display:"flex", alignItems:"center",
                              gap:"6px", marginBottom:"5px", cursor:"pointer" }}
                              onClick={() => { setSelId(sc.id); setView("main");
                                setStageId(getCurrStage(sc.id).id); setDetailFilter("all"); }}>
                              <div style={{ width:"20px", height:"20px", borderRadius:"50%",
                                background:cs.color, display:"flex", alignItems:"center",
                                justifyContent:"center", fontSize:"9px", fontWeight:"700",
                                color:"#fff", flexShrink:0 }}>{sc.name[0]}</div>
                              <div style={{ flex:1, minWidth:0 }}>
                                <div style={{ fontSize:"11px", fontWeight:"600",
                                  color:DS.text.primary, overflow:"hidden",
                                  textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{sc.name}</div>
                                <div style={{ fontSize:"9.5px", color:cs.color }}>{ts.pct}% total</div>
                              </div>
                              <span onClick={e => { e.stopPropagation(); togglePriority(sc.id); }}
                                className="rv-priority-star" style={{ color:DS.brand.gold }}>★</span>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            );
          })()}

          {/* ── VISTA: REPORTES (Fase 10, solo admin) ───────────────────────── */}
          {/* ── VISTA: CONFIGURACIÓN (Fase 11, solo admin) ─────────────────── */}
          {view === "settings" && isAdmin && (() => {
            const today = new Date().toLocaleString("es-GT");
            const TABS = [
              { id:"institution", icon:"🏛", label:"Institución"   },
              { id:"stages",      icon:"📋", label:"Progresión"    },
              { id:"system",      icon:"📊", label:"Estado"        },
              { id:"audit",       icon:"📋", label:"Actividad"     },
              { id:"trash",       icon:"🗑", label:`Papelera${trash.length>0?" ("+trash.length+")":""}` },
              { id:"a11y",        icon:"♿", label:"Accesibilidad"  },
              { id:"help",        icon:"❓", label:"Ayuda"         },
            ];
            // RF-11: lazy-load audit/trash only when tab opens
            const handleTabChange = (id) => {
              setSettingsTab(id);
              if (id === "audit") loadAudit();
              if (id === "trash") loadTrash();
            };
            const tabStyle = (id) => ({
              padding:"6px 12px", borderRadius:"7px", cursor:"pointer",
              fontSize:"11px", fontWeight:"600", fontFamily:"'Source Sans 3',sans-serif",
              background: settingsTab===id ? "rgba(197,162,83,0.15)" : "transparent",
              border: `1px solid ${settingsTab===id ? "rgba(197,162,83,0.35)" : "transparent"}`,
              color: settingsTab===id ? DS.brand.gold : DS.text.tertiary,
              transition:"all 0.12s",
              display:"flex", alignItems:"center", gap:"5px",
            });

            return (
              <div className="fade-in" style={{ flex:1, overflow:"auto", padding:"16px 18px",
                display:"flex", flexDirection:"column", gap:"12px" }}>

                <div style={{ fontFamily:DS.font.display, fontSize:"15px",
                  fontWeight:"700", color:DS.text.primary }}>
                  ⚙ Configuración del Sistema
                </div>

                {/* Tab bar */}
                <div style={{ display:"flex", gap:"4px", flexWrap:"wrap",
                  padding:"8px", background:"rgba(255,255,255,0.02)",
                  border:"1px solid rgba(255,255,255,0.04)", borderRadius:"10px" }}>
                  {TABS.map(t => (
                    <button key={t.id} style={tabStyle(t.id)}
                      onClick={() => handleTabChange(t.id)}>
                      <span>{t.icon}</span>{t.label}
                    </button>
                  ))}
                </div>

                {/* ── RF-02: Información institucional ── */}
                {settingsTab === "institution" && (
                  <div style={{ display:"flex", flexDirection:"column", gap:"10px" }}>
                    <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                      color:DS.brand.goldMuted, letterSpacing:"1.5px",
                      fontWeight:"700" }}>INFORMACIÓN INSTITUCIONAL</div>
                    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",
                      gap:"10px" }}>
                      {[
                        { key:"groupName",  label:"Nombre del Grupo Scout", ph:"Ej. Grupo Scout No. 12" },
                        { key:"clanName",   label:"Nombre del Clan",        ph:"Ej. Clan Takeda" },
                        { key:"district",   label:"Distrito / Región",      ph:"Ej. Guatemala Centro" },
                        { key:"contact",    label:"Correo de contacto",     ph:"Ej. clan@scouts.gt" },
                        { key:"logoUrl",    label:"URL del logotipo",       ph:"https://..." },
                        { key:"website",    label:"Sitio web",              ph:"https://..." },
                      ].map(f => (
                        <div key={f.key}>
                          <label style={{ fontSize:"10px", color:DS.text.tertiary,
                            fontFamily:DS.font.body, display:"block", marginBottom:"4px" }}>
                            {f.label}
                          </label>
                          <input className="ni"
                            style={{ width:"100%", padding:"7px 9px" }}
                            placeholder={f.ph}
                            value={sysConfig[f.key]||""}
                            onChange={e => saveSysConfig({ [f.key]: e.target.value })} />
                        </div>
                      ))}
                    </div>
                    {sysConfig.logoUrl && (
                      <div style={{ display:"flex", alignItems:"center", gap:"10px",
                        padding:"10px 14px", background:"rgba(255,255,255,0.02)",
                        border:"1px solid rgba(255,255,255,0.05)", borderRadius:"8px" }}>
                        <img src={sysConfig.logoUrl} alt="Logo"
                          style={{ height:"40px", objectFit:"contain", borderRadius:"4px" }}
                          onError={e => e.target.style.display="none"} />
                        <span style={{ fontSize:"11px", color:DS.text.tertiary,
                          fontFamily:DS.font.body }}>Vista previa del logotipo</span>
                      </div>
                    )}
                    <div style={{ fontSize:"10px", color:DS.text.muted,
                      fontFamily:DS.font.body, padding:"6px 0" }}>
                      Los cambios se aplican automáticamente al guardar. La configuración se almacena localmente en este dispositivo.
                    </div>

                    {/* RF-06: Backup */}
                    <div style={{ padding:"14px", background:"rgba(14,7,2,0.97)",
                      border:"1px solid rgba(255,255,255,0.05)", borderRadius:"10px" }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                        color:DS.brand.goldMuted, letterSpacing:"1.5px",
                        fontWeight:"700", marginBottom:"8px" }}>COPIA DE SEGURIDAD</div>
                      <div style={{ fontSize:"11px", color:DS.text.tertiary,
                        fontFamily:DS.font.body, marginBottom:"10px" }}>
                        Descarga una copia JSON de todos los datos: Rovers, progresión, fechas y asistencia.
                      </div>
                      <button className="btn-action" style={{ fontSize:"11px" }}
                        onClick={() => {
                          const bk = {
                            exportedAt: new Date().toISOString(),
                            version: "Cartilla 2.0",
                            config: sysConfig,
                            scouts: dataRef.current.scouts,
                            prog: dataRef.current.prog,
                            dates: dataRef.current.dates,
                            invDates: dataRef.current.invDates,
                            attendance: dataRef.current.attendance,
                          };
                          const blob = new Blob([JSON.stringify(bk, null, 2)],
                            { type:"application/json" });
                          const a = document.createElement("a");
                          a.href = URL.createObjectURL(blob);
                          a.download = `Backup_Clan_${new Date().toISOString().split("T")[0]}.json`;
                          a.click();
                          addToast("Copia de seguridad descargada", "success", 2200);
                          addAudit("Copia de seguridad generada");
                          saveSysConfig({ lastBackup: new Date().toISOString() });
                        }}>
                        ↓ Descargar copia de seguridad
                      </button>
                      {sysConfig.lastBackup && (
                        <div style={{ fontSize:"9.5px", color:DS.text.muted,
                          fontFamily:DS.font.body, marginTop:"6px" }}>
                          Último backup: {new Date(sysConfig.lastBackup).toLocaleString("es-GT")}
                        </div>
                      )}
                    </div>

                    {/* RF-03: Dirigentes — informativo */}
                    <div style={{ padding:"14px", background:"rgba(14,7,2,0.97)",
                      border:"1px solid rgba(255,255,255,0.05)", borderRadius:"10px" }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                        color:DS.brand.goldMuted, letterSpacing:"1.5px",
                        fontWeight:"700", marginBottom:"8px" }}>ACCESO DE DIRIGENTES</div>
                      <div style={{ display:"flex", alignItems:"center", gap:"10px",
                        marginBottom:"8px" }}>
                        <div style={{ width:"32px", height:"32px", borderRadius:"50%",
                          background:DS.brand.red, display:"flex", alignItems:"center",
                          justifyContent:"center", color:"#fff", fontSize:"13px",
                          fontWeight:"700", flexShrink:0 }}>
                          {(auth.currentUser?.email||"A")[0].toUpperCase()}
                        </div>
                        <div>
                          <div style={{ fontSize:"12px", fontWeight:"600",
                            color:DS.text.primary, fontFamily:DS.font.body }}>
                            {auth.currentUser?.email || "Administrador"}
                          </div>
                          <span className="rv-badge done" style={{ marginTop:"2px" }}>
                            ✓ Jefe de Clan activo
                          </span>
                        </div>
                      </div>
                      <div style={{ fontSize:"10px", color:DS.text.muted,
                        fontFamily:DS.font.body }}>
                        Para agregar dirigentes adicionales, gestiona los usuarios en la consola de Firebase Authentication.
                      </div>
                    </div>
                  </div>
                )}

                {/* ── RF-01/RF-03/RF-06: Administración de Etapas y Objetivos ── */}
                {settingsTab === "stages" && stagesConfig && (() => {
                  const cfgList = Object.values(stagesConfig)
                    .sort((a,b) => (a.order||a.id)-(b.order||b.id));
                  return (
                    <div style={{ display:"flex", flexDirection:"column", gap:"10px" }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                        color:DS.brand.goldMuted, letterSpacing:"1.5px", fontWeight:"700" }}>
                        ADMINISTRACIÓN DE PROGRESIÓN
                      </div>
                      <div style={{ fontSize:"11px", color:DS.text.tertiary,
                        fontFamily:DS.font.body, padding:"8px 12px",
                        background:"rgba(255,255,255,0.02)",
                        border:"1px solid rgba(255,255,255,0.04)", borderRadius:"8px" }}>
                        Los IDs internos nunca cambian. Renombrar etapas u objetivos no afecta el progreso registrado.
                      </div>

                      {cfgList.map(stageCfg => {
                        const hasProgress = scouts.some(sc =>
                          stageStats(sc.id, stageCfg).d > 0
                        );
                        const totalObj = stageCfg.sections
                          ? stageCfg.sections.reduce((a,s) => a + (s.items||[]).length, 0)
                          : 0;
                        return (
                          <div key={stageCfg.stableId}
                            style={{ background:"rgba(14,7,2,0.97)",
                              border:`1px solid ${stageCfg.active!==false?stageCfg.color+"33":"rgba(255,255,255,0.04)"}`,
                              borderRadius:"12px", overflow:"hidden",
                              opacity: stageCfg.active===false ? 0.55 : 1 }}>

                            {/* Cabecera de etapa */}
                            <div style={{ padding:"12px 14px",
                              background:`rgba(${parseInt(stageCfg.color.slice(1,3),16)},${parseInt(stageCfg.color.slice(3,5),16)},${parseInt(stageCfg.color.slice(5,7),16)},0.08)`,
                              display:"flex", alignItems:"center", gap:"10px" }}>
                              <span style={{ fontFamily:DS.font.display, fontSize:"16px",
                                fontWeight:"700", color:stageCfg.color }}>{stageCfg.numeral}</span>
                              <div style={{ flex:1 }}>
                                <div style={{ fontSize:"13px", fontWeight:"700",
                                  color:DS.text.primary, fontFamily:DS.font.body }}>
                                  {stageCfg.name}
                                </div>
                                <div style={{ fontSize:"9.5px", color:DS.text.muted,
                                  fontFamily:DS.font.body }}>
                                  {totalObj} objetivos · {hasProgress?"Con progreso registrado":"Sin progreso aún"}
                                </div>
                              </div>
                              <div style={{ display:"flex", gap:"6px", flexShrink:0, flexWrap:"wrap" }}>
                                {/* RF-03: reorder stages */}
                                <button className="rv-qbtn"
                                  title="Subir etapa" aria-label="Subir etapa"
                                  style={{ fontSize:"10px", padding:"4px 7px" }}
                                  onClick={() => reorderStage(stageCfg.stableId, -1)}>▲</button>
                                <button className="rv-qbtn"
                                  title="Bajar etapa" aria-label="Bajar etapa"
                                  style={{ fontSize:"10px", padding:"4px 7px" }}
                                  onClick={() => reorderStage(stageCfg.stableId, 1)}>▼</button>
                                {/* RF-02: activate/deactivate */}
                                <button className="btn-action"
                                  style={{ fontSize:"10px", padding:"4px 9px" }}
                                  onClick={() => updateStage(stageCfg.stableId, {
                                    active: stageCfg.active === false ? true : false
                                  })}>
                                  {stageCfg.active === false ? "Activar" : "Desactivar"}
                                </button>
                              </div>
                            </div>

                            {/* Edición de nombre */}
                            <div style={{ padding:"10px 14px",
                              borderBottom:"1px solid rgba(255,255,255,0.03)" }}>
                              <label style={{ fontSize:"9px", color:DS.text.muted,
                                fontFamily:DS.font.display, letterSpacing:"1px",
                                fontWeight:"700", display:"block", marginBottom:"4px" }}>
                                NOMBRE
                              </label>
                              <input className="ni"
                                style={{ width:"100%", padding:"6px 9px" }}
                                defaultValue={stageCfg.name}
                                onBlur={e => {
                                  if (e.target.value.trim() && e.target.value !== stageCfg.name)
                                    updateStage(stageCfg.stableId, { name: e.target.value.trim() });
                                }} />
                            </div>

                            {/* Objetivos por sección */}
                            {(stageCfg.sections||[]).map((sec, si) => (
                              <div key={sec.stableId||si}
                                style={{ padding:"10px 14px",
                                  borderBottom:"1px solid rgba(255,255,255,0.03)" }}>
                                {/* P2-1: Section header — editable name/area + reorder */}
                                <div style={{ display:"flex", alignItems:"flex-start",
                                  gap:"6px", marginBottom:"8px", flexWrap:"wrap" }}>
                                  <div style={{ flex:1, minWidth:"180px", display:"flex",
                                    flexDirection:"column", gap:"4px" }}>
                                    {/* P2-1: Edit section name */}
                                    <div style={{ display:"flex", gap:"4px", alignItems:"center" }}>
                                      <span style={{ fontSize:"9px", color:DS.brand.goldMuted,
                                        fontFamily:DS.font.display, flexShrink:0 }}>Sección:</span>
                                      <input className="ni"
                                        style={{ flex:1, padding:"3px 7px", fontSize:"11px",
                                          fontWeight:"700", color:DS.brand.gold }}
                                        defaultValue={sec.name}
                                        onBlur={e => {
                                          if (e.target.value.trim() && e.target.value !== sec.name)
                                            updateSection(stageCfg.stableId, sec.stableId||sec.name,
                                              { name: e.target.value.trim() });
                                        }} />
                                    </div>
                                    {/* P2-1: Edit section area */}
                                    <div style={{ display:"flex", gap:"4px", alignItems:"center" }}>
                                      <span style={{ fontSize:"9px", color:DS.text.muted,
                                        fontFamily:DS.font.display, flexShrink:0 }}>Área:</span>
                                      <input className="ni"
                                        style={{ flex:1, padding:"3px 7px", fontSize:"10px",
                                          color:stageCfg.color }}
                                        defaultValue={sec.area}
                                        onBlur={e => {
                                          if (e.target.value.trim() && e.target.value !== sec.area)
                                            updateSection(stageCfg.stableId, sec.stableId||sec.name,
                                              { area: e.target.value.trim() });
                                        }} />
                                    </div>
                                  </div>
                                  {/* P2-2: Reorder section buttons — stableId stays */}
                                  <div style={{ display:"flex", flexDirection:"column",
                                    gap:"2px", flexShrink:0 }}>
                                    <button className="rv-qbtn"
                                      title="Subir sección"
                                      style={{ fontSize:"9px", padding:"2px 6px" }}
                                      onClick={() => reorderSection(stageCfg.stableId,
                                        sec.stableId||sec.name, -1)}>▲</button>
                                    <button className="rv-qbtn"
                                      title="Bajar sección"
                                      style={{ fontSize:"9px", padding:"2px 6px" }}
                                      onClick={() => reorderSection(stageCfg.stableId,
                                        sec.stableId||sec.name, 1)}>▼</button>
                                  </div>
                                </div>
                                <div style={{ display:"flex", flexDirection:"column", gap:"4px" }}>
                                  {(sec.items||[]).map((it, ii) => {
                                    const itemText = typeof it === 'string' ? it : it.text;
                                    const itemActive = typeof it === 'object' ? it.active !== false : true;
                                    const hasItemProgress = scouts.some(sc =>
                                      !!prog[pKey(sc.id, stageCfg.id, sec.name, ii)]
                                    );
                                    return (
                                      <div key={ii} style={{ display:"flex", alignItems:"center",
                                        gap:"7px", opacity: itemActive ? 1 : 0.45 }}>
                                        <span style={{ fontSize:"9px", color:stageCfg.color,
                                          flexShrink:0 }}>◦</span>
                                        <input className="ni"
                                          style={{ flex:1, padding:"4px 7px", fontSize:"11px" }}
                                          defaultValue={itemText}
                                          disabled={!itemActive}
                                          onBlur={e => {
                                            if (e.target.value.trim() && e.target.value !== itemText && typeof it === "object" && it.stableId)
                                              updateObjective(stageCfg.stableId, sec.stableId, it.stableId,
                                                { text: e.target.value.trim() });
                                          }} />
                                        {/* RF-04: no eliminar si tiene progreso */}
                                        {/* RF-03: reorder objective */}
                                        <button className="rv-qbtn"
                                          title="Subir objetivo" style={{ fontSize:"9px", padding:"2px 5px" }}
                                          onClick={() => reorderObjective(stageCfg.stableId, sec.stableId||sec.name, ii, Math.max(0,ii-1))}>
                                          ▲
                                        </button>
                                        <button className="rv-qbtn"
                                          title="Bajar objetivo" style={{ fontSize:"9px", padding:"2px 5px" }}
                                          onClick={() => reorderObjective(stageCfg.stableId, sec.stableId||sec.name, ii, Math.min((sec.items||[]).length-1,ii+1))}>
                                          ▼
                                        </button>
                                        {/* RF-02: deactivate */}
                                        <button className="rv-qbtn"
                                          title={hasItemProgress ? "Tiene progreso — solo se puede desactivar" : ""}
                                          style={{ fontSize:"9px", flexShrink:0,
                                            color: itemActive ? DS.text.muted : DS.state.done }}
                                          onClick={() => {
                                            if (typeof it === "object" && it.stableId)
                                              updateObjective(stageCfg.stableId, sec.stableId, it.stableId,
                                                { active: !itemActive });
                                          }}>
                                          {itemActive ? "✕" : "↩"}
                                        </button>
                                      </div>
                                    );
                                  })}
                                  {/* Agregar objetivo */}
                                  <div style={{ display:"flex", gap:"6px", marginTop:"4px" }}>
                                    <input className="ni"
                                      id={`new-obj-${stageCfg.stableId}-${si}`}
                                      style={{ flex:1, padding:"4px 7px", fontSize:"11px" }}
                                      placeholder="Nuevo objetivo…"
                                      onKeyDown={e => {
                                        if (e.key === "Enter" && e.target.value.trim()) {
                                          addObjective(stageCfg.stableId, sec.stableId || sec.name, e.target.value);
                                          e.target.value = "";
                                        }
                                      }} />
                                    <button className="btn-action"
                                      style={{ fontSize:"10px", padding:"4px 9px" }}
                                      onClick={() => {
                                        const inp = document.getElementById(`new-obj-${stageCfg.stableId}-${si}`);
                                        if (inp && inp.value.trim()) {
                                          addObjective(stageCfg.stableId, sec.stableId || sec.name, inp.value);
                                          inp.value = "";
                                        }
                                      }}>+ Agregar</button>
                                  </div>
                                </div>
                              </div>
                            ))}
                          {/* P2-1: Add new section to this stage */}
                          <div style={{ padding:"10px 14px",
                            borderTop:"1px solid rgba(255,255,255,0.04)",
                            background:"rgba(255,255,255,0.015)" }}>
                            <div style={{ fontSize:"9px", color:DS.text.muted,
                              fontFamily:DS.font.display, letterSpacing:"1px",
                              fontWeight:"700", marginBottom:"6px" }}>+ NUEVA SECCIÓN</div>
                            <div style={{ display:"flex", gap:"5px", flexWrap:"wrap" }}>
                              <input className="ni"
                                id={`new-sec-name-${stageCfg.stableId}`}
                                style={{ flex:1, minWidth:"120px", padding:"5px 8px",
                                  fontSize:"11px" }}
                                placeholder="Nombre de sección…" />
                              <input className="ni"
                                id={`new-sec-area-${stageCfg.stableId}`}
                                style={{ flex:1, minWidth:"100px", padding:"5px 8px",
                                  fontSize:"11px" }}
                                placeholder="Área…" />
                              <button className="btn-action"
                                style={{ fontSize:"10px", padding:"5px 10px" }}
                                onClick={() => {
                                  const nInp = document.getElementById(`new-sec-name-${stageCfg.stableId}`);
                                  const aInp = document.getElementById(`new-sec-area-${stageCfg.stableId}`);
                                  if (nInp?.value.trim()) {
                                    addSection(stageCfg.stableId, nInp.value, aInp?.value||nInp.value);
                                    nInp.value = ""; if (aInp) aInp.value = "";
                                  }
                                }}>+ Agregar</button>
                            </div>
                          </div>

                        </div>
                        );
                      })}
                      {/* RF-03: Create new stage */}
                      <div style={{ padding:"12px 14px", background:"rgba(14,7,2,0.97)",
                        border:"1px solid rgba(197,162,83,0.12)", borderRadius:"12px" }}>
                        <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                          color:DS.brand.goldMuted, letterSpacing:"1.5px",
                          fontWeight:"700", marginBottom:"8px" }}>CREAR NUEVA ETAPA</div>
                        <div style={{ display:"flex", gap:"6px" }}>
                          <input className="ni" id="new-stage-name"
                            style={{ flex:1, padding:"6px 9px", fontSize:"12px" }}
                            placeholder="Nombre de la nueva etapa…"
                            onKeyDown={e => {
                              if (e.key==="Enter" && e.target.value.trim()) {
                                createStage(e.target.value);
                                e.target.value = "";
                              }
                            }} />
                          <button className="btn-primary"
                            style={{ fontSize:"11px", padding:"6px 12px" }}
                            onClick={() => {
                              const inp = document.getElementById("new-stage-name");
                              if (inp?.value.trim()) { createStage(inp.value); inp.value=""; }
                            }}>+ Etapa</button>
                        </div>
                        <div style={{ fontSize:"9.5px", color:DS.text.muted,
                          fontFamily:DS.font.body, marginTop:"6px" }}>
                          La nueva etapa aparecerá en el dashboard inmediatamente.
                          Agrega secciones y objetivos desde el panel de cada etapa.
                        </div>
                      </div>
                    </div>
                  );
                })()}

                {/* ── RF-07/RF-11: Estado del sistema ── */}
                {settingsTab === "system" && (() => {
                  const totalD = scouts.reduce((a,s) => a+totalStats(s.id).d, 0);
                  const totalT = scouts.reduce((a,s) => a+totalStats(s.id).t, 0);
                  return (
                    <div style={{ display:"flex", flexDirection:"column", gap:"10px" }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                        color:DS.brand.goldMuted, letterSpacing:"1.5px", fontWeight:"700" }}>
                        ESTADO DEL SISTEMA
                      </div>
                      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",
                        gap:"10px" }}>
                        {[
                          { label:"Versión del Dashboard", value:"Cartilla 2.0 · Fase 11", icon:"🏷", color:DS.brand.gold },
                          { label:"Estado de Firebase", value:saveErr?"⚠ Sin conexión":"✓ Conectado", icon:"🔥",
                            color:saveErr?DS.state.error:DS.state.done },
                          { label:"Último guardado", value:syncing?"Guardando…":"Sincronizado", icon:"☁",
                            color:DS.state.done },
                          { label:"Rovers registrados", value:scouts.length, icon:"👤", color:DS.brand.gold },
                          { label:"Objetivos completados", value:`${totalD}/${totalT}`, icon:"✓", color:DS.state.done },
                          { label:"Último backup", value:sysConfig.lastBackup
                            ? new Date(sysConfig.lastBackup).toLocaleDateString("es-GT")
                            : "Sin backup", icon:"💾", color:DS.text.secondary },
                          { label:"Acciones en auditoría", value:auditLog.length, icon:"📋", color:DS.state.info },
                          { label:"Elementos en papelera", value:trash.length, icon:"🗑", color:DS.text.tertiary },
                          { label:"Fecha del sistema", value:today, icon:"📅", color:DS.text.secondary },
                        ].map((item, i) => (
                          <div key={i} style={{ padding:"12px 14px",
                            background:"rgba(14,7,2,0.97)",
                            border:"1px solid rgba(255,255,255,0.05)", borderRadius:"10px" }}>
                            <div style={{ fontSize:"18px", marginBottom:"6px" }}>{item.icon}</div>
                            <div style={{ fontFamily:DS.font.display, fontSize:"14px",
                              fontWeight:"700", color:item.color, lineHeight:1,
                              marginBottom:"3px" }}>{item.value}</div>
                            <div style={{ fontSize:"9.5px", color:DS.text.muted,
                              fontFamily:DS.font.body }}>{item.label}</div>
                          </div>
                        ))}
                      </div>

                      {/* RF-04/RF-05: Etapas — informativo */}
                      <div style={{ padding:"14px", background:"rgba(14,7,2,0.97)",
                        border:"1px solid rgba(255,255,255,0.05)", borderRadius:"10px",
                        marginTop:"4px" }}>
                        <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                          color:DS.brand.goldMuted, letterSpacing:"1.5px",
                          fontWeight:"700", marginBottom:"10px" }}>ETAPAS CONFIGURADAS</div>
                        {activeStagesList.map(st => {
                          const clanSt = scouts.filter(s => getCurrStage(s.id).id===st.id).length;
                          const totalStD = scouts.reduce((a,s) => a+stageStats(s.id,st).d, 0);
                          const totalStT = scouts.reduce((a,s) => a+stageStats(s.id,st).t, 0);
                          return (
                            <div key={st.id} style={{ display:"flex", alignItems:"center",
                              gap:"12px", padding:"8px 0",
                              borderBottom:"1px solid rgba(255,255,255,0.03)" }}>
                              <span style={{ fontFamily:DS.font.display, fontSize:"14px",
                                fontWeight:"700", color:st.color, width:"22px",
                                flexShrink:0 }}>{st.numeral}</span>
                              <div style={{ flex:1 }}>
                                <div style={{ fontSize:"12px", fontWeight:"600",
                                  color:DS.text.primary }}>{st.name}</div>
                                <div style={{ fontSize:"9.5px", color:DS.text.muted,
                                  fontFamily:DS.font.body }}>{st.subtitle}</div>
                              </div>
                              <div style={{ fontSize:"10px", color:DS.text.tertiary,
                                fontFamily:DS.font.body, textAlign:"right" }}>
                                <div>{clanSt} Rover{clanSt!==1?"s":""} en esta etapa</div>
                                <div style={{ color:DS.text.muted }}>{totalStD}/{totalStT} obj. completados</div>
                              </div>
                              <span className="rv-badge done" style={{ flexShrink:0 }}>Activa</span>
                            </div>
                          );
                        })}
                        <div style={{ fontSize:"10px", color:DS.text.muted,
                          fontFamily:DS.font.body, marginTop:"8px" }}>
                          Las etapas y objetivos se administran desde Configuración → Progresión. Los cambios se guardan en Firebase automáticamente.
                        </div>
                      </div>

                      {/* RF-11: Panel de monitoreo de rendimiento */}
                      {(() => {
                        const p = perfRef.current;
                        const loadMs = p.loadEnd ? Math.round(p.loadEnd - p.loadStart) : null;
                        const avgSync = p.syncTimes.length
                          ? Math.round(p.syncTimes.reduce((a,b)=>a+b,0)/p.syncTimes.length)
                          : null;
                        const perfItems = [
                          { label:"Tiempo de carga inicial (real)", value: loadMs!=null?`${loadMs}ms`:"Midiendo…",
                            color: loadMs!=null?(loadMs<2000?DS.state.done:loadMs<4000?DS.state.inprogress:DS.state.error):DS.text.muted },
                          { label:"Snapshots Firebase recibidos", value: p.firebaseReads,
                            color:DS.state.info },
                          { label:"Recálculos evitados (useMemo)", value: p.cacheHits,
                            color:DS.state.done },
                          { label:"Guardados sincronizados", value: p.syncTimes.length,
                            color:DS.state.info },
                          { label:"Tiempo últ. guardado (est.)", value:avgSync!=null?`~${avgSync}ms`:"—",
                            color:DS.text.secondary },
                          { label:"Estado Firebase", value:saveErr?"⚠ Error de sync":"✓ Sincronizado",
                            color:saveErr?DS.state.error:DS.state.done },
                          { label:"Rovers en memoria", value: scouts.length, color:DS.text.secondary },
                          { label:"Etapas activas", value: activeStagesList.length, color:DS.text.secondary },
                        ];
                        return (
                          <div style={{ padding:"14px", background:"rgba(14,7,2,0.97)",
                            border:"1px solid rgba(197,162,83,0.12)", borderRadius:"10px",
                            marginTop:"10px" }}>
                            <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                              color:DS.brand.goldMuted, letterSpacing:"1.5px",
                              fontWeight:"700", marginBottom:"10px" }}>
                              MONITOREO DE RENDIMIENTO
                            </div>
                            <div style={{ display:"grid",
                              gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",
                              gap:"8px" }}>
                              {perfItems.map((item,i) => (
                                <div key={i} style={{ padding:"9px 11px",
                                  background:"rgba(255,255,255,0.02)",
                                  border:"1px solid rgba(255,255,255,0.04)",
                                  borderRadius:"8px" }}>
                                  <div style={{ fontFamily:DS.font.display, fontSize:"16px",
                                    fontWeight:"700", color:item.color,
                                    lineHeight:1, marginBottom:"3px" }}>{item.value}</div>
                                  <div style={{ fontSize:"9px", color:DS.text.muted,
                                    fontFamily:DS.font.body }}>{item.label}</div>
                                </div>
                              ))}
                            </div>
                            <div style={{ fontSize:"9px", color:DS.text.muted,
                              fontFamily:DS.font.body, marginTop:"8px" }}>
                              Los contadores se reinician al recargar la página.
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  );
                })()}

                {/* ── RF-01/RF-03/RF-04: Registro de actividad — Firebase ── */}
                {settingsTab === "audit" && (() => {
                  const [auditTypeFilter, setAuditTypeFilter] = React.useState("all");
                  const filtered = auditLog.filter(e =>
                    auditTypeFilter === "all" || (e.type && e.type === auditTypeFilter)
                  );
                  return (
                    <div style={{ display:"flex", flexDirection:"column", gap:"8px" }}>
                      <div style={{ display:"flex", alignItems:"center",
                        justifyContent:"space-between", flexWrap:"wrap", gap:"8px" }}>
                        <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                          color:DS.brand.goldMuted, letterSpacing:"1.5px", fontWeight:"700" }}>
                          REGISTRO DE ACTIVIDAD ({auditLog.length} entradas)
                          <span style={{ color:DS.text.muted, marginLeft:"6px",
                            fontFamily:DS.font.body, fontWeight:"400", fontSize:"9px" }}>
                            · Firebase · persiste entre dispositivos
                          </span>
                        </div>
                      </div>
                      {/* RF-03: Filtros básicos */}
                      <div style={{ display:"flex", gap:"5px", flexWrap:"wrap" }}>
                        {["all","trash","restore","stage","objective","scout"].map(f => (
                          <button key={f}
                            className={`rv-filter-chip${auditTypeFilter===f?" on":""}`}
                            onClick={() => setAuditTypeFilter(f)}
                            style={{ background: auditTypeFilter===f ? "rgba(197,162,83,0.12)":undefined,
                              borderColor: auditTypeFilter===f ? "rgba(197,162,83,0.35)":undefined,
                              color: auditTypeFilter===f ? DS.brand.gold:undefined }}>
                            {f==="all"?"Todos":f==="trash"?"Papelera":f==="restore"?"Restaurados":
                             f==="stage"?"Etapas":f==="objective"?"Objetivos":"Rovers"}
                          </button>
                        ))}
                      </div>
                      <div style={{ background:"rgba(14,7,2,0.97)",
                        border:"1px solid rgba(255,255,255,0.05)", borderRadius:"10px",
                        overflow:"hidden" }}>
                        {!auditLoaded ? (
                          <div style={{ padding:"20px", textAlign:"center",
                            fontSize:"11px", color:DS.text.muted, fontFamily:DS.font.body }}>
                            <span className="rv-sync-dot" style={{ marginRight:"6px" }} />
                            Cargando desde Firebase…
                          </div>
                        ) : filtered.length === 0 ? (
                          <div style={{ padding:"24px", textAlign:"center",
                            fontSize:"11px", color:DS.text.muted, fontFamily:DS.font.body }}>
                            Sin actividad registrada aún
                          </div>
                        ) : filtered.map((e, i) => (
                          <div key={e.id||i} style={{ display:"flex", alignItems:"flex-start",
                            gap:"10px", padding:"8px 12px",
                            borderBottom:"1px solid rgba(255,255,255,0.03)",
                            background: i===0 ? "rgba(197,162,83,0.04)" : "transparent" }}>
                            <div style={{ width:"5px", height:"5px", borderRadius:"50%",
                              background: e.type==="trash" ? DS.state.inprogress :
                                e.type==="restore" ? DS.state.done : DS.brand.gold,
                              flexShrink:0, marginTop:"5px" }} />
                            <div style={{ flex:1, minWidth:0 }}>
                              <div style={{ fontSize:"12px", color:DS.text.primary,
                                fontFamily:DS.font.body }}>{e.action}</div>
                              <div style={{ fontSize:"9.5px", color:DS.text.muted,
                                fontFamily:DS.font.body, marginTop:"1px" }}>
                                {new Date(e.ts).toLocaleString("es-GT")} · {e.user}
                                {e.type && <span style={{ marginLeft:"6px",
                                  color:DS.brand.goldMuted }}>#{e.type}</span>}
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}

                {/* ── RF-05/RF-06/RF-07: Papelera — Firebase persistente ── */}
                {settingsTab === "trash" && (
                  <div style={{ display:"flex", flexDirection:"column", gap:"8px" }}>
                    <div style={{ display:"flex", alignItems:"center",
                      justifyContent:"space-between", flexWrap:"wrap", gap:"6px" }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                        color:DS.brand.goldMuted, letterSpacing:"1.5px", fontWeight:"700" }}>
                        PAPELERA ({trash.length} elementos)
                        <span style={{ color:DS.text.muted, marginLeft:"6px",
                          fontFamily:DS.font.body, fontWeight:"400", fontSize:"9px" }}>
                          · Firebase · persiste entre dispositivos
                        </span>
                      </div>
                    </div>
                    {!trashLoaded ? (
                      <div style={{ padding:"24px", textAlign:"center",
                        fontSize:"11px", color:DS.text.muted, fontFamily:DS.font.body }}>
                        <span className="rv-sync-dot" style={{ marginRight:"6px" }} />
                        Cargando desde Firebase…
                      </div>
                    ) : trash.length === 0 ? (
                      <div style={{ padding:"32px", textAlign:"center",
                        background:"rgba(14,7,2,0.97)",
                        border:"1px solid rgba(255,255,255,0.04)", borderRadius:"10px",
                        fontSize:"12px", color:DS.text.muted, fontFamily:DS.font.body }}>
                        La papelera está vacía
                      </div>
                    ) : trash.map((item, i) => {
                      // RF-07: check if has historical progress
                      const hasHistory = Object.keys(dataRef.current.prog||{})
                        .some(k => k.startsWith(item.id+"|") || k.startsWith(`p:${item.id}|`));
                      return (
                        <div key={item.id||i} style={{ display:"flex", alignItems:"center",
                          gap:"10px", padding:"10px 14px",
                          background:"rgba(14,7,2,0.97)",
                          border:"1px solid rgba(255,255,255,0.05)", borderRadius:"10px" }}>
                          <div style={{ width:"32px", height:"32px", borderRadius:"50%",
                            background:"rgba(255,255,255,0.1)", display:"flex",
                            alignItems:"center", justifyContent:"center",
                            fontSize:"13px", fontWeight:"700", color:DS.text.secondary,
                            flexShrink:0 }}>{(item.name||"?")[0].toUpperCase()}</div>
                          <div style={{ flex:1, minWidth:0 }}>
                            <div style={{ fontSize:"12.5px", fontWeight:"600",
                              color:DS.text.primary }}>{item.name}</div>
                            {item.group && <div style={{ fontSize:"10px",
                              color:DS.text.muted }}>{item.group}</div>}
                            <div style={{ fontSize:"9.5px", color:DS.text.tertiary,
                              fontFamily:DS.font.body, marginTop:"1px" }}>
                              Eliminado: {new Date(item.deletedAt).toLocaleString("es-GT")}
                              {item.deletedBy && ` · ${item.deletedBy}`}
                            </div>
                            {hasHistory && (
                              <span className="rv-badge info" style={{ marginTop:"3px" }}>
                                Con progreso registrado
                              </span>
                            )}
                          </div>
                          <button className="btn-action" style={{ fontSize:"10px", padding:"5px 10px" }}
                            onClick={() => restoreFromTrash(item.id)}>
                            ↩ Restaurar
                          </button>
                          {/* RF-07: block permanent delete if has history */}
                          <button className="btn-danger" style={{ fontSize:"10px", padding:"5px 10px",
                            opacity: hasHistory ? 0.4 : 1 }}
                            title={hasHistory ? "No se puede eliminar: tiene progreso registrado" : ""}
                            onClick={() => {
                              if (hasHistory) {
                                addToast("No se puede eliminar: tiene progreso registrado", "error", 3000);
                                addAudit(`Eliminación bloqueada por historial: ${item.name}`,
                                  { type:"blocked", elementId: item.id });
                                return;
                              }
                              showConfirm(
                                `¿Eliminar permanentemente a ${item.name}? Esta acción es irreversible.`,
                                () => {
                                  updateDoc(TRASH_DOC_REF, { [item.id]: deleteField() }).catch(()=>{});
                                  setTrash(prev => prev.filter(t => t.id !== item.id));
                                  addAudit(`Rover eliminado permanentemente: ${item.name}`,
                                    { type:"permanent_delete", elementId: item.id });
                                  addToast(`${item.name} eliminado permanentemente`, "info", 2000);
                                }
                              );
                            }}>
                            ✕ Borrar
                          </button>
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* ── RF-10: Preferencias de accesibilidad ── */}
                {settingsTab === "a11y" && (
                  <div style={{ display:"flex", flexDirection:"column", gap:"10px" }}>
                    <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                      color:DS.brand.goldMuted, letterSpacing:"1.5px", fontWeight:"700" }}>
                      PREFERENCIAS DE ACCESIBILIDAD
                    </div>
                    {[
                      { key:"largeText",    label:"Texto más grande",
                        desc:"Aumenta el tamaño del texto en un 15% en toda la aplicación.", icon:"🔤" },
                      { key:"reduceMotion", label:"Reducir animaciones",
                        desc:"Desactiva transiciones y animaciones para usuarios sensibles al movimiento.", icon:"⏸" },
                      { key:"highContrast", label:"Mayor contraste",
                        desc:"Aumenta el contraste de bordes, tablas y barras de progreso.", icon:"◑" },
                    ].map(pref => (
                      <div key={pref.key} style={{ display:"flex", alignItems:"center",
                        gap:"12px", padding:"12px 14px",
                        background:"rgba(14,7,2,0.97)",
                        border:`1px solid ${a11yPrefs[pref.key]?"rgba(197,162,83,0.3)":"rgba(255,255,255,0.05)"}`,
                        borderRadius:"10px", cursor:"pointer",
                        transition:"border-color 0.15s" }}
                        onClick={() => saveA11y({ [pref.key]: !a11yPrefs[pref.key] })}>
                        <span style={{ fontSize:"20px" }}>{pref.icon}</span>
                        <div style={{ flex:1 }}>
                          <div style={{ fontSize:"12.5px", fontWeight:"700",
                            color:DS.text.primary, fontFamily:DS.font.body,
                            marginBottom:"2px" }}>{pref.label}</div>
                          <div style={{ fontSize:"10.5px", color:DS.text.tertiary,
                            fontFamily:DS.font.body }}>{pref.desc}</div>
                        </div>
                        <div style={{ width:"36px", height:"20px", borderRadius:"10px",
                          background: a11yPrefs[pref.key]
                            ? `linear-gradient(90deg,${DS.brand.red},${DS.brand.gold})`
                            : "rgba(255,255,255,0.1)",
                          position:"relative", flexShrink:0,
                          transition:"background 0.2s",
                          cursor:"pointer" }}>
                          <div style={{ position:"absolute", top:"2px",
                            left: a11yPrefs[pref.key] ? "18px" : "2px",
                            width:"16px", height:"16px", borderRadius:"50%",
                            background:"#fff", transition:"left 0.2s",
                            boxShadow:"0 1px 3px rgba(0,0,0,0.3)" }} />
                        </div>
                      </div>
                    ))}
                    {/* RF-09: keyboard shortcuts reference */}
                    <div style={{ padding:"12px 14px",
                      background:"rgba(14,7,2,0.97)",
                      border:"1px solid rgba(255,255,255,0.05)",
                      borderRadius:"10px", marginTop:"4px" }}>
                      <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                        color:DS.brand.goldMuted, letterSpacing:"1.5px",
                        fontWeight:"700", marginBottom:"10px" }}>
                        ATAJOS DE TECLADO
                      </div>
                      {[
                        ["Ctrl+K",  "Enfocar búsqueda de Rovers"],
                        ["Ctrl+1",  "Ir a Dashboard principal"],
                        ["Ctrl+2",  "Ir a Panel del Clan"],
                        ["Ctrl+3",  "Ir a Reportes"],
                        ["Ctrl+4",  "Ir a Asistencia"],
                        ["Ctrl+5",  "Ir a Configuración"],
                        ["Escape",  "Cerrar modal / diálogo activo"],
                      ].map(([key, desc]) => (
                        <div key={key} style={{ display:"flex", alignItems:"center",
                          gap:"10px", padding:"5px 0",
                          borderBottom:"1px solid rgba(255,255,255,0.03)" }}>
                          <span className="rv-kbd">{key}</span>
                          <span style={{ fontSize:"11px", color:DS.text.tertiary,
                            fontFamily:DS.font.body }}>{desc}</span>
                        </div>
                      ))}
                      <div style={{ fontSize:"9.5px", color:DS.text.muted,
                        fontFamily:DS.font.body, marginTop:"8px" }}>
                        En macOS usa ⌘ (Cmd) en lugar de Ctrl.
                      </div>
                    </div>
                  </div>
                )}

                {/* ── RF-08: Centro de ayuda ── */}
                {settingsTab === "help" && (
                  <div style={{ display:"flex", flexDirection:"column", gap:"10px" }}>
                    <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                      color:DS.brand.goldMuted, letterSpacing:"1.5px", fontWeight:"700" }}>
                      CENTRO DE AYUDA
                    </div>
                    {[
                      { title:"¿Cómo marcar un objetivo como completado?",
                        body:"En la vista principal, selecciona un Rover desde el panel lateral. Ve a la etapa correspondiente y haz clic en el checkbox junto al objetivo. Solo los dirigentes autenticados pueden marcar objetivos." },
                      { title:"¿Cómo registrar la fecha de cumplimiento?",
                        body:"Haz clic en el ícono ▼ a la derecha de cualquier objetivo para expandir el campo de fecha. Selecciona la fecha y se guardará automáticamente." },
                      { title:"¿Cómo generar una ficha oficial?",
                        body:"Ve a 📄 Reportes en el menú superior. Selecciona 'Ficha Oficial Individual' y haz clic en el nombre del Rover. Se mostrará una vista previa que puedes imprimir o guardar como PDF." },
                      { title:"¿Cómo filtrar Rovers en el Panel del Clan?",
                        body:"En la vista ⚑ Clan, usa los chips de filtro (Con avance, Inactivos, ★ Prioritarios) o escribe en el buscador. Los filtros se combinan automáticamente." },
                      { title:"¿Cómo hacer una copia de seguridad?",
                        body:"Ve a ⚙ Configuración → Institución. En la sección 'Copia de seguridad', haz clic en 'Descargar copia de seguridad'. Se descargará un archivo JSON con todos los datos." },
                      { title:"¿Cómo restaurar un Rover eliminado?",
                        body:"Ve a ⚙ Configuración → Papelera. Allí encontrarás los Rovers eliminados recientemente. Haz clic en '↩ Restaurar' para devolverlos al sistema." },
                      { title:"¿Qué es el Seguimiento Prioritario?",
                        body:"En el Panel del Clan (⚑), puedes marcar Rovers con ★ para identificarlos como prioritarios. Esto es una herramienta de organización interna que no modifica la progresión." },
                      { title:"¿Los Rovers pueden editar su progresión?",
                        body:"No. Los Rovers acceden al sistema sin credenciales y solo pueden consultar su progresión. Únicamente el Jefe de Clan autenticado puede marcar objetivos y registrar fechas." },
                    ].map((item, i) => (
                      <div key={i} style={{ padding:"12px 14px",
                        background:"rgba(14,7,2,0.97)",
                        border:"1px solid rgba(255,255,255,0.05)", borderRadius:"10px" }}>
                        <div style={{ fontSize:"12px", fontWeight:"700",
                          color:DS.text.primary, marginBottom:"5px",
                          fontFamily:DS.font.body }}>{item.title}</div>
                        <div style={{ fontSize:"11px", color:DS.text.tertiary,
                          fontFamily:DS.font.body, lineHeight:"1.5" }}>{item.body}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })()}

          {view === "reports" && isAdmin && (() => {
            const today = new Date().toLocaleDateString("es-GT",{day:"2-digit",month:"long",year:"numeric"});
            const qrBase = window.location.origin + window.location.pathname + "?rover=";
            const INACTIVE_DAYS = 21;

            // RF-07: filtrar scouts para el reporte
            const getFilteredScouts = (filter) => {
              if (filter === "all") return scouts;
              if (filter === "priority") return scouts.filter(s => priorityIds.has(s.id));
              if (filter === "active") return scouts.filter(s => totalStats(s.id).d > 0);
              if (filter === "inactive") return scouts.filter(s => {
                const allDates = Object.entries(dates)
                  .filter(([k]) => k.startsWith(`d:${s.id}|`))
                  .map(([,v]) => v).filter(Boolean).sort().reverse();
                const last = allDates[0];
                return !last || Math.floor((new Date()-new Date(last))/86400000) >= INACTIVE_DAYS;
              });
              return scouts.filter(s => getCurrStage(s.id).id === filter);
            };

            const openPreview = (html, title) => setReportPreview({ html, title });

            // Separated: PDF | HTML/Print | XLSX | CSV
            const REPORT_TYPES = [
              {
                id:"pdf_general", icon:"🔴", title:"Descargar PDF — Reporte General",
                desc:"Genera directamente un archivo PDF real con todos los Rovers filtrados. No requiere dialog de impresión.",
                action: async () => {
                  const filtered = getFilteredScouts(reportFilter);
                  setPdfProgress("Generando PDF…");
                  try {
                    await exportPDFGeneral(filtered, prog, dates, attendance, activeStagesList, priorityIds,
                      msg => setPdfProgress(msg));
                    addToast("PDF Reporte General descargado", "success", 2500);
                    addAudit(`PDF Reporte General — ${filtered.length} Rovers`, { type:"report" });
                  } catch {
                    addToast("Error al generar PDF", "error", 3000);
                  } finally { setPdfProgress(null); }
                }
              },
              {
                id:"general", icon:"📋", title:"Vista previa HTML / Imprimir",
                desc:"Abre vista previa en nueva pestaña. Usa Ctrl+P para imprimir o guardar como PDF.",
                action: () => {
                  const filtered = getFilteredScouts(reportFilter);
                  openPreview(exportAllHTML(filtered, prog, dates, attendance, filtered.map(s=>s.id), activeStagesList), "Reporte General del Clan");
                  addAudit(`Reporte HTML generado — ${filtered.length} Rovers`, { type:"report" });
                }
              },
              {
                id:"xlsx", icon:"📗", title:"Exportar a Excel (.xlsx) — OOXML real",
                desc:"Archivo Excel real (Office Open XML, vía SheetJS) con 3 hojas: Resumen, Progresión, Asistencia.",
                action: () => {
                  try {
                    exportXLSX(getFilteredScouts(reportFilter), prog, dates, attendance,
                      activeStagesList, priorityIds);
                    addToast("Excel descargado", "success", 2000);
                    addAudit(`Excel exportado — ${getFilteredScouts(reportFilter).length} Rovers`, { type:"export" });
                  } catch(e) { addToast("Error al generar Excel", "error", 3000); console.error(e); }
                }
              },
              {
                id:"csv", icon:"📄", title:"Exportar CSV (progresión)",
                desc:"Formato CSV ligero e independiente. No es Excel.",
                action: () => {
                  exportAllCSV(getFilteredScouts(reportFilter), prog, dates, activeStagesList);
                  addToast("CSV descargado", "success", 2000);
                  addAudit(`CSV exportado — ${getFilteredScouts(reportFilter).length} Rovers`, { type:"export" });
                }
              },
              {
                id:"attendance", icon:"📅", title:"Exportar Asistencia CSV",
                desc:"Historial completo de asistencia en formato CSV.",
                action: () => {
                  exportAllAttendanceCSV(scouts, attendance);
                  addToast("Asistencia descargada", "success", 2000);
                  addAudit("Asistencia exportada", { type:"export" });
                }
              },
            ];

            return (
              <div className="fade-in" style={{ flex:1, overflow:"auto", padding:"16px 18px",
                display:"flex", flexDirection:"column", gap:"14px" }}>

                <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
                  flexWrap:"wrap", gap:"10px" }}>
                  <div>
                    <div style={{ fontFamily:DS.font.display, fontSize:"15px",
                      fontWeight:"700", color:DS.text.primary }}>📄 Reportes y Exportación</div>
                    <div style={{ fontSize:"11px", color:DS.text.tertiary,
                      fontFamily:DS.font.body, marginTop:"2px" }}>Grupo Scout No. 12 · San Pedro Apóstol</div>
                  </div>
                </div>

                {/* RF-07: filtros de contenido */}
                <div style={{ background:"rgba(255,255,255,0.02)",
                  border:"1px solid rgba(255,255,255,0.04)", borderRadius:"10px",
                  padding:"12px 14px" }}>
                  <div style={{ fontSize:"9px", fontFamily:DS.font.display,
                    color:DS.brand.goldMuted, letterSpacing:"1.5px", fontWeight:"700",
                    marginBottom:"8px" }}>INCLUIR EN EL REPORTE</div>
                  <div style={{ display:"flex", gap:"6px", flexWrap:"wrap" }}>
                    {[
                      { id:"all", label:`Todos (${scouts.length})` },
                      ...activeStagesList.map(s => ({ id:s.id, label:`${s.numeral} ${s.name}` })),
                      { id:"active", label:"Con avance" },
                      { id:"inactive", label:"⚠ Inactivos" },
                      { id:"priority", label:"★ Prioritarios" },
                    ].map(f => (
                      <button key={f.id}
                        className={`rv-filter-chip${reportFilter===f.id?" on":""}`}
                        onClick={() => setReportFilter(f.id)}>
                        {f.label}
                        {reportFilter===f.id && f.id!=="all" &&
                          <span style={{ marginLeft:"4px", opacity:.7 }}>
                            ({getFilteredScouts(f.id).length})
                          </span>}
                      </button>
                    ))}
                  </div>
                </div>

                {/* RF-01/RF-02: Reportes principales */}
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",
                  gap:"10px" }}>
                  {REPORT_TYPES.map(r => (
                    <div key={r.id} className="rv-report-card">
                      <div style={{ fontSize:"22px", marginBottom:"8px" }}>{r.icon}</div>
                      <div style={{ fontFamily:DS.font.display, fontSize:"12px",
                        fontWeight:"700", color:DS.text.primary, marginBottom:"4px" }}>{r.title}</div>
                      <div style={{ fontSize:"11px", color:DS.text.tertiary,
                        fontFamily:DS.font.body, marginBottom:"12px",
                        lineHeight:"1.5" }}>{r.desc}</div>
                      <button className="btn-primary"
                        style={{ fontSize:"11px", padding:"6px 14px" }}
                        onClick={r.action}>
                        {r.id==="csv"||r.id==="attendance" ? "↓ Descargar" : "👁 Vista previa"}
                      </button>
                    </div>
                  ))}

                  {/* RF-01/RF-03: Fichas individuales */}
                  <div className="rv-report-card" style={{ gridColumn:"1/-1" }}>
                    <div style={{ fontSize:"22px", marginBottom:"8px" }}>📋</div>
                    <div style={{ fontFamily:DS.font.display, fontSize:"12px",
                      fontWeight:"700", color:DS.text.primary, marginBottom:"4px" }}>
                      Ficha Individual de Progresión
                    </div>
                    <div style={{ fontSize:"11px", color:DS.text.tertiary,
                      fontFamily:DS.font.body, marginBottom:"12px" }}>
                      Genera la ficha completa de un Rover: etapas, objetivos, fechas, QR y firmas.
                      Vista previa con botón Imprimir / Guardar como PDF.
                    </div>
                    <div style={{ display:"flex", gap:"7px", flexWrap:"wrap" }}>
                      {getFilteredScouts(reportFilter).map(sc => (
                        <div key={sc.id} style={{ display:"flex", gap:"4px", alignItems:"center" }}>
                          {/* PDF directo */}
                          <button className="btn-primary"
                            style={{ fontSize:"10px", padding:"4px 9px" }}
                            disabled={!!pdfProgress}
                            title={`Descargar PDF de ${sc.name}`}
                            onClick={async () => {
                              setPdfProgress(`Generando PDF de ${sc.name}…`);
                              try {
                                await exportPDFFicha(sc, prog, dates, attendance, activeStagesList,
                                  msg => setPdfProgress(msg));
                                addToast(`PDF de ${sc.name} descargado`, "success", 2500);
                                addAudit(`Ficha PDF descargada: ${sc.name}`, { type:"report", elementId: sc.id });
                              } catch {
                                addToast("Error al generar PDF", "error", 3000);
                              } finally { setPdfProgress(null); }
                            }}>
                            ↓ PDF
                          </button>
                          {/* Vista previa HTML */}
                          <button className="btn-action"
                            style={{ fontSize:"10px", padding:"4px 9px" }}
                            onClick={() => {
                              openPreview(
                                buildRoverFicha(sc, prog, dates, attendance, qrBase),
                                `Ficha — ${sc.name}`
                              );
                              addAudit(`Ficha HTML generada: ${sc.name}`, { type:"report", elementId: sc.id });
                            }}>
                            👁 {sc.name}
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* RF-08/RF-09: Backup y Restauración */}
                  <div className="rv-report-card" style={{ gridColumn:"1/-1" }}>
                    <div style={{ fontSize:"22px", marginBottom:"8px" }}>💾</div>
                    <div style={{ fontFamily:DS.font.display, fontSize:"12px",
                      fontWeight:"700", color:DS.text.primary, marginBottom:"4px" }}>
                      Copia de Seguridad y Restauración
                    </div>

                    {/* RF-08: Generate backup */}
                    <div style={{ display:"flex", gap:"10px", flexWrap:"wrap",
                      marginBottom:"12px", alignItems:"flex-start" }}>
                      <div style={{ flex:1, minWidth:"200px" }}>
                        <div style={{ fontSize:"11px", color:DS.text.tertiary,
                          fontFamily:DS.font.body, marginBottom:"8px" }}>
                          Genera un archivo JSON con todos los datos: Rovers, progresión,
                          etapas, configuración y asistencia.
                        </div>
                        <button className="btn-primary"
                          disabled={backupProcessing}
                          style={{ fontSize:"11px", padding:"7px 14px" }}
                          onClick={() => {
                            setBackupProcessing(true);
                            try {
                              generateBackup({
                                scouts: dataRef.current.scouts,
                                prog: dataRef.current.prog,
                                dates: dataRef.current.dates,
                                invDates: dataRef.current.invDates,
                                attendance: dataRef.current.attendance,
                                stagesConfig, sysConfig,
                              });
                              addToast("Copia de seguridad descargada", "success", 2500);
                              addAudit("Copia de seguridad generada", { type:"backup" });
                              saveSysConfig({ lastBackup: new Date().toISOString() });
                            } catch(e) {
                              addToast("Error al generar backup", "error", 3000);
                            } finally {
                              setBackupProcessing(false);
                            }
                          }}>
                          {backupProcessing ? "Generando…" : "↓ Descargar backup"}
                        </button>
                      </div>
                    </div>

                    {/* RF-09: Restore flow */}
                    <div style={{ borderTop:"1px solid rgba(255,255,255,0.05)",
                      paddingTop:"12px" }}>
                      <div style={{ fontSize:"11px", fontWeight:"700",
                        color:DS.text.secondary, fontFamily:DS.font.body,
                        marginBottom:"8px" }}>Restaurar copia de seguridad</div>

                      {/* Step 1: file select */}
                      {!restoreState && (
                        <div>
                          <div style={{ fontSize:"11px", color:DS.text.tertiary,
                            fontFamily:DS.font.body, marginBottom:"8px" }}>
                            Selecciona un archivo .json generado por este sistema.
                            Se validará antes de aplicar cualquier cambio.
                          </div>
                          <label className="btn-action"
                            style={{ fontSize:"11px", padding:"7px 14px", cursor:"pointer" }}
                            role="button" tabIndex={0}
                            onKeyDown={e => e.key==="Enter" && e.currentTarget.click()}>
                            📂 Seleccionar archivo de backup
                            <input type="file" accept=".json" style={{ display:"none" }}
                              onChange={e => {
                                const file = e.target.files[0];
                                if (!file) return;
                                const reader = new FileReader();
                                reader.onload = (ev) => {
                                  try {
                                    const parsed = JSON.parse(ev.target.result);
                                    const validation = validateBackup(parsed);
                                    if (!validation.valid) {
                                      addToast(`Archivo inválido: ${validation.error}`, "error", 4000);
                                      return;
                                    }
                                    // Full conflict analysis (pure, no Firebase writes)
                                    const analysis = analyzeConflicts(
                                      parsed,
                                      dataRef.current,
                                      stagesConfig
                                    );
                                    setRestoreState({ step:"preview", parsed, validation, analysis });
                                    addAudit("Archivo de backup cargado para revisión", { type:"restore" });
                                  } catch (err) {
                                    addToast("El archivo no es un JSON válido", "error", 3000);
                                    console.error("Backup parse error:", err);
                                  }
                                };
                                reader.readAsText(file);
                                e.target.value = "";
                              }} />
                          </label>
                        </div>
                      )}

                      {/* Step 2: Full conflict preview — no Firebase writes yet */}
                      {restoreState?.step === "preview" && (() => {
                        const { analysis, validation, parsed } = restoreState;
                        const { conflicts, autoMerge, identical } = analysis;
                        const hasRealConflicts = conflicts.length > 0;
                        const typeColors = {
                          rover:"#E67E22", prog:"#C0392B", dates:"#8E44AD",
                          attendance:"#2980B9", stage:"#27AE60", section:"#16A085",
                          objective:"#D35400"
                        };
                        return (
                          <div className="fade-in" style={{ display:"flex",
                            flexDirection:"column", gap:"10px" }}>

                            {/* Metadata */}
                            <div style={{ padding:"10px 14px",
                              background:"rgba(255,255,255,0.02)",
                              border:"1px solid rgba(255,255,255,0.06)",
                              borderRadius:"9px" }}>
                              <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                                color:DS.brand.goldMuted, letterSpacing:"1.5px",
                                fontWeight:"700", marginBottom:"8px" }}>
                                VISTA PREVIA DEL BACKUP
                              </div>
                              {[
                                ["Versión", validation.version],
                                ["Generado el", new Date(validation.generatedAt).toLocaleString("es-GT")],
                                ["Rovers en backup", validation.summary.scouts],
                                ["Etapas", validation.summary.stages],
                                ["Progresión", validation.summary.progKeys + " registros"],
                                ["Asistencia", (validation.summary.attKeys||0) + " registros"],
                              ].map(([k,v]) => (
                                <div key={k} style={{ display:"flex",
                                  justifyContent:"space-between", fontSize:"11px",
                                  fontFamily:DS.font.body, padding:"3px 0",
                                  borderBottom:"1px solid rgba(255,255,255,0.03)" }}>
                                  <span style={{ color:DS.text.tertiary }}>{k}</span>
                                  <span style={{ color:DS.text.primary, fontWeight:"600" }}>{v}</span>
                                </div>
                              ))}
                            </div>

                            {/* Analysis summary */}
                            <div style={{ display:"flex", gap:"8px", flexWrap:"wrap" }}>
                              {[
                                { label:"Nuevos / Agregar", count:autoMerge.length,
                                  color:DS.state.done, bg:"rgba(39,174,96,0.08)" },
                                { label:"Idénticos / Saltar", count:identical.length,
                                  color:DS.text.secondary, bg:"rgba(255,255,255,0.03)" },
                                { label:"⚠ Conflictos reales", count:conflicts.length,
                                  color:conflicts.length>0?DS.state.error:DS.text.muted,
                                  bg:conflicts.length>0?"rgba(192,57,43,0.08)":"rgba(255,255,255,0.02)" },
                              ].map(c => (
                                <div key={c.label} style={{ flex:1, minWidth:"90px",
                                  padding:"8px 10px", background:c.bg,
                                  border:`1px solid ${c.color}33`, borderRadius:"8px",
                                  textAlign:"center" }}>
                                  <div style={{ fontSize:"20px", fontWeight:"700",
                                    color:c.color, fontFamily:DS.font.display }}>{c.count}</div>
                                  <div style={{ fontSize:"9px", color:DS.text.muted,
                                    fontFamily:DS.font.body, marginTop:"2px" }}>{c.label}</div>
                                </div>
                              ))}
                            </div>

                            {/* Real conflicts — require decision */}
                            {conflicts.length > 0 && (
                              <div style={{ padding:"10px 12px",
                                background:"rgba(192,57,43,0.07)",
                                border:"1px solid rgba(192,57,43,0.25)",
                                borderRadius:"8px", display:"flex",
                                flexDirection:"column", gap:"6px" }}>
                                <div style={{ fontSize:"9px", fontFamily:DS.font.display,
                                  color:DS.state.error, letterSpacing:"1.5px",
                                  fontWeight:"700" }}>
                                  CONFLICTOS QUE REQUIEREN DECISIÓN
                                </div>
                                {conflicts.map((c, i) => (
                                  <div key={i} style={{ fontSize:"10.5px",
                                    fontFamily:DS.font.body, padding:"6px 8px",
                                    background:"rgba(255,255,255,0.02)",
                                    borderRadius:"6px",
                                    borderLeft:`3px solid ${typeColors[c.type]||DS.state.error}` }}>
                                    <div style={{ fontWeight:"700",
                                      color:typeColors[c.type]||DS.state.error,
                                      marginBottom:"2px" }}>{c.label}</div>
                                    {c.current && <div style={{ color:DS.text.tertiary }}>
                                      Actual: <span style={{ color:DS.text.secondary }}>{c.current.slice(0,60)}</span>
                                    </div>}
                                    {c.backup && <div style={{ color:DS.text.tertiary }}>
                                      Backup: <span style={{ color:DS.text.secondary }}>{c.backup.slice(0,60)}</span>
                                    </div>}
                                    <div style={{ fontSize:"9.5px", color:DS.brand.goldMuted,
                                      marginTop:"2px" }}>⚠ {c.impact}</div>
                                  </div>
                                ))}
                              </div>
                            )}

                            {/* Auto-merge items (safe) */}
                            {autoMerge.length > 0 && (
                              <div style={{ padding:"8px 12px",
                                background:"rgba(39,174,96,0.05)",
                                border:"1px solid rgba(39,174,96,0.2)",
                                borderRadius:"8px" }}>
                                <div style={{ fontSize:"9px", fontFamily:DS.font.display,
                                  color:DS.state.done, letterSpacing:"1.5px",
                                  fontWeight:"700", marginBottom:"5px" }}>
                                  SE AGREGARÁN AUTOMÁTICAMENTE (sin conflicto)
                                </div>
                                {autoMerge.slice(0,6).map((m,i) => (
                                  <div key={i} style={{ fontSize:"10.5px",
                                    color:DS.text.tertiary, fontFamily:DS.font.body,
                                    padding:"2px 0" }}>
                                    ✓ {m.label}
                                  </div>
                                ))}
                                {autoMerge.length > 6 && (
                                  <div style={{ fontSize:"10px", color:DS.text.muted,
                                    fontFamily:DS.font.body }}>
                                    ...y {autoMerge.length-6} más
                                  </div>
                                )}
                              </div>
                            )}

                            {/* Warning */}
                            <div style={{ fontSize:"10px", color:DS.text.tertiary,
                              fontFamily:DS.font.body, padding:"7px 10px",
                              background:"rgba(197,162,83,0.06)",
                              border:"1px solid rgba(197,162,83,0.2)",
                              borderRadius:"7px" }}>
                              ⚠ Se recomienda generar un backup del estado actual antes de restaurar.
                              {hasRealConflicts && " Existen conflictos que no se resolverán automáticamente."}
                            </div>

                            {/* Action buttons */}
                            <div style={{ display:"flex", gap:"8px", flexWrap:"wrap" }}>
                              {/* Merge safe only (respects conflicts) */}
                              <button className="btn-primary"
                                style={{ fontSize:"11px", padding:"7px 14px" }}
                                onClick={() => {
                                  if (hasRealConflicts) {
                                    setRestoreState(s => ({ ...s, step:"conflict_decision" }));
                                    return;
                                  }
                                  // No real conflicts — safe auto-merge
                                  const d = parsed.data;
                                  const cur = dataRef.current;
                                  const existingIds = new Set(cur.scouts.map(s=>s.id));
                                  const newScouts = [...cur.scouts,
                                    ...(d.scouts||[]).filter(sc => !existingIds.has(sc.id))];
                                  const newProg  = { ...d.prog,  ...cur.prog };
                                  const newDates = { ...d.dates, ...cur.dates };
                                  const newAtt   = { ...d.attendance, ...cur.attendance };
                                  setDoc(DOC_REF, { scouts:newScouts, prog:newProg,
                                    dates:newDates, invDates:{...(d.invDates||{}), ...(cur.invDates||{})},
                                    attendance:newAtt })
                                    .then(() => {
                                      addToast("Fusión completada sin pérdida de datos", "success", 3000);
                                      addAudit("Fusión segura desde backup", { type:"restore" });
                                      setRestoreState(null);
                                    })
                                    .catch(() => addToast("Error al fusionar", "error", 4000));
                                }}>
                                {hasRealConflicts ? "⚠ Revisar conflictos →" : "↩ Fusionar (seguro)"}
                              </button>
                              {/* Full replace — always warns */}
                              <button className="btn-danger"
                                style={{ fontSize:"11px", padding:"7px 14px" }}
                                onClick={() => showConfirm(
                                  "¿REEMPLAZAR todos los datos actuales con el backup? Esta acción eliminará todo el progreso actual. Es irreversible.",
                                  () => {
                                    // Guard: stagesConfig is null while Firebase load is pending.
                                    // stagesConfigSynced is set to true only after the first confirmed
                                    // snapshot (onSnapshot handler, L2504/L2517). Until then we cannot
                                    // distinguish null (not yet loaded) from confirmed-empty, so the
                                    // curHasStages branch would be unreliable. Block the action.
                                    if (!stagesConfigSynced) {
                                      addToast("Espera a que la configuración de etapas termine de cargar antes de restaurar.", "warning", 4000);
                                      return;
                                    }
                                    const d = parsed.data;
                                    const hasStages = d.stagesConfig && Object.keys(d.stagesConfig).length > 0;
                                    // stagesConfigSynced is true here, so stagesConfig is a confirmed
                                    // value: either null→{} (empty) or a populated object.
                                    // stagesConfig===null is impossible at this point.
                                    const curHasStages = stagesConfig != null && Object.keys(stagesConfig).length > 0;
                                    // F2 FIX: both writes are chained and any failure is
                                    // reported explicitly. No .catch(()=>{}) silencing.
                                    // Note: these are two separate Firestore writes — not atomic.
                                    // If DOC_REF succeeds and STAGES_DOC_REF fails, the partial
                                    // state is reported to the admin with a specific message.
                                    setDoc(DOC_REF, {
                                      scouts:d.scouts||[], prog:d.prog||{},
                                      dates:d.dates||{}, invDates:d.invDates||{},
                                      attendance:d.attendance||{} })
                                      .then(() => {
                                        if (!hasStages) {
                                          // Backup has no stages. If current config has stages,
                                          // they are NOT removed — this is not a full replacement.
                                          // Report honestly: rovers restored, stages unchanged.
                                          if (curHasStages) {
                                            addToast("Rovers y progresión restaurados. Las etapas NO se modificaron: el backup no las incluía.", "warning", 7000);
                                            addAudit("Restauración parcial: backup sin etapas — STAGES_DOC_REF no modificado", { type:"restore_warning" });
                                          } else {
                                            // Neither backup nor current has stages.
                                            // The onSnapshot listener on STAGES_DOC_REF will
                                            // auto-initialize default stages — that was not part of
                                            // this backup, so we cannot claim a full restore.
                                            addToast("Rovers y progresión restaurados. Las etapas se inicializarán a los valores predeterminados del sistema (el backup no las incluía).", "info", 7000);
                                            addAudit("Restauración de rovers completa; etapas predeterminadas serán inicializadas por el sistema", { type:"restore_warning" });
                                          }
                                          setRestoreState(null);
                                          return;
                                        }
                                        return setDoc(STAGES_DOC_REF, d.stagesConfig)
                                          .then(() => {
                                            addToast("Restauración completa aplicada", "success", 3000);
                                            addAudit("Restauración completa (reemplazar)", { type:"restore" });
                                            setRestoreState(null);
                                          })
                                          .catch(() => {
                                            // Rovers/prog saved — stages failed. Partial state.
                                            addToast("Rovers y progresión restaurados, pero las etapas NO pudieron guardarse. Intenta de nuevo o contacta al administrador.", "error", 7000);
                                            addAudit("Restauración parcial: DOC_REF guardado, STAGES_DOC_REF falló", { type:"restore_error" });
                                          });
                                      })
                                      .catch(() => {
                                        addToast("Error al restaurar — no se guardó ningún dato.", "error", 4000);
                                        addAudit("Error al restaurar (DOC_REF falló)", { type:"restore_error" });
                                      });
                                  }
                                )}>
                                ⚠ Reemplazar todo
                              </button>
                              <button className="btn-secondary"
                                style={{ fontSize:"11px", padding:"7px 14px" }}
                                onClick={() => {
                                  setRestoreState(null);
                                  addAudit("Restauración cancelada", { type:"restore" });
                                }}>
                                Cancelar
                              </button>
                            </div>
                          </div>
                        );
                      })()}

                      {/* Step 3: Conflict decision per conflict */}
                      {restoreState?.step === "conflict_decision" && (() => {
                        const { analysis, parsed } = restoreState;
                        const { conflicts, autoMerge } = analysis;
                        return (
                          <div className="fade-in" style={{ display:"flex",
                            flexDirection:"column", gap:"10px" }}>
                            <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                              color:DS.state.error, letterSpacing:"1.5px", fontWeight:"700" }}>
                              DECISIÓN SOBRE CONFLICTOS ({conflicts.length})
                            </div>
                            <div style={{ fontSize:"11px", color:DS.text.tertiary,
                              fontFamily:DS.font.body }}>
                              Se encontraron {conflicts.length} conflicto(s) que no pueden resolverse automáticamente.
                              Elige cómo proceder:
                            </div>
                            <div style={{ display:"flex", gap:"8px", flexWrap:"wrap" }}>
                              {/* Keep current — ignore backup conflicting data */}
                              <button className="btn-secondary"
                                style={{ fontSize:"11px", padding:"7px 14px" }}
                                onClick={() => {
                                  // Merge only auto-merge items, skip conflicts.
                                  // Current wins on every key that exists in both.
                                  const d = parsed.data;
                                  const cur = dataRef.current;
                                  const existingIds = new Set(cur.scouts.map(s=>s.id));
                                  const newScouts = [...cur.scouts,
                                    ...(d.scouts||[]).filter(sc => !existingIds.has(sc.id))];
                                  const newProg  = { ...d.prog,  ...cur.prog };
                                  const newDates = { ...d.dates, ...cur.dates };
                                  // F1 FIX: merge stagesConfig — current wins on conflict sids,
                                  // backup-only stages (new) are added automatically.
                                  // Key is the stageId (sid); no array positions used.
                                  const curStages = stagesConfig || {};
                                  const bkStages  = d.stagesConfig || {};
                                  const mergedStages = mergeStagesConfig(curStages, bkStages, "cur"); // deep merge — cur wins conflicts
                                  const hasStages = Object.keys(mergedStages).length > 0;
                                  setDoc(DOC_REF, { scouts:newScouts, prog:newProg,
                                    dates:newDates, invDates:{...(d.invDates||{}),...(cur.invDates||{})},
                                    attendance:{...d.attendance,...cur.attendance} })
                                    .then(() => {
                                      if (!hasStages) {
                                        addToast(`Fusión aplicada — datos actuales conservados en ${conflicts.length} conflicto(s)`, "success", 3500);
                                        addAudit(`Fusión con conflictos: datos actuales conservados (${conflicts.length} conflictos)`, { type:"restore" });
                                        setRestoreState(null);
                                        return;
                                      }
                                      return setDoc(STAGES_DOC_REF, mergedStages)
                                        .then(() => {
                                          addToast(`Fusión aplicada — datos actuales conservados en ${conflicts.length} conflicto(s)`, "success", 3500);
                                          addAudit(`Fusión con conflictos: datos actuales conservados (${conflicts.length} conflictos)`, { type:"restore" });
                                          setRestoreState(null);
                                        })
                                        .catch(() => {
                                          addToast("Rovers y progresión guardados. Error al guardar etapas — revisa la conexión.", "error", 5000);
                                          addAudit("Error al guardar etapas en fusión (conservar actuales)", { type:"restore_error" });
                                        });
                                    }).catch(() => {
                                      addToast("Error al aplicar fusión — no se guardó ningún dato.", "error", 4000);
                                      addAudit("Error al aplicar fusión (conservar actuales — DOC_REF falló)", { type:"restore_error" });
                                    });
                                }}>
                                ✓ Conservar datos actuales
                              </button>
                              {/* Use backup — backup overwrites conflicts */}
                              <button className="btn-danger"
                                style={{ fontSize:"11px", padding:"7px 14px" }}
                                onClick={() => showConfirm(
                                  `¿Aplicar datos del backup en los ${conflicts.length} conflicto(s)? Los datos actuales en conflicto serán reemplazados.`,
                                  () => {
                                    const d = parsed.data;
                                    const cur = dataRef.current;
                                    const existingIds = new Set(cur.scouts.map(s=>s.id));
                                    // Backup wins on conflicts
                                    const newScouts = [
                                      ...(d.scouts||[]),
                                      ...cur.scouts.filter(sc => !existingIds.has(sc.id))
                                    ];
                                    // F1 FIX: merge stagesConfig — backup wins on conflict sids,
                                    // current-only stages (absent from backup) are preserved.
                                    // Key is the stageId (sid); no array positions used.
                                    const curStages = stagesConfig || {};
                                    const bkStages  = d.stagesConfig || {};
                                    const mergedStages = mergeStagesConfig(curStages, bkStages, "bk"); // deep merge — backup wins conflicts
                                    const hasStages = Object.keys(mergedStages).length > 0;
                                    setDoc(DOC_REF, {
                                      scouts:newScouts, prog:{...cur.prog,...d.prog},
                                      dates:{...cur.dates,...d.dates},
                                      invDates:{...cur.invDates,...d.invDates},
                                      attendance:{...cur.attendance,...d.attendance} })
                                      .then(() => {
                                        if (!hasStages) {
                                          addToast("Fusión con datos de backup aplicada", "success", 3000);
                                          addAudit(`Fusión: datos de backup usados en conflictos`, { type:"restore" });
                                          setRestoreState(null);
                                          return;
                                        }
                                        return setDoc(STAGES_DOC_REF, mergedStages)
                                          .then(() => {
                                            addToast("Fusión con datos de backup aplicada", "success", 3000);
                                            addAudit(`Fusión: datos de backup usados en conflictos`, { type:"restore" });
                                            setRestoreState(null);
                                          })
                                          .catch(() => {
                                            addToast("Rovers y progresión guardados. Error al guardar etapas — revisa la conexión.", "error", 5000);
                                            addAudit("Error al guardar etapas en fusión (usar backup)", { type:"restore_error" });
                                          });
                                      }).catch(() => {
                                        addToast("Error al aplicar fusión — no se guardó ningún dato.", "error", 4000);
                                        addAudit("Error al aplicar fusión (usar backup — DOC_REF falló)", { type:"restore_error" });
                                      });
                                  }
                                )}>
                                ↩ Usar datos del backup
                              </button>
                              <button className="btn-secondary"
                                style={{ fontSize:"11px", padding:"7px 14px" }}
                                onClick={() => setRestoreState(s => ({...s, step:"preview"}))}>
                                ← Volver
                              </button>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                </div>

                {/* RF-08: Estadísticas automáticas del clan */}
                <div style={{ background:"rgba(14,7,2,0.97)",
                  border:"1px solid rgba(255,255,255,0.05)", borderRadius:"12px",
                  padding:"14px 16px" }}>
                  <div style={{ fontFamily:DS.font.display, fontSize:"9px",
                    color:DS.brand.goldMuted, letterSpacing:"1.5px",
                    fontWeight:"700", marginBottom:"10px" }}>ESTADÍSTICAS DEL CLAN</div>
                  {(() => {
                    const filtered = getFilteredScouts(reportFilter);
                    const totalD = filtered.reduce((a,s) => a+totalStats(s.id).d, 0);
                    const totalT = filtered.reduce((a,s) => a+totalStats(s.id).t, 0);
                    const avgPct = filtered.length
                      ? Math.round(filtered.reduce((a,s) => a+totalStats(s.id).pct, 0)/filtered.length)
                      : 0;
                    const completed = filtered.filter(s => totalStats(s.id).pct===100).length;
                    const stats = [
                      { label:"Rovers en reporte", value:filtered.length, color:DS.brand.gold },
                      { label:"Promedio de avance", value:avgPct+"%", color:DS.state.done },
                      { label:"Progresión completa", value:completed, color:DS.state.done },
                      { label:"Objetivos completados", value:totalD, color:DS.state.info },
                      { label:"Objetivos pendientes", value:totalT-totalD, color:DS.text.tertiary },
                      { label:"Total de objetivos", value:totalT, color:DS.text.secondary },
                    ];
                    return (
                      <div style={{ display:"flex", gap:"10px", flexWrap:"wrap" }}>
                        {stats.map((s,i) => (
                          <div key={i} style={{ flex:"1", minWidth:"100px" }}>
                            <div style={{ fontFamily:DS.font.display, fontSize:"18px",
                              fontWeight:"700", color:s.color, lineHeight:1 }}>{s.value}</div>
                            <div style={{ fontSize:"9.5px", color:DS.text.muted,
                              fontFamily:DS.font.body, marginTop:"2px" }}>{s.label}</div>
                          </div>
                        ))}
                      </div>
                    );
                  })()}
                </div>
              </div>
            );
          })()}

          {/* ── VISTA: ASISTENCIA ─────────────────────────────────────────────── */}
          {view === "attendance" && (() => {
            const { y, m } = attMonth;
            const monthLabel = new Date(y, m, 1).toLocaleDateString("es-GT", { month: "long", year: "numeric" });
            const daysInMonth = new Date(y, m + 1, 0).getDate();
            const todayStr = new Date().toISOString().split("T")[0];
            const pad2 = (n) => String(n).padStart(2, "0");
            const days = Array.from({ length: daysInMonth }, (_, i) => i + 1)
              .filter(d => [5, 6, 0].includes(new Date(y, m, d).getDay()));
            const dateOf = (d) => `${y}-${pad2(m + 1)}-${pad2(d)}`;
            const dowShort = (d) => ["D","L","M","X","J","V","S"][new Date(y, m, d).getDay()];

            return (
              <div className="fade-in" style={{ flex:1, overflow:"auto", padding:"14px 18px",
                display:"flex", flexDirection:"column", gap:"10px" }}>

                <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
                  flexWrap:"wrap", gap:"8px" }}>
                  <div style={{ fontFamily:"'Cinzel',serif", fontSize:"15px",
                    fontWeight:"700", color:DS.state.infoLight }}>📋 Cuadro de Asistencia — Todos los Rovers <span style={{ fontSize:"11px", color:"#5A7A8A", fontWeight:"600" }}>(Vie · Sáb · Dom)</span></div>
                  <div style={{ display:"flex", alignItems:"center", gap:"8px" }}>
                    <button className="att-nav btn-secondary"
                      onClick={() => { setAttMonth(({y,m}) => m===0 ? {y:y-1,m:11} : {y,m:m-1}); }}
                      style={{ padding:"4px 10px", fontSize:"12px",
                        color:DS.state.info, borderColor:"rgba(91,143,168,0.3)" }}>←</button>
                    <div style={{ fontFamily:"'Cinzel',serif", fontSize:"12.5px",
                      fontWeight:"700", color:DS.text.primary, textTransform:"capitalize",
                      minWidth:"130px", textAlign:"center" }}>{monthLabel}</div>
                    <button className="att-nav btn-secondary"
                      onClick={() => { setAttMonth(({y,m}) => m===11 ? {y:y+1,m:0} : {y,m:m+1}); }}
                      style={{ padding:"4px 10px", fontSize:"12px",
                        color:DS.state.info, borderColor:"rgba(91,143,168,0.3)" }}>→</button>
                  </div>
                </div>

                <div style={{ display:"flex", gap:"12px", flexWrap:"wrap", padding:"4px 9px",
                  background:"rgba(255,255,255,0.015)", borderRadius:"6px",
                  border:"1px solid rgba(255,255,255,0.04)" }}>
                  {Object.values(ATT_STATES).map(s => (
                    <span key={s.label} style={{ fontSize:"10.5px", color:s.color,
                      fontFamily:"'Source Sans 3',sans-serif" }}>● {s.short} = {s.label}</span>
                  ))}
                  {isAdmin && <span style={{ fontSize:"10.5px", color:DS.text.tertiary,
                    fontFamily:"'Source Sans 3',sans-serif" }}>— clic en la celda para marcar/cambiar</span>}
                </div>

                {scouts.length === 0 ? (
                  <div style={{ textAlign:"center", padding:"36px 12px", opacity:.6 }}>
                    <div style={{ fontSize:"32px", opacity:.2, marginBottom:"8px" }}>⚜</div>
                    <div style={{ fontSize:"12px", color:DS.text.tertiary,
                      fontFamily:"'Source Sans 3',sans-serif" }}>Aún no hay Rovers registrados</div>
                  </div>
                ) : (
                  <div className="att-table-wrap" style={{ overflowX:"auto", border:"1px solid rgba(91,143,168,0.18)",
                    borderRadius:"9px", background:"rgba(14,7,2,0.97)" }}>
                    <table style={{ borderCollapse:"collapse", width:"100%", fontSize:"10.5px" }}>
                      <thead>
                        <tr>
                          <th scope="col"
                            style={{ position:"sticky", left:0, background:"rgba(13,14,16,0.99)",
                            zIndex:2, textAlign:"left", padding:"6px 10px", minWidth:"118px",
                            borderBottom:"1px solid rgba(255,255,255,0.06)",
                            borderRight:"1px solid rgba(255,255,255,0.06)",
                            fontFamily:"'Cinzel',serif", fontSize:"9.5px", color:DS.text.tertiary,
                            letterSpacing:"0.5px" }}>ROVER</th>
                          {days.map((d, di) => {
                            const isToday = dateOf(d) === todayStr;
                            const isFri = new Date(y, m, d).getDay() === 5;
                            const weekStart = isFri && di > 0;
                            return (
                              <th key={d} style={{ padding:"4px 2px", textAlign:"center",
                                minWidth:"22px", borderBottom:"1px solid rgba(255,255,255,0.06)",
                                borderLeft: weekStart ? "2px solid rgba(197,162,83,0.25)" : "none",
                                background: isToday ? "rgba(197,162,83,0.1)" : "transparent" }}>
                                <div style={{ color: isToday ? DS.brand.gold : DS.text.tertiary,
                                  fontWeight: isToday ? "700" : "600",
                                  fontFamily:"'Source Sans 3',sans-serif" }}>{d}</div>
                                <div style={{ color:DS.text.muted, fontSize:"8px",
                                  fontFamily:"'Source Sans 3',sans-serif" }}>{dowShort(d)}</div>
                              </th>
                            );
                          })}
                          <th style={{ padding:"4px 8px", textAlign:"center", minWidth:"46px",
                            borderBottom:"1px solid rgba(255,255,255,0.06)",
                            borderLeft:"1px solid rgba(255,255,255,0.06)",
                            fontFamily:"'Cinzel',serif", fontSize:"9px", color:DS.text.tertiary }}>%</th>
                        </tr>
                      </thead>
                      <tbody>
                        {scouts.map((sc, ri) => {
                          const stats = attendanceStats(sc.id);
                          return (
                            <tr key={sc.id} style={{ background: ri % 2 ? "rgba(255,255,255,0.012)" : "transparent" }}>
                              <td style={{ position:"sticky", left:0,
                                background: ri % 2 ? "rgba(20,11,5,0.99)" : "rgba(16,9,4,0.99)",
                                zIndex:1, padding:"4px 10px", whiteSpace:"nowrap",
                                borderRight:"1px solid rgba(255,255,255,0.06)",
                                borderBottom:"1px solid rgba(255,255,255,0.03)",
                                color:"#D8D0C8", fontWeight:"600", fontSize:"11px",
                                fontFamily:"'Source Sans 3',sans-serif", cursor:"pointer" }}
                                onClick={() => { setSelId(sc.id); setView("main"); setStageId(getCurrStage(sc.id).id); setDetailFilter("all"); }}
                                title="Ver progresión">
                                {sc.name}
                              </td>
                              {days.map((d, di) => {
                                const date = dateOf(d);
                                const rec = attendance[attKey(sc.id, date)];
                                const st = rec ? ATT_STATES[rec.status] : null;
                                const weekStart = new Date(y, m, d).getDay() === 5 && di > 0;
                                return (
                                  <td key={d} style={{ padding:"2px", textAlign:"center",
                                    borderBottom:"1px solid rgba(255,255,255,0.03)",
                                    borderLeft: weekStart ? "2px solid rgba(197,162,83,0.12)" : "none" }}>
                                    <div className="att-cell" onClick={() => cycleAttStatus(sc.id, date)}
                                      title={rec?.note || (isAdmin ? "Clic para marcar" : "")}
                                      style={{ width:"19px", height:"19px", margin:"0 auto",
                                        borderRadius:"4px", cursor: isAdmin ? "pointer" : "default",
                                        display:"flex", alignItems:"center", justifyContent:"center",
                                        background: st ? `${st.color}30` : "rgba(255,255,255,0.025)",
                                        border: st ? `1px solid ${st.color}77` : "1px solid rgba(255,255,255,0.05)",
                                        color: st ? st.color : "transparent",
                                        fontSize:"9px", fontWeight:"700",
                                        fontFamily:"'Source Sans 3',sans-serif",
                                        transition:"transform 0.1s" }}>
                                      {st ? st.short : ""}
                                    </div>
                                  </td>
                                );
                              })}
                              <td style={{ padding:"4px 8px", textAlign:"center",
                                borderLeft:"1px solid rgba(255,255,255,0.06)",
                                borderBottom:"1px solid rgba(255,255,255,0.03)",
                                color:DS.state.infoLight, fontWeight:"700", fontSize:"11px",
                                fontFamily:"'Source Sans 3',sans-serif" }}>
                                {stats.t > 0 ? `${stats.pct}%` : "—"}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}

                <div style={{ fontSize:"10.5px", color:DS.text.tertiary,
                  fontFamily:"'Source Sans 3',sans-serif" }}>
                  Clic en el nombre de un Rover para ver su progresión completa.
                </div>
              </div>
            );
          })()}

          {/* ── VISTA: RADAR ──────────────────────────────────────────────────── */}
          {view === "radar" && selScout && (
            <div className="fade-in" style={{ flex:1, overflow:"auto", padding:"16px 20px" }}>
              <button className="back-btn btn-secondary"
                onClick={() => setView("main")}
                style={{ padding:"6px 12px", fontSize:"12px", marginBottom:"14px" }}>
                ← Volver
              </button>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",
                gap:"18px" }}>
                {activeStagesList.map(st => (
                  <div key={st.id} style={{ background:"rgba(14,7,2,0.97)",
                    border:`1px solid ${st.color}25`, borderRadius:"12px", padding:"16px" }}>
                    <div style={{ textAlign:"center", marginBottom:"6px" }}>
                      <span style={{ fontFamily:"'Cinzel',serif", fontSize:"13px",
                        fontWeight:"700", color:st.color }}>{st.name}</span>
                    </div>
                    <div className="radar-svg-wrap">
                      <RadarChart stage={st} scoutId={selId} prog={prog} size={220} />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ── VISTA PRINCIPAL: Vacío (sin rover seleccionado) ───────────────── */}
          {view === "main" && !selScout && (
            <div style={{ flex:1, display:"flex", flexDirection:"column",
              alignItems:"center", justifyContent:"center", gap:"14px",
              padding:"20px", textAlign:"center" }}>
              <div style={{ fontSize:"80px", opacity:.05 }}>⚜</div>
              <div style={{ fontFamily:"'Cinzel',serif", fontSize:"18px",
                color:DS.brand.red, letterSpacing:"1px" }}>Selecciona un Rover</div>
              <div style={{ fontSize:"13px", color:DS.text.tertiary,
                fontFamily:"'Source Sans 3',sans-serif" }}>
                {isAdmin ? "o agrega uno en el panel lateral" : "elige uno en el panel lateral para ver su progresión"}
              </div>
            </div>
          )}

          {/* ── VISTA PRINCIPAL: layout con nav de etapas (RF-01 a RF-09) ─────── */}
          {view === "main" && selScout && (() => {
            const currStage = getCurrStage(selId);

            // RF-05: la última etapa visitada persiste en stageId (ya es estado React)
            // Al seleccionar un Rover nuevo, getCurrStage determina la etapa inicial

            // helpers de estado visual (sin tocar Firebase)
            const itemStatus = (stId, secName, idx) => {
              const k  = pKey(selId, stId, secName, idx);
              const dk = dKey(selId, stId, secName, idx);
              if (prog[k]) return "done";
              if (dates[dk]) return "inprogress";
              return "pending";
            };
            const STATUS_CFG = {
              done:       { label:"Completado",  dot:DS.state.done, icon:"✓" },
              inprogress: { label:"En proceso",  dot:DS.state.inprogress, icon:"◐" },
              pending:    { label:"No iniciado", dot:"rgba(255,255,255,0.18)", icon:"" },
            };
            const FILTERS = [
              { id:"all",        label:"Todos"       },
              { id:"pending",    label:"Pendientes"  },
              { id:"inprogress", label:"En proceso"  },
              { id:"done",       label:"Completados" },
            ];

            const isStageCollapsed = (st) => {
              if (collapsedStages[st.id] !== undefined) return collapsedStages[st.id];
              return st.id !== currStage.id;
            };
            const toggleStageCollapse = (stId) =>
              setCollapsedStages(p => ({
                ...p,
                [stId]: !isStageCollapsed(activeStagesList.find(s => s.id === stId) || activeStagesList[0])
              }));

            const selectedSt = activeStagesList.find(s => s.id === stageId) || currStage;
            const selStats = stageStats(selId, selectedSt);
            const inv = invDates[invKey(selId, selectedSt.id)];

            // RF-02: resumen de etapa seleccionada
            let doneCount = 0, inprogCount = 0, pendCount = 0;
            selectedSt.sections.forEach(sec => sec.items.forEach((_, idx) => {
              const st = itemStatus(selectedSt.id, sec.name, idx);
              if (st === "done") doneCount++;
              else if (st === "inprogress") inprogCount++;
              else pendCount++;
            }));

            return (
              <div className="rv-content-area" style={{ flex:1, overflow:"hidden",
                display:"flex", flexDirection:"row" }}>

                {/* ── RF-01: Panel de navegación por etapas — permanente ─────────── */}
                <nav className="rv-stage-nav" aria-label="Navegación por etapas"
                  style={{ width:"160px", flexShrink:0,
                  background:DS.bg.surface,
                  borderRight:"1px solid rgba(197,162,83,0.08)",
                  overflowY:"auto", overflowX:"hidden",
                  display:"flex", flexDirection:"column", gap:"3px",
                  padding:"10px 6px" }}>

                  {/* RF-06: Acción de exportar — siempre visible en nav */}
                  <div style={{ padding:"0 2px 8px", borderBottom:"1px solid rgba(255,255,255,0.04)",
                    marginBottom:"4px" }}>
                    <div style={{ fontSize:"9px", fontFamily:"'Cinzel',serif",
                      color:DS.text.muted, letterSpacing:"1.5px", fontWeight:"700",
                      marginBottom:"6px" }}>ETAPAS</div>

                    {/* RF-01/RF-02: botones de etapa con % individual */}
                    {activeStagesList.map(st => {
                      const s = stageStats(selId, st);
                      const isActive = st.id === currStage.id;
                      const isSel = st.id === stageId;
                      return (
                        <button key={st.id}
                          className={`stage-nav-btn rv-stage-nav-item${isSel ? " active-stage" : ""}`}
                          onClick={() => { setStageId(st.id); setDetailFilter("all"); }}
                          style={{ width:"100%", display:"flex", flexDirection:"column",
                            alignItems:"flex-start", gap:"4px", padding:"8px 8px",
                            background: isSel ? `${st.color}18` : "transparent",
                            border:`1px solid ${isSel ? st.color+"44" : "transparent"}`,
                            borderRadius:"8px", cursor:"pointer", marginBottom:"2px",
                            transition:"all 0.12s",
                            opacity: (!isActive && s.pct === 0) ? 0.45 : 1 }}>
                          <div style={{ display:"flex", alignItems:"center",
                            justifyContent:"space-between", width:"100%" }}>
                            <span style={{ fontFamily:"'Cinzel',serif", fontSize:"9px",
                              fontWeight:"700", color:st.color, opacity:.8 }}>{st.numeral}</span>
                            {/* RF-01: porcentaje individual siempre visible */}
                            <span className="rv-stage-nav-pct"
                              style={{ fontSize:"12px", fontWeight:"700",
                                color: isSel ? st.color : `${st.color}88`,
                                fontFamily:"'Cinzel',serif" }}>{s.pct}%</span>
                          </div>
                          <span className="rv-stage-name"
                            style={{ fontSize:"10.5px", fontWeight:"600", textAlign:"left",
                              color: isSel ? DS.text.primary : DS.text.tertiary, lineHeight:"1.3",
                              fontFamily:"'Source Sans 3',sans-serif" }}>{st.name}</span>
                          {/* RF-01/RF-02/RF-04: barra + contadores */}
                          <div className="rv-prog-track" style={{ width:"100%", height:"4px", marginTop:"4px" }}>
                            <div className={`rv-prog-fill${s.pct===100?" done":""}`}
                              style={{ width:`${s.pct}%`,
                                background: s.pct===100 ? undefined : st.color,
                                opacity: isSel ? 1 : 0.45 }} />
                          </div>
                          <div style={{ display:"flex", alignItems:"center",
                            justifyContent:"space-between", width:"100%", marginTop:"2px" }}>
                            <span style={{ fontSize:"8.5px", color:DS.text.muted,
                              fontFamily:"'Source Sans 3',sans-serif" }}>{s.d}/{s.t}</span>
                            {s.pct === 100
                              ? <span style={{ fontSize:"8px", color:DS.state.done, lineHeight:1 }}>✓</span>
                              : isActive
                                ? <span style={{ fontSize:"8px", color:st.color, lineHeight:1, opacity:.9 }}>●</span>
                                : null
                            }
                          </div>
                        </button>
                      );
                    })}
                  </div>

                  {/* RF-06: Acciones permanentes */}
                  <button onClick={() => setView("radar")}
                    className="btn-action"
                    style={{ width:"100%", padding:"7px 8px",
                      fontSize:"10.5px", textAlign:"left", marginBottom:"3px" }}>
                    📊 Radar
                  </button>

                  {/* Export — RF-06 */}
                  <div style={{ position:"relative" }}>
                    <button onClick={(e) => { e.stopPropagation(); setExportMenu(m => !m); }}
                      className="btn-action"
                      style={{ width:"100%", padding:"7px 8px",
                        fontSize:"10.5px", textAlign:"left" }}>
                      ↓ Exportar
                    </button>
                    {exportMenu && (
                      <div style={{ position:"fixed", left:"168px", top:"auto",
                        background:DS.bg.card, border:"1px solid rgba(197,162,83,0.25)",
                        borderRadius:"9px", padding:"6px", zIndex:300, minWidth:"240px",
                        boxShadow:"0 8px 24px rgba(0,0,0,0.5)" }}
                        onClick={e => e.stopPropagation()}>
                        {[
                          ["📄 Reporte HTML / PDF", () => { const html = exportAllHTML(scouts, prog, dates, attendance, null, activeStagesList); const w = window.open("", "_blank"); w.document.write(html); w.document.close(); setExportMenu(false); }],
                          ["📊 CSV — Progresión", () => { exportAllCSV(scouts, prog, dates, activeStagesList); setExportMenu(false); }],
                          ["📋 CSV — Asistencia", () => { exportAllAttendanceCSV(scouts, attendance); setExportMenu(false); }],
                        ].map(([lbl, fn], i) => (
                          <button key={i} className="ebtn" onClick={fn}
                            style={{ display:"block", width:"100%", textAlign:"left",
                              padding:"9px 10px", background:"none", border:"none",
                              borderRadius:"6px", color:"#D8D0C8", fontSize:"12px",
                              cursor:"pointer", fontFamily:"'Source Sans 3',sans-serif",
                              transition:"background 0.15s" }}>{lbl}</button>
                        ))}
                      </div>
                    )}
                  </div>
                </nav>

                {/* ── RF-03/RF-04: Contenido de la etapa seleccionada ───────────── */}
                <div key={stageId} className="rv-stage-content"
                  style={{ flex:1, overflow:"hidden", display:"flex",
                  flexDirection:"column", minWidth:0, minHeight:0 }}>

                  {/* Encabezado de etapa — fijo arriba */}
                  <div style={{ padding:"12px 16px",
                    borderBottom:`1px solid ${selectedSt.color}22`,
                    background:`linear-gradient(135deg, ${selectedSt.bg}CC, rgba(10,6,4,0.99))`,
                    flexShrink:0 }}>
                    <div style={{ display:"flex", alignItems:"flex-start",
                      justifyContent:"space-between", gap:"10px" }}>
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:"8px",
                          flexWrap:"wrap", marginBottom:"4px" }}>
                          <span style={{ fontFamily:"'Cinzel',serif", fontSize:"13px",
                            fontWeight:"700", color:selectedSt.color, opacity:.7 }}>{selectedSt.numeral}</span>
                          <span style={{ fontFamily:"'Cinzel',serif", fontSize:"16px",
                            fontWeight:"700", color:DS.text.primary }}>{selectedSt.name}</span>
                          {selectedSt.id === currStage.id && (
                            <span className="rv-badge active">● ACTIVA</span>
                          )}
                          {selStats.pct === 100 && (
                            <span className="rv-badge complete">✓ COMPLETA</span>
                          )}
                        </div>
                        <div style={{ fontSize:"11px", color:DS.text.tertiary,
                          fontFamily:"'Source Sans 3',sans-serif" }}>{selectedSt.subtitle}</div>
                      </div>
                      {/* RF-01/RF-02/RF-04/RF-05: porcentaje prominente */}
                      <div style={{ textAlign:"right", flexShrink:0 }}>
                        <div style={{ fontFamily:"'Cinzel',serif", fontSize:"34px",
                          fontWeight:"700", lineHeight:1,
                          color: selStats.pct===100 ? DS.state.done : selectedSt.color,
                          transition:"color 0.25s" }}>{selStats.pct}%</div>
                        <div style={{ fontSize:"10px", color:DS.text.muted,
                          fontFamily:"'Source Sans 3',sans-serif",
                          marginTop:"3px" }}>{selStats.d}/{selStats.t} req.</div>
                      </div>
                    </div>
                    {/* RF-10: última modificación */}
                    {(() => {
                      const lastMod = Object.entries(dates)
                        .filter(([k]) => k.startsWith(`d:${selId}|${selectedSt.id}|`))
                        .map(([,v]) => v).filter(Boolean).sort().reverse()[0];
                      return lastMod ? (
                        <div style={{ fontSize:"9.5px", color:DS.text.muted,
                          fontFamily:"'Source Sans 3',sans-serif", marginTop:"4px" }}>
                          Última actividad: <span style={{ color:DS.text.tertiary }}>{lastMod}</span>
                        </div>
                      ) : null;
                    })()}
                    {/* RF-01/RF-02/RF-03/RF-05: barra principal + métricas */}
                    <div style={{ marginTop:"10px" }}>
                      <div className="rv-prog-track" style={{ height:"8px" }}>
                        <div className={`rv-prog-fill${selStats.pct===100?" done":""}`}
                          style={{ width:`${selStats.pct}%`,
                            background: selStats.pct===100 ? undefined
                              : `linear-gradient(90deg, ${selectedSt.color}BB, ${selectedSt.color})` }} />
                      </div>
                      <div style={{ display:"flex", alignItems:"center",
                        justifyContent:"space-between", marginTop:"6px",
                        flexWrap:"wrap", gap:"4px" }}>
                        <div style={{ display:"flex", gap:"12px", flexWrap:"wrap" }}>
                          <span style={{ fontSize:"11px", fontFamily:"'Source Sans 3',sans-serif",
                            color:DS.text.tertiary }}>
                            <span style={{ color:DS.state.done, fontWeight:"700" }}>{doneCount}</span>
                            {" de "}{selStats.t}{" completados"}
                          </span>
                          {pendCount > 0 && (
                            <span style={{ fontSize:"11px", fontFamily:"'Source Sans 3',sans-serif",
                              color:DS.text.tertiary }}>
                              <span style={{ color:selectedSt.color, fontWeight:"700" }}>{pendCount}</span>
                              {" pendientes"}
                            </span>
                          )}
                          {inprogCount > 0 && (
                            <span style={{ fontSize:"11px", fontFamily:"'Source Sans 3',sans-serif",
                              color:DS.text.tertiary }}>
                              <span style={{ color:DS.state.inprogress, fontWeight:"700" }}>{inprogCount}</span>
                              {" en proceso"}
                            </span>
                          )}
                        </div>
                        {selStats.pct === 100 && (
                          <span style={{ fontSize:"10px", color:DS.state.done,
                            fontFamily:"'Source Sans 3',sans-serif", fontWeight:"700" }}>
                            ✓ Etapa completada
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* RF-06/RF-08: Filtros — posición y estilo consistentes */}
                  <div style={{ display:"flex", gap:"5px", flexWrap:"wrap",
                    padding:"8px 14px",
                    borderBottom:"1px solid rgba(255,255,255,0.03)",
                    alignItems:"center", flexShrink:0 }}>
                    <span style={{ fontSize:"9px", fontFamily:"'Cinzel',serif",
                      color:DS.text.muted, letterSpacing:"1px", fontWeight:"700",
                      marginRight:"4px" }}>FILTRAR:</span>
                    {FILTERS.map(f => {
                      const isAct = detailFilter === f.id;
                      return (
                        <button key={f.id}
                          className={`rv-filter-btn${isAct ? " active" : ""}`}
                          aria-pressed={isAct}
                          aria-label={`Filtrar: ${f.label}`}
                          onClick={() => setDetailFilter(f.id)}
                          style={{
                            background: isAct ? `${selectedSt.color}20` : undefined,
                            borderColor: isAct ? `${selectedSt.color}55` : undefined,
                            color: isAct ? selectedSt.color : undefined,
                          }}>{f.label}</button>
                      );
                    })}
                  </div>

                  {/* Contenido scrollable — único contenedor con scroll */}
                  <div ref={scrollRef}
                    style={{ flex:1, overflowY:"auto", overflowX:"hidden",
                    minHeight:0, padding:"12px 14px",
                    display:"flex", flexDirection:"column", gap:"10px" }}>

                    {/* Investidura */}
                    <div onClick={() => toggleInv(selectedSt.id)}
                      style={{ display:"flex", alignItems:"center", gap:"8px",
                        padding:"8px 12px",
                        background: inv ? `${selectedSt.color}12` : "rgba(255,255,255,0.015)",
                        border:`1px solid ${inv ? selectedSt.color+"33" : "rgba(255,255,255,0.04)"}`,
                        borderRadius:"7px", cursor: isAdmin ? "pointer" : "default",
                        flexShrink:0 }}>
                      <div style={{ width:"15px", height:"15px", borderRadius:"4px",
                        border:`1.5px solid ${inv ? selectedSt.color : "rgba(255,255,255,0.15)"}`,
                        background: inv ? selectedSt.color : "transparent",
                        display:"flex", alignItems:"center", justifyContent:"center",
                        flexShrink:0, fontSize:"9px", color:"#fff" }}>{inv && "✓"}</div>
                      <span style={{ fontSize:"11.5px", fontWeight:"700",
                        color: inv ? selectedSt.color : DS.text.tertiary,
                        fontFamily:"'Source Sans 3',sans-serif" }}>🎖️ {selectedSt.investidura}</span>
                      {inv && <span style={{ fontSize:"10px", color:DS.text.tertiary,
                        fontFamily:"'Source Sans 3',sans-serif", marginLeft:"4px" }}>· {inv}</span>}
                    </div>

                    {/* Secciones de la etapa seleccionada */}
                    {selectedSt.sections.map(sec => {
                      let sd = 0;
                      sec.items.forEach((_, i) => {
                        if (prog[pKey(selId, selectedSt.id, sec.name, i)]) sd++;
                      });
                      const secKey = `${selectedSt.id}|${sec.name}`;
                      const collapsed = !!collapsedSecs[secKey];

                      const filteredItems = sec.items.map((rawItem, idx) => {
                        const itemObj = typeof rawItem === "object" ? rawItem : null;
                        const itemText = itemObj ? (itemObj.text || rawItem) : rawItem;
                        // Use stable read: getItemProg for done, legacy itemStatus for inprogress
                        const isDone = getItemProg(selId, selectedSt, sec, itemObj || { legacyIdx: idx, legacySecName: sec.name }, prog);
                        const dateVal = getItemDate(selId, selectedSt, sec, itemObj || { legacyIdx: idx, legacySecName: sec.name }, dates);
                        const status = isDone ? "done"
                          : dateVal ? "inprogress"
                          : "pending";
                        return { item: itemText, idx, itemObj, status, dateVal };
                      }).filter(({ status }) => {
                        if (detailFilter === "all") return true;
                        return status === detailFilter;
                      });

                      if (filteredItems.length === 0) return null;

                      // RF-10: sección de la etapa activa recibe más protagonismo si tiene pendientes
                      const secPct = Math.round(sd/sec.items.length*100);
                      const secIsPrimary = selectedSt.id === currStage.id && secPct < 100;

                      return (
                        // RF-01/RF-05/RF-08/RF-11: tarjeta unificada, superficie interactiva
                        <div key={sec.name}
                          className={`rv-card${secIsPrimary ? "" : ""}`}
                          style={{
                            borderColor: secIsPrimary
                              ? `${selectedSt.color}22`
                              : "rgba(255,255,255,0.04)",
                            boxShadow: secIsPrimary
                              ? `0 2px 16px ${selectedSt.color}08`
                              : "none",
                          }}>

                          {/* RF-04/RF-06/RF-11: cabecera con acciones rápidas */}
                          <div className="rv-card-hdr sec-hdr"
                            role="button" tabIndex={0}
                            aria-expanded={!collapsed}
                            aria-label={`${sec.name} — ${collapsed?"expandir":"contraer"}`}
                            onKeyDown={e => (e.key==="Enter"||e.key===" ") && toggleSec(secKey)}
                            onClick={() => toggleSec(secKey)}>
                            <div style={{ flex:1, minWidth:0 }}>
                              <div style={{ fontFamily:"'Cinzel',serif", fontSize:"9.5px",
                                fontWeight:"700", color:DS.brand.gold, letterSpacing:"1px",
                                textTransform:"uppercase", marginBottom:"2px" }}>{sec.name}</div>
                              <div style={{ fontSize:"9.5px", color:selectedSt.color,
                                fontFamily:"'Source Sans 3',sans-serif",
                                opacity:.8 }}>Área: {sec.area}</div>
                            </div>
                            <div style={{ display:"flex", alignItems:"center",
                              gap:"8px", flexShrink:0 }}>
                              {/* RF-11: acción rápida — solo admin, visible en hover */}
                              {isAdmin && !collapsed && (
                                <div className="rv-sec-actions" onClick={e => e.stopPropagation()}>
                                  <button className="rv-qbtn"
                                    title="Marcar todos como completados"
                                    onClick={() => {
                                      sec.items.forEach((rawIt, i) => {
                                        const itObj = typeof rawIt === "object" ? rawIt : null;
                                        const done = getItemProg(selId, selectedSt, sec,
                                          itObj || { legacyIdx: i, legacySecName: sec.name }, prog);
                                        if (!done)
                                          toggleItem(selectedSt.id, sec.name, i, itObj, selectedSt, sec);
                                      });
                                    }}>✓ todos</button>
                                </div>
                              )}
                              <div style={{ textAlign:"right" }}>
                                <div style={{ fontSize:"13px", fontWeight:"700",
                                  color:selectedSt.color, fontFamily:"'Cinzel',serif",
                                  lineHeight:1 }}>{secPct}%</div>
                                <div style={{ fontSize:"9px", color:DS.text.muted,
                                  fontFamily:"'Source Sans 3',sans-serif",
                                  marginTop:"1px" }}>{sd}/{sec.items.length}</div>
                              </div>
                              <div style={{ width:"20px", height:"20px",
                                display:"flex", alignItems:"center", justifyContent:"center",
                                borderRadius:"5px", background:"rgba(255,255,255,0.03)",
                                color:DS.text.muted, fontSize:"10px",
                                flexShrink:0 }}>{collapsed ? "▼" : "▲"}</div>
                            </div>
                          </div>

                          {/* RF-01: barra de progreso de sección */}
                          <div className="rv-prog-track" style={{ height:"3px",
                            margin:"0", borderRadius:"0",
                            background:"rgba(255,255,255,0.04)" }}>
                            <div className={`rv-prog-fill${secPct===100?" done":""}`}
                              style={{ width:`${secPct}%`,
                                background: secPct===100 ? undefined : selectedSt.color }} />
                          </div>

                          {/* RF-09: contenido colapsable */}
                          {!collapsed && (
                            <div className="rv-card-body" style={{ padding:"6px 10px 10px",
                              display:"flex", flexDirection:"column", gap:"2px" }}>
                              {filteredItems.map(({ item, idx, status }) => {
                                // Use pre-computed dateVal from filteredItems (stable+legacy)
                                const itemDate = dateVal || "";
                                const expK = `${selectedSt.id}|${sec.name}|${idx}`;
                                const isExp = expItem === expK;
                                const cfg = STATUS_CFG[status];
                                const isDone = status === "done";
                                const isInProg = status === "inprogress";

                                return (
                                  // RF-05/RF-08: fila completa interactiva con estados hover/active
                                  <div key={idx}>
                                    <div
                                      className={`rv-item-row check-row${isDone ? " done" : ""}${isExp ? " expanded" : ""}`}
                                      style={{ borderColor: isExp ? `${selectedSt.color}22` : "transparent",
                                        background: isExp ? `${selectedSt.color}06` : "transparent" }}>

                                      {/* RF-06/RF-11: checkbox con checkPop animation */}
                                      <div className={`rv-chk${poppingItems.has(
                                          itemObj?.stableId && selectedSt.stableId && sec.stableId
                                            ? pKeyS(selId, selectedSt.stableId, sec.stableId, itemObj.stableId)
                                            : pKey(selId, selectedSt.id, sec.name, idx)
                                        ) ? " popping" : ""}`}
                                        onClick={() => toggleItem(selectedSt.id, sec.name, idx, itemObj, selectedSt, sec)}
                                        title={cfg.label}
                                        style={{
                                          borderColor: isDone ? selectedSt.color : isInProg ? DS.state.inprogress : "rgba(255,255,255,0.18)",
                                          background: isDone ? selectedSt.color : isInProg ? "rgba(230,126,34,0.15)" : "transparent",
                                          cursor: isAdmin ? "pointer" : "default",
                                          width:"18px", height:"18px", borderRadius:"5px",
                                          fontSize:"10px",
                                          boxShadow: isDone ? `0 0 6px ${selectedSt.color}55` : "none",
                                          transition:"all 0.2s ease",
                                        }}>
                                        {isDone && "✓"}
                                        {isInProg && <span style={{ color:DS.state.inprogress, fontSize:"8px" }}>◐</span>}
                                      </div>

                                      {/* RF-04: nombre del objetivo — elemento principal */}
                                      <span onClick={() => toggleItem(selectedSt.id, sec.name, idx, itemObj, selectedSt, sec)}
                                        style={{ fontSize:"12.5px",
                                          color: isDone ? DS.text.muted : "#D8D0C8",
                                          textDecoration: isDone ? "line-through" : "none",
                                          lineHeight:"1.4", flex:1,
                                          cursor: isAdmin ? "pointer" : "default",
                                          fontFamily:"'Source Sans 3',sans-serif" }}>
                                        {item}
                                      </span>

                                      {/* RF-02: fecha solo si existe, info secundaria compacta */}
                                      {itemDate && (
                                        <span style={{ fontSize:"9px",
                                          color: isDone ? "#2A1808" : DS.state.inprogress,
                                          fontFamily:"'Source Sans 3',sans-serif",
                                          flexShrink:0, whiteSpace:"nowrap" }}>
                                          {itemDate}
                                        </span>
                                      )}

                                      {/* RF-06: botón expandir — posición y tamaño consistentes */}
                                      <button className="expbtn"
                                        style={{ background:"none", border:"none",
                                          color: isExp ? selectedSt.color : "#2A1808",
                                          fontSize:"9px", cursor:"pointer",
                                          padding:"3px 5px", flexShrink:0,
                                          borderRadius:"4px",
                                          transition:"color 0.12s, background 0.12s" }}
                                        title="Fecha de cumplimiento"
                                        onClick={() => setExpItem(isExp ? null : expK)}>
                                        {isExp ? "▲" : "▼"}
                                      </button>
                                    </div>

                                    {isExp && (
                                      // RF-09: panel de detalle — visible solo al expandir
                                      <div className="fade-in rv-item-detail"
                                        style={{ borderColor:`${selectedSt.color}18` }}>
                                        <span style={{ fontSize:"10px", color:"#5A7A8A",
                                          whiteSpace:"nowrap",
                                          fontFamily:"'Source Sans 3',sans-serif" }}>📅 Fecha:</span>
                                        <input type="date" className="di"
                                          disabled={!isAdmin}
                                          value={itemDate}
                                          onChange={e => setItemDate(selectedSt.id, sec.name, idx, e.target.value, itemObj, selectedSt, sec)} />
                                        {itemDate && isAdmin && (
                                          <button onClick={() => setItemDate(selectedSt.id, sec.name, idx, "", itemObj, selectedSt, sec)}
                                            style={{ background:"none", border:"none",
                                              color:DS.text.muted, fontSize:"10px",
                                              cursor:"pointer", padding:"0 3px" }}>✕</button>
                                        )}
                                        <span style={{ fontSize:"10px", marginLeft:"auto",
                                          color: cfg.dot === "rgba(255,255,255,0.18)" ? DS.text.muted : cfg.dot,
                                          fontFamily:"'Source Sans 3',sans-serif",
                                          fontWeight:"600", whiteSpace:"nowrap" }}>
                                          {cfg.icon} {cfg.label}
                                        </span>
                                      </div>
                                    )}
                                  </div>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            );
          })()}
        </main>
      </div>
      {/* PDF progress indicator */}
      {pdfProgress && (
        <div style={{ position:"fixed", bottom:"70px", right:"18px", zIndex:9998,
          background:DS.bg.cardAlt, border:"1px solid rgba(197,162,83,0.35)",
          borderRadius:"10px", padding:"10px 16px", display:"flex",
          alignItems:"center", gap:"10px", fontSize:"12px",
          color:DS.text.primary, fontFamily:DS.font.body,
          boxShadow:"0 6px 20px rgba(0,0,0,0.5)" }}>
          <span className="rv-sync-dot" aria-hidden="true" />
          {pdfProgress}
        </div>
      )}

      {/* RF-05/RF-08: Toast stack — aria-live for screen readers */}
      <div className="rv-toast-stack" role="status" aria-live="polite" aria-atomic="false">
        {toasts.map(t => (
          <div key={t.id} className={`rv-toast ${t.type}`}>
            <span className="rv-toast-icon" style={{ color:TOAST_COLORS_MAP[t.type] }}>
              {TOAST_ICONS[t.type]}
            </span>
            <span style={{ flex:1, lineHeight:"1.4" }}>{t.message}</span>
            <button className="rv-toast-dismiss" onClick={() => dismissToast(t.id)}>✕</button>
          </div>
        ))}
      </div>

      {/* RF-05 (Fase 10): Vista previa de reporte */}
      {reportPreview && (
        <div className="rv-preview-overlay">
          <div className="rv-preview-toolbar">
            <div style={{ fontFamily:DS.font.display, fontSize:"12px",
              fontWeight:"700", color:DS.text.primary, flex:1 }}>
              {reportPreview.title}
            </div>
            <button className="btn-action" style={{ fontSize:"11px", padding:"5px 12px" }}
              onClick={() => {
                const w = window.open("","_blank");
                w.document.write(reportPreview.html);
                w.document.close();
              }}>
              🖨️ Abrir / Imprimir
            </button>
            <button className="btn-secondary" style={{ fontSize:"11px", padding:"5px 12px" }}
              autoFocus
              aria-label="Cerrar vista previa"
              onClick={() => setReportPreview(null)}>
              ✕ Cerrar
            </button>
          </div>
          <iframe className="rv-preview-frame"
            srcDoc={reportPreview.html}
            title={reportPreview.title}
            style={{ flex:1, width:"100%", maxWidth:"960px", border:"none" }} />
        </div>
      )}

      {/* RF-05: Confirm dialog — solo acciones críticas */}
      {confirmDlg && (
        <div className="rv-confirm-overlay"
          role="alertdialog" aria-modal="true"
          aria-label="Confirmar acción"
          onClick={closeConfirm}>
          <div className="rv-confirm-box" onClick={e => e.stopPropagation()}
            onKeyDown={e => e.key === "Tab" && e.preventDefault()}>
            <div style={{ fontSize:"22px", marginBottom:"10px",
              color: confirmDlg.danger ? DS.state.error : DS.state.info }}>
              {confirmDlg.danger ? "⚠" : "●"}
            </div>
            <div style={{ fontSize:"13.5px", color:DS.text.primary, lineHeight:"1.5",
              fontFamily:"'Source Sans 3',sans-serif", marginBottom:"18px" }}>
              {confirmDlg.message}
            </div>
            <div style={{ display:"flex", gap:"8px", justifyContent:"flex-end" }}>
              <button className="btn-secondary"
                style={{ padding:"7px 14px", fontSize:"12px" }}
                onClick={closeConfirm}>
                Cancelar
              </button>
              <button className={confirmDlg.danger ? "btn-danger" : "btn-primary"}
                autoFocus
                style={{ padding:"7px 14px", fontSize:"12px" }}
                onClick={() => { confirmDlg.onConfirm(); closeConfirm(); }}>
                {confirmDlg.danger ? "Eliminar" : "Confirmar"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
