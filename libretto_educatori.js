const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, PageBreak,
  HeadingLevel, UnderlineType, ImageRun, convertInchesToTwip
} = require("docx");
const fs = require("fs");
const path = require("path");

// ── Palette ──────────────────────────────────────────────────────────────────
const BLU = "5BAFD6", BLU_CHIARO = "D6EAF8", BLU_SCURO = "1A6FA0";
const MARRONE = "7A5230", MARRONE_CHIARO = "F5EFE6";
const VERDE = "5A9B4A", VERDE_CHIARO = "E9F7EF";
const GRIGIO = "F2F3F4", GRIGIO_SCURO = "888888", BIANCO = "FFFFFF";
const TESTO = "2C2C2C";
const GIALLO = "FEF9E7", GIALLO_SCURO = "9A7D0A";
const ARANCIO = "FEF3E2", ARANCIO_SCURO = "B7650A";

// ── Page dims ─────────────────────────────────────────────────────────────────
const W = 11906, H = 16838;
const ML = 1440 + 720, MR = 1440, MT = 1440, MB = 1440;
const BODY_W = W - ML - MR; // usable width

// ── Helpers ───────────────────────────────────────────────────────────────────
const sp = (before, after) => new Paragraph({ spacing: { before, after }, children: [] });

function T(t, bold=false, size=20, color=TESTO, font="Arial", italic=false) {
  return new TextRun({ text: t, bold, size, color, font, italics: italic });
}

function body(text, bold=false, size=21, color=TESTO) {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text, bold, size, color, font: "Arial" })]
  });
}

function bodySm(text, bold=false, color=TESTO) {
  return new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text, bold, size: 20, color, font: "Arial" })]
  });
}

function bodyItalic(text, color=MARRONE) {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text, italics: true, size: 21, color, font: "Arial" })]
  });
}

function heading1(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, color: BLU_SCURO, font: "Arial" })]
  });
}

function heading2(text, color=BLU_SCURO) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, color, font: "Arial" })]
  });
}

function heading3(text, color=TESTO) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22, color, font: "Arial" })]
  });
}

function pb() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Colored box (table trick)
function colorBox(paras, bg=BLU_CHIARO) {
  return new Table({
    width: { size: BODY_W, type: WidthType.DXA },
    columnWidths: [BODY_W],
    borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE} },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: BODY_W, type: WidthType.DXA },
      shading: { fill: bg, type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 160, bottom: 160, left: 180, right: 180 },
      borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
      children: paras
    })]})]
  });
}

// Placeholder box
function placeholder(label) {
  return colorBox([
    new Paragraph({ spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "⏳  " + label, bold: true, size: 22, color: "8B6500", font: "Arial" })] })
  ], "FFF8DC");
}

// C: / T: script lines
function C(text) {
  return new Paragraph({ spacing: { before: 40, after: 40 }, indent: { left: 120 },
    children: [
      new TextRun({ text: "C: ", bold: true, size: 21, color: BLU_SCURO, font: "Arial" }),
      new TextRun({ text, size: 21, color: TESTO, font: "Arial" })
    ]
  });
}
function Tall(text) {
  return new Paragraph({ spacing: { before: 40, after: 40 }, indent: { left: 120 },
    children: [
      new TextRun({ text: "T: ", bold: true, size: 21, color: VERDE, font: "Arial" }),
      new TextRun({ text, bold: true, size: 21, color: TESTO, font: "Arial" })
    ]
  });
}

// Section divider: time + label
function timeBox(time, label, color=BLU_SCURO, bg=BLU_CHIARO) {
  const LEFT = 1600;
  return new Table({
    width: { size: BODY_W, type: WidthType.DXA },
    columnWidths: [LEFT, BODY_W - LEFT],
    borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE} },
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: LEFT, type: WidthType.DXA },
        shading: { fill: color, type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 100, bottom: 100, left: 140, right: 100 },
        borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: time, bold: true, size: 22, color: BIANCO, font: "Arial" })] })]
      }),
      new TableCell({
        width: { size: BODY_W - LEFT, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 100, bottom: 100, left: 160, right: 120 },
        borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 22, color, font: "Arial" })] })]
      })
    ]})]
  });
}

// Day header
function dayHeader(n, icon, theme, date, citation) {
  const COLORS = [BLU_SCURO,"1A6FA0","1B7A3C","8B5A00","5A2D82"];
  const bg = COLORS[n-1] || BLU_SCURO;
  return [
    new Table({
      width: { size: BODY_W, type: WidthType.DXA },
      columnWidths: [BODY_W],
      borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE} },
      rows: [new TableRow({ children: [new TableCell({
        width: { size: BODY_W, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 180, bottom: 180, left: 220, right: 220 },
        borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${icon}  Giorno ${n} — ${theme}`, bold: true, size: 38, color: BIANCO, font: "Arial" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: date, size: 22, color: "DDDDDD", font: "Arial" })] })
        ]
      })]})]
    }),
    colorBox([
      new Paragraph({ spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: "«" + citation + "»", italics: true, size: 21, color: MARRONE, font: "Arial" })]
      }),
      new Paragraph({ spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: "— Pier Giorgio Frassati", size: 19, color: GRIGIO_SCURO, font: "Arial" })]
      })
    ], MARRONE_CHIARO)
  ];
}

// Guida box (educator instructions)
function guidaBox(title, paras, bg=VERDE_CHIARO, titleColor=VERDE) {
  return colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "📋  " + title, bold: true, size: 22, color: titleColor, font: "Arial" })]
    }),
    ...paras
  ], bg);
}

// Note box
function noteBox(paras) {
  return colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "💡  Note per l'educatore", bold: true, size: 21, color: "8B6500", font: "Arial" })]
    }),
    ...paras
  ], GIALLO);
}

// Obiettivo box
function obiettivoBox(text) {
  return colorBox([
    new Paragraph({ spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "🎯  Obiettivo", bold: true, size: 21, color: BLU_SCURO, font: "Arial" })]
    }),
    new Paragraph({ spacing: { before: 0, after: 0 },
      children: [new TextRun({ text, size: 20, color: TESTO, font: "Arial" })]
    })
  ], BLU_CHIARO);
}

// Bullet list item
function bullet(text, indent=240) {
  return new Paragraph({
    spacing: { before: 0, after: 60 },
    indent: { left: indent, hanging: 240 },
    children: [new TextRun({ text: "•  " + text, size: 20, color: TESTO, font: "Arial" })]
  });
}

// Salmo verses  
function salmoVerses(title, ref, verses) {
  const vItems = verses.map(v =>
    v === "" ? sp(40, 0) :
    new Paragraph({ spacing: { before: 0, after: 30 }, indent: { left: 200 },
      children: [new TextRun({ text: v, size: 20, color: TESTO, font: "Arial", italics: true })]
    })
  );
  return colorBox([
    new Paragraph({ spacing: { before: 0, after: 60 },
      children: [
        new TextRun({ text: title + "  ", bold: true, size: 21, color: BLU_SCURO, font: "Arial" }),
        new TextRun({ text: "(" + ref + ")", size: 19, color: GRIGIO_SCURO, font: "Arial" })
      ]
    }),
    ...vItems
  ], GRIGIO);
}

// Vangelo block
function vangeloBlock(title, verses) {
  const items = verses.map(v => new Paragraph({ spacing: { before: 0, after: 50 },
    children: [new TextRun({ text: v, size: 20, color: TESTO, font: "Arial", italics: true })]
  }));
  return colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "📖  " + title, bold: true, size: 21, color: MARRONE, font: "Arial" })]
    }),
    ...items
  ], MARRONE_CHIARO);
}

// Two-col table for game rules
function twoColTable(rows) {
  const CW = Math.floor(BODY_W / 2);
  return new Table({
    width: { size: BODY_W, type: WidthType.DXA },
    columnWidths: [CW, CW],
    borders: { top:{style:BorderStyle.SINGLE, size:4, color:"DDDDDD"}, bottom:{style:BorderStyle.SINGLE, size:4, color:"DDDDDD"}, left:{style:BorderStyle.SINGLE, size:4, color:"DDDDDD"}, right:{style:BorderStyle.SINGLE, size:4, color:"DDDDDD"}, insideH:{style:BorderStyle.SINGLE, size:4, color:"DDDDDD"}, insideV:{style:BorderStyle.SINGLE, size:4, color:"DDDDDD"} },
    rows: rows.map(([l,r]) => new TableRow({ children: [
      new TableCell({ width:{size:CW,type:WidthType.DXA}, margins:{top:100,bottom:100,left:140,right:100},
        children:[new Paragraph({children:[new TextRun({text:l,bold:true,size:20,color:TESTO,font:"Arial"})]})] }),
      new TableCell({ width:{size:CW,type:WidthType.DXA}, margins:{top:100,bottom:100,left:140,right:100},
        children:[new Paragraph({children:[new TextRun({text:r,size:20,color:TESTO,font:"Arial"})]})] })
    ]}))
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// PAGE BUILDERS
// ─────────────────────────────────────────────────────────────────────────────

function pagCopertina() {
  const imgPath = path.join(__dirname, "../../..", "Documents/CLAUDE/campo terza media/materiale guida/COPERTINA.jpg");
  let imgRun = null;
  try {
    const imgData = fs.readFileSync(imgPath);
    imgRun = new ImageRun({ data: imgData, transformation: { width: 400, height: 320 }, type: "jpg" });
  } catch(e) {}

  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 160 },
      children: imgRun ? [imgRun] : [new TextRun({ text: "✝", size: 80, color: BLU_SCURO, font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "Vivere, non vivacchiare!", bold: true, size: 52, color: BLU_SCURO, font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: "Guida Educatori", bold: true, size: 36, color: MARRONE, font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "San Pier Giorgio Frassati", italics: true, size: 26, color: MARRONE, font: "Arial" })] }),
    sp(120, 0),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "Camposcuola Diocesano ACR  ·  Terza Media", size: 22, color: "555555", font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "Montecalvo Irpino  ·  18 – 22 agosto 2026", size: 22, color: "555555", font: "Arial" })] }),
    pb()
  ];
}

function pagIntro() {
  return [
    heading1("Come usare questo libretto"),
    body("Questo libretto è la tua guida operativa per il campo. Ogni giornata è organizzata in ordine cronologico: trovi i contenuti nell'ordine in cui li vivrai con i ragazzi."),
    body("Per ogni momento trovi lo script completo (C: = Celebrante/Educatore, T: = Tutti/Ragazzi), le istruzioni pratiche e le note per gestire situazioni particolari."),
    sp(60, 0),
    colorBox([
      new Paragraph({ spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Legenda riquadri", bold: true, size: 21, color: BLU_SCURO, font: "Arial" })] }),
      bullet("📋  Riquadro verde = istruzioni pratiche per l'educatore"),
      bullet("🎯  Riquadro blu = obiettivo del momento"),
      bullet("💡  Riquadro giallo = note e suggerimenti"),
      bullet("⏳  Riquadro arancio = segnaposto materiale mancante"),
      bullet("C: = parole del Celebrante / Educatore che guida"),
      bullet("T: = risposta di Tutti i ragazzi (in grassetto)"),
    ], BLU_CHIARO),
    pb()
  ];
}

function pagProgramma() {
  const CW1 = 1400, CW2 = 1400, CW3 = BODY_W - CW1 - CW2;
  const hdr = (t, bg=BLU_SCURO) => new TableCell({
    width:{size:CW1,type:WidthType.DXA},
    shading:{fill:bg,type:ShadingType.CLEAR,color:"auto"},
    margins:{top:80,bottom:80,left:120,right:80},
    borders:{top:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},bottom:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
    children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:t,bold:true,size:19,color:BIANCO,font:"Arial"})]})]
  });
  const cell = (t, bg="FFFFFF", bold=false) => new TableCell({
    width:{size:CW1,type:WidthType.DXA},
    shading:{fill:bg,type:ShadingType.CLEAR,color:"auto"},
    margins:{top:60,bottom:60,left:120,right:80},
    borders:{top:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},bottom:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
    children:[new Paragraph({children:[new TextRun({text:t,bold,size:19,color:TESTO,font:"Arial"})]})]
  });

  const giorni = [
    ["18/08  Martedì","⚓ CERCARE","Arrivo · Presentazione · Catechesi · G.Studio G1 · Preghiera sera · Serata Social Retrò"],
    ["19/08  Mercoledì","⛰️ ASCOLTARE","Pregh. mattino · Catechesi · G.Studio G2 (Deserto) · Giochi conoscenza · Preghiera sera · Serata [⏳]"],
    ["20/08  Giovedì","🏳️ RESTARE","Pregh. mattino · Catechesi · G.Studio G3 · Escursione · Preghiera sera [⏳] · Veglia"],
    ["21/08  Venerdì","🤝 PRENDERSI CURA","Pregh. mattino · Catechesi · G.Studio G4 · Attività libretti · Preghiera sera · Caccia al Tesoro"],
    ["22/08  Sabato","→ ANDARE","Pregh. mattino · Catechesi · G.Studio G5 · Messa conclusiva · Partenza"],
  ];

  const bgs = [BLU_CHIARO,"E8F5E9","FEF9E7","FCE4EC","EDE7F6"];
  const themeBgs = [BLU_SCURO,"1B7A3C","9A7D0A","8B1A2E","5A2D82"];

  return [
    heading1("Programma del campo"),
    body("Montecalvo Irpino · 18–22 agosto 2026 · 73 ragazzi, 16 educatori"),
    sp(60, 0),
    new Table({
      width:{size:BODY_W,type:WidthType.DXA},
      columnWidths:[CW1,CW2,CW3],
      borders:{top:{style:BorderStyle.SINGLE,size:6,color:"CCCCCC"},bottom:{style:BorderStyle.SINGLE,size:6,color:"CCCCCC"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},insideH:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},insideV:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"}},
      rows: [
        new TableRow({ children: [hdr("Data"), hdr("Tema"), hdr("Contenuto")] }),
        ...giorni.map(([data, tema, contenuto], i) =>
          new TableRow({ children: [
            new TableCell({width:{size:CW1,type:WidthType.DXA},shading:{fill:bgs[i],type:ShadingType.CLEAR,color:"auto"},margins:{top:80,bottom:80,left:120,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},bottom:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:data,bold:true,size:19,color:TESTO,font:"Arial"})]})]}),
            new TableCell({width:{size:CW2,type:WidthType.DXA},shading:{fill:themeBgs[i],type:ShadingType.CLEAR,color:"auto"},margins:{top:80,bottom:80,left:120,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},bottom:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:tema,bold:true,size:19,color:BIANCO,font:"Arial"})]})]}),
            new TableCell({width:{size:CW3,type:WidthType.DXA},shading:{fill:bgs[i],type:ShadingType.CLEAR,color:"auto"},margins:{top:80,bottom:80,left:120,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},bottom:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:contenuto,size:18,color:TESTO,font:"Arial"})]})]})
          ]})
        )
      ]
    }),
    sp(120, 0),
    heading1("Filo rosso dei 5 giorni"),
    sp(60, 0),
    new Table({
      width:{size:BODY_W,type:WidthType.DXA},
      columnWidths:[800,1300,1600,2400,2200],
      borders:{top:{style:BorderStyle.SINGLE,size:6,color:"CCCCCC"},bottom:{style:BorderStyle.SINGLE,size:6,color:"CCCCCC"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},insideH:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},insideV:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"}},
      rows: [
        new TableRow({ children: [
          ...(["G","Tema","Segno","Citazione Frassati","Vangelo"].map((h,i) =>
            new TableCell({shading:{fill:BLU_SCURO,type:ShadingType.CLEAR,color:"auto"},margins:{top:80,bottom:80,left:100,right:80},borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:h,bold:true,size:18,color:BIANCO,font:"Arial"})]})]})))
        ]}),
        ...[
          ["1","CERCARE","Ancora ⚓",'"La fede è l\'unica ancora di salvezza…"',"Gv 4, 1-42"],
          ["2","ASCOLTARE","Montagna ⛰️",'"Nella solitudine della montagna…"',"Mt 5, 1-12"],
          ["3","RESTARE","Bandiera AC",'"Ogni giorno la comunione mi dà la forza…"',"Lc 24, 13-53"],
          ["4","PRENDERSI CURA","Abbraccio 🤝",'"Preghiera, azione e sacrificio"',"Gv 19, 25-30"],
          ["5","ANDARE","Pagina bianca 📄",'"Verso l\'Alto!"',"Mt 28, 16-20"],
        ].map(([g,t,s,c,v], i) => new TableRow({ children: [
          new TableCell({shading:{fill:i%2===0?BLU_CHIARO:BIANCO,type:ShadingType.CLEAR,color:"auto"},margins:{top:70,bottom:70,left:100,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},bottom:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:g,bold:true,size:19,color:BLU_SCURO,font:"Arial"})]})]}),
          new TableCell({shading:{fill:i%2===0?BLU_CHIARO:BIANCO,type:ShadingType.CLEAR,color:"auto"},margins:{top:70,bottom:70,left:100,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},bottom:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:t,bold:true,size:19,color:TESTO,font:"Arial"})]})]}),
          new TableCell({shading:{fill:i%2===0?BLU_CHIARO:BIANCO,type:ShadingType.CLEAR,color:"auto"},margins:{top:70,bottom:70,left:100,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},bottom:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:s,size:19,color:TESTO,font:"Arial"})]})]}),
          new TableCell({shading:{fill:i%2===0?BLU_CHIARO:BIANCO,type:ShadingType.CLEAR,color:"auto"},margins:{top:70,bottom:70,left:100,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},bottom:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:c,italics:true,size:18,color:MARRONE,font:"Arial"})]})]}),
          new TableCell({shading:{fill:i%2===0?BLU_CHIARO:BIANCO,type:ShadingType.CLEAR,color:"auto"},margins:{top:70,bottom:70,left:100,right:80},borders:{top:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},bottom:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},children:[new Paragraph({children:[new TextRun({text:v,size:18,color:TESTO,font:"Arial"})]})]}),
        ]}))
      ]
    }),
    pb()
  ];
}


// ── GIORNO 1 — CERCARE (18 agosto) ──────────────────────────────────────────
function pagGiorno1() {
  return [
    ...dayHeader(1,"⚓","CERCARE","18 agosto 2026","La fede è l'unica ancora di salvezza… senza di essa che cosa sarebbe la nostra vita?"),
    sp(100, 0),

    // 12:30 Arrivo
    timeBox("12:30", "ARRIVO E ACCOGLIENZA"),
    sp(60, 0),
    guidaBox("Arrivo in struttura", [
      bodySm("Accogliere i ragazzi e accompagnarli nelle stanze. Sistemazione bagagli."),
      bodySm("Ore 12:15 — Presentazione del campo da parte dei responsabili."),
      bodySm("Ore 12:30 — Pranzo."),
    ]),
    sp(100, 0),

    // 15:00 Catechesi
    timeBox("15:00", "CATECHESI — Vangelo del giorno"),
    sp(60, 0),
    vangeloBlock("Gv 4, 1-42 — La Samaritana (testo di riferimento)", [
      "Gesù stanco si siede presso un pozzo a mezzogiorno. Una donna samaritana viene ad attingere acqua. Gesù le chiede da bere. La donna si stupisce: un giudeo parla con lei, emarginata per razza e storia personale.",
      "«Se tu conoscessi il dono di Dio e chi è colui che ti dice: 'Dammi da bere', tu stessa gliene avresti chiesto ed egli ti avrebbe dato acqua viva.»",
      "Gesù rivela alla donna la sua storia, la sua sete più profonda — non d'acqua ma di verità. La donna riconosce il Messia e porta tutto il villaggio da lui.",
      "(Rif. completo: Gv 4, 1-42)"
    ]),
    guidaBox("Nota catechesi", [
      bodySm("La catechesi è condotta dal sacerdote. Il tuo ruolo è ascoltare con i ragazzi e raccogliere eventuali domande per il gruppo studio."),
      bodySm("Tema chiave: Gesù CERCA la Samaritana — non il contrario. Spesso è Dio che fa il primo passo."),
    ]),
    sp(100, 0),

    // 15:30 Gruppi Studio
    timeBox("15:30", "GRUPPI STUDIO — Giorno 1: CERCARE"),
    sp(60, 0),
    obiettivoBox("Aiutare i ragazzi a riconoscere che la vita è una ricerca continua di senso e di Bellezza. Come Pier Giorgio cercava Dio nelle vette e nello studio, stimolare la curiosità verso le cose grandi."),
    sp(80, 0),
    guidaBox("1. Apertura con il segno — L'ancora", [
      bodySm("Mostra il segno dell'ancora (portane una fisica o disegnata su cartellone)."),
      bodySm("Domande di apertura:"),
      bullet("A cosa serve un'ancora?"),
      bullet("Nella tua vita c'è qualcosa che fa da ancora?"),
      bodySm("Collega: l'ancora non blocca la barca, la tiene ferma mentre intorno tutto si muove. Così funziona la fede."),
    ]),
    sp(60, 0),
    guidaBox("2. Attività — I tre pezzi dell'ancora", [
      bodySm("Distribuisci a ogni ragazzo tre pezzi di carta (o le tre parti di un'ancora), ognuno con una domanda:"),
      sp(40,0),
      twoColTable([["Pezzo","Domanda"],["CAMPO","Che cosa cerchi da questo camposcuola?"],["IO","Nelle scelte di quest'anno, cosa hai cercato?"],["DIO","Cosa cerchi in Dio?"]]),
      sp(60,0),
      bodySm("I ragazzi scrivono individualmente in silenzio, poi condividono ciò che vogliono."),
      bodySm("Domanda bonus (se il gruppo è pronto): «Ti sei mai sentito cercato da Dio?»"),
      bodySm("Collega: come Gesù cerca la Samaritana — una persona che non conta nulla per la società di allora — così Lui ci cerca, sempre per primo."),
    ]),
    sp(60, 0),
    guidaBox("3. Conclusione", [
      bodyItalic("\"Cercare è già un atto di fede: non bisogna avere tutto chiaro per mettersi in cammino. Pier Giorgio non aspettava di 'sentirsi pronto' — saliva in montagna, pregava, si metteva in gioco. Il campo è un'occasione per cercare, non per trovare tutte le risposte.\""),
    ]),
    sp(60, 0),
    noteBox([
      bullet("Se il gruppo fatica a rispondere, parti tu con un esempio personale."),
      bullet("La domanda su Dio può generare imbarazzo: normalizza il fatto che non avere risposte è ok."),
    ]),
    sp(100, 0),

    // 19:30 Preghiera della Sera
    timeBox("19:30", "PREGHIERA DELLA SERA — 18 agosto"),
    sp(60, 0),
    guidaBox("Preparazione", [
      bodySm("Raccogliere i ragazzi in uno spazio tranquillo. Luce soffusa se possibile. Preparare un cartellone con una sagoma di ancora su cui i ragazzi attaccheranno i loro pezzi alla fine."),
      bodySm("Tema serale: la sete — come la Samaritana, ognuno porta al pozzo la propria sete di senso."),
    ]),
    sp(60, 0),
    C("Nel nome del Padre, del Figlio e dello Spirito Santo."), Tall("Amen."),
    C("Il Signore, che ci cerca per primo e conosce la nostra sete, sia con tutti voi."), Tall("E con il tuo spirito."),
    sp(60, 0),
    C("Siamo arrivati alla fine di questa prima giornata. Ma cosa abbiamo cercato davvero? Stasera ci mettiamo accanto a una donna che è andata al pozzo per cercare acqua e ha trovato molto di più. Lasciamo che Gesù si sieda accanto a noi e ci chieda: «Dammi da bere»."),
    sp(60, 0),
    salmoVerses("Salmo 41 — Come la cerva anela", "Sal 41",
      ["Come la cerva anela","ai corsi d'acqua,","così l'anima mia anela a te, o Dio.","",
       "L'anima mia ha sete di Dio,","del Dio vivente:","quando verrò e vedrò","il volto di Dio?","",
       "Le lacrime sono il mio pane","giorno e notte,","mentre mi dicono sempre:","\"Dov'è il tuo Dio?\"","",
       "Perché ti rattristi, anima mia, / perché ti agiti in me?","Spera in Dio: ancora potrò lodarlo,","lui, salvezza del mio volto e mio Dio."]
    ),
    sp(60, 0),
    vangeloBlock("Gv 4, 5-15 — La Samaritana al pozzo", [
      "Giunse dunque a una città della Samaria, chiamata Sicar. Gesù dunque, stanco del cammino, stava seduto presso la fonte. Era circa l'ora sesta.",
      "Una donna della Samaria venne ad attingere l'acqua. Gesù le disse: «Dammi da bere».",
      "La donna samaritana allora gli disse: «Come mai tu che sei Giudeo chiedi da bere a me, che sono una donna samaritana?».",
      "Gesù le rispose: «Se tu conoscessi il dono di Dio e chi è che ti dice: \"Dammi da bere\", tu stessa gliene avresti chiesto, ed egli ti avrebbe dato dell'acqua viva».",
      "La donna gli disse: «Signore, tu non hai nulla per attingere, e il pozzo è profondo; da dove avresti quest'acqua viva?»",
      "Gesù le rispose: «Chiunque beve di quest'acqua avrà di nuovo sete; ma chi beve dell'acqua che io gli darò non avrà mai più sete; anzi, l'acqua che io gli darò diventerà in lui una fonte d'acqua che scaturisce in vita eterna».",
      "La donna gli disse: «Signore, dammi di quest'acqua, affinché io non abbia più sete.»"
    ]),
    sp(60, 0),
    C("Portiamo al Signore la nostra sete e le nostre ricerche e diciamo insieme:"), Tall("Signore, dacci dell'acqua viva."),
    bodySm("Per quando cerchiamo la felicità nelle cose che passano e ci dimentichiamo di Te, preghiamo."), Tall("Signore, dacci dell'acqua viva."),
    bodySm("Per chi si sente \"arido\" e svuotato, perché trovi in questo campo una sorgente di amicizia, preghiamo."), Tall("Signore, dacci dell'acqua viva."),
    bodySm("Per la nostra ricerca di futuro e di sogni: guidaci Tu sulla strada giusta, preghiamo."), Tall("Signore, dacci dell'acqua viva."),
    sp(60, 0),
    C("Padre Nostro che sei nei cieli..."), Tall("Padre nostro che sei nei cieli, / sia santificato il tuo nome, / venga il tuo regno, / sia fatta la tua volontà / come in cielo così in terra. / Dacci oggi il nostro pane quotidiano, / rimetti a noi i nostri debiti / come noi li rimettiamo ai nostri debitori, / e non abbandonarci alla tentazione, / ma liberaci dal male. Amen."),
    sp(60, 0),
    colorBox([
      bodySm("Segno: ogni ragazzo attacca i propri pezzi dell'ancora sul cartellone comune. L'ancora cresce insieme.", true, BLU_SCURO),
    ], BLU_CHIARO),
    sp(60, 0),
    C("Il Signore sia il vostro pozzo di acqua fresca nei momenti di stanchezza."), Tall("Amen."),
    C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), Tall("Amen."),
    sp(100, 0),

    // 21:30 Serata
    timeBox("21:30", "SERATA — Social Retrò"),
    sp(60, 0),
    guidaBox("Serata Social Retrò — Istruzioni generali", [
      bodySm("La serata si svolge a squadre. Preparare in anticipo video-input con tematiche baresi. 5 giochi a rotazione."),
      bodySm("Materiali necessari: proiettore o schermo, video preregistrati, cartelloni per punteggi, fischietto."),
      bodySm("Durata stimata: 90 minuti."),
    ]),
    sp(60, 0),
    heading3("Gioco 1 — Blind Test Musicale"),
    bodySm("Squadre sentono spezzoni di canzoni popolari baresi/pugliesi anni '80-'90. Prima squadra a riconoscere titolo e artista fa punto."),
    sp(40, 0),
    heading3("Gioco 2 — Chi è? (Personaggi baresi)"),
    bodySm("Mostrare foto o descrizioni di personaggi famosi di Bisceglie/Trani/Puglia. Le squadre indovinano il nome."),
    sp(40, 0),
    heading3("Gioco 3 — Dialetto Barese"),
    bodySm("Tradurre parole o frasi in dialetto barese. 1 punto per risposta corretta. Ridere è consentito."),
    sp(40, 0),
    heading3("Gioco 4 — Il Prezzo è Giusto — Versione Pugliese"),
    bodySm("Indovinare il prezzo di prodotti tipici pugliesi (taralli, focaccia barese, orecchiette, ecc.). Vince chi si avvicina di più."),
    sp(40, 0),
    heading3("Gioco 5 — Mimo Barese"),
    bodySm("Un membro della squadra mima un gesto tipico barese o una scena di vita locale. La squadra indovina."),
    sp(60, 0),
    noteBox([
      bullet("Dividere i ragazzi in squadre miste (non per parrocchia) per favorire la conoscenza."),
      bullet("Un educatore funge da arbitro, uno tiene il punteggio."),
      bullet("Terminare con classifica, premiazione simbolica e saluto della buonanotte."),
    ]),
    pb()
  ];
}

// ── GIORNO 2 — ASCOLTARE (19 agosto) ────────────────────────────────────────
function pagGiorno2() {
  return [
    ...dayHeader(2,"⛰️","ASCOLTARE","19 agosto 2026","Nella solitudine della montagna l'anima si purifica e si eleva più facilmente a Dio."),
    sp(100, 0),

    timeBox("8:30", "PREGHIERA DEL MATTINO — 19 agosto"),
    sp(60, 0),
    guidaBox("Preparazione", [
      bodySm("Raduno in cappella o in uno spazio raccolto. Atmosfera silenziosa — invitare i ragazzi a entrare in silenzio."),
      bodySm("Tema: l'ascolto. Iniziare con 30 secondi di silenzio prima della preghiera."),
    ]),
    sp(60, 0),
    C("Nel nome del Padre, del Figlio e dello Spirito Santo."), Tall("Amen."),
    C("Il Signore, che ci parla nel silenzio dell'aurora e nel rumore dei nostri passi, sia con tutti voi."), Tall("E con il tuo spirito."),
    C("Buongiorno! Spesso pensiamo che ascoltare significhi solo 'sentire' dei suoni. Ma l'ascolto vero è un'arte: è fare spazio a qualcuno dentro di noi. Iniziamo la giornata chiedendo il dono di un ascolto attento."),
    sp(60, 0),
    salmoVerses("Salmo 16 — Preservami, o Dio", "Sal 16",
      ["Preservami, o Dio, perché io confido in te.","",
       "Io ho detto all'Eterno: «Tu sei il mio Signore;","io non ho altro bene all'infuori di te».","",
       "L'Eterno è la mia parte di eredità e il mio calice;","tu mantieni quel che m'è toccato in sorte.","La sorte è caduta per me in luoghi dilettevoli;","una bella eredità mi è pur toccata!","",
       "Io benedirò l'Eterno che mi consiglia;","anche di notte le mie reni mi ammaestrano.","Io ho sempre il Signore davanti a me;","poiché egli è alla mia destra, non sarò smosso.","",
       "Tu mi mostrerai il sentiero della vita;","vi sono gioie a sazietà alla tua presenza;","vi sono delizie perpetue alla tua destra."]
    ),
    sp(60, 0),
    colorBox([
      heading3("Preghiera Corale", BLU_SCURO),
      new Paragraph({ spacing:{before:0,after:60}, children:[new TextRun({text:"Signore Gesù,",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({ spacing:{before:0,after:60}, children:[new TextRun({text:"mentre il mondo si sveglia e si riempie di rumori, noi ti chiediamo un momento di ascolto vero.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({ spacing:{before:0,after:60}, children:[new TextRun({text:"Aiutaci oggi ad ascoltare davvero: non solo le parole, ma anche i silenzi dei nostri amici.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({ spacing:{before:0,after:60}, children:[new TextRun({text:"Liberaci dalla fretta di rispondere, donaci la pazienza di capire.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({ spacing:{before:0,after:0}, children:[new TextRun({text:"Che il nostro cuore sia come una terra buona, pronta a ricevere la tua Parola. Amen.",size:20,color:TESTO,font:"Arial",italics:true})]})
    ], GRIGIO),
    sp(60, 0),
    C("Il Signore apra il vostro cuore all'ascolto e vi renda capaci di ascoltare gli altri."), Tall("Amen."),
    C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), Tall("Amen."),
    sp(100, 0),

    timeBox("10:00", "CATECHESI — Vangelo del giorno"),
    sp(60, 0),
    vangeloBlock("Mt 5, 1-12 — Le Beatitudini (testo di riferimento)", [
      "«Beati i poveri in spirito, perché di essi è il regno dei cieli. Beati gli afflitti, perché saranno consolati.",
      "Beati i miti, perché erediteranno la terra. Beati quelli che hanno fame e sete della giustizia, perché saranno saziati.",
      "Beati i misericordiosi, perché troveranno misericordia. Beati i puri di cuore, perché vedranno Dio.",
      "Beati gli operatori di pace, perché saranno chiamati figli di Dio.",
      "Beati i perseguitati per causa della giustizia, perché di essi è il regno dei cieli.»"
    ]),
    sp(100, 0),

    timeBox("10:30", "GRUPPI STUDIO — Giorno 2: ASCOLTARE (Deserto)"),
    sp(60, 0),
    obiettivoBox("Educare al silenzio e all'ascolto: della Parola di Dio, della propria coscienza, degli altri. Passare dal 'sentire' distratto all'ascolto profondo."),
    sp(80, 0),
    guidaBox("1. Introduzione (5 min)", [
      bodySm("Cita Frassati: sulla montagna si è soli, lontani dal rumore. Oggi facciamo un po' di 'montagna' anche noi."),
      bodySm("Domanda iniziale: «Quanto spazio ha il silenzio nella tua giornata? Quando sei in silenzio, cosa senti?»"),
    ]),
    sp(60, 0),
    guidaBox("2. Entrata nel silenzio (2 min)", [
      bodySm("Prima di disperdersi, far fermare tutti. Chiudere gli occhi, due respiri profondi. Poi leggi lentamente:"),
      bodyItalic("\"Chiudi gli occhi. Dove sei, in questo istante? Cosa senti — i suoni, il corpo, il respiro? Nella tua giornata tipica, quante ore passi davvero in silenzio? Ricordi un momento recente in cui ti sei davvero fermato ad ascoltare qualcuno?\""),
    ]),
    sp(60, 0),
    guidaBox("3. Deserto guidato (25 min)", [
      bodySm("I ragazzi si allontanano in solitaria. Ogni ragazzo ha il foglio con le domande (vedi sotto). Riportali insieme dopo 20-25 min."),
      sp(40,0),
      heading3("Domande per la riflessione (consegnare su foglio)"),
      bullet("Gesù dice 'Beati i poveri in spirito'. Chi sono, secondo te? Ti ci riconosci?"),
      bullet("'Beati i miti': la mitezza è apprezzata nel mondo che conosci — a scuola, sui social? Essere miti ti sembra coraggioso o debolezza?"),
      bullet("'Beati i misericordiosi': c'è qualcuno nella tua vita a cui fai fatica a perdonare? E qualcuno che ti ha perdonato qualcosa di grosso?"),
      bullet("'Beati i puri di cuore, perché vedranno Dio': cosa vuol dire per te avere un cuore puro?"),
      bullet("'Beati gli operatori di pace': sei mai stato tu quello che ha fatto da ponte, che ha calmato una lite?"),
      bullet("Quale beatitudine ti sembra più difficile da vivere? Quale invece ti sembra più vicina a come sei già?"),
      bullet("Se Gesù ti stesse parlando direttamente — ti sta descrivendo, ti sta chiamando, o ti sta sfidando?"),
      sp(60,0),
      heading3("Il tassello del mosaico"),
      bodySm("Ogni ragazzo realizza un tassello: pezzo di cartoncino decorato con la parola/frase che lo ha colpito + sentimento provato. I tasselli si incollano sulla sagoma della montagna del gruppo."),
    ]),
    sp(60, 0),
    guidaBox("4. Dialogo sulla capacità di ascolto (15 min)", [
      bullet("Chi ascolti nella tua vita? Chi influenza di più le tue scelte?"),
      bullet("Quando ascolti gli altri, dai peso solo alle cose positive o anche a quelle negative?"),
      bullet("Quanto conta il pensiero degli altri rispetto al tuo, nel momento in cui devi decidere?"),
      bullet("Ascolti il tuo cuore — che parole emergono in questo momento?"),
      sp(40,0),
      bodySm("Conclusione: saper ascoltare non è passività. È un atto coraggioso — perché ci espone a sentire cose che non vorremmo."),
    ]),
    sp(60, 0),
    guidaBox("5. Conclusione — Incollare i tasselli", [
      bodySm("Ogni ragazzo incolla il proprio tassello sulla montagna del gruppo. La montagna cresce insieme."),
    ]),
    sp(60, 0),
    noteBox([
      bullet("I tasselli non devono essere 'belli' — incoraggia l'espressione autentica."),
      bullet("Se qualcuno non vuole condividere, rispetta i suoi tempi."),
    ]),
    sp(100, 0),

    timeBox("15:30", "POMERIGGIO — Giochi di conoscenza"),
    sp(60, 0),
    obiettivoBox("Favorire la conoscenza reciproca tra ragazzi di parrocchie diverse attraverso il gioco. Creare un clima di gruppo positivo."),
    sp(60, 0),
    guidaBox("Gioco 1 — Bingo Campisti (25 min)", [
      bodySm("Ogni ragazzo riceve una griglia 4×4 con caratteristiche (es. 'ha i capelli ricci', 'suona uno strumento', 'è di Trani', ecc.)."),
      bodySm("I ragazzi girano per la stanza e trovano persone che corrispondono alle caratteristiche, facendosi firmare la casella."),
      bodySm("Chi completa una riga grida 'BINGO!'. Si può continuare fino a cartella piena."),
    ]),
    sp(60, 0),
    guidaBox("Gioco 2 — Palla Avvelenata (20 min)", [
      bodySm("Due squadre divise da una linea centrale. Una palla viene lanciata nell'area avversaria. Chi viene colpito dalla palla si 'avvelena' e deve uscire."),
      bodySm("Vince la squadra che elimina tutti gli avversari. Variante: i 'colpiti' possono tornare in gioco se un compagno afferra la palla al volo."),
    ]),
    sp(60, 0),
    guidaBox("Gioco 3 — Indovina Chi (20 min)", [
      bodySm("Un ragazzo ha un post-it sulla fronte con il nome di un personaggio (storico, biblico, personaggio noto). Può fare solo domande con risposta sì/no."),
      bodySm("A turno, ogni ragazzo pone domande al gruppo. Vince chi indovina con meno domande."),
    ]),
    sp(60, 0),
    guidaBox("Bonus — La Ghigliottina (se rimane tempo)", [
      bodySm("Un educatore dà 5 parole. I ragazzi devono trovare la parola che le collega tutte e cinque (come nel programma televisivo)."),
      bodySm("Esempio: MARE · VELA · PORTO · PESCATORE · ONDE → risposta: BARCA"),
    ]),
    sp(100, 0),

    timeBox("19:30", "PREGHIERA DELLA SERA — 19 agosto"),
    sp(60, 0),
    guidaBox("Preparazione", [
      bodySm("Tema: il Signore è la mia luce e salvezza. Stasera si fa memoria di questa certezza dopo un giorno di ascolto e gioco."),
      bodySm("Segno: la storia 'L'eco della montagna' narrata dall'educatore."),
    ]),
    sp(60, 0),
    C("Nel nome del Padre, del Figlio e dello Spirito Santo."), Tall("Amen."),
    C("Il Signore, che parla al cuore di chi lo ascolta con fiducia, sia con tutti voi."), Tall("E con il tuo spirito."),
    sp(60, 0),
    salmoVerses("Salmo 27 — Il Signore è la mia luce", "Sal 27",
      ["Il Signore è la mia luce e la mia salvezza; di chi temerò?","Il Signore è il baluardo della mia vita; di chi mi spaventerò?","",
       "Una cosa ho chiesto al Signore, e quella ricerco:","abitare nella casa del Signore tutti i giorni della mia vita,","per contemplare la bellezza del Signore.","",
       "O Signore, ascolta la mia voce quando t'invoco;","abbi pietà di me, e rispondimi.","Il mio cuore mi dice di cercare il tuo volto;","o Signore, il tuo volto io cerco.","",
       "Spera nel Signore! Sii forte,","il tuo cuore si rinfranchi;","sì, spera nel Signore!"]
    ),
    sp(60, 0),
    vangeloBlock("Mt 5, 1-12 — Le Beatitudini", [
      "«Beati i poveri in spirito, perché di essi è il regno dei cieli. Beati gli afflitti, perché saranno consolati.",
      "Beati i miti, perché erediteranno la terra. Beati quelli che hanno fame e sete della giustizia, perché saranno saziati.",
      "Beati i misericordiosi, perché troveranno misericordia. Beati i puri di cuore, perché vedranno Dio.",
      "Beati gli operatori di pace, perché saranno chiamati figli di Dio.»"
    ]),
    sp(60, 0),
    colorBox([
      heading3("La storia dell'eco della montagna", MARRONE),
      bodyItalic("Un bambino, arrabbiatissimo, corre su una montagna e urla: 'Ti odio!' La montagna risponde: 'Ti odio!' Il bambino grida: 'Sei stupido!' La montagna risponde: 'Sei stupido!' Spaventato, il bambino corre dalla madre. Lei gli dice: 'Vai e di' alla montagna: ti voglio bene.' Il bambino grida: 'Ti voglio bene!' La montagna risponde: 'Ti voglio bene!' La madre spiega: 'La vita è un eco — quello che dai ti ritorna. Se dai amore, ricevi amore.'"),
      sp(40,0),
      bodySm("Cosa rimanda la tua vita come un eco? Cosa stai dando agli altri?"),
    ], MARRONE_CHIARO),
    sp(60, 0),
    colorBox([
      heading3("Gesto degli educatori", VERDE),
      bodySm("Gli educatori fanno un gesto simbolico verso i ragazzi (es. un abbraccio di gruppo, una parola personale a ciascuno, un segnale di incoraggiamento concordato). Il gesto dice: 'Ti stiamo ascoltando davvero.'"),
    ], VERDE_CHIARO),
    sp(60, 0),
    C("Chiediamo al Signore la capacità di ascoltare queste parole non con le orecchie, ma con il cuore. Preghiamo:"),
    Tall("Signore, apri le nostre orecchie."),
    C("Il Signore benedica i vostri cuori perché siano pronti all'ascolto."), Tall("Amen."),
    C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), Tall("Amen."),
    sp(100, 0),

    timeBox("21:30", "SERATA — 19 agosto"),
    sp(60, 0),
    placeholder("SERATA 19/08 — Materiale non pervenuto dalla commissione serate. Segnaposto da aggiornare quando il materiale sarà disponibile."),
    pb()
  ];
}


// ── GIORNO 3 — RESTARE (20 agosto) ──────────────────────────────────────────
function pagGiorno3() {
  return [
    ...dayHeader(3,"🏳️","RESTARE","20 agosto 2026","Ogni giorno la comunione mi dà la forza di sopportare i sacrifici."),
    sp(100, 0),

    timeBox("8:30", "PREGHIERA DEL MATTINO — 20 agosto"),
    sp(60, 0),
    guidaBox("Preparazione", [
      bodySm("Tema: restare — fedeltà, costanza, appartenenza. Questo è il giorno più denso: la sera c'è la Veglia."),
      bodySm("Creare un'atmosfera di raccoglimento. Richiamare brevemente il tema della giornata prima di iniziare."),
    ]),
    sp(60, 0),
    C("Nel nome del Padre, del Figlio e dello Spirito Santo."), Tall("Amen."),
    C("Il Signore, che rimane fedele per sempre e ci invita a dimorare nel suo amore, sia con tutti voi."), Tall("E con il tuo spirito."),
    sp(60, 0),
    salmoVerses("Salmo 90 — Rifugio nell'Altissimo", "Sal 90",
      ["Chi abita al riparo dell'Altissimo","passerà la notte all'ombra dell'Onnipotente.","Io dico al Signore: «Mio rifugio e mia fortezza,","mio Dio in cui confido».","",
       "Egli ti libererà dal laccio del cacciatore, / dalla peste che distrugge.","Ti coprirà con le sue penne,","sotto le sue ali troverai rifugio.","",
       "Non temerai il terrore della notte","né la freccia che vola di giorno,","la peste che vaga nelle tenebre.","",
       "«Lo libererò, perché a me si è legato,","lo porrò al sicuro, perché ha conosciuto il mio nome.","Mi invocherà e io gli risponderò;","sarò con lui nella sventura.»"]
    ),
    sp(60, 0),
    colorBox([
      heading3("Preghiera Corale", BLU_SCURO),
      new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:"Signore Gesù, spesso siamo cercatori di novità, sempre pronti a correre verso qualcosa di diverso.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:"Stamattina ti chiediamo il coraggio di restare. Aiutaci a restare nelle fatiche di oggi senza scappare.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:"Insegnaci a dimorare nel tuo amore, come i tralci restano uniti alla vite per dare frutto. Amen.",size:20,color:TESTO,font:"Arial",italics:true})]})
    ], GRIGIO),
    sp(60, 0),
    C("Il Signore vi dia la forza di restare saldi nel bene. Vi custodisca nell'unità."), Tall("Amen."),
    C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), Tall("Amen."),
    sp(100, 0),

    timeBox("10:00", "CATECHESI — Vangelo del giorno"),
    sp(60, 0),
    vangeloBlock("Lc 24, 13-35 — I discepoli di Emmaus (testo di riferimento)", [
      "Due discepoli camminano verso Emmaus, delusi, dopo la morte di Gesù. Un viandante si unisce a loro, ma non lo riconoscono.",
      "Il viandante spiega le Scritture. La sera i discepoli lo invitano: «Resta con noi, perché si fa sera.»",
      "Mentre spezza il pane, i discepoli lo riconoscono — ed Egli sparisce. Corrono a Gerusalemme: «Ardeva il nostro cuore mentre ci parlava lungo la via.»",
      "(Rif. completo: Lc 24, 13-53)"
    ]),
    sp(100, 0),

    timeBox("10:30", "GRUPPI STUDIO — Giorno 3: RESTARE"),
    sp(60, 0),
    obiettivoBox("Comprendere l'importanza della fedeltà e della sosta. 'Restare' in Cristo attraverso l'Eucaristia e la vita associativa. Creare un ponte verso la Veglia serale."),
    sp(80, 0),
    guidaBox("1. Canzone — Resta con me (Bambole di Pezza)", [
      bodySm("Ascoltate insieme la canzone (preparare l'audio in anticipo). Poi chiedere:"),
      bullet("Cosa ti ha colpito del testo? C'è una frase che ti è rimasta?"),
    ]),
    sp(60, 0),
    guidaBox("2. Brainstorming — Cosa vuol dire 'restare'?", [
      bodySm("Su un foglio grande o lavagna, raccogliere tutte le risposte senza giudicare. Poi orientare:"),
      bullet("Ci sono situazioni nella tua vita in cui vuoi restare? (amici, casa, gruppo, squadra…)"),
      bullet("Quando invece vuoi che gli altri restino con te?"),
      bullet("Chi resta con te, anche quando è difficile?"),
    ]),
    sp(60, 0),
    guidaBox("3. Collegamento con Emmaus", [
      bodySm("Racconta l'episodio: due persone deluse che non riconoscono Gesù accanto a loro, finché lui 'resta' a cena con loro."),
      bodySm("Chiedi: «Quando è stato Dio a 'restare' con te, anche se non te ne accorgevi?»"),
    ]),
    sp(60, 0),
    guidaBox("4. Verso la Veglia", [
      bodySm("Prepara i ragazzi alla veglia serale. Di' loro: «Questa sera, nella veglia, faremo l'esperienza di chi resta con noi per eccellenza: Gesù nell'Eucaristia.»"),
      bodySm("Chiedi a ciascuno di portarsi UNA PAROLA: cosa vuol dire per me 'restare'?"),
    ]),
    sp(60, 0),
    noteBox([
      bullet("Il tema di oggi è profondamente legato alla veglia: crea un ponte esplicito tra il gruppo studio e il momento liturgico serale."),
      bullet("'Restare' può toccare corde sensibili (separazioni familiari, amici persi). Stai attento e lascia spazio."),
      bullet("La bandiera AC può diventare un gancio per parlare di appartenenza alla comunità anche fuori dal campo."),
    ]),
    sp(100, 0),

    timeBox("15:30", "POMERIGGIO — Escursione"),
    sp(60, 0),
    guidaBox("Escursione guidata dalla Pro Loco di Montecalvo Irpino", [
      bodySm("L'escursione è organizzata e guidata dalla Pro Loco di Montecalvo Irpino. Ruolo degli educatori: accompagnare i ragazzi, mantenere il gruppo unito, garantire sicurezza."),
      bodySm("Materiali da portare: acqua, scarpe adatte, cappellino, eventuali medicinali personali."),
      bodySm("Durata stimata: 2-2.5 ore."),
    ]),
    noteBox([
      bullet("Cogliere l'occasione del cammino nella natura per richiamare il tema del giorno: restare, camminare insieme, non lasciare indietro nessuno."),
      bullet("Dopo l'escursione: docce e tempo libero prima della cena."),
    ]),
    sp(100, 0),

    timeBox("19:30", "PREGHIERA DELLA SERA — 20 agosto"),
    sp(60, 0),
    placeholder("PREGHIERA DELLA SERA 20/08 — La commissione liturgia indica 'Liturgia delle Ore CEI' per questo momento. Il testo specifico non è stato fornito. Sostituire con il testo quando disponibile."),
    sp(100, 0),

    timeBox("21:30", "VEGLIA DI PREGHIERA — 20 agosto"),
    sp(60, 0),
    guidaBox("Preparazione della Veglia", [
      bodySm("Questo è il momento liturgico più importante del campo. Preparare lo spazio con cura: luci soffuse o candele, spazio per sedersi in cerchio attorno all'ostensorio o al tabernacolo aperto."),
      bodySm("Titolo della Veglia: «Beato chi mangia di Te e vive per Te!»"),
      bodySm("Struttura: 3 fasi (durata totale circa 75 minuti)."),
    ]),
    sp(60, 0),
    colorBox([
      new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:"1ª Fase — Restare nell'AC (15 min)",bold:true,size:22,color:BLU_SCURO,font:"Arial"})]}),
      bodySm("Introduzione al tema dell'Eucaristia come luogo in cui Gesù 'resta'. Collegamento con Emmaus e con la giornata vissuta."),
      bodySm("L'educatore o il sacerdote introduce con poche parole. Poi momento di silenzio."),
    ], BLU_CHIARO),
    sp(60, 0),
    colorBox([
      new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:"2ª Fase — Restare in ascolto (30 min)",bold:true,size:22,color:VERDE,font:"Arial"})]}),
      salmoVerses("Salmo 148 — Lodate il Signore dai cieli", "Sal 148",
        ["Alleluia. Lodate il Signore dai cieli, lodatelo nell'alto dei cieli.","",
         "Lodatelo, voi tutti, suoi angeli, lodatelo, voi tutte, sue schiere.","Lodatelo, sole e luna, lodatelo, voi tutte, stelle fulgenti.","",
         "Lodino tutti il nome del Signore, / perché egli disse e furono creati.","Li ha stabiliti per sempre nei secoli;","",
         "Lodate il Signore dalla terra, mostri marini e voi tutti abissi,","fuoco e grandine, neve e nebbia, vento impetuoso che esegui i suoi ordini;","",
         "I re della terra e i popoli tutti, / i governanti e i giudici della terra,","i giovani e le fanciulle,","gli anziani insieme ai bambini","lodino il nome del Signore.","",
         "Egli ha sollevato la potenza del suo popolo. / È canto di lode per tutti i suoi fedeli."]
      ),
      sp(60,0),
      vangeloBlock("Gv 6, 53-57 — Il pane della vita", [
        "«In verità, in verità io vi dico: se non mangiate la carne del Figlio dell'uomo e non bevete il suo sangue, non avete in voi la vita.",
        "Chi mangia la mia carne e beve il mio sangue ha la vita eterna e io lo risusciterò nell'ultimo giorno.",
        "La mia carne è vero cibo e il mio sangue vera bevanda.",
        "Chi mangia la mia carne e beve il mio sangue rimane in me e io in lui.",
        "Come il Padre, che ha la vita, ha mandato me e io vivo per il Padre, così anche colui che mangia me vivrà per me.»"
      ]),
    ], VERDE_CHIARO),
    sp(60, 0),
    colorBox([
      new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:"3ª Fase — Rendere Grazie (30 min)",bold:true,size:22,color:MARRONE,font:"Arial"})]}),
      C("Con il cuore colmo di gioia, affidiamo a Dio chi ci insegna a rimanere nella Fede e nella Vita piena."),
      bodySm("San Pier Giorgio Frassati era un giovane come noi. Diceva sempre: perché staccarsi dalla vita?"),
      C("Preghiamo per le nostre famiglie, per i nostri genitori. Grazie a loro rimaniamo nelle nostre case e nei nostri quartieri."), Tall("Ti preghiamo, Signore."),
      C("Vogliamo ringraziarTi, o Padre, per il dono dell'amicizia. Permettici di restare nei nostri legami buoni."), Tall("Ti preghiamo, Signore."),
      C("Signore, ti benediciamo per le comunità parrocchiali di cui facciamo parte. Custodisci le nostre parrocchie."), Tall("Ti preghiamo, Signore."),
      sp(60,0),
      C("Benedizione finale."), Tall("Amen."),
    ], MARRONE_CHIARO),
    sp(60, 0),
    noteBox([
      bullet("Il silenzio è essenziale in questa veglia: non riempire ogni spazio con parole."),
      bullet("Se c'è adorazione eucaristica, prepararla con il sacerdote in anticipo."),
      bullet("Dopo la veglia: buonanotte breve, invitare al silenzio nelle stanze."),
    ]),
    pb()
  ];
}

// ── GIORNO 4 — PRENDERSI CURA (21 agosto) ───────────────────────────────────
function pagGiorno4() {
  return [
    ...dayHeader(4,"🤝","PRENDERSI CURA","21 agosto 2026","Preghiera, azione e sacrificio."),
    sp(100, 0),

    timeBox("8:30", "PREGHIERA DEL MATTINO — 21 agosto"),
    sp(60, 0),
    C("Nel nome del Padre, del Figlio e dello Spirito Santo."), Tall("Amen."),
    C("Il Signore, che si prende cura di ogni sua creatura con amore infinito, sia con tutti voi."), Tall("E con il tuo spirito."),
    sp(60, 0),
    salmoVerses("Salmo 22 — Il Signore è il mio pastore", "Sal 22",
      ["Il Signore è il mio pastore:","non manco di nulla;","su pascoli erbosi mi fa riposare,","ad acque tranquille mi conduce.","",
       "Mi rinfranca, mi guida per il giusto cammino, / per amore del suo nome.","Se dovessi camminare in una valle oscura,","non temerei alcun male, perché tu sei con me.","",
       "Il tuo bastone e il tuo vincastro / mi danno sicurezza.","Davanti a me tu prepari una mensa / sotto gli occhi dei miei nemici.","",
       "Felicità e grazia mi saranno compagne","tutti i giorni della mia vita,","e abiterò nella casa del Signore per lunghissimi anni."]
    ),
    sp(60, 0),
    colorBox([
      heading3("Preghiera Corale", BLU_SCURO),
      new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:"Signore Gesù, grazie perché Tu ti prendi cura di noi ogni istante.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:"Oggi ti chiediamo di prestarci il tuo sguardo. Insegnaci a prenderci cura di chi ci sta vicino, senza aspettare di essere ricambiati.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:"Aiutaci a non sprecare le occasioni di bene e a trattare ogni persona con delicatezza. Amen.",size:20,color:TESTO,font:"Arial",italics:true})]})
    ], GRIGIO),
    sp(60, 0),
    C("Il Signore vi renda capaci di gesti di cura gratuiti e gioiosi."), Tall("Amen."),
    C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), Tall("Amen."),
    sp(100, 0),

    timeBox("10:00", "CATECHESI — Vangelo del giorno"),
    sp(60, 0),
    vangeloBlock("Gv 19, 25-30 — Gesù affida Maria al discepolo amato", [
      "Stavano presso la croce di Gesù sua madre, la sorella di sua madre, Maria di Cleofa e Maria di Màgdala.",
      "Gesù allora, vedendo la madre e accanto a lei il discepolo che egli amava, disse alla madre: «Donna, ecco tuo figlio!»",
      "Poi disse al discepolo: «Ecco tua madre!» E da quell'ora il discepolo l'accolse con sé.",
      "Persino nel momento più buio, sulla croce, Gesù si prende cura."
    ]),
    sp(100, 0),

    timeBox("10:30", "GRUPPI STUDIO — Giorno 4: PRENDERSI CURA"),
    sp(60, 0),
    obiettivoBox("Imparare la carità concreta sul modello di Pier Giorgio Frassati, che si prendeva cura dei poveri con gratuità totale. Scoprire che prendersi cura non è solo un gesto — è uno stile di vita."),
    sp(80, 0),
    guidaBox("Prima parte — Presentazione di sé (20 min)", [
      bodySm("Ogni ragazzo si presenta al gruppo attraverso quattro elementi scritti su un foglio:"),
      sp(40,0),
      twoColTable([
        ["Elemento","Cosa scrivere"],
        ["Passioni","Cosa ti piace fare, cosa ti appassiona"],
        ["Talenti","In cosa sei bravo (anche cose piccole)"],
        ["Mancanze ⚠️","Una cosa in cui fai fatica o di cui hai bisogno"],
        ["Valori","Cosa è importante per te nella vita"],
      ]),
      sp(60,0),
      bodySm("Ogni ragazzo condivide con il gruppo. INCORAGGIA LA SINCERITÀ sulle mancanze — è la parte più preziosa."),
    ]),
    sp(60, 0),
    guidaBox("Seconda parte — Le coppie di cura (25 min)", [
      bodySm("Dopo le presentazioni, chiedi a tutto il gruppo:"),
      bodyItalic("«Sulla base di quello che abbiamo sentito, chi sceglieresti di curarti oggi?»"),
      bodySm("I ragazzi si scelgono a coppie (o tris). Una volta formate, le coppie discutono:"),
      bullet("Cosa vuol dire secondo te 'prendersi cura' di qualcuno?"),
      bullet("Di chi mi prendo cura nella mia vita?"),
      bullet("C'è già qualcuno che si prende cura delle mie mancanze?"),
      bullet("Viste le vostre descrizioni, come potete prendervi cura vicendevolmente? (→ da condividere col gruppo)"),
    ]),
    sp(60, 0),
    guidaBox("Conclusione dell'educatore", [
      bodyItalic("\"Per noi cristiani, prendersi cura non è un obbligo né uno scambio. È un dono gratuito. Pier Giorgio non si prendeva cura dei poveri perché si aspettava qualcosa in cambio — lo faceva per amore. Il segno di oggi è l'abbraccio: un gesto che non costa nulla e che cambia tutto.\""),
    ]),
    sp(60, 0),
    noteBox([
      bullet("La parte delle 'mancanze' è delicata: dai l'esempio tu per primo, se il gruppo è bloccato."),
      bullet("Le coppie devono formarsi liberamente — non assegnarle tu."),
      bullet("Se emergono esclusioni o tensioni, gestiscile con calma e usale come occasione educativa."),
    ]),
    sp(100, 0),

    timeBox("15:30", "POMERIGGIO — Attività libretti (dediche)"),
    sp(60, 0),
    obiettivoBox("Spazio libero e guidato per dediche e riflessione finale. I ragazzi scrivono nel libretto bambini dediche ai compagni, agli educatori e a Don Michele."),
    sp(60, 0),
    guidaBox("Svolgimento", [
      bodySm("Distribuire i libretti bambini (o fogli appositi). Ogni ragazzo ha 30-40 minuti per scrivere dediche personalizzate."),
      bodySm("Destinatari suggeriti: un compagno di gruppo, un educatore, don Michele, se stesso."),
      bodySm("Musica di sottofondo leggera, atmosfera rilassata. Gli educatori circolano e si rendono disponibili per un momento di dialogo personale."),
    ]),
    noteBox([
      bullet("Non forzare chi non vuole scrivere — anche stare in silenzio o disegnare è ok."),
      bullet("Questo momento può generare emozioni: alcune ragazze piangono, alcuni ragazzi diventano insolitamente silenziosi. È normale e positivo."),
    ]),
    sp(100, 0),

    timeBox("19:30", "PREGHIERA DELLA SERA — 21 agosto"),
    sp(60, 0),
    guidaBox("Preparazione", [
      bodySm("Segno speciale: audio di persone care. Preparare in anticipo brevi registrazioni audio di genitori o persone significative dei ragazzi (coordinate con le famiglie prima del campo)."),
      bodySm("Far ascoltare i messaggi durante la preghiera, prima della benedizione finale."),
    ]),
    sp(60, 0),
    C("Nel nome del Padre, del Figlio e dello Spirito Santo."), Tall("Amen."),
    C("Il Signore, che ci fascia le ferite e si prende cura della nostra stanchezza, sia con tutti voi."), Tall("E con il tuo spirito."),
    sp(60, 0),
    salmoVerses("Salmo 145 — Loda il Signore, anima mia", "Sal 145",
      ["Loda il Signore, anima mia:","loderò il Signore finché ho vita,","canterò inni al mio Dio finché esisto.","",
       "Non confidate nei potenti,","in un uomo che non può salvare.","",
       "Beato chi ha per aiuto il Dio di Giacobbe:","la sua speranza è nel Signore suo Dio,","che ha fatto il cielo e la terra,","il mare e quanto contiene.","",
       "Il Signore libera i prigionieri,","il Signore ridona la vista ai ciechi,","il Signore rialza chi è caduto,","il Signore ama i giusti.","",
       "Il Signore regna per sempre,","il tuo Dio, o Sion, per ogni generazione."]
    ),
    sp(60, 0),
    vangeloBlock("Gv 19, 25-30 — Gesù affida Maria", [
      "«Donna, ecco tuo figlio!» — «Ecco tua madre!»",
      "Sulla croce, nel momento più difficile, Gesù non pensa a sé. Si prende cura. Ci lascia l'una all'altro.",
      "Anche noi, in questo campo, siamo stati affidati gli uni agli altri."
    ]),
    sp(60, 0),
    colorBox([
      heading3("Ascolto dei messaggi audio", MARRONE),
      bodySm("Far ascoltare i messaggi delle persone care. Dopo ogni messaggio, un momento di silenzio."),
      bodySm("Se non ci sono messaggi disponibili: invitare ciascun ragazzo a pensare in silenzio a chi si prende cura di lui ogni giorno."),
    ], MARRONE_CHIARO),
    sp(60, 0),
    C("Preghiamo insieme il Signore che non ci lascia mai soli nelle nostre fatiche:"), Tall("Grazie, Signore, per chi si prende cura di noi."),
    sp(60, 0),
    C("Il Signore vi dia un cuore attento e occhi che sanno vedere oltre l'apparenza."), Tall("Amen."),
    C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), Tall("Amen."),
    sp(100, 0),

    timeBox("21:30", "SERATA — Caccia al Tesoro «Il Castello dei Misteri»"),
    sp(60, 0),
    guidaBox("Preparazione", [
      bodySm("La caccia al tesoro si svolge in 5 giochi. Tra i ragazzi ci sono dei TRADITORI nascosti (educatori o ragazzi selezionati) che cercano di fare perdere le squadre dall'interno."),
      bodySm("Preparare in anticipo: indizi scritti, oggetti di scena, postazioni delimitate. Dividere i ragazzi in squadre miste."),
    ]),
    sp(60, 0),
    guidaBox("Gioco 1 — La Camera dei Segreti", [
      bodySm("Ogni squadra riceve una busta sigillata con un indizio. Devono decodificarlo (cifrario semplice o rebus) per trovare la prossima postazione."),
      bodySm("Il traditore nella squadra riceve un indizio falso e deve convincere il gruppo a seguirlo."),
    ]),
    sp(60, 0),
    guidaBox("Gioco 2 — Il Labirinto Bendato", [
      bodySm("Un membro della squadra (bendato) deve attraversare un percorso guidato solo dalle istruzioni verbali dei compagni."),
      bodySm("Il traditore dà istruzioni sbagliate di tanto in tanto. Il gruppo deve capire chi è il sabotatore."),
    ]),
    sp(60, 0),
    guidaBox("Gioco 3 — Il Codice del Castello", [
      bodySm("Le squadre devono risolvere 5 enigmi logici per trovare un codice a 5 cifre che apre un 'forziere' (scatola con il prossimo indizio)."),
      bodySm("Il traditore ha una delle cifre del codice errate — la squadra deve confrontarsi per trovare l'errore."),
    ]),
    sp(60, 0),
    guidaBox("Gioco 4 — La Torre dei Testimoni", [
      bodySm("Una serie di affermazioni su eventi del campo viene letta. I ragazzi devono dire se sono VERE o FALSE."),
      bodySm("Chi risponde correttamente fa avanzare la squadra. Il traditore risponde sempre il contrario."),
    ]),
    sp(60, 0),
    guidaBox("Gioco 5 — Il Tesoro Finale", [
      bodySm("Le squadre giungono alla postazione finale. L'ultima sfida: costruire una torre con i materiali trovati lungo la caccia (biglietti, oggetti, ecc.)."),
      bodySm("Vince la squadra con la torre più alta. Il 'tesoro' è un premietto simbolico (dolci, gadget del campo)."),
    ]),
    sp(60, 0),
    noteBox([
      bullet("Svelare i traditori alla fine, non durante — la rivelazione è il momento più divertente."),
      bullet("Assicurarsi che nessun ragazzo si senta davvero escluso o umiliato dal ruolo del traditore."),
      bullet("Terminare con classifica e saluto della buonanotte: domani è l'ultimo giorno."),
    ]),
    pb()
  ];
}

// ── GIORNO 5 — ANDARE (22 agosto) ───────────────────────────────────────────
function pagGiorno5() {
  return [
    ...dayHeader(5,"→","ANDARE","22 agosto 2026","Verso l'Alto!"),
    sp(100, 0),

    timeBox("8:30", "PREGHIERA DEL MATTINO — 22 agosto"),
    sp(60, 0),
    guidaBox("Preparazione", [
      bodySm("Ultimo mattino del campo. Atmosfera di gratitudine e invio. Richiamare il motto 'Verso l'Alto' come orizzonte della vita intera, non solo del campo."),
    ]),
    sp(60, 0),
    C("Nel nome del Padre, del Figlio e dello Spirito Santo."), Tall("Amen."),
    C("Il Signore, che ci chiama a metterci in viaggio e guida i nostri passi verso la gioia, sia con tutti voi."), Tall("E con il tuo spirito."),
    sp(60, 0),
    salmoVerses("Salmo 138 — Signore, tu mi scruti e mi conosci", "Sal 138",
      ["Signore, tu mi scruti e mi conosci,","tu sai quando seggo e quando mi alzo.","Penetri da lontano i miei pensieri,","osservi il mio cammino e il mio riposo.","",
       "Ti sono note tutte le mie vie;","la mia parola non è ancora sulla lingua","e tu, Signore, già la conosci tutta.","",
       "Dove andare lontano dal tuo spirito,","dove fuggire dalla tua presenza?","Se salgo in cielo, là tu sei;","se scendo negli inferi, eccoti.","",
       "Se prendo le ali dell'aurora","per abitare all'estremità del mare,","anche là mi guida la tua mano.","",
       "Sei tu che hai creato le mie viscere","e mi hai tessuto nel seno di mia madre.","Ti lodo, perché mi hai fatto come un prodigio."]
    ),
    sp(60, 0),
    colorBox([
      heading3("Preghiera Corale", BLU_SCURO),
      new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:"Signore Gesù, mettiti Tu alla guida dei nostri passi.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:"Insegnaci ad andare verso chi è solo, ad andare oltre i nostri pregiudizi, ad andare avanti anche quando la strada è in salita.",size:20,color:TESTO,font:"Arial",italics:true})]}),
      new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:"Non permettere che la paura ci blocchi a terra, ma donaci lo slancio per seguire i nostri sogni. Amen.",size:20,color:TESTO,font:"Arial",italics:true})]})
    ], GRIGIO),
    sp(60, 0),
    C("Il Signore vi benedica nel vostro andare e nel vostro tornare. Sia Lui la vostra strada, la vostra meta, la vostra forza."), Tall("Amen."),
    C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), Tall("Amen."),
    sp(100, 0),

    timeBox("10:00", "CATECHESI — Vangelo del giorno"),
    sp(60, 0),
    vangeloBlock("Mt 28, 16-20 — Il mandato missionario", [
      "Gli undici discepoli andarono in Galilea, sul monte che Gesù aveva loro indicato.",
      "Quando lo videro, si prostrarono. Lui si avvicinò e disse loro:",
      "«A me è stato dato ogni potere in cielo e sulla terra.",
      "Andate dunque e fate discepoli tutti i popoli, battezzandoli nel nome del Padre e del Figlio e dello Spirito Santo,",
      "insegnando loro a osservare tutto ciò che vi ho comandato.",
      "Ed ecco, io sono con voi tutti i giorni, fino alla fine del mondo.»"
    ]),
    sp(100, 0),

    timeBox("10:30", "GRUPPI STUDIO — Giorno 5: ANDARE"),
    sp(60, 0),
    obiettivoBox("Fare una verifica del campo e aiutare i ragazzi a portare a casa qualcosa di concreto. Il campo non è un mondo a parte: deve nutrire la vita di tutti i giorni."),
    sp(80, 0),
    guidaBox("Step 1 — Cosa ha funzionato (aspetti pratici, 10 min)", [
      bodySm("⚠️ Questo step riguarda l'organizzazione, NON i temi catechetici."),
      bullet("Cosa ti è piaciuto di questo campo sul piano pratico?"),
      bullet("Cosa cambieresti o miglioreresti?"),
      bodySm("Prendere nota delle risposte: serviranno per il feedback organizzativo post-campo."),
    ]),
    sp(60, 0),
    guidaBox("Step 2 — Cosa ti ha colpito (contenuti, 15 min)", [
      bodyItalic("«In questi giorni, cosa ti ha colpito di più? Una tematica, una frase, un simbolo, un gesto, una catechesi, una conversazione…»"),
      bodySm("Lascia che ognuno condivida liberamente. Non commentare troppo — ascolta."),
    ]),
    sp(60, 0),
    guidaBox("Step 3 — Come lo porto a casa (15 min)", [
      bodyItalic("«In che modo puoi mettere in pratica nella vita quotidiana quello che hai vissuto qui?»"),
      bodySm("Ogni ragazzo scrive su un cartoncino UNA COSA CONCRETA che si impegna a fare tornando a casa."),
      bodySm("I cartoncini vengono assemblati in un CARTELLONE da portare durante la Messa conclusiva."),
      sp(40,0),
      colorBox([
        bodySm("Il cartellone rappresenta la risposta del gruppo al mandato missionario — «Andate e fate discepoli» (Mt 28, 19). Sarà portato in processione all'ingresso della Messa.", true, MARRONE),
      ], MARRONE_CHIARO),
    ]),
    sp(60, 0),
    guidaBox("Conclusione dell'educatore", [
      bodyItalic("\"Quello che abbiamo vissuto in questi giorni non può restare un mondo a parte. Il motto di Frassati era 'Verso l'Alto' — ma l'alto non è lontano dalla vita: è la vita vissuta con pienezza, con fede, con amore. Adesso tocca a voi.\""),
      bodySm("Richiama i segni dei giorni precedenti: ancora (cercare) → montagna (ascoltare) → bandiera (restare) → abbraccio (prendersi cura) → pagina bianca (andare). Mostra come si connettono."),
    ]),
    sp(60, 0),
    noteBox([
      bullet("Il cartellone è il segno più visibile di questo giorno: curane la realizzazione, sarà esposto durante la Messa."),
      bullet("Incoraggia impegni piccoli e concreti — meglio una cosa sola davvero fatta che dieci propositi dimenticati."),
    ]),
    sp(100, 0),

    timeBox("15:30", "MESSA CONCLUSIVA"),
    sp(60, 0),
    guidaBox("Preparazione della Messa", [
      bodySm("La Messa è celebrata da Don Michele Torre. Ruolo degli educatori: coordinare i ragazzi, portare il cartellone in processione all'ingresso."),
      bodySm("Processione iniziale: un rappresentante per ogni gruppo porta una parte del cartellone/i cartoncini degli impegni."),
      bodySm("Durata stimata: 60-75 minuti."),
    ]),
    colorBox([
      body("Al termine della Messa: partenza verso le rispettive città. Orario indicativo: ore 17:00.", false, 21, BLU_SCURO),
      bodySm("Raccogliere eventuali oggetti dimenticati. Salutare la struttura e ringraziare il personale."),
    ], BLU_CHIARO),
    pb()
  ];
}

// ── INFO PRATICHE ────────────────────────────────────────────────────────────
function pagInfoPratiche() {
  const CW1 = 2800, CW2 = 2200, CW3 = BODY_W - CW1 - CW2;
  const hCell = (t) => new TableCell({
    shading:{fill:BLU_SCURO,type:ShadingType.CLEAR,color:"auto"},
    margins:{top:80,bottom:80,left:120,right:80},
    borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
    children:[new Paragraph({children:[new TextRun({text:t,bold:true,size:19,color:BIANCO,font:"Arial"})]})]
  });
  const dCell = (t, bg="FFFFFF") => new TableCell({
    shading:{fill:bg,type:ShadingType.CLEAR,color:"auto"},
    margins:{top:60,bottom:60,left:120,right:80},
    borders:{top:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},bottom:{style:BorderStyle.SINGLE,size:4,color:"EEEEEE"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
    children:[new Paragraph({children:[new TextRun({text:t,size:19,color:TESTO,font:"Arial"})]})]
  });

  const educatoriRows = [
    ["Costantinopoli (Bisceglie)","Adriano Misino, Don Michele Torre","15"],
    ["Passavia (Bisceglie)","Adriano Cantarone, Federica Di Pilato","21"],
    ["Santi Matteo e Nicolò (Bisceglie)","—","4"],
    ["Santa Chiara (Trani)","Andrea Valenziano","8"],
    ["SMG (Trani)","Federica Lasorsa, Francesca Suriano, Miriam Soldano","10"],
    ["Spirito Santo (Trani)","Adriana Ricchiuti, F. Di Pilato, F. Di Meo, G. Cacamo, M. Binetti, M. Giardino, Paki P.","15"],
  ];

  return [
    heading1("Informazioni pratiche"),
    sp(60, 0),
    body("Camposcuola Diocesano ACR — Terza Media · Montecalvo Irpino · 18–22 agosto 2026"),
    body("Totale partecipanti: 93 (73 ragazzi · 16 educatori · 4 cuochi · 1 assistente)"),
    sp(80, 0),
    heading2("Parrocchie e educatori"),
    new Table({
      width:{size:BODY_W,type:WidthType.DXA},
      columnWidths:[CW1,CW2,CW3],
      borders:{top:{style:BorderStyle.SINGLE,size:6,color:"CCCCCC"},bottom:{style:BorderStyle.SINGLE,size:6,color:"CCCCCC"},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},insideH:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"},insideV:{style:BorderStyle.SINGLE,size:4,color:"DDDDDD"}},
      rows:[
        new TableRow({children:[hCell("Parrocchia"),hCell("Educatori"),hCell("Ragazzi")]}),
        ...educatoriRows.map(([p,e,r],i) => new TableRow({children:[
          dCell(p, i%2===0?BLU_CHIARO:"FFFFFF"),
          dCell(e, i%2===0?BLU_CHIARO:"FFFFFF"),
          dCell(r, i%2===0?BLU_CHIARO:"FFFFFF"),
        ]}))
      ]
    }),
    sp(120, 0),
    heading2("Logistica primo giorno (18 agosto)"),
    twoColTable([
      ["Orario","Evento"],
      ["8:00","Raduno a Bisceglie"],
      ["8:15","Raduno a Trani"],
      ["8:30 / 8:45","Partenza da Bisceglie / Trani"],
      ["11:30","Arrivo in struttura"],
      ["12:15","Presentazione del campo"],
      ["12:30","Pranzo"],
    ]),
    sp(120, 0),
    heading2("Programma giornate tipo (19–21 agosto)"),
    twoColTable([
      ["Orario","Momento"],
      ["7:30","Sveglia"],
      ["8:00","Colazione"],
      ["8:30","Preghiera del mattino"],
      ["9:00","Pulizie"],
      ["10:00","Catechesi (sacerdote)"],
      ["10:30","Gruppi studio"],
      ["13:00","Pranzo"],
      ["15:30","Attività pomeridiane"],
      ["18:00","Docce"],
      ["19:30","Preghiera della sera"],
      ["20:00","Cena"],
      ["21:30","Serata (20/08 → Veglia)"],
      ["23:30","Buonanotte"],
    ]),
    sp(120, 0),
    heading2("Ultimo giorno (22 agosto)"),
    twoColTable([
      ["Orario","Momento"],
      ["8:30","Preghiera del mattino"],
      ["10:00","Catechesi"],
      ["10:30","Gruppi studio — ANDARE + Cartellone Messa"],
      ["13:00","Pranzo"],
      ["15:30","Messa conclusiva"],
      ["17:00","Partenza"],
    ]),
    sp(120, 0),
    colorBox([
      heading3("Note generali", BLU_SCURO),
      bullet("Ogni educatore è responsabile del proprio gruppo durante i gruppi studio e le attività."),
      bullet("In caso di emergenza medica: contattare immediatamente il responsabile del campo."),
      bullet("Buonanotte tassativa alle 23:30 — silenzio nelle stanze."),
      bullet("Telefoni ritirati all'arrivo (modalità concordata dai responsabili)."),
    ], BLU_CHIARO),
  ];
}

// ── MAIN ─────────────────────────────────────────────────────────────────────
(async () => {
  const sections = [
    ...pagCopertina(),
    ...pagIntro(),
    ...pagProgramma(),
    ...pagGiorno1(),
    ...pagGiorno2(),
    ...pagGiorno3(),
    ...pagGiorno4(),
    ...pagGiorno5(),
    ...pagInfoPratiche(),
  ];

  const doc = new Document({
    sections: [{ properties: {
      page: { size: { width: W, height: H }, margin: { top: MT, bottom: MB, left: ML, right: MR } }
    }, children: sections }]
  });

  const buf = await Packer.toBuffer(doc);
  const outPath = "/tmp/docx_work/Libretto_Educatori_2026_NEW.docx";
  fs.writeFileSync(outPath, buf);
  console.log(`OK - Libretto Educatori: ${buf.length} bytes`);
})();
