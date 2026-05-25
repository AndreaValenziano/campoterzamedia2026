'use strict';
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak
} = require('docx');

// ===== PERCORSI =====
const IMG_DIR   = '/sessions/confident-hopeful-rubin/mnt/campo terza media/materiale guida/';
const BASE_DIR  = '/sessions/confident-hopeful-rubin/mnt/campo terza media/';
const OUT_FILE  = '/sessions/confident-hopeful-rubin/mnt/campo terza media/Libretto_Bambini_2026.docx';

// ===== DIMENSIONI PAGINA A4 (DXA) =====
const A4W = 11906, A4H = 16838;
// Margini per GUIDA sezione 6: sin 2.5cm, destra 2cm, sup 2cm, inf 2cm
const ML = 1418, MR = 1134, MT = 1134, MB = 1134;
const CW = A4W - ML - MR; // 9354 DXA

// ===== COLORI (GUIDA sezione 1) =====
const BLU       = "1A6FA0";   // Blu ACR — titoli, intestazioni
const BLU_MED   = "5BAFD6";   // Sfondo header giornata (tono intermedio)
const BLU_LIGHT = "D5E8F0";   // Sfondo box informativi/istruzioni
const MARRONE   = "7A5230";   // Terra Frassati — citazioni, "Verso l'Alto!"
const MARR_BG   = "F5EDE0";   // Sfondo box citazioni Frassati
const VERDE     = "5A9B4A";   // Verde montagna — accenti secondari
const VERD_BG   = "E8F5E4";   // Sfondo box attività ragazzi
const GRIGIO_BG = "F2F3F4";   // Sfondo neutro (righe note)
const TESTO     = "333333";   // Grigio scuro — corpo testo (Candara)
const TESTO_SEC = "555555";   // Grigio medio — note, didascalie, salmi
const BIANCO    = "FFFFFF";

// ===== FONT (GUIDA sezione 2) =====
const FT = "More Sugar";          // Titoli sezioni, enfasi calligrafica
const FC = "Badger Script Bold";  // Solo titolo principale copertina
const FB = "Candara";             // Corpo testo, paragrafi, C:/T:

// ===== BORDI =====
const nb = { style: BorderStyle.NONE, size: 0, color: BIANCO };
const noBorders = { top: nb, bottom: nb, left: nb, right: nb };

// ===== UTILITY =====
function loadImg(f, base) {
  const dir = base || IMG_DIR;
  return fs.readFileSync(path.join(dir, f));
}
function pb() { return new Paragraph({ children: [new PageBreak()] }); }
function sp(before = 0, after = 120) {
  return new Paragraph({ spacing: { before, after }, children: [new TextRun("")] });
}
function img(filename, w, h, base) {
  const ext = path.extname(filename).toLowerCase().replace('.', '');
  const type = (ext === 'jpeg') ? 'jpg' : ext;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [new ImageRun({
      type,
      data: loadImg(filename, base),
      transformation: { width: w, height: h },
      altText: { title: filename, description: filename, name: filename }
    })]
  });
}

function colorBox(children, fill, borderColor, pad = { top: 120, bottom: 120, left: 200, right: 200 }) {
  const bc = borderColor || fill;
  const border = { style: BorderStyle.SINGLE, size: 6, color: bc };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    rows: [new TableRow({ children: [new TableCell({
      borders,
      shading: { fill, type: ShadingType.CLEAR },
      margins: pad,
      width: { size: CW, type: WidthType.DXA },
      children
    })] })]
  });
}

// ===== INTESTAZIONE GIORNATA (GUIDA sezione 4 — Sezioni giornata) =====
// Titolo: More Sugar 24pt (#1A6FA0), data: Candara 11pt italic (#555555)
// Citazione: More Sugar 12pt italic (#7A5230) su sfondo beige caldo
function dayHeader(num, icon, theme, date, citation) {
  return [
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], rows: [
      // Riga 1: box blu con "Giorno N — TEMA"
      new TableRow({ children: [new TableCell({
        borders: noBorders,
        shading: { fill: BLU_MED, type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 60, left: 200, right: 200 },
        width: { size: CW, type: WidthType.DXA },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 40 },
            children: [new TextRun({ text: `${icon}  Giorno ${num} — ${theme}`, size: 48, color: BIANCO, font: FT })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 80 },
            children: [new TextRun({ text: date, italics: true, size: 22, color: "E0F0FA", font: FB })]
          })
        ]
      })] }),
      // Riga 2: box marrone chiaro con citazione Frassati
      new TableRow({ children: [new TableCell({
        borders: noBorders,
        shading: { fill: MARR_BG, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        width: { size: CW, type: WidthType.DXA },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `“${citation}”`, italics: true, size: 24, color: MARRONE, font: FT })]
        })]
      })] })
    ] }),
    sp(160, 0)
  ];
}

// ===== BARRA LITURGIA =====
// Font titolo barra: More Sugar (è un elemento di intestazione), 13pt
function liturgyBar(emoji, title, fill = BLU_LIGHT, tc = BLU) {
  return [
    sp(160, 0),
    colorBox(
      [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `${emoji}  ${title}`, bold: true, size: 26, color: tc, font: FT })]
      })],
      fill, tc, { top: 100, bottom: 100, left: 160, right: 160 }
    ),
    sp(80, 0)
  ];
}

// ===== TESTI LITURGICI C: / T: =====
// Candara 12pt (#333333), C: bold blu, T: bold verde (GUIDA sezione 2 e 4)
function C(text) {
  return new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [
      new TextRun({ text: "C: ", bold: true, size: 24, color: BLU, font: FB }),
      new TextRun({ text, size: 24, color: TESTO, font: FB })
    ]
  });
}
function T(text) {
  return new Paragraph({
    spacing: { before: 40, after: 80 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: "T: ", bold: true, size: 24, color: VERDE, font: FB }),
      new TextRun({ text, bold: true, size: 24, color: TESTO, font: FB })
    ]
  });
}

// ===== SALMO (versetti) =====
// Titolo salmo: Candara 12pt bold marrone; versi: Candara 11pt italic #555555
function salmoVerses(title, verses) {
  const paras = [new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text: title, bold: true, size: 24, color: MARRONE, font: FB })]
  })];
  verses.forEach(v => {
    if (v === "") paras.push(sp(0, 60));
    else paras.push(new Paragraph({
      spacing: { before: 0, after: 20 },
      indent: { left: 360 },
      children: [new TextRun({ text: v, italics: true, size: 22, color: TESTO_SEC, font: FB })]
    }));
  });
  paras.push(T("Gloria al Padre, al Figlio e allo Spirito Santo, come era nel principio, ora e sempre nei secoli dei secoli. Amen."));
  return paras;
}

// ===== BRANO SCRITTURISTICO =====
// Riferimento: Candara 12pt bold marrone; versi: Candara 11pt italic #555555
function scriptureBlock(ref, verses) {
  const paras = [new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text: ref, bold: true, size: 24, color: MARRONE, font: FB })]
  })];
  verses.forEach(v => paras.push(new Paragraph({
    spacing: { before: 0, after: 60 },
    indent: { left: 360 },
    children: [new TextRun({ text: v, italics: true, size: 22, color: TESTO_SEC, font: FB })]
  })));
  return paras;
}

// ===== SEGNAPOSTO CANTO =====
function canto(label = "Canto") {
  return new Paragraph({
    spacing: { before: 120, after: 80 },
    children: [
      new TextRun({ text: `🎵 ${label}: `, bold: true, size: 24, color: BLU, font: FB }),
      new TextRun({ text: "________________________________", size: 24, color: "AAAAAA", font: FB })
    ]
  });
}

// ===== SEGNO DEL GIORNO (box verde) =====
function segno(text) {
  return [
    sp(80, 0),
    colorBox(
      [new Paragraph({ children: [new TextRun({ text, size: 24, color: TESTO, font: FB })] })],
      VERD_BG, VERDE, { top: 100, bottom: 100, left: 160, right: 160 }
    ),
    sp(80, 0)
  ];
}

// ===== RIGHE PER NOTE =====
// Etichetta: Candara 12pt bold; righe: bordo inferiore grigio chiaro
function noteLines(label, n = 4) {
  const paras = [new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: label, bold: true, size: 24, color: TESTO, font: FB })]
  })];
  for (let i = 0; i < n; i++) {
    paras.push(new Paragraph({
      spacing: { before: 0, after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" } },
      children: [new TextRun({ text: "", size: 24, font: FB })]
    }));
  }
  return colorBox(paras, GRIGIO_BG, "CCCCCC", { top: 100, bottom: 100, left: 160, right: 160 });
}

// ===== SEGNAPOSTO (materiale mancante) =====
function placeholder(text) {
  return [
    sp(120, 0),
    colorBox(
      [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "⏳", size: 48, font: FB })] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 0 },
          children: [new TextRun({ text, bold: true, size: 24, color: BLU, font: FB })]
        })
      ],
      "EBF5FB", BLU, { top: 180, bottom: 180, left: 200, right: 200 }
    ),
    sp(120, 0)
  ];
}

// ===== CORPO TESTO =====
// Candara 12pt #333333
function bodySm(text, before = 60, after = 60) {
  return new Paragraph({
    spacing: { before, after },
    children: [new TextRun({ text, size: 24, color: TESTO, font: FB })]
  });
}

// ===== PREGHIERA CORALE (box marrone chiaro) =====
function preghieraCorale(lines) {
  const paras = [new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: "Preghiera Corale", bold: true, size: 24, color: MARRONE, font: FT })]
  })];
  lines.forEach(l => paras.push(new Paragraph({
    spacing: { before: 0, after: 20 },
    indent: { left: 200 },
    children: [new TextRun({ text: l, italics: true, size: 22, color: TESTO, font: FB })]
  })));
  return colorBox(paras, MARR_BG, MARRONE, { top: 120, bottom: 120, left: 200, right: 200 });
}

// ===== PADRE NOSTRO =====
function padreNostro() {
  return [
    sp(80, 0),
    new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "Padre Nostro", bold: true, size: 24, color: TESTO, font: FB })]
    }),
    sp(40, 0)
  ];
}

// ===========================
// ===== LITURGIA =====
// ===========================

function sera18() { return [
  ...liturgyBar("🌙", "PREGHIERA DELLA SERA  —  18 agosto"),
  canto("Canto iniziale"), sp(40, 0),
  C("Nel nome del Padre, del Figlio e dello Spirito Santo."), T("Amen."),
  C("Il Signore, che ci cerca per primo e conosce la nostra sete, sia con tutti voi."), T("E con il tuo spirito."),
  C("Siamo arrivati alla fine di questa giornata. Ma cosa abbiamo cercato davvero? Stasera ci mettiamo accanto a una donna che è andata al pozzo per cercare acqua e ha trovato molto di più. Lasciamo che Gesù si sieda accanto a noi e ci chieda: “Dammi da bere”."),
  sp(60, 0),
  ...salmoVerses("Salmo 41", [
    "Come la cerva anela", "ai corsi d’acqua,", "così l’anima mia anela a te, o Dio.", "",
    "L’anima mia ha sete di Dio,", "del Dio vivente:", "quando verrò e vedrò", "il volto di Dio?", "",
    "Le lacrime sono il mio pane", "giorno e notte,", "mentre mi dicono sempre:", "“Dov’è il tuo Dio?”", "",
    "Questo io ricordo / e l’anima mia si strugge:", "avanzavo tra la folla,", "la precedevo fino alla casa di Dio,", "fra canti di gioia e di lode / di una moltitudine in festa.", "",
    "Perché ti rattristi, anima mia, / perché ti agiti in me?", "Spera in Dio: ancora potrò lodarlo,", "lui, salvezza del mio volto e mio Dio.", "",
    "In me si rattrista l’anima mia;", "perciò di te mi ricordo", "dalla terra del Giordano e dell’Ermon, / dal monte Misar.", "",
    "Un abisso chiama l’abisso", "al fragore delle tue cascate;", "tutti i tuoi flutti e le tue onde", "sopra di me sono passati.", "",
    "Di giorno il Signore mi dona il suo amore", "e di notte il suo canto è con me,", "preghiera al Dio della mia vita.", "",
    "Dirò a Dio: “Mia roccia! / Perché mi hai dimenticato?", "Perché triste me ne vado, / oppresso dal nemico?”", "",
    "Perché ti rattristi, anima mia, / perché ti agiti in me?", "Spera in Dio: ancora potrò lodarlo,", "lui, salvezza del mio volto e mio Dio."
  ]),
  sp(60, 0),
  ...scriptureBlock("Vangelo secondo Giovanni (Gv 4, 5-15)", [
    "Giunse dunque a una città della Samaria, chiamata Sicar. Gesù dunque, stanco del cammino, stava sedere presso la fonte. Era circa l’ora sesta.",
    "Una donna della Samaria venne ad attingere l’acqua. Gesù le disse: «Dammi da bere».",
    "La donna samaritana allora gli disse: «Come mai tu che sei Giudeo chiedi da bere a me, che sono una donna samaritana?».",
    "Gesù le rispose: «Se tu conoscessi il dono di Dio e chi è che ti dice: “Dammi da bere”, tu stessa gliene avresti chiesto, ed egli ti avrebbe dato dell’acqua viva».",
    "La donna gli disse: «Signore, tu non hai nulla per attingere, e il pozzo è profondo; da dove avresti quest’acqua viva? Sei tu più grande di Giacobbe, nostro padre, che ci diede questo pozzo?»",
    "Gesù le rispose: «Chiunque beve di quest’acqua avrà di nuovo sete; ma chi beve dell’acqua che io gli darò, non avrà mai più sete; anzi, l’acqua che io gli darò diventerà in lui una fonte d’acqua che scaturisce in vita eterna».",
    "La donna gli disse: «Signore, dammi di quest’acqua, affinché io non abbia più sete e non venga più fin qui ad attingere»."
  ]),
  sp(60, 0), bodySm("Riflessione del Celebrante", 80, 120), sp(40, 0),
  C("Portiamo al Signore la nostra sete e le nostre ricerche e diciamo insieme:"), T("Signore, dacci dell’acqua viva."),
  bodySm("Per quando cerchiamo la felicità nelle cose che passano e ci dimentichiamo di Te, preghiamo."), T("Signore, dacci dell’acqua viva."),
  bodySm("Per chi si sente “arido” e svuotato, perché trovi in questo campo una sorgente di amicizia, preghiamo."), T("Signore, dacci dell’acqua viva."),
  bodySm("Perché sappiamo cercare il bene in ogni compagno, anche in chi ci è meno simpatico, preghiamo."), T("Signore, dacci dell’acqua viva."),
  bodySm("Per la nostra ricerca di futuro e di sogni: guidaci Tu sulla strada giusta, preghiamo."), T("Signore, dacci dell’acqua viva."),
  bodySm("Per questo camposcuola e per chi ci attende a casa, preghiamo."), T("Signore, dacci dell’acqua viva."),
  ...segno("⚓  Segno del giorno: ogni ragazzo scrive il nome di una persona che è la propria ancora nella fede o nella vita"),
  ...padreNostro(),
  C("Il Signore sia il vostro pozzo di acqua fresca nei momenti di stanchezza."), T("Amen."),
  C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), T("Amen."),
  canto("Canto finale"),
]; }

function mattino19() { return [
  ...liturgyBar("☀️", "PREGHIERA DEL MATTINO  —  19 agosto"),
  canto("Canto iniziale"), sp(40, 0),
  C("Nel nome del Padre, del Figlio e dello Spirito Santo."), T("Amen."),
  C("Il Signore, che ci parla nel silenzio dell’aurora e nel rumore dei nostri passi, sia con tutti voi."), T("E con il tuo spirito."),
  C("Buongiorno! Spesso pensiamo che ascoltare significhi solo “sentire” dei suoni. Ma l’ascolto vero è un’arte: è fare spazio a qualcuno dentro di noi. Oggi il campo ci sfida a questo: ascoltare il compagno di stanza, ascoltare il bosco, ascoltare la voce di Dio che ci sussurra che siamo amati. Iniziamo la giornata chiedendo il dono di un ascolto attento."),
  sp(60, 0),
  ...salmoVerses("Salmo 16", [
    "Preservami, o Dio, perché io confido in te.", "",
    "Io ho detto all’Eterno: «Tu sei il mio Signore;", "io non ho altro bene all’infuori di te»;", "",
    "L’Eterno è la mia parte di eredità e il mio calice;", "tu mantieni quel che m’è toccato in sorte.", "La sorte è caduta per me in luoghi dilettevoli;", "una bella eredità mi è pur toccata!", "",
    "Io benedirò l’Eterno che mi consiglia;", "anche la notte le mie reni mi ammaestrano.", "Io ho sempre posto l’Eterno davanti agli occhi miei;", "poiché egli è alla mia destra, io non sarò affatto smosso.", "",
    "Perciò il mio cuore si rallegra e l’anima mia festeggia;", "anche la mia carne dimorerà al sicuro;", "poiché tu non abbandonerai l’anima mia in potere della morte.", "",
    "Tu mi mostrerai il sentiero della vita;", "vi sono gioie a sazietà alla tua presenza;", "vi sono delizie alla tua destra in eterno."
  ]),
  sp(60, 0),
  preghieraCorale([
    "Signore Gesù,", "mentre il mondo si sveglia e si riempie di rumori,", "noi ti chiediamo un momento di silenzio.",
    "Aiutaci oggi ad ascoltare davvero:", "non solo le parole, ma anche i silenzi dei nostri amici;", "non solo chi ci è attorno, ma la tua voce che ci guida.",
    "Liberaci dalla fretta di rispondere, donaci la pazienza di capire", "e aiutaci a dare il giusto peso alle parole.",
    "Che il nostro cuore sia come una terra buona", "pronta a ricevere la tua Parola. Amen."
  ]),
  sp(60, 0),
  C("Il Signore apra il vostro cuore all’ascolto e vi renda capaci di ascoltare gli altri."), T("Amen."),
  C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), T("Amen."),
  canto("Canto finale"),
]; }

function sera19() { return [
  ...liturgyBar("🌙", "PREGHIERA DELLA SERA  —  19 agosto"),
  canto("Canto iniziale"), sp(40, 0),
  C("Nel nome del Padre, del Figlio e dello Spirito Santo."), T("Amen."),
  C("Il Signore, che parla al cuore di chi lo ascolta con fiducia, sia con tutti voi."), T("E con il tuo spirito."),
  C("Dopo una giornata piena di voci, rumori e risate, stasera vogliamo fare silenzio. Ascoltare non è solo ricevere informazioni, ma sintonizzarsi sulla frequenza di Dio. Gesù oggi ci parla di ascolto attraverso le Beatitudini: parole che a volte sembrano strane o difficili, ma che, se ascoltate bene, sono la via per una gioia che non finisce mai."),
  sp(60, 0),
  ...salmoVerses("Salmo 27", [
    "Il Signore è la mia luce e la mia salvezza; di chi temerò?", "Il Signore è il baluardo della mia vita; di chi avrò paura?", "",
    "Se un esercito si accampasse contro di me,", "il mio cuore non avrebbe paura;", "se infuriasse la battaglia contro di me,", "anche allora sarei fiducioso.", "",
    "Una cosa ho chiesto al Signore, e quella ricerco:", "abitare nella casa del Signore tutti i giorni della mia vita,", "per contemplare la bellezza del Signore", "e meditare nel suo tempio.", "",
    "Poiché egli mi nasconderà nella sua tenda in giorno di sventura,", "mi custodirà nel luogo più segreto della sua dimora,", "mi porterà in alto sopra una roccia.", "",
    "O Signore, ascolta la mia voce quando t’invoco;", "abbi pietà di me, e rispondimi.", "Il mio cuore mi dice da parte tua: «Cercate il mio volto!»", "Io cerco il tuo volto, o Signore.", "",
    "Non nascondermi il tuo volto,", "non respingere con ira il tuo servo;", "non lasciarmi e non abbandonarmi, o Dio della mia salvezza!", "",
    "Qualora mio padre e mia madre m’abbandonino,", "il Signore mi accoglierà.", "O Signore, insegnami la tua via,", "guidami per un sentiero diritto.", "",
    "Spera nel Signore! Sii forte,", "il tuo cuore si rinfranchi;", "sì, spera nel Signore!"
  ]),
  sp(60, 0),
  ...scriptureBlock("Vangelo secondo Matteo (Mt 5, 1-12 — Le Beatitudini)", [
    "Vedendo le folle, Gesù salì sulla montagna e, messosi a sedere, gli si avvicinarono i suoi discepoli. Prendendo allora la parola, li ammaestrava dicendo:",
    "«Beati i poveri in spirito, perché di essi è il regno dei cieli.", "Beati gli afflitti, perché saranno consolati.",
    "Beati i miti, perché erediteranno la terra.", "Beati quelli che hanno fame e sete della giustizia, perché saranno saziati.",
    "Beati i misericordiosi, perché troveranno misericordia.", "Beati i puri di cuore, perché vedranno Dio.",
    "Beati gli operatori di pace, perché saranno chiamati figli di Dio.", "Beati i perseguitati per causa della giustizia, perché di essi è il regno dei cieli.",
    "Beati voi quando vi insulteranno, vi perseguiteranno e, mentendo, diranno ogni sorta di male contro di voi per causa mia. Rallegratevi ed esultate, perché grande è la vostra ricompensa nei cieli.»"
  ]),
  sp(60, 0), bodySm("Riflessione del Celebrante", 80, 120),
  C("Chiediamo al Signore la capacità di ascoltare queste parole non con le orecchie, ma con il cuore e diciamo insieme:"), T("Signore, aiutaci ad ascoltarti."),
  bodySm("Preghiere spontanee", 60, 100),
  ...segno("⛰️  Segno del giorno: gli educatori si avvicinano e sussurrano all’orecchio di ciascuno qualcosa di bello legato alla giornata"),
  ...padreNostro(),
  C("Il Signore benedica i vostri cuori perché siano pronti all’ascolto. Vi conceda una notte serena."), T("Amen."),
  C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), T("Amen."),
  canto("Canto finale"),
]; }

function mattino20() { return [
  ...liturgyBar("☀️", "PREGHIERA DEL MATTINO  —  20 agosto"),
  canto("Canto iniziale"), sp(40, 0),
  C("Nel nome del Padre, del Figlio e dello Spirito Santo."), T("Amen."),
  C("Il Signore, che rimane fedele per sempre e ci invita a dimorare nel suo amore, sia con tutti voi."), T("E con il tuo spirito."),
  C("Buongiorno! Dopo i primi giorni di campo, la tentazione è quella di lasciarsi trascinare dalla stanchezza. Ma oggi la parola d’ordine è “restare”. Restare nel gioco anche quando si perde, restare amici anche quando si discute, restare con Gesù anche quando non sentiamo “i fuochi d’artificio”. Rimanere è l’unico modo per mettere radici profonde."),
  sp(60, 0),
  ...salmoVerses("Salmo 90", [
    "Chi abita al riparo dell’Altissimo", "passerà la notte all’ombra dell’Onnipotente.", "Io dico al Signore: «Mio rifugio e mia fortezza,", "mio Dio in cui confido».", "",
    "Egli ti libererà dal laccio del cacciatore, / dalla peste che distrugge.", "Ti coprirà con le sue penne,", "sotto le sue ali troverai rifugio;", "la sua fedeltà ti sarà scudo e corazza.", "",
    "Non temerai il terrore della notte", "né la freccia che vola di giorno,", "la peste che vaga nelle tenebre,", "lo sterminio che devasta a mezzogiorno.", "",
    "«Lo libererò, perché a me si è legato,", "lo porrò al sicuro, perché ha conosciuto il mio nome.", "Mi invocherà e io gli darò risposta;", "nell’angoscia io sarò con lui,", "lo libererò e lo renderò glorioso.", "Lo sazierò di lunghi giorni", "e gli farò vedere la mia salvezza»."
  ]),
  sp(60, 0),
  preghieraCorale([
    "Signore Gesù,", "spesso siamo cercatori di novità,", "sempre pronti a correre verso qualcosa di diverso.",
    "Stamattina ti chiediamo il coraggio di restare.", "Aiutaci a restare nelle fatiche di oggi senza scappare,", "a restare fedeli ai nostri impegni e ai nostri amici.",
    "Insegnaci a dimorare nel tuo amore,", "come i tralci restano uniti alla vite per dare frutto.",
    "Benedici il nostro legame; che resti negli anni e ci dia la gioia di stare insieme. Amen."
  ]),
  sp(60, 0),
  C("Il Signore vi dia la forza di restare saldi nel bene. Vi custodisca nell’unità."), T("Amen."),
  C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), T("Amen."),
  canto("Canto finale"),
]; }

function veglia20() { return [
  ...liturgyBar("⭐", "VEGLIA DI PREGHIERA  —  20 agosto tarda sera", "FEF9E7", "9A7D0A"),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: "“Beato chi mangia di Te e vive per Te!”", bold: true, italics: true, size: 26, color: MARRONE, font: FT })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: "Montecalvo Irpino, 20 Agosto 2026", size: 20, color: TESTO_SEC, font: FB })]
  }),
  canto("Canto iniziale"), sp(40, 0),
  C("Saluto del celebrante"), sp(80, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "1ª Fase — Restare nell’AC", bold: true, size: 24, color: BLU, font: FT })] }),
    bodySm("All’inizio di questa veglia vogliamo collegarci con i gruppi studio di questa mattina. Come i discepoli di Emmaus, anche noi chiediamo al Signore: resta con noi, perché si fa sera.", 0, 60),
    bodySm("Per noi campisti restare significa rimanere saldi in Cristo nell’Eucaristia e coltivare la nostra fede nell’Azione Cattolica.", 0, 60),
  ], BLU_LIGHT, BLU, { top: 140, bottom: 140, left: 200, right: 200 }),
  sp(80, 0),
  new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "Il custode della stazione", bold: true, italics: true, size: 24, color: MARRONE, font: FB })] }),
  bodySm("In un piccolo paese di montagna c’era una vecchia stazione ferroviaria. I treni passavano ancora, ma sempre più raramente. Il custode era rimasto. Ogni mattina arrivava alla stessa ora, apriva le porte, spazzava il pavimento, controllava i binari come se dovesse arrivare qualcuno da un momento all’altro.", 0, 40),
  bodySm("«Non viene più nessuno», gli dicevano. Lui rispondeva solo: «Qualcuno potrebbe tornare.» E restava.", 0, 40),
  bodySm("Un inverno particolarmente freddo, il paese fu isolato dalla neve. Corsero alla vecchia stazione senza molte speranze. La trovarono aperta. Le luci accese. Il custode, seduto sulla panchina come sempre. Si alzò soltanto e disse: «Entrate. Qui almeno siete al sicuro.»", 0, 40),
  bodySm("A volte il valore di un luogo non è chi ci passa… ma chi sceglie di restare anche quando sembra inutile.", 0, 80),
  sp(80, 0),
  colorBox([new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "2ª Fase — Restare in Cristo", bold: true, size: 24, color: BLU, font: FT })] })],
    BLU_LIGHT, BLU, { top: 100, bottom: 100, left: 200, right: 200 }),
  sp(60, 0),
  ...salmoVerses("Salmo 148", [
    "Alleluia. Lodate il Signore dai cieli, lodatelo nell’alto dei cieli.", "",
    "Lodatelo, voi tutti, suoi angeli, lodatelo, voi tutte, sue schiere.", "Lodatelo, sole e luna, lodatelo, voi tutte, fulgide stelle.", "Lodatelo, cieli dei cieli, voi acque al di sopra dei cieli.", "",
    "Lodino tutti il nome del Signore, / perché egli disse e furono creati.", "Li ha stabiliti per sempre, / ha posto una legge che non passa.", "",
    "Lodate il Signore dalla terra, mostri marini e voi tutti abissi,", "fuoco e grandine, neve e nebbia, / vento di bufera che obbedisce alla sua parola,", "monti e voi tutte, colline, alberi da frutto e tutti voi, cedri,", "voi fiere e tutte le bestie, rettili e uccelli alati.", "",
    "I re della terra e i popoli tutti, / i governanti e i giudici della terra,", "i giovani e le fanciulle, / i vecchi insieme ai bambini", "lodino il nome del Signore:", "perché solo il suo nome è sublime, / la sua gloria risplende sulla terra e nei cieli.", "",
    "Egli ha sollevato la potenza del suo popolo. / È canto di lode per tutti i suoi fedeli,", "per i figli di Israele, popolo che egli ama. Alleluia."
  ]),
  sp(60, 0),
  ...scriptureBlock("Vangelo secondo Giovanni (Gv 6, 53-57 — L’Eucaristia)", [
    "Gesù disse: «In verità, in verità vi dico: se non mangiate la carne del Figlio dell’uomo e non bevete il suo sangue, non avrete in voi la vita.",
    "Chi mangia la mia carne e beve il mio sangue ha la vita eterna e io lo risusciterò nell’ultimo giorno.",
    "Perché la mia carne è vero cibo e il mio sangue vera bevanda.",
    "Chi mangia la mia carne e beve il mio sangue dimora in me e io in lui.",
    "Come il Padre, che ha la vita, ha mandato me e io vivo per il Padre, così anche colui che mangia di me vivrà per me.»"
  ]),
  bodySm("Riflessione del celebrante", 80, 120), sp(80, 0),
  colorBox([new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "3ª Fase — Rendere Grazie", bold: true, size: 24, color: BLU, font: FT })] })],
    BLU_LIGHT, BLU, { top: 100, bottom: 100, left: 200, right: 200 }),
  sp(60, 0),
  C("Con il cuore colmo di gioia, affidiamo a Dio chi ci insegna a rimanere nella Fede e nella Vita piena. Preghiamo insieme e diciamo:"), T("Ti ringraziamo, o Padre."),
  bodySm("San Pier Giorgio Frassati era un giovane come noi. Diceva sempre: perché staccarsi dalla vita? È così bella! Signore, ogni istante della vita è un dono Tuo. Preghiamo."), T("Ti ringraziamo, o Padre."),
  bodySm("Ti affidiamo Signore le nostre famiglie, a cominciare dai nostri genitori. Grazie a loro rimaniamo fedeli all’impegno di amore. Preghiamo."), T("Ti ringraziamo, o Padre."),
  bodySm("Vogliamo ringraziarTi, o Padre, per il dono dell’amicizia. Permettici di restare nei nostri rapporti umani. Preghiamo."), T("Ti ringraziamo, o Padre."),
  bodySm("Signore, ti benediciamo per le comunità parrocchiali di cui facciamo parte. Custodisci le nostre parrocchie, i nostri parroci, il nostro Vescovo Leonardo e l’Azione Cattolica Italiana. Preghiamo."), T("Ti ringraziamo, o Padre."),
  ...segno("🏳  Segno del giorno: scrivi il nome di una persona che ti mostra come restare — da affidare al Signore — sulla bandiera dell’AC"),
  C("Benedizione finale"), T("Amen."),
  canto("Canto finale"),
]; }

function mattino21() { return [
  ...liturgyBar("☀️", "PREGHIERA DEL MATTINO  —  21 agosto"),
  canto("Canto iniziale"), sp(40, 0),
  C("Nel nome del Padre, del Figlio e dello Spirito Santo."), T("Amen."),
  C("Il Signore, che si prende cura di ogni sua creatura con amore infinito, sia con tutti voi."), T("E con il tuo spirito."),
  C("Buongiorno! Il Signore si prende ogni giorno cura di noi. Non significa solo fare qualcosa per qualcuno, ma accorgersi che l’altro è un dono prezioso da custodire. Chiediamo al Signore di donarci mani operose e cuori attenti."),
  sp(60, 0),
  ...salmoVerses("Salmo 22", [
    "Il Signore è il mio pastore:", "non manco di nulla;", "su pascoli erbosi mi fa riposare,", "ad acque tranquille mi conduce.", "",
    "Mi rinfranca, mi guida per il giusto cammino, / per amore del suo nome.", "Se dovessi camminare in una valle oscura,", "non temerei alcun male, perché tu sei con me.", "",
    "Il tuo bastone e il tuo vincastro / mi danno sicurezza.", "Davanti a me tu prepari una mensa / sotto gli occhi dei miei nemici;", "cospargi di olio il mio capo. / Il mio calice trabocca.", "",
    "Felicità e grazia mi saranno compagne", "tutti i giorni della mia vita,", "e abiterò nella casa del Signore", "per lunghissimi anni."
  ]),
  sp(60, 0),
  preghieraCorale([
    "Signore Gesù,", "grazie perché Tu ti prendi cura di noi ogni istante.",
    "Oggi ti chiediamo di prestarci il tuo sguardo.", "Insegnaci a prenderci cura di chi ci sta vicino,", "specialmente di chi oggi è più stanco o si sente solo.",
    "Aiutaci a non sprecare le occasioni di bene", "e a trattare ogni persona con delicatezza.",
    "Fa’ che le nostre mani siano strumenti della tua tenerezza. Amen."
  ]),
  sp(60, 0),
  C("Il Signore vi renda capaci di gesti di cura gratuiti e gioiosi."), T("Amen."),
  C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), T("Amen."),
  canto("Canto finale"),
]; }

function sera21() { return [
  ...liturgyBar("🌙", "PREGHIERA DELLA SERA  —  21 agosto"),
  canto("Canto iniziale"), sp(40, 0),
  C("Nel nome del Padre, del Figlio e dello Spirito Santo."), T("Amen."),
  C("Il Signore, che ci fascia le ferite e si prende cura della nostra stanchezza, sia con tutti voi."), T("E con il tuo spirito."),
  C("Siamo arrivati alla fine di questa giornata dedicata alla cura. La vita spesso ci affida delle persone: materiale da maneggiare con cura! Ascoltiamo la Parola del Signore che ci ricorda la bellezza dello stare accanto agli altri."),
  sp(60, 0),
  ...salmoVerses("Salmo 145", [
    "Loda il Signore, anima mia:", "loderò il Signore finché ho vita,", "canterò inni al mio Dio finché esisto.", "",
    "Non confidate nei potenti,", "in un uomo che non può salvare.", "",
    "Beato chi ha per aiuto il Dio di Giacobbe:", "la sua speranza è nel Signore suo Dio,", "che ha fatto il cielo e la terra,", "il mare e quanto contiene,", "che rimane fedele per sempre,", "rende giustizia agli oppressi,", "dà il pane agli affamati.", "",
    "Il Signore libera i prigionieri,", "il Signore ridona la vista ai ciechi,", "il Signore rialza chi è caduto,", "il Signore ama i giusti,", "il Signore protegge i forestieri,", "egli sostiene l’orfano e la vedova.", "",
    "Il Signore regna per sempre, / il tuo Dio, o Sion, / di generazione in generazione."
  ]),
  sp(60, 0),
  ...scriptureBlock("Vangelo secondo Giovanni (Gv 19, 25-30)", [
    "Stavano presso la croce di Gesù sua madre, la sorella di sua madre, Maria di Clèofa e Maria di Màgdala.",
    "Gesù allora, vedendo la madre e lì accanto a lei il discepolo che egli amava, disse alla madre: «Donna, ecco il tuo figlio!». Poi disse al discepolo: «Ecco la tua madre!». E da quel momento il discepolo la prese nella sua casa.",
    "Dopo questo, Gesù, sapendo che ogni cosa era stata ormai compiuta, disse: «Ho sete». Gliela accostarono alla bocca. E dopo aver ricevuto l’aceto, Gesù disse: «Tutto è compiuto!». E, chinato il capo, spirò."
  ]),
  sp(60, 0), bodySm("Riflessione del Celebrante", 80, 120),
  C("Preghiamo insieme il Signore che non ci lascia mai soli nelle nostre fatiche e diciamo insieme:"), T("Prenditi cura di noi, Signore."),
  bodySm("Preghiere spontanee", 60, 100),
  ...segno("🤝  Segno del giorno: ascoltiamo le voci di chi ci vuole bene a casa — nominati uno a uno"),
  ...padreNostro(),
  C("Il Signore vi dia un cuore attento e occhi che sanno vedere oltre l’apparenza."), T("Amen."),
  C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), T("Amen."),
  canto("Canto finale"),
]; }

function mattino22() { return [
  ...liturgyBar("☀️", "PREGHIERA DEL MATTINO  —  22 agosto"),
  canto("Canto iniziale"), sp(40, 0),
  C("Nel nome del Padre, del Figlio e dello Spirito Santo."), T("Amen."),
  C("Il Signore, che ci chiama a metterci in viaggio e guida i nostri passi verso la gioia, sia con tutti voi."), T("E con il tuo spirito."),
  C("Buongiorno! La vita è un movimento continuo. Oggi vogliamo riflettere sulla parola “Andare”. Non è solo spostarsi da un posto all’altro, ma è il coraggio di uscire dalle nostre comodità per incontrare gli altri. Come i discepoli, anche noi oggi siamo invitati a camminare non da soli, pronti a scoprire dove il Signore ci sta portando."),
  sp(60, 0),
  ...salmoVerses("Salmo 138", [
    "Signore, tu mi scruti e mi conosci,", "tu sai quando seggo e quando mi alzo.", "Penetri da lontano i miei pensieri,", "mi scruti quando cammino e quando riposo.", "",
    "Ti sono note tutte le mie vie;", "la mia parola non è ancora sulla lingua", "e tu, Signore, già la conosci tutta.", "",
    "Dove andare lontano dal tuo spirito,", "dove fuggire dalla tua presenza?", "Se salgo in cielo, là tu sei, / se scendo negli inferi, eccoti.", "",
    "Se prendo le ali dell’aurora", "per abitare all’estremità del mare,", "anche là mi guida la tua mano", "e mi afferra la tua destra.", "",
    "Sei tu che hai creato le mie viscere", "e mi hai tessuto nel seno di mia madre.", "Ti lodo, perché mi hai fatto come un prodigio;", "sono stupende le tue opere, / tu mi conosci fino in fondo.", "",
    "Quanto profondi per me i tuoi pensieri,", "quanto grande il loro numero, o Dio;", "se li conto sono più della sabbia,", "se li credo finiti, con te sono ancora."
  ]),
  sp(60, 0),
  preghieraCorale([
    "Signore Gesù,", "mettiti Tu alla guida dei nostri passi.",
    "Insegnaci ad andare verso chi è solo,", "ad andare oltre i nostri pregiudizi,", "ad andare avanti anche quando la strada si fa in salita.",
    "Non permettere che la paura ci blocchi a terra,", "ma donaci lo slancio per seguire i nostri sogni.",
    "Rendici costanti nell’annunciare la Tua Parola, mandaci nel Mondo come tuoi servi. Amen."
  ]),
  sp(60, 0),
  C("Il Signore vi benedica nel vostro andare e nel vostro tornare. Sia Lui la vostra strada, la vostra mappa e la vostra meta."), T("Amen."),
  C("E la benedizione di Dio onnipotente, Padre e Figlio e Spirito Santo, discenda su di voi e con voi rimanga sempre."), T("Amen."),
  canto("Canto finale"),
]; }

// ===========================
// ===== PAGINE =====
// ===========================

// GUIDA sezione 4 — Pagina 1: immagine full-page senza testo
// GUIDA sezione 4 — Pagina 2: titolo "Vivere, non vivacchiare!" Badger Script Bold 72pt #1A6FA0
function pagCopertina() { return [
  // — Pagina 1: copertina full-page (COPERTINA LIBRETTO BAMBINI.png) —
  img("COPERTINA LIBRETTO BAMBINI.png", 620, 878, BASE_DIR),
  pb(),
  // — Pagina 2: titolo e dati —
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text: "Vivere, non vivacchiare!", bold: true, size: 144, color: BLU, font: FC })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: "San Pier Giorgio Frassati", italics: true, size: 28, color: MARRONE, font: FT })]
  }),
  sp(80, 0),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: "Camposcuola Diocesano ACR  ·  Terza Media", size: 24, color: TESTO_SEC, font: FT })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "Montecalvo Irpino  ·  18 – 22 agosto 2026", size: 24, color: TESTO_SEC, font: FT })]
  }),
  sp(80, 0),
  // Box nome / parrocchia / gruppo — More Sugar 11pt
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Nome: _______________________________________", size: 22, font: FT, color: TESTO })] }),
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Parrocchia: _________________________________", size: 22, font: FT, color: TESTO })] }),
    new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "Gruppo studio: ______________________________", size: 22, font: FT, color: TESTO })] })
  ], GRIGIO_BG, "CCCCCC", { top: 120, bottom: 120, left: 200, right: 200 }),
]; }

// GUIDA sezione 4 — Chi era Frassati?
// Titolo: More Sugar 24pt #1A6FA0; corpo: Candara 12pt; "Verso l'Alto!": More Sugar 18pt #7A5230
function pagFrassati() { return [pb(),
  new Paragraph({
    spacing: { before: 0, after: 120 },
    children: [new TextRun({ text: "Chi era Pier Giorgio Frassati?", bold: true, size: 48, color: BLU, font: FT })]
  }),
  bodySm("Pier Giorgio Frassati nacque a Torino nel 1901. Figlio di una famiglia benestante, avrebbe potuto “vivacchiare” comodamente — ma scelse di vivere davvero.", 0, 80),
  bodySm("Amava la montagna, lo sci, l’alpinismo. Si alzava presto la mattina per andare a Messa ogni giorno. Studiava da ingegnere e nel tempo libero portava cibo, medicine e soldi ai poveri dei quartieri più difficili di Torino — spesso tornando a casa senza il cappotto perché l’aveva dato a qualcuno che aveva più freddo di lui.", 0, 80),
  bodySm("Era allegro. Era concreto. Non predicava — faceva. Il suo motto era semplice e rivoluzionario:", 0, 40),
  // "Verso l'Alto!" — More Sugar 18pt #7A5230 (MARRONE), NON rosso (GUIDA regola 1)
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "«Verso l’Alto!»", bold: true, italics: true, size: 36, color: MARRONE, font: FT })]
  }),
  bodySm("Morì a soli 24 anni, stroncato da una malattia che aveva preso da un bambino povero che assisteva. Giovanni Paolo II lo beatificò nel 1990. È il santo dei giovani, dello sport, dei social — della vita vissuta con tutto se stessi.", 0, 0),
]; }

// GUIDA sezione 4 — Il cammino dei 5 giorni
function pag5Giorni() {
  const dati = [
    ["⚓", "1", "CERCARE",        "18 agosto", "Gv 4, 1-42"],
    ["⛰️",  "2", "ASCOLTARE",     "19 agosto", "Mt 5, 1-12"],
    ["🏳", "3", "RESTARE",        "20 agosto", "Lc 24 + Veglia"],
    ["🤝", "4", "PRENDERSI CURA", "21 agosto", "Gv 19, 25-30"],
    ["→",  "5", "ANDARE",         "22 agosto", "Mt 28, 16-20"],
  ];
  const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const W1 = 900, W2 = 2200, W3 = CW - W1 - W2; // 6254
  const rows = dati.map(([icon, g, theme, date, vang]) => new TableRow({ children: [
    new TableCell({
      borders, shading: { fill: BLU_MED, type: ShadingType.CLEAR },
      width: { size: W1, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 80, right: 80 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: icon, size: 32, font: FB })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `G.${g}`, bold: true, size: 18, color: BIANCO, font: FT })] })
      ]
    }),
    new TableCell({
      borders, shading: { fill: BLU_LIGHT, type: ShadingType.CLEAR },
      width: { size: W2, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [
        new Paragraph({ children: [new TextRun({ text: theme, bold: true, size: 22, color: BLU, font: FT })] }),
        new Paragraph({ children: [new TextRun({ text: date, size: 18, color: TESTO_SEC, font: FB })] })
      ]
    }),
    new TableCell({
      borders,
      width: { size: W3, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: vang, size: 20, color: MARRONE, font: FB })] })]
    }),
  ]}));
  return [pb(),
    new Paragraph({
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: "Il cammino dei 5 giorni", bold: true, size: 48, color: BLU, font: FT })]
    }),
    new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [W1, W2, W3], rows }),
  ];
}

// GUIDA sezione 7 mapping: Giorno 1 → GIORNO 1.jpg
function pagGiorno1() { return [pb(),
  img("GIORNO 1.jpg", 280, 395),
  sp(120, 0),
  ...dayHeader(1, "⚓", "CERCARE", "18 agosto 2026", "La fede è l’unica ancora di salvezza…"),
  colorBox([new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "Oggi siamo arrivati! I momenti di preghiera iniziano nel pomeriggio.", size: 20, italics: true, color: TESTO_SEC, font: FB })] })],
    GRIGIO_BG, "CCCCCC", { top: 80, bottom: 80, left: 160, right: 160 }),
  sp(120, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "⚓  I tre pezzi dell’ancora", bold: true, size: 26, color: BLU, font: FT })] }),
    bodySm("Scrivi le tue risposte — poi condividi col gruppo quello che vuoi.", 0, 60),
  ], BLU_LIGHT, BLU, { top: 120, bottom: 80, left: 200, right: 200 }),
  sp(60, 0),
  noteLines("CAMPO — Che cosa cerchi da questo camposcuola?", 4), sp(60, 0),
  noteLines("IO — Nelle scelte di quest’anno, cosa hai cercato?", 4), sp(60, 0),
  noteLines("DIO — Cosa cerchi in Dio?", 4),
  sp(120, 0), ...sera18(),
]; }

// GUIDA sezione 7 mapping: Giorno 2 → GIORNO 2.jpg
function pagGiorno2() { return [pb(),
  img("GIORNO 2.jpg", 280, 395),
  sp(120, 0),
  ...dayHeader(2, "⛰️", "ASCOLTARE", "19 agosto 2026", "Nella solitudine della montagna l’anima si purifica…"),
  sp(80, 0), ...mattino19(), sp(120, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "⛰️  Momento di deserto — Le Beatitudini (Mt 5, 1-12)", bold: true, size: 26, color: BLU, font: FT })] }),
    bodySm("Trova un posto tranquillo. Chiudi gli occhi e fai due respiri profondi. Ascolta lentamente le parole di Gesù. Poi rispondi a qualcuna di queste domande.", 0, 60),
  ], BLU_LIGHT, BLU, { top: 120, bottom: 80, left: 200, right: 200 }),
  sp(60, 0),
  noteLines("«Beati i poveri in spirito» — chi sono secondo te? Ti ci riconosci?", 3), sp(60, 0),
  noteLines("«Beati i miti» — la mitezza ti sembra coraggio o debolezza?", 3), sp(60, 0),
  noteLines("«Beati i misericordiosi» — c’è qualcuno a cui fai fatica a perdonare?", 3), sp(60, 0),
  noteLines("Chi ascolti davvero nella tua vita? Chi influenza di più le tue scelte?", 3), sp(80, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "💭  Il mio tassello", bold: true, size: 22, color: TESTO, font: FB })] }),
    bodySm("Su un pezzo di cartoncino scrivi, colora o disegna cosa ti rimane da questo momento. Poi incollalo sulla montagna del tuo gruppo.", 0, 0),
  ], GRIGIO_BG, "CCCCCC", { top: 100, bottom: 100, left: 160, right: 160 }),
  sp(120, 0), ...sera19(), sp(120, 0),
  ...placeholder("SERATA  —  19 agosto  (in definizione)"),
]; }

// GUIDA sezione 7 mapping: Giorno 3 prima parte → GIORNO 6.jpg, veglia → GIORNO 4.jpg
function pagGiorno3() { return [pb(),
  img("GIORNO 6.jpg", 280, 395),
  sp(120, 0),
  ...dayHeader(3, "🏳", "RESTARE", "20 agosto 2026", "Ogni giorno la comunione mi dà la forza di sopportare i sacrifici."),
  sp(80, 0), ...mattino20(), sp(120, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "🎵  Canzone del giorno — Resta con me", bold: true, size: 26, color: BLU, font: FT })] }),
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Bambole di Pezza — Sanremo 2026", italics: true, size: 20, color: TESTO_SEC, font: FB })] }),
    ...[
      "Volevo dirti in queste notti / Ti penso ancora",
      "Che la mia vita da quel giorno / È un’altra storia",
      "E a volte per cambiare tutto / Basta una parola", "",
      "Sono una donna che non guarda in faccia niente",
      "Mi hanno guardato male ma è il giudizio della gente", "",
      "Resta con me in questi tempi di odio",
      "Tu resta con me",
      "Anche se tutto questo ci cambierà",
      "Adesso sono io / A dirti che ho bisogno",
      "A dirti in questo posto sembra tutto una follia",
      "Resta con me / Resta con me", "",
      "Ho fatto sogni senza mai / Chiudere gli occhi",
      "Vissuto vite che non sai se immaginarti", "",
      "E allora dimmi se conosci un modo per dimenticare i guai",
      "Per noi che siamo stati sempre appesi a un filo",
      "Aspettami nell’alba di questo mattino",
      "Ho superato anni come questi / E avrei voluto dirti", "",
      "Resta con me in questi tempi di odio", "Tu resta con me…",
    ].map(l => l === "" ? sp(0, 30) : new Paragraph({
      spacing: { before: 0, after: 20 },
      indent: { left: 160 },
      children: [new TextRun({ text: l, size: 20, color: TESTO, font: FB })]
    })),
  ], GRIGIO_BG, "CCCCCC", { top: 120, bottom: 120, left: 200, right: 200 }),
  sp(60, 0), noteLines("Cosa ti rimane di questa canzone?", 3), sp(120, 0),
  ...liturgyBar("🌙", "PREGHIERA DELLA SERA  —  20 agosto"),
  ...placeholder("Testo preghiera della sera 20 agosto (in arrivo dalla commissione liturgia)"),
  sp(80, 0), img("GIORNO 4.jpg", 280, 395), sp(120, 0),
  ...veglia20(),
]; }

// GUIDA sezione 7 mapping: Giorno 4 → GIORNO 5.jpg
function pagGiorno4() { return [pb(),
  img("GIORNO 5.jpg", 280, 395),
  sp(120, 0),
  ...dayHeader(4, "🤝", "PRENDERSI CURA", "21 agosto 2026", "Preghiera, azione e sacrificio."),
  sp(80, 0), ...mattino21(), sp(120, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "🤝  Chi sono io?", bold: true, size: 26, color: BLU, font: FT })] }),
    bodySm("Presentati al gruppo attraverso questi quattro elementi.", 0, 60),
  ], BLU_LIGHT, BLU, { top: 120, bottom: 80, left: 200, right: 200 }),
  sp(60, 0),
  noteLines("🔥  Le mie PASSIONI — cosa ti piace fare, cosa ti appassiona", 3), sp(60, 0),
  noteLines("💪  I miei TALENTI — in cosa sei bravo/a (anche cose piccole!)", 3), sp(60, 0),
  noteLines("💔  Le mie MANCANZE (obbligatorio!) — una cosa in cui fai fatica o di cui hai bisogno", 3), sp(60, 0),
  noteLines("⭐  I miei VALORI — cosa è importante per te nella vita", 3),
  sp(120, 0), ...sera21(), sp(120, 0),
  colorBox([
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "📖  Spazio dediche — Pomeriggio del 21 agosto", bold: true, size: 28, color: BLU, font: FT })] }),
    bodySm("Cercare, ascoltare, restare, prendersi cura… cosa ci resta?", 0, 60),
    bodySm("Vai dai tuoi compagni, dagli educatori, da Don Michele e di’ loro quel che vuoi. Puoi scrivere qui o dirtelo a voce.", 0, 60),
  ], GRIGIO_BG, "CCCCCC", { top: 120, bottom: 80, left: 200, right: 200 }),
  sp(60, 0),
  noteLines("Da _________________ per _________________", 5), sp(60, 0),
  noteLines("Da _________________ per _________________", 5), sp(60, 0),
  noteLines("Da _________________ per _________________", 5),
]; }

// GUIDA sezione 7 mapping: Giorno 5 → GIORNO 3.jpg + GIORNO 7.jpg
function pagGiorno5() { return [pb(),
  img("GIORNO 3.jpg", 280, 395),
  sp(120, 0),
  ...dayHeader(5, "→", "ANDARE", "22 agosto 2026", "Verso l’Alto!"),
  sp(80, 0), ...mattino22(), sp(120, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "→  Verifica del campo", bold: true, size: 26, color: BLU, font: FT })] }),
    bodySm("In questi giorni, cosa ti ha colpito di più? Una tematica, una frase, un simbolo, un gesto, una conversazione…", 0, 60),
  ], BLU_LIGHT, BLU, { top: 120, bottom: 80, left: 200, right: 200 }),
  sp(60, 0), noteLines("La cosa che porterò a casa è…", 5), sp(80, 0),
  colorBox([
    new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "✏️  Il mio impegno concreto", bold: true, size: 26, color: VERDE, font: FT })] }),
    bodySm("Una cosa sola — piccola e concreta — che mi impegno a fare tornando a casa:", 0, 60),
    bodySm("Questo impegno entra in chiesa durante la Messa conclusiva. Scrivi bene!", 0, 0),
  ], VERD_BG, VERDE, { top: 120, bottom: 80, left: 200, right: 200 }),
  sp(60, 0), noteLines("Il mio impegno:", 4),
]; }

function pagFinale() { return [pb(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text: "Un campo per non dimenticare", bold: true, size: 40, color: BLU, font: FT })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: "Montecalvo Irpino  ·  18–22 agosto 2026", italics: true, size: 22, color: TESTO_SEC, font: FB })]
  }),
  noteLines("Lo spazio più importante — scrivi quello che vuoi ricordare:", 12),
  sp(120, 0),
  // "Verso l'Alto!" finale — More Sugar 18pt #7A5230 (GUIDA regola 1: NON rosso)
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 40 },
    children: [new TextRun({ text: "«Verso l’Alto!»", bold: true, italics: true, size: 44, color: MARRONE, font: FT })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: "Beato Pier Giorgio Frassati", size: 22, color: TESTO_SEC, font: FB })]
  })
]; }

// ===== ASSEMBLAGGIO =====
function flatten(arr) {
  const r = [];
  function w(x) { if (!x) return; if (Array.isArray(x)) { x.forEach(w); } else { r.push(x); } }
  w(arr); return r;
}

const children = flatten([
  pagCopertina(), pagFrassati(), pag5Giorni(),
  pagGiorno1(), pagGiorno2(), pagGiorno3(), pagGiorno4(), pagGiorno5(),
  pagFinale(),
]);

// ===== DOCUMENTO =====
// Interlinea 1.4 (276 twip) per il libretto bambini (GUIDA sezione 6)
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FB, size: 24, color: TESTO }
      }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: A4W, height: A4H },
        margin: { top: MT, bottom: MB, left: ML, right: MR }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT_FILE, buf);
  console.log('OK — Libretto Bambini 2026:', buf.length, 'bytes →', OUT_FILE);
}).catch(e => { console.error(e); process.exit(1); });
