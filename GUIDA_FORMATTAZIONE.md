# Guida alla Formattazione — Libretti Campo 2026
## "Vivere, non vivacchiare!" · Camposcuola Diocesano ACR · Terza Media

> Documento di riferimento per chiunque aggiunga contenuto ai libretti. Seguire queste specifiche garantisce coerenza grafica tra tutte le pagine e tra i due documenti.

---

## 1. Palette Colori

| Ruolo | Nome | Codice Hex | Uso |
|---|---|---|---|
| **Blu primario** | Blu ACR | `#1A6FA0` | Titoli sezioni, intestazioni, elementi principali |
| **Marrone** | Terra Frassati | `#7A5230` | Citazioni di Frassati, sottotitoli, motto "Verso l'Alto!" |
| **Verde** | Verde montagna | `#5A9B4A` | Accenti secondari, box verdi, elementi natura |
| **Grigio testo** | Grigio scuro | `#333333` | Corpo testo normale (applicato via font Candara) |
| **Grigio neutro** | Grigio medio | `#555555` | Info secondarie, note, didascalie |
| **Sfondo box blu** | Azzurro chiaro | `#D5E8F0` | Sfondo box informativi / istruzioni educatore |
| **Sfondo box verde** | Verde chiaro | `#E8F5E4` | Sfondo box attività / schede ragazzi |
| **Sfondo box marrone** | Beige caldo | `#F5EDE0` | Sfondo citazioni Frassati |

---

## 2. Tipografia — Font

### Font principali in uso

| Font | Ruolo | Disponibilità |
|---|---|---|
| **More Sugar** | Titoli sezioni, enfasi calligrafica | Da installare (Google Fonts / file) |
| **Badger Script Bold** | Solo titolo copertina | Da installare (file) |
| **Candara** | Corpo testo, paragrafi | Preinstallato su macOS/Windows |

> ⚠️ **Importante:** More Sugar e Badger Script sono font non standard. Chi apre il documento su un PC senza questi font li vedrà sostituiti. Prima della stampa finale, verificare che il documento venga aperto su una macchina con i font installati, oppure esportare in PDF.

### Gerarchia tipografica completa

| Elemento | Font | Dimensione | Stile | Colore |
|---|---|---|---|---|
| **Titolo copertina** ("Vivere, non vivacchiare!") | Badger Script Bold | 72 pt | Bold | `#1A6FA0` |
| **Sottotitolo copertina** (nome santo, info campo) | More Sugar | 14 pt | Normal | `#7A5230` / `#555555` |
| **Box nome/parrocchia/gruppo** (copertina) | More Sugar | 11 pt | Normal | `#333333` |
| **Titolo giornata / sezione principale** | More Sugar | 24 pt | Normal | `#1A6FA0` |
| **Citazione Frassati** (testo breve in evidenza) | More Sugar | 12 pt | Normal | `#1A6FA0` |
| **Corpo testo** (paragrafi normali) | Candara | 12 pt | Normal | `#333333` |
| **Motto "Verso l'Alto!"** | More Sugar | 18 pt | Normal | `#7A5230` |
| **Testo liturgico C: / T:** | Candara | 12 pt | Normal | `#333333` |
| **Testo salmo / scrittura** | Candara | 11 pt | Italic | `#555555` |
| **Note operative (solo educatori)** | Candara | 10.5 pt | Normal | `#555555` |

---

## 3. Stili Word definiti nel documento

Questi stili sono già stati creati/modificati nel file `.docx` e si applicano selezionando il testo e scegliendo lo stile dal pannello Stili di Word.

| ID Stile (interno) | Nome visibile | Definizione | Dove usarlo |
|---|---|---|---|
| `Normale2` | Normale (corpo) | Candara 12pt, colore automatico | Tutto il corpo testo |
| `Titolo2` | Titolo sezione | More Sugar 24pt, `#1A6FA0`, spaziatura dopo: 120 | Titoli giornata e sezioni principali |
| `Enfasigrassetto1` | Citazione/Enfasi | More Sugar 12pt, `#1A6FA0`, basato su Normale | Citazioni di Frassati in evidenza |

> **Come applicare uno stile:** selezionare il paragrafo → Scheda Home → pannello Stili → cliccare il nome dello stile.

---

## 4. Struttura pagine — Libretto Bambini

### Pagina 1 — Copertina (immagine full-page)
- Immagine `COPERTINA LIBRETTO BAMBINI.png` ancorata a tutta la pagina
- Nessun testo su questa pagina

### Pagina 2 — Titolo
- **"Vivere, non vivacchiare!"** → Badger Script Bold, 72pt, `#1A6FA0`, centrato
- **"San Pier Giorgio Frassati"** → More Sugar, 14pt, `#7A5230`, centrato
- **Info campo** (luogo, date, organizzazione) → More Sugar, 12pt, `#555555`, centrato
- **Box nome/parrocchia/gruppo** → More Sugar, 11pt, bordo leggero

### Pagine successive — Frassati & Introduzione
- Titolo ("Chi era Pier Giorgio Frassati?") → Stile `Titolo2` (More Sugar 24pt, `#1A6FA0`)
- Corpo testo → Stile `Normale2` (Candara 12pt)
- **"Verso l'Alto!"** → More Sugar, 18pt, `#7A5230` (marrone, NON rosso)

### Pagina cammino dei 5 giorni
- Titolo ("Il cammino dei 5 giorni") → Stile `Titolo2`
- Tabella 5 verbi → Candara, colori palette

### Sezioni giornata (Giorno 1–5)
- **Immagine giornata**: posizionata in alto, larghezza intera colonna (~40% altezza pagina)
- **"Giorno N — VERBO"** → Stile `Titolo2` (More Sugar 24pt, `#1A6FA0`)
- **Data** → Candara 11pt, `#555555`, italic
- **Citazione Frassati** → Stile `Enfasigrassetto1` (More Sugar 12pt, `#1A6FA0`), in corsivo, preceduta da `>`
- **Testi liturgici** (C: / T:) → Candara 12pt
- **Salmi e brani biblici** → Candara 11pt italic, `#555555`
- **Spazio note personali** → riga tratteggiata, etichetta in More Sugar 11pt

---

## 5. Struttura pagine — Libretto Educatori

### Copertina
- Stessa immagine della copertina bambini (versione tinta/più sobria se possibile)
- Titolo → Badger Script Bold, 60pt, `#1A6FA0`
- Sottotitolo "Guida Educatori" → More Sugar, 18pt, `#7A5230`

### Pagine operative
- Titoli sezione → Stile `Titolo2`
- Corpo testo → Stile `Normale2` (Candara 12pt)
- **Box istruzioni operative** (sfondo `#D5E8F0`) → Candara 11pt, bordo sinistro `#1A6FA0` 3pt
- **Script C: / T:** → Candara 12pt, `C:` in bold `#1A6FA0`, `T:` in normale
- **Note per l'educatore** → Candara 10.5pt italic, `#555555`
- Tabelle programma → righe alternate bianco / `#F0F7FB`

---

## 6. Impostazioni pagina

| Parametro | Libretto Bambini | Libretto Educatori |
|---|---|---|
| **Formato** | A4 (11906 × 16838 DXA) | A4 (11906 × 16838 DXA) |
| **Margine superiore** | 2 cm (1134 DXA) | 2.5 cm (1418 DXA) |
| **Margine inferiore** | 2 cm (1134 DXA) | 2.5 cm (1418 DXA) |
| **Margine sinistro** | 2.5 cm (1418 DXA) | 3 cm (1701 DXA) |
| **Margine destro** | 2 cm (1134 DXA) | 2.5 cm (1418 DXA) |
| **Interlinea** | 1.4 (276 twip) | 1.3 (260 twip) |

---

## 7. Mapping immagini per giornata

| Giornata | File immagine | Tema visivo |
|---|---|---|
| Copertina | `COPERTINA.jpg` | Immagine generale campo |
| Giorno 1 — CERCARE | `GIORNO 1.jpg` | Ricerca / ancora |
| Giorno 2 — ASCOLTARE | `GIORNO 2.jpg` | Ascolto / silenzio |
| Giorno 3 — RESTARE (prima parte) | `GIORNO 6.jpg` | Comunità / restare |
| Giorno 3 — Veglia | `GIORNO 4.jpg` | Ostensorio / adorazione |
| Giorno 4 — PRENDERSI CURA | `GIORNO 5.jpg` | Cura / solidarietà |
| Giorno 5 — ANDARE (doppia) | `GIORNO 3.jpg` + `GIORNO 7.jpg` | Partenza / missione |

Tutte le immagini si trovano in `materiale guida/`.

---

## 8. Regole grafiche da rispettare

1. **Non usare rosso** (`#C00000` o simili) — riservato solo a eventuali avvisi critici. Il motto "Verso l'Alto!" va in marrone `#7A5230`.
2. **Non mischiare Badger Script con More Sugar** nello stesso blocco di testo — Badger Script è solo per il titolo principale di copertina.
3. **Le citazioni di Frassati** vanno sempre con virgolette tipografiche (`"..."`) e in corsivo.
4. **I box colorati** usano `ShadingType.CLEAR` nel codice JS, non `SOLID` (evita sfondi neri in alcuni visualizzatori).
5. **Le tabelle** hanno sempre `columnWidths` e `width` su ogni cella in DXA — mai percentuali.
6. **I canti** (`🎵 Canto iniziale / finale`) rimangono come segnaposto con riga tratteggiata finché non arrivano i testi.

---

## 9. Segnaposti attivi (materiale mancante)

| Segnaposto | Dove | Stato |
|---|---|---|
| Testi canti (iniziale/finale ogni preghiera) | Entrambi i libretti | ⏳ Non pervenuti |
| Serata 19/08 | Libretto Educatori | ⚠️ Mancante dalla commissione serate |
| Testi veglia completi | Libretto Bambini + Educatori | ⏳ Parziali |
| Attività pomeridiane guidate | Libretto Educatori | ⏳ Non pervenute |

---

*Guida aggiornata al 15 maggio 2026 — da aggiornare se cambiano font o palette*
