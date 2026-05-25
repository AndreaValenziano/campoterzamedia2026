# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Progetto

Generazione programmatica di due libretti Word (.docx) per il Camposcuola Diocesano ACR Terza Media 2026 — **"Vivere, non vivacchiare!"** (18–22 agosto, Montecalvo Irpino). I libretti sono prodotti da script Node.js che usano la libreria `docx`.

## Eseguire gli script

```bash
# Installare le dipendenze (la prima volta)
npm install docx

# Generare il libretto bambini
node libretto_bambini.js

# Generare il libretto educatori
node libretto_educatori.js
```

> **Attenzione — percorsi hardcoded**: entrambi gli script contengono percorsi assoluti `/sessions/confident-hopeful-rubin/mnt/...` che devono essere aggiornati al percorso locale prima di eseguire. Cercare le costanti `IMG_DIR`, `BASE_DIR`, `OUT_FILE` all'inizio di ogni file.

## Architettura

### Script principali

- **`libretto_bambini.js`** — genera `Libretto_Bambini_2026.docx`. Grafica vivace, immagini Frassati grandi, contiene solo i momenti in cui il ragazzo partecipa attivamente (testi liturgici, domande, spazi note).
- **`libretto_educatori.js`** — genera `Libretto_Educatori_2026.docx`. Struttura funzionale, guide complete per condurre ogni momento della giornata.

Entrambi gli script seguono la **struttura cronologica per giornata** (non per sezioni tematiche): preghiera mattino → catechesi → gruppi studio → pomeriggio → preghiera sera → serata/veglia.

### Documentazione di riferimento

Prima di modificare qualsiasi script leggere:

- **`CONTESTO_CAMPO_2026.md`** — contesto completo: partecipanti, programma, arco tematico 5 giorni, mapping immagini guida → giorni campo, palette colori illustrazioni.
- **`GUIDA_FORMATTAZIONE.md`** — specifica tecnica definitiva: palette hex, gerarchia tipografica, stili Word (`Normale2`, `Titolo2`, `Enfasigrassetto1`), layout pagina, regole grafiche da rispettare.
- **`PIANO_LAVORO.md`** — stato del materiale per commissione e prossimi passi.

### Palette e font (da GUIDA_FORMATTAZIONE.md)

| Costante | Hex | Uso |
|---|---|---|
| `BLU` / `BLU_SCURO` | `#1A6FA0` | Titoli, intestazioni |
| `MARRONE` | `#7A5230` | Citazioni Frassati, "Verso l'Alto!" |
| `VERDE` | `#5A9B4A` | Accenti secondari |
| `BLU_LIGHT` | `#D5E8F0` | Sfondo box istruzioni |
| `MARR_BG` | `#F5EDE0` | Sfondo citazioni |
| `VERD_BG` | `#E8F5E4` | Sfondo attività ragazzi |

Font: **More Sugar** (titoli sezioni), **Badger Script Bold** (solo titolo copertina), **Candara** (corpo testo). More Sugar e Badger Script non sono standard — devono essere installati prima della stampa finale.

### Immagini

Tutte le illustrazioni flat design di Frassati sono in `materiale guida/`. Il mapping file → giornata campo è documentato in `GUIDA_FORMATTAZIONE.md` sezione 7 e `CONTESTO_CAMPO_2026.md` sezione 7.

### Regole grafiche critiche

1. **Non usare rosso** — "Verso l'Alto!" va sempre in marrone `#7A5230`.
2. **Box colorati**: usare `ShadingType.CLEAR`, non `SOLID`.
3. **Tabelle**: sempre `columnWidths` e `width` in DXA — mai percentuali.
4. **Citazioni Frassati**: virgolette tipografiche (`"..."`) e corsivo.
5. **Badger Script** solo per il titolo principale di copertina, non mischiare con More Sugar.

### Materiale mancante (segnaposti attivi)

I segnaposti nel documento hanno il formato `⏳ [SEGNAPOSTO — SEZIONE]`. Materiale ancora da integrare:
- Testi canti (iniziale/finale per ogni preghiera)
- Serata 19/08 (commissione serate non ha ancora inviato)
- Testi completi veglia (parziali)
