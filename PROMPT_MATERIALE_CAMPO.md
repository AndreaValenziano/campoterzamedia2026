# Prompt per generare il "Documento Materiale" — Campo Terza Media 2026

Il documento finale (`Materiale_Campo_2026.docx`) raccoglie **tutto il materiale operativo**:
cosa serve, chi lo prepara, come si usa. È diverso dal Libretto Educatori (che spiega *come
condurre*): qui si risponde a **"cosa devo mettere in valigia / stampare / comprare"**.

Lavoro in **5 sessioni separate**: 4 di raccolta (una per commissione) + 1 di unificazione.

---

## 0. Blocco comune — incollare all'inizio di OGNI sessione

```
CONTESTO CAMPO
- Camposcuola Diocesano ACR Terza Media 2026 — "Vivere, non vivacchiare!"
- 18–22 agosto 2026, Montecalvo Irpino. Figura ispiratrice: San Pier Giorgio Frassati.
- 93 persone totali: 73 ragazzi, 16 educatori, 4 cuochi, 1 assistente.
- I ragazzi sono divisi in 6 GRUPPI STUDIO fissi (~12 ragazzi ciascuno).
  Gli STESSI 6 gruppi si usano anche per i giochi del pomeriggio e per le serate.
- Arco tematico:
  G1 18/08 CERCARE (ancora) | G2 19/08 ASCOLTARE (montagna) | G3 20/08 RESTARE (bandiera AC)
  G4 21/08 PRENDERSI CURA (abbraccio) | G5 22/08 ANDARE (pagina bianca)
- Riferimenti in repo: CONTESTO_CAMPO_2026.md (programma e orari completi),
  GUIDA_FORMATTAZIONE.md (palette e stili), Libretto_Educatori_2026.docx (testo unificato).

OBIETTIVO GENERALE
Sto costruendo il "Documento Materiale": l'elenco completo e operativo di tutto ciò che
serve per il campo, organizzato PER GIORNATA. Questa sessione riguarda una sola commissione.
Il risultato di questa sessione è un file Markdown intermedio, NON il documento finale.

REGOLE
- Italiano, linguaggio semplice da contesto parrocchiale.
- Dimensionare SEMPRE le quantità su 73 ragazzi / 6 gruppi (~12 a gruppo) e dire il calcolo,
  es. "6 cartelloni (1 per gruppo)", "80 copie (73 ragazzi + scorte)".
- Non inventare: se un dato manca nel .docx della commissione, scrivere
  `⏳ [DA CHIEDERE ALLA COMMISSIONE — <cosa>]` invece di riempire a caso.
- Distinguere sempre: materiale DA COMPRARE / DA STAMPARE / GIÀ DISPONIBILE (da portare).
- Segnalare il materiale riutilizzato in più giornate (va comprato una volta sola).
```

---

## 1. Sessione LITURGIA

```
<blocco comune>

Leggi "materiale commissioni/liturgia/Campo Medie 2026 - Commissione Liturgia.docx".
Estrai TUTTO il materiale necessario per i momenti di preghiera (mattino e sera),
la veglia del 20/08 e la Messa conclusiva del 22/08.

Per ogni momento indica:
- Giorno e fascia oraria (es. G2 19/08 — preghiera del mattino, 8:30)
- Materiale liturgico: candele, croce, icone, teli, segni del giorno, cesti, incenso...
- Fogli da stampare: quanti e per chi (ragazzi / educatori / animatore del momento)
- Testi dei canti: se assenti, segnaposto
- Chi prepara e cosa va allestito prima (es. "spazio preghiera montato la sera prima")
- Serve un impianto audio / chitarra / proiettore?

Output: scrivi il file "materiale/01_liturgia.md" con una sezione per giornata
(G1…G5) e, dentro ciascuna, i momenti in ordine cronologico. Chiudi con una
tabella riepilogativa "Materiale liturgia — totale campo".
```

---

## 2. Sessione GRUPPI STUDIO

```
<blocco comune>

Leggi "materiale commissioni/gruppi studio/Campo Medie 2026 - Commissione Gruppi Studio.docx"
e "guida_educatori_gruppi_studio.md".

Per ciascuno dei 5 gruppi studio (uno al giorno, G1 CERCARE → G5 ANDARE) indica:
- Materiale per gruppo × 6 gruppi (cartelloni, pennarelli, colla, forbici, scotch...)
- Materiale individuale × 73 ragazzi (schede, biglietti, post-it, oggetti simbolici)
- Schede/allegati da stampare: nome, quantità, fronte-retro sì/no, colore o B/N
- Oggetti particolari da procurare in anticipo (es. pezzi dell'ancora del G1)
- Spazi necessari: 6 spazi separati? interni o esterni? servono sedie/tavoli?
- Cosa resta appeso o conservato per i giorni successivi

Output: scrivi il file "materiale/02_gruppi_studio.md", una sezione per giornata,
più una tabella finale "Kit gruppo studio" (il kit base identico per ognuno dei 6 gruppi).
```

---

## 3. Sessione POMERIGGI

```
<blocco comune>

Leggi "materiale commissioni/pomeriggi/Campo Medie 2026 - Commissione Pomeriggi.docx".

Le attività pomeridiane (15:30) si giocano con gli STESSI 6 gruppi studio.
Per ogni pomeriggio indica:
- Nome del gioco/attività e giornata
- Materiale per postazione/prova, moltiplicato per il numero di postazioni
- Materiale di consumo (palloncini, farina, acqua, gessetti, sacchi...) con quantità
- Materiale da recuperare/riusare, e cosa si rovina e va comprato in eccesso
- Segnaposti, cartelli, tabelloni punteggi, fischietto, cronometro
- Spazi: campo, prato, zone d'ombra; piano B in caso di pioggia e relativo materiale
- Educatori necessari per gestire l'attività (arbitri, postazioni)

Output: scrivi il file "materiale/03_pomeriggi.md", una sezione per giornata,
più una lista "Da comprare al supermercato/cartoleria" con quantità aggregate.
```

---

## 4. Sessione SERATE

```
<blocco comune>

Leggi "materiale commissioni/serate/Campo Medie 2026 - Commissione Serate.docx".

Serate alle 21:30. Il 20/08 al posto della serata c'è la VEGLIA (materiale già coperto
dalla commissione liturgia: qui indica solo cosa serve in più lato scenografia/tecnica).
Anche le serate usano i 6 gruppi.

Per ogni serata indica:
- Titolo e giornata
- Materiale scenografico e costumi
- Tecnica: audio, microfoni, casse, proiettore, luci, prolunghe, playlist/tracce
- Materiale per i ragazzi (bigliettini, cartelli, oggetti di scena per gruppo)
- Premi/gadget se previsti
- Allestimento: quando e chi lo monta, quanto tempo serve

Output: scrivi il file "materiale/04_serate.md", una sezione per giornata,
più una tabella "Materiale tecnico — da portare da casa" con indicato chi lo porta.
```

---

## 5. Sessione FINALE — unificazione

```
<blocco comune>

Ho completato le 4 sessioni di raccolta. Leggi:
  materiale/01_liturgia.md
  materiale/02_gruppi_studio.md
  materiale/03_pomeriggi.md
  materiale/04_serate.md

Genera il documento unificato "Materiale_Campo_2026.docx" con uno script Node.js
"materiale_campo.js" che usa la libreria `docx`, seguendo GUIDA_FORMATTAZIONE.md
(palette BLU #1A6FA0, MARRONE #7A5230, VERDE #5A9B4A; font Candara per il corpo,
More Sugar per i titoli; box con ShadingType.CLEAR; tabelle con columnWidths in DXA,
mai percentuali; niente rosso).

STRUTTURA DEL DOCUMENTO — organizzato PER GIORNATA, non per commissione:

  Copertina
  1. Come si usa questo documento + legenda (🛒 comprare, 🖨️ stampare, 📦 portare)
  2. Elenco dei 6 gruppi studio e loro uso trasversale (studio + pomeriggi + serate)
  3. GIORNO 1 — 18/08 CERCARE
       3.1 Preghiera del mattino/sera   [liturgia]
       3.2 Gruppi studio                [gruppi studio]
       3.3 Pomeriggio                   [pomeriggi]
       3.4 Serata                       [serate]
       3.5 Riepilogo materiale del giorno (tabella unica)
  4. GIORNO 2 — 19/08 ASCOLTARE   (stessa sotto-struttura)
  5. GIORNO 3 — 20/08 RESTARE     (la serata è la Veglia)
  6. GIORNO 4 — 21/08 PRENDERSI CURA
  7. GIORNO 5 — 22/08 ANDARE      (include la Messa conclusiva)
  8. LISTA DELLA SPESA COMPLESSIVA — aggregata su tutto il campo, deduplicata,
     divisa per: cartoleria / supermercato / ferramenta / materiale liturgico / tecnica
  9. LISTA STAMPE — tabella: documento, quantità, B/N o colore, fronte-retro, giorno d'uso
  10. MATERIALE DA PORTARE DA CASA — con colonna "chi lo porta"
  11. SEGNAPOSTI APERTI — tutti i ⏳ raccolti, con la commissione a cui chiederli

REGOLE DI UNIFICAZIONE
- Deduplica: se lo stesso materiale serve in più giornate, resta nella giornata
  (con nota "già presente dal G2") ma nella lista della spesa compare UNA volta sola,
  con la quantità massima necessaria.
- Non riscrivere le attività: qui serve solo il materiale, non le istruzioni di conduzione
  (quelle stanno nel Libretto Educatori).
- Prima di scrivere lo script aggiorna i percorsi hardcoded (IMG_DIR, BASE_DIR, OUT_FILE)
  al percorso locale.
- A fine generazione stampa un riepilogo: quante voci di spesa, quante stampe,
  quanti segnaposti aperti.
```

---

## Note operative

- Crea la cartella `materiale/` prima della sessione 1.
- Le 4 sessioni di raccolta sono indipendenti: si possono fare in qualsiasi ordine
  o in parallelo.
- Se una commissione consegna materiale nuovo, si rigenera solo il suo `.md`
  e si rilancia la sessione 5.
