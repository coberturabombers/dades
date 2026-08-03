# INSTRUCTIU A PROVA DE TONTOS
## Com posar en marxa el web de Cobertura de Bombers i fer la prova de 5 dies

Aquest document et porta de la mà, pas a pas. No cal saber programar.
Es divideix en 4 parts:

- **PART A** — Preparar-ho tot (GitHub + Netlify). Es fa un sol cop. ~20 minuts.
- **PART B** — La prova de 5 dies (carregar Excels a mà).
- **PART C** — Comprovar que quadra amb el TRI en PDF del Telegram.
- **PART D** — (Opcional) Activar la càrrega automàtica diària.

Necessitaràs: un compte de GitHub (gratuït) i un de Netlify (gratuït). Els pots
crear amb el teu correu en 2 minuts cadascun.

---

# PART A — Preparar-ho tot (un sol cop)

## A1. Instal·la Python al teu ordinador (si no el tens)

El necessitem per llegir els Excels a la prova manual.

- **Windows:** ves a https://www.python.org/downloads/ → botó groc "Download Python".
  Executa l'instal·lador i **MOLT IMPORTANT: marca la casella "Add Python to PATH"**
  a baix de tot abans de clicar Install.
- **Mac:** normalment ja el porta. Si no, descarrega'l del mateix enllaç.

Per comprovar que està: obre el **Terminal** (Mac) o **Símbol del sistema / CMD**
(Windows) i escriu:
```
python --version
```
Ha de sortir un número (per exemple `Python 3.12`). Si a Windows no funciona, prova `py --version`.

Després, instal·la la llibreria que llegeix Excels. Al mateix terminal:
```
pip install openpyxl
```

## A2. Crea un compte a GitHub

1. Ves a https://github.com i registra't (Sign up) amb el teu correu.
2. Un cop dins, clica el **+** a dalt a la dreta → **New repository**.
3. Nom del repositori: `coberturabombers` (o el que vulguis).
4. Deixa'l **Public**. No marquis res més.
5. Clica **Create repository**. Deixa aquesta pàgina oberta, la faràs servir al pas A4.

## A3. Descomprimeix el paquet

Descomprimeix `coberturabombers_paquet.zip` en una carpeta del teu ordinador que
recordis (per exemple a l'Escriptori). Tindràs una carpeta `coberturabombers` amb
tot a dins.

## A4. Puja el projecte a GitHub (la manera fàcil, sense comandes)

GitHub deixa pujar fitxers arrossegant, sense escriure res:

1. A la pàgina del teu repositori nou (la del pas A2), busca l'enllaç
   **"uploading an existing file"** (surt al mig de la pàgina quan el repo és buit).
2. **Obre la carpeta `coberturabombers`** al teu ordinador.
3. **Selecciona tot el que hi ha dins** (no la carpeta, sinó el seu contingut:
   les carpetes `scripts`, `data`, `site`, `.github`, i els fitxers `README.md`,
   `netlify.toml`, `.gitignore`) i **arrossega-ho** a la finestra de GitHub.

   > ⚠️ Si la carpeta `.github` no es deixa arrossegar (de vegades Windows amaga les
   > carpetes que comencen per punt), no pateixis: acaba la resta i mira la NOTA al
   > final d'aquesta part per pujar-la a part. És la que fa l'automatisme (PART D),
   > no la necessites per a la prova manual.

4. A baix, on diu **"Commit changes"**, clica el botó verd **Commit changes**.
5. Espera que es pugin. Quan acabi, veuràs les carpetes al repositori. ✅

## A5. Publica el web a Netlify

1. Ves a https://www.netlify.com i registra't (Sign up) — tria **"Sign up with GitHub"**,
   així queda tot connectat.
2. Un cop dins, clica **Add new site → Import an existing project**.
3. Tria **GitHub** i autoritza. Busca i selecciona el repositori `coberturabombers`.
4. Et demanarà la configuració. **No toquis res** (el fitxer `netlify.toml` ja ho
   configura tot): deixa el "Build command" buit i el "Publish directory" com estigui.
5. Clica **Deploy**.
6. Al cap d'un minut, Netlify et donarà una adreça tipus
   `https://un-nom-aleatori.netlify.app`. **Obre-la.**

Hauries de veure el web amb el missatge "Encara no hi ha dades disponibles" (normal:
encara no hem carregat cap dia). Si ho veus, **tot funciona.** 🎉

> Pots canviar el nom aleatori per un de més bonic a Netlify:
> **Site configuration → Change site name.**

**NOTA (només si no vas poder pujar la carpeta `.github` al pas A4):**
Al teu repositori de GitHub, clica **Add file → Create new file**. Al camp del nom,
escriu exactament: `.github/workflows/actualitza-tri.yml` (GitHub crearà les carpetes
soles). Copia-hi el contingut del fitxer del mateix nom del paquet, i clica
**Commit changes**.

---

# PART B — Carregar les dades (pujar l'Excel al repositori)

Amb aquest sistema no cal executar res al teu ordinador. Només cal pujar l'Excel.

## B1. Prepara l'Excel del dia
Agafa l'Excel del TRI del dia. No cal reanomenar-lo: el sistema accepta qualsevol fitxer .xlsx.

## B2. Puja'l al repositori
1. Al repositori de GitHub, entra a la carpeta **`entrada`**.
2. **Add file → Upload files** → arrossega l'Excel → **Commit changes**.
   (Cada dia el substituiràs pel nou; és el que volem.)

## B3. El sistema fa la resta sol
En pujar l'Excel, s'engega la tasca automàtica (pestanya **Actions**): llegeix el
fitxer, fa la foto del dia i actualitza `data/history.json`. Netlify republica el
web sol en un minut. Refresca el teu `.netlify.app` i hi surt el dia, amb el mapa ple. ✅

> Alternativa per a proves ràpides al teu ordinador (opcional, cal Python):
> `python scripts/parse_tri.py --input entrada/Personal de guàrdia.xlsx --date 2026-08-02`


---

# PART C — Comprovar que quadra amb el TRI en PDF del Telegram

Aquesta és la prova de foc: que les dades del web coincideixin amb el TRI oficial.

Per a cada dia de la prova:
1. Obre el web i selecciona aquell dia amb el selector de dates.
2. Mira els números de dalt (parcs sota mínims, tancats, diferència).
3. Obre el PDF del TRI d'aquell mateix dia del grup de Telegram.
4. Compara:
   - **Parcs sota mínims** del web ↔ el que digui el TRI.
   - **Parcs tancats** ↔ TRI.
   - Pots entrar al detall d'un parc concret al mapa (clica'l) i comprovar que el
     seu mínim i reals coincideixen amb el PDF.

Si quadra → el sistema funciona i podeu confiar-hi. Si algun número balla, apunta
quin dia i quin parc, i m'ho dius: segurament serà un detall d'alguna columna que
puc ajustar al parser.

> Recorda: si l'Excel que fas servir és una **previsió** (generada dies abans),
> pot no coincidir amb el TRI **real** d'aquell dia. Per a la comparació, fes servir
> l'Excel tal com estava **el mateix dia** que el PDF.

---

# PART D — Com funciona la càrrega diària (tot dins de GitHub)

Amb aquest sistema NO cal cap Drive ni cap enllaç extern. L'Excel viu dins del
mateix repositori, a la carpeta `entrada` (qualsevol nom .xlsx serveix).

**Cada dia:**
1. Agafa l'Excel del TRI del dia (amb el seu nom original, per exemple `Personal de guàrdia.xlsx`).
2. Al repositori de GitHub, entra a la carpeta **`entrada`** → **Add file →
   Upload files** → arrossega'l → **Commit changes**.
3. En pujar-lo, el sistema s'engega SOL: llegeix l'Excel, fa la foto del dia i
   actualitza les dades. Ho veus a la pestanya **Actions** (cercle groc → ✅ verd).
4. Netlify republica el web sol. Refresca i hi surt el dia nou.

**Si vols que es faci sol a una hora fixa** (sense pujar-lo tu), el sistema també
té programada una execució diària (~10:10 h de Barcelona). Però com que l'Excel el
tens al teu ordinador, per a la prova el més pràctic és pujar-lo tu cada dia.

> Per a la prova de 5 dies detallada, segueix el document **PLA_PROVA_5_DIES.md**.


---

# Si alguna cosa no va

- **El web diu "Encara no hi ha dades":** encara no has pujat el `history.json` amb
  dades (PART B3), o el fitxer és buit `{}`.
- **`python` no es reconeix (Windows):** prova `py` en lloc de `python`.
- **`pip install openpyxl` falla:** prova `python -m pip install openpyxl`.
- **Un número no quadra amb el PDF:** apunta dia + parc i m'ho passes.
- **Qualsevol pas se't fa bola:** fes-me captura de pantalla i t'ho desencallo.

Sort amb la prova! 💪
