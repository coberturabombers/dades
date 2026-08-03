# Cobertura Bombers Catalunya — Paquet complet

Eina web que mostra la cobertura diària dels parcs de bombers de Catalunya
(mínims de torn vs efectius reals), amb mapa, detall per regió, línia temporal
i generació de resums per compartir. Feta per a l'Assemblea de Bombers de Catalunya.

Aquest paquet està pensat perquè el desplegament sigui gairebé "clau en mà":
tot el codi hi és; només cal connectar la font de dades (l'Excel) i publicar.

---

## Com funciona (arquitectura)

Dues peces, totes dues gratuïtes:

1. **El motor diari (GitHub Actions).** Un cop al dia, un procés automàtic:
   - baixa l'Excel del TRI des d'un enllaç de GitHub (carpeta entrada/),
   - n'extreu NOMÉS les dades de cobertura (parc, mínim, reals) amb `scripts/parse_tri.py`,
   - desa la "foto" del dia a `data/history.json`,
   - fa commit al repositori.

2. **La cara visible (Netlify).** Un web estàtic (`site/index.html`) que llegeix
   `data/history.json` i el mostra: mapa, KPIs, taula per regió, línia temporal i
   resums per compartir. Netlify es republica sol cada cop que el motor fa commit.

```
GitHub (carpeta entrada/) (Excel TRI)  ──▶  GitHub Actions (parse_tri.py)  ──▶  data/history.json  ──▶  Netlify (web)
        ▲ s'actualitza          ▲ cada dia a l'hora fixada          ▲ commit                ▲ es republica sol
          durant el dia
```

**Per què aquesta arquitectura:** un web estàtic no pot, tot sol, ni executar
tasques programades ni desar històric. GitHub Actions aporta el "cada dia a les X"
i el repositori Git fa de base de dades senzilla (l'històric és un fitxer JSON
versionat). Tot gratuït dins dels límits habituals d'un projecte com aquest.

---

## Estructura del projecte

```
coberturabombers/
├── scripts/
│   └── parse_tri.py                 # Motor: llegeix l'Excel i genera les dades
├── data/
│   ├── history.json                 # Històric de tots els dies (base de dades)
│   └── latest.json                  # Últim dia (còpia de comoditat)
├── site/
│   └── index.html                   # El dashboard web
├── .github/workflows/
│   └── actualitza-tri.yml           # Tasca diària automàtica
├── netlify.toml                     # Configuració de Netlify
├── .gitignore
└── README.md                        # Aquest fitxer
```

---

## Desplegament pas a pas

### Requisit previ: l'enllaç de l'Excel

Pujar l'Excel del TRI ("Personal de guàrdia") a GitHub (carpeta entrada/) i obtenir un
**enllaç de descàrrega directa**. A GitHub (carpeta entrada/):
1. Comparteix el fitxer → "Qualsevol persona amb l'enllaç pot veure".
2. Converteix l'enllaç en descàrrega directa. Un enllaç de GitHub (carpeta entrada/) del tipus
   `https://onedrive.live.com/...` es pot transformar afegint `&download=1`, o bé
   fent servir el format `https://api.onedrive.com/v1.0/shares/<token>/root/content`.
   (El desenvolupador ho pot ajustar; el que importa és que l'enllaç retorni el
   `.xlsx` directament, no una pàgina web.)

> Nota de privadesa: l'Excel conté molta informació interna (comandaments, menús,
> vehicles…). El parser **només** llegeix parc/mínim/reals i descarta la resta, així
> que res sensible arriba mai al web. Tot i això, l'enllaç de GitHub (carpeta entrada/) dona accés al
> fitxer sencer: val la pena valorar compartir-lo de manera restringida, o pujar a
> GitHub (carpeta entrada/) una còpia que només contingui el full "Dades".

### 1. Puja el projecte a GitHub

```bash
cd coberturabombers
git init
git add .
git commit -m "Primera versió"
git branch -M main
git remote add origin https://github.com/<usuari>/coberturabombers.git
git push -u origin main
```

### 2. Configura el secret amb l'enllaç de l'Excel

A GitHub: **Settings → Secrets and variables → Actions → New repository secret**
- Nom: `ENLLAC_TRI`
- Valor: l'enllaç de descàrrega directa de l'Excel a GitHub (carpeta entrada/).

### 3. Activa i prova el motor diari

- A la pestanya **Actions** del repositori, activa els workflows si cal.
- Executa'l manualment un cop: workflow **"Actualitza dades TRI cada dia"** →
  **Run workflow**. Comprova que `data/history.json` s'actualitza amb el dia d'avui.
- A partir d'aquí, s'executarà sol cada dia (per defecte a les ~10:10 hora de
  Barcelona; ajusta el `cron` a `.github/workflows/actualitza-tri.yml` si vols
  una altra hora).

### 4. Publica el web a Netlify

- A Netlify: **Add new site → Import an existing project → GitHub** → tria el repositori.
- Build command: (buit). Publish directory: `.` (arrel). El `netlify.toml` ja ho configura.
- Netlify et donarà una URL tipus `https://<nom>.netlify.app`.

### 5. (Opcional) Domini propi

- A Netlify: **Domain settings → Add a custom domain** → `coberturabombers.com`
  (o el que tingueu). Netlify guia la configuració del DNS. El domini es compra a
  part (uns 10–15 €/any un `.com`; un `.cat` una mica més).

---

## Prova en local (sense desplegar res)

```bash
pip install openpyxl
python scripts/parse_tri.py --input "Personal_de_guàrdia.xlsx"   # genera data/history.json
cd . && python -m http.server 8000                                # obre http://localhost:8000/site/
```

---

## Manteniment i coses a tenir en compte

- **Si canvia l'estructura de l'Excel**, el parser té una comprovació de sumes de
  control (els mínims han de sumar 354). Si no quadra, avisa amb un warning: cal
  revisar `REGION_START_COLS` a `scripts/parse_tri.py`.
- **Coordenades aproximades**: 3 parcs (Puig-reig, Barcelona-Gros, el Papiol) tenen
  ubicació aproximada. Es poden afinar a `PARC_COORDS` dins del parser.
- **Hora del cron en UTC**: el workflow usa UTC. A l'estiu (CEST) resta 2 h; a
  l'hivern (CET) resta 1 h. Ajusta si vols una hora local exacta tot l'any.
- **Dades històriques de mostra**: `data/history.json` ve amb uns dies d'exemple
  (20/07 i 25/07–02/08) perquè el web tingui contingut des del primer moment.
  Es poden esborrar quan el motor comenci a omplir dades reals.

---

## Resum de decisions ja preses

- Font de dades: **Excel del TRI a GitHub (carpeta entrada/)** (una sola font, 7 regions completes,
  s'actualitza sola). Descartats el PDF (no fiable) i els Google Sheets (3 regions
  sense accés públic).
- Es desa **només** parc + mínim + reals. Cap dada personal ni operativa.
- Els mínims són fixos; cada dia només canvien els reals.
