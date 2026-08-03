# PLA DE PROVA REAL — 5 dies seguits
## Cobertura Bombers Catalunya (automàtic des del Google Drive)

L'objectiu és validar que cada dia el sistema fa SOL la "foto" de la cobertura i
que coincideix amb el TRI en PDF del Telegram.

Idea clau: **l'app consulta sola el fitxer del Drive a una hora fixa cada dia.**
Tu no has de pujar res al web cada dia. L'únic que ha d'estar actualitzat és la
còpia neta al Drive.

---

## Com queda el flux (un cop muntat)

```
El teu PC:  Excel del TRI  ──(script neteja)──▶  cobertura_neta.xlsx
                                                       │  (el puges/sobreescrius al Drive)
                                                       ▼
Google Drive de l'Assemblea:  cobertura_neta.xlsx  (enllaç fix, sempre el mateix)
                                                       │
                                                       ▼  cada dia a les 10:10, sol
GitHub Actions  ──llegeix el Drive──▶  fa la foto  ──▶  data/history.json
                                                       │
                                                       ▼
Netlify  ──republica el web sol──▶  el mapa i la línia temporal s'actualitzen
```

---

## IMPORTANT — Privadesa (llegeix això primer)

L'Excel original conté informació interna sensible (comandaments, menús, vehicles,
serveis...). **NO pugis l'Excel original al Drive.** Puja-hi només una còpia neta
que conté NOMÉS parc/mínim/reals. Te la genera un script (PAS 1).

Així, encara que algú trobés l'enllaç del Drive, només veuria dades de cobertura,
res comprometedor.

---

## PAS 1 — Generar la còpia neta (al teu PC)

Necessites Python instal·lat (mira la PART A de l'INSTRUCTIU si no el tens).

1. Posa l'Excel original i la carpeta `scripts` a la mateixa carpeta.
2. Obre el Terminal / CMD en aquella carpeta i executa:
   ```
   python scripts/neteja_excel.py "Personal de guàrdia.xlsx"
   ```
   (o fes doble clic a `neteja_excel.py` i arrossega-hi l'Excel quan t'ho demani)
3. Es crea `cobertura_neta.xlsx`. **Aquest és el que pujaràs al Drive.**
   Passa de ~720 KB a ~8 KB perquè només té les dades de cobertura.

Cada dia repetiràs aquest pas amb l'Excel del dia (és molt ràpid).

---

## PAS 2 — Preparar el Google Drive (un sol cop)

1. Entra al Google Drive de l'Assemblea i puja-hi `cobertura_neta.xlsx`.
2. Botó dret sobre el fitxer → **Compartir → Compartir** → canvia a
   **"Qualsevol persona amb l'enllaç"**, permís **Lector**. Copia l'enllaç.
3. L'enllaç serà tipus:
   `https://drive.google.com/file/d/XXXXXXXXXXXX/view?usp=sharing`
   Anota el tros `XXXXXXXXXXXX` (l'**ID del fitxer**).
4. L'enllaç que fa servir el sistema és:
   `https://drive.google.com/uc?export=download&id=XXXXXXXXXXXX`
   (el mateix ID). Guarda'l per al PAS 4.

> Mentre no canviïs el fitxer de lloc ni el nom al Drive, aquest enllaç serà
> SEMPRE EL MATEIX cada dia. El que canvia és el contingut, no l'adreça.

> NO escampis aquest enllaç. Que només el coneguin el sistema i tu.

---

## PAS 3 — Muntar GitHub + Netlify (un sol cop)

Segueix la PART A de l'INSTRUCTIU.md (crear repositori PRIVAT, pujar el projecte,
i publicar a Netlify).

---

## PAS 4 — Connectar el Drive amb el sistema

1. Al teu repositori de GitHub: **Settings → Secrets and variables → Actions →
   New repository secret.**
   - Nom: `ENLLAC_TRI`
   - Valor: l'enllaç de descàrrega directa del PAS 2.4.
2. **Add secret.**

## PAS 5 — Provar que funciona ara mateix

1. Al repositori, pestanya **Actions**.
2. Workflow **"Actualitza dades TRI cada dia"** → **Run workflow**.
3. Espera un minut. Si surt ✅ verd, ha llegit el Drive i ha desat el dia.
4. Refresca el web: hi surt el dia, amb el mapa ple.

---

## PAS 6 — La prova de 5 dies

A partir d'aquí, el sistema ho farà SOL cada dia a les ~10:10. Durant 5 dies:

**Cada dia (només has de mantenir el Drive al dia):**
1. Genera la còpia neta de l'Excel del dia (PAS 1).
2. Puja-la al Drive **substituint** l'anterior (mateix nom → l'enllaç no canvia).
   - A Google Drive: clica el fitxer → dalt a la dreta els 3 punts, o simplement
     arrossega el nou a sobre → "Substituir fitxer existent".
3. No cal fer res més: a les 10:10 el sistema fa la foto sol.
   (Si vols comprovar-ho a l'instant, ves a Actions → Run workflow.)
4. Obre el web i el PDF del TRI d'aquell dia al Telegram, i **compara**.

### Taula de comprovació

| Dia | Data | Web: sota mínims | PDF: sota mínims | Web: tancats | PDF: tancats | Quadra? |
|-----|------|------------------|------------------|--------------|--------------|---------|
| 1   |      |                  |                  |              |              |         |
| 2   |      |                  |                  |              |              |         |
| 3   |      |                  |                  |              |              |         |
| 4   |      |                  |                  |              |              |         |
| 5   |      |                  |                  |              |              |         |

---

## Què validem al final

- Que la foto es fa SOLA cada dia (un cop el Drive està al dia).
- Que els números del web coincideixen amb el PDF oficial.
- Que la línia temporal creix cada dia.

---

## Sobre l'hora

El TRI del dia següent s'omple entre les 15:00 i les 18:00. La foto es fa a les
10:10, així que captura la situació del dia en curs. Si algun dia a les 10:10
l'Excel encara no reflecteix el dia, es pot canviar l'hora al fitxer del workflow
(línia `cron`). M'ho dius i t'ajudo.

## Si algun número no quadra

Apunta dia + parc + web vs PDF, i m'ho passes. El sistema ja avisa sol si la suma
de mínims no dona 354 (senyal que l'Excel ha canviat d'estructura).

---

## Es pot fer del tot automàtic (sense generar la còpia neta cada dia)?

Sí, més endavant. Si el teu Excel se sincronitzés sol a un núvol, es podria fer que
el sistema el netegés ell mateix. De moment, generar la còpia neta cada dia és el
més senzill i segur per a la prova. Quan la prova funcioni, ho podem automatitzar més.
