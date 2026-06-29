# Task 1 — xls-reader MCP : Setup initial

## Contexte

Créer un nouveau serveur MCP **sans interface graphique** dans `xls-reader/` capable de fournir à
l'agent des outils pour lire le contenu d'un fichier Excel (`.xlsx` / `.xls`).

Fonctionnalités attendues :
- Lister les onglets d'un fichier
- Lire un onglet et le retourner au format tableau Markdown
- Gestion intelligente du contenu : ligne vide = séparateur de section, cellules vides ignorées
- Limite configurable : par défaut, le budget cellules (`max_cols × max_rows`) ne dépasse pas **2000**

### Choix technique : SheetJS (`xlsx`)

**SheetJS** (`xlsx` sur npm) est retenu comme bibliothèque de lecture Excel :
- Pur JS, **aucune dépendance native** (contrairement à `exceljs` qui peut nécessiter des binaires)
- Supporte `.xlsx`, `.xls`, `.ods`, `.csv`
- API minimaliste : `readFile()` + `utils.sheet_to_json()` ou accès direct aux cellules
- ~3M téléchargements hebdomadaires, stable et maintenu

> **Budget cellules — justification du défaut à 2000** : à ~20 caractères/cellule en moyenne,
> 2000 cellules représentent ~40 000 caractères, bien en-deçà des limites de contexte des LLMs
> modernes (100k+ tokens). Avec 10 colonnes cela donne 200 lignes ; avec 20 colonnes, 100 lignes —
> deux vues raisonnables sur une feuille de calcul. L'agent peut toujours surcharger via les
> paramètres `max_cols` / `max_rows` / `cell_budget`.

### Patterns à respecter

Suivre les mêmes conventions que les autres MCPs du workspace (`simple-scraper-mcp`, `raindrop-mcp`) :

- `McpServer` + `server.tool()` du SDK MCP (API haut niveau)
- Validation des entrées avec **Zod**
- Logger centralisé avec fichier de log persistant (copier le pattern existant)
- `rootDir` calculé via `import.meta.url`
- `dotenv` avec `override: true`
- `"type": "module"` dans `package.json` (ESM)
- Redirection `console.log` → `console.error` dans le processus MASTER

---

## Structure du projet

```
xls-reader/
├── src/
│   ├── index.ts       # Point d'entrée MCP + définition des tools
│   ├── excel.ts       # Logique de lecture Excel (listSheets, readSheet)
│   └── logger.ts      # Logger centralisé (copier depuis un MCP existant)
├── .env.example
├── .gitignore
├── package.json
└── tsconfig.json
```

---

## Phase 1 — Initialisation du projet

### 1.1 — `package.json`

```json
{
  "name": "xls-reader",
  "type": "module",
  "main": "dist/index.js",
  "scripts": {
    "build": "tsc",
    "start": "node dist/index.js",
    "dev": "tsx src/index.ts"
  },
  "engines": { "node": ">=18.0.0" }
}
```

**Dépendances** :

```bash
npm install @modelcontextprotocol/sdk xlsx dotenv zod
npm install -D typescript @types/node tsx
```

### 1.2 — `tsconfig.json`

Copier exactement le pattern des autres MCPs du workspace :

```json
{
  "compilerOptions": {
    "target": "ESNext",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "outDir": "./dist",
    "rootDir": "./src",
    "strict": true,
    "skipLibCheck": true,
    "esModuleInterop": true,
    "forceConsistentCasingInFileNames": true
  },
  "include": ["src/**/*"]
}
```

### 1.3 — `.gitignore`

```
node_modules/
dist/
.env
```

---

## Phase 2 — Logger (`src/logger.ts`)

Copier `logger.ts` depuis `simple-scraper-mcp/src/logger.ts` en adaptant :
- Nom du dossier de log : `xls-reader`
- Même API : `setupLogging(prefix)`, redirection `console.log` → `console.error`

---

## Phase 3 — Logique Excel (`src/excel.ts`)

Ce module contient toute la logique métier, indépendante du serveur MCP.

### 3.1 — `listSheets(filePath: string): string[]`

Ouvre le fichier et retourne la liste des noms d'onglets dans l'ordre.

```typescript
import * as XLSX from 'xlsx';

export function listSheets(filePath: string): string[] {
  const workbook = XLSX.readFile(filePath);
  return workbook.SheetNames;
}
```

### 3.2 — `readSheet(filePath, sheetName, options): string`

Retourne le contenu d'un onglet au format Markdown.

**Signature** :

```typescript
export interface ReadSheetOptions {
  cellBudget?: number;   // max cols × rows (défaut : 200)
  maxCols?: number;      // surcharge du nombre de colonnes
  maxRows?: number;      // surcharge du nombre de lignes
  headerRow?: boolean;   // première ligne = en-tête (défaut : true)
}

export function readSheet(
  filePath: string,
  sheetName: string,
  options: ReadSheetOptions = {}
): string
```

**Implémentation** :

1. Lire le workbook avec `XLSX.readFile(filePath)`
2. Récupérer le worksheet : `workbook.Sheets[sheetName]`
3. Convertir en tableau 2D avec `XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })` — cela retourne `string[][]`
4. Appliquer les limites (voir §3.3)
5. Convertir en Markdown (voir §3.4)

### 3.3 — Calcul des limites

```typescript
const { cellBudget = 2000, maxCols, maxRows, headerRow = true } = options;

const totalCols = rows[0]?.length ?? 0;
const effectiveCols = Math.min(totalCols, maxCols ?? totalCols);
const effectiveRows = Math.min(
  rows.length,
  maxRows ?? Math.floor(cellBudget / Math.max(effectiveCols, 1))
);

const truncatedCols = effectiveCols < totalCols;
const truncatedRows = effectiveRows < rows.length;
```

Si les données sont tronquées, ajouter un avertissement en fin de sortie :

```
> ⚠ Affichage limité : X colonnes sur Y, Z lignes sur W. Utiliser max_cols / max_rows pour ajuster.
```

### 3.4 — Conversion en Markdown

**Règle clé** : une ligne **entièrement vide** est interprétée comme un **séparateur de section**
— le tableau en cours est fermé et un nouveau tableau commence après.

Algorithme :

```
sections = [[]]   // liste de groupes de lignes non-vides

pour chaque ligne dans les données :
  si ligne est entièrement vide :
    si sections.dernière n'est pas vide :
      sections.ajouter([])   // nouvelle section
  sinon :
    sections.dernière.ajouter(ligne)

retourner sections.non-vides.map(sectionToMarkdownTable).join("\n\n")
```

**Conversion section → tableau Markdown** :

- Si `headerRow = true`, la première ligne de la *première* section devient l'en-tête ; les sections
  suivantes reprennent le même en-tête (utile quand les sections partagent la même structure).
- Les cellules vides deviennent des chaînes vides (` ` dans le tableau).
- Chaque cellule est convertie via `String(value).trim()`.

Exemple de sortie attendue :

```markdown
| Nom | Prénom | Age |
|-----|--------|-----|
| Dupont | Jean | 42 |
| Martin | Claire | 35 |

| Dupont | Jean | 42 |
| Martin | Claire | 35 |
```

> **Remarque** : pour les nombres et les dates, SheetJS retourne des valeurs JS natives (`number`,
> `Date`). Formater les dates avec `date.toLocaleDateString('fr-FR')` ou laisser le `String()` par
> défaut selon le besoin — à décider à l'implémentation.

---

## Phase 4 — Serveur MCP (`src/index.ts`)

### Structure

1. Imports + `rootDir` + `dotenv` + `setupLogging`
2. Création du `McpServer`
3. Définition des outils via `server.tool()`
4. Fonction `main()` avec `StdioServerTransport`

### Imports clés

```typescript
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { listSheets, readSheet } from './excel.js';
```

---

### Outil 1 : `list_sheets`

**Description** : List all sheet names (tabs) in an Excel file.

**Paramètres** :

| Param | Type | Requis | Description |
|-------|------|--------|-------------|
| `file_path` | `z.string()` | Oui | Chemin absolu ou relatif vers le fichier Excel |

**Retour** : liste des noms d'onglets séparés par des sauts de ligne.

```typescript
return { content: [{ type: 'text', text: sheets.join('\n') }] };
```

---

### Outil 2 : `get_sheet`

**Description** : Read an Excel sheet and return its content as a Markdown table. Empty rows are treated as section separators. Data is limited by a cell budget (cols × rows).

**Paramètres** :

| Param | Type | Requis | Description |
|-------|------|--------|-------------|
| `file_path` | `z.string()` | Oui | Chemin vers le fichier Excel |
| `sheet_name` | `z.string()` | Oui | Nom de l'onglet à lire |
| `cell_budget` | `z.number().optional()` | Non | Budget cellules max (défaut : 2000) |
| `max_cols` | `z.number().optional()` | Non | Nombre max de colonnes |
| `max_rows` | `z.number().optional()` | Non | Nombre max de lignes |
| `header_row` | `z.boolean().optional()` | Non | Première ligne = en-tête (défaut : true) |

**Retour** : tableau(x) Markdown + avertissement de troncature si applicable.

---

### Gestion des erreurs

Pattern identique aux autres MCPs :

```typescript
try {
  // ...
} catch (error) {
  return {
    content: [{ type: 'text', text: (error as Error).message }],
    isError: true,
  };
}
```

Erreurs à gérer explicitement :
- Fichier introuvable → message clair : `"File not found: <path>"`
- Onglet introuvable → `"Sheet '<name>' not found. Available: <list>"`

---

## Récapitulatif des fichiers

| Fichier | Action |
|---------|--------|
| `src/index.ts` | Nouveau — serveur MCP avec 2 outils (`list_sheets`, `get_sheet`) |
| `src/excel.ts` | Nouveau — logique de lecture Excel (`listSheets`, `readSheet`) |
| `src/logger.ts` | Nouveau — copier depuis `simple-scraper-mcp/src/logger.ts` (adapter le nom) |
| `package.json` | Nouveau — `npm init -y` puis adapter |
| `tsconfig.json` | Nouveau — config TypeScript standard ESM |
| `.gitignore` | Nouveau |
| `.env.example` | Nouveau — documenter les variables (aucune requise pour l'instant) |

---

## Ordre de développement recommandé

1. `npm init -y` + dépendances + `package.json` / `tsconfig.json` / `.gitignore`
2. Copier et adapter `src/logger.ts`
3. Implémenter `src/excel.ts` (tester en isolation avec un fichier Excel de test)
4. Implémenter `src/index.ts` avec les 2 outils
5. Tester avec `npm run dev` ou `mcp-inspector`

---

## Propositions d'amélioration (hors scope initial)

Ces fonctionnalités peuvent être ajoutées dans de prochaines tâches :

1. **Pagination** : paramètres `start_row` et `start_col` pour naviguer dans les grandes feuilles
   sans changer le budget — utile quand un agent lit une feuille par morceaux.

2. **Outil `search_in_sheet`** : recherche textuelle dans un onglet, retourne les lignes
   correspondantes avec leur numéro (utile pour les feuilles de données volumineuses).

3. **Outil `get_cell_range`** : lecture d'une plage nommée (e.g. `A1:D10`) pour un accès précis
   sans passer par le budget global.

4. **Formatage des types** : option pour formater les dates selon une locale / un pattern
   (`date_format: 'ISO' | 'FR' | 'US'`) et les nombres selon une locale.

5. **Support des formules** : SheetJS peut retourner la valeur calculée **ou** la formule brute.
   Un paramètre `raw_formulas: boolean` exposerait les formules à l'agent.

6. **Écriture** : SheetJS supporte l'écriture de fichiers Excel — un outil `patch_cell` ou
   `append_rows` permettrait à l'agent de modifier des fichiers (scope différent, à évaluer).
