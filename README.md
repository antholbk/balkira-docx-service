# Balkira DocX Service

Microservice Node.js/Express qui génère de vrais fichiers `.docx` Word stylisés (palette Balkira / GxP) à partir de contenu Markdown.

## Stack

- **Express** — serveur HTTP
- **docx** (v9) — génération Word
- **cors** — headers CORS

## Endpoint

### `POST /generate-docx`

**Body JSON :**
```json
{
  "content": "## Titre\n\nContenu markdown...",
  "filename": "BALKIRA_ANOMALY_FORM_V1_REQ-xxx.docx",
  "template_id": "anomaly_form",
  "version": "V1",
  "request_id": "REQ-xxx",
  "language": "fr"
}
```

**Réponse :**
```json
{
  "base64": "UEsDBAoA...",
  "filename": "BALKIRA_ANOMALY_FORM_V1_REQ-xxx.docx",
  "size_kb": 45
}
```

### `GET /health`

Retourne `{ "status": "ok" }`.

## Test local

```bash
npm install
node server.js
# Dans un autre terminal :
curl -X POST http://localhost:3000/generate-docx \
  -H "Content-Type: application/json" \
  -d '{"content":"# Test\n\n## Section 1\n\nTexte **important**\n\n⚠ À VÉRIFIER : compléter\n\n| Col1 | Col2 |\n|------|------|\n| Val1 | Val2 |","filename":"test.docx","template_id":"anomaly_form","version":"V1","request_id":"REQ-TEST-001","language":"fr"}'
```

## Déploiement sur Railway (< 5 minutes)

### Prérequis
- Compte Railway gratuit : https://railway.app
- Git installé

### Étapes

```bash
# 1. Initialise un dépôt Git dans ce dossier
cd docx-service
git init
git add .
git commit -m "init balkira docx service"

# 2. Pousse sur GitHub (optionnel mais recommandé)
gh repo create balkira-docx-service --public --push --source=.

# 3. Dans Railway :
#    - New Project → Deploy from GitHub repo → sélectionne balkira-docx-service
#    - Railway détecte le Procfile automatiquement
#    - Variables d'environnement : aucune requise
#    - Le service est en ligne en ~2 minutes
```

### Alternative : Railway CLI

```bash
npm install -g @railway/cli
railway login
railway init           # "Create new project"
railway up             # déploie depuis le dossier courant
railway domain         # obtient l'URL publique
```

### URL finale

```
https://balkira-docx-service-xxxx.up.railway.app/generate-docx
```

Remplace l'URL dans ton workflow n8n dans le nœud **"⚡ Mistral AI — Generate Document"** → ajoute un nœud HTTP Request vers ce service pour convertir le Markdown généré en `.docx`.

## Markdown supporté

| Syntaxe | Rendu Word |
|---------|------------|
| `# Titre` | Heading 1, couleur navy |
| `## Titre` | Heading 2, couleur navy |
| `### Titre` | Heading 3 |
| `**gras**` | TextRun bold |
| `- item` ou `* item` | Bullet point orange |
| `1. item` | Liste numérotée navy |
| `⚠ À VÉRIFIER` | Bloc amber avec bordure orange |
| `\| col \| col \|` | Tableau Word avec header navy |
| `---` | Ligne horizontale orange |

## Templates supportés

`anomaly_form` · `interface_spec` · `urs` · `dira_pdfm` · `sop` · `iq_oq_pq` · `uat_design` · `capa` · `change_control`
