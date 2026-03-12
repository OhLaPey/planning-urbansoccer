# Automatisation SharePoint → GitHub

Ce guide explique comment automatiser l'upload des plannings depuis le dossier
SharePoint `Documents > 02. USP - TOS - General > RH > Plannings` vers ce repo GitHub.

## Architecture

```
SharePoint (Plannings/*.xlsx)
        │
        ├──── Option A : Power Automate (recommandé)
        │         └─► GitHub (commit via API)
        │
        └──── Option B : GitHub Actions (cron)
                  └─► Microsoft Graph API → télécharge → commit
```

Les deux options déclenchent ensuite le workflow `generate.yml` existant
qui génère les ICS/HTML et déploie sur GitHub Pages.

---

## Option A : Power Automate (recommandé ⭐)

C'est la solution la plus simple, accessible depuis le menu SharePoint
**Automatiser > Flux**.

### Étapes de création du flux

1. **Depuis SharePoint**, cliquez sur **Automatiser > Flux > Créer un flux**

2. **Déclencheur** : `Quand un fichier est créé ou modifié dans un dossier`
   - **Site** : `USP - TOS`
   - **ID de dossier** : `/General/RH/Plannings`

3. **Condition** : Le nom du fichier se termine par `.xlsx`
   ET le nom du fichier commence par `Plannings`
   ```
   Condition :
     @endsWith(triggerOutputs()?['headers/x-ms-file-name'], '.xlsx')
     AND
     @startsWith(triggerOutputs()?['headers/x-ms-file-name'], 'Plannings')
   ```

4. **Action** : `Obtenir le contenu du fichier` (SharePoint)
   - Identifiant de fichier : `ID` du déclencheur

5. **Action** : `Créer ou mettre à jour un fichier` (GitHub)
   - **Connexion** : Se connecter avec un compte GitHub ayant accès au repo
   - **Repository** : `OhLaPey/planning-urbansoccer`
   - **Branche** : `main`
   - **Chemin** : Le nom du fichier depuis le déclencheur
   - **Contenu** : Le contenu du fichier (en base64)
   - **Message de commit** : `📅 Sync planning @{triggerOutputs()?['headers/x-ms-file-name']}`

### Flux Power Automate — JSON exportable

Vous pouvez importer ce flux directement dans Power Automate :

```json
{
  "definition": {
    "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
    "triggers": {
      "When_a_file_is_created_or_modified": {
        "type": "ApiConnection",
        "inputs": {
          "host": {
            "connection": { "name": "@parameters('$connections')['sharepointonline']['connectionId']" }
          },
          "method": "get",
          "path": "/datasets/@{encodeURIComponent(encodeURIComponent('https://votre-tenant.sharepoint.com/sites/USP-TOS'))}/triggers/onupdatedfile",
          "queries": {
            "folderId": "/Shared Documents/General/RH/Plannings",
            "includeFileContent": true
          }
        },
        "recurrence": { "frequency": "Minute", "interval": 5 }
      }
    },
    "actions": {
      "Condition_fichier_planning": {
        "type": "If",
        "expression": {
          "and": [
            { "endsWith": ["@triggerOutputs()?['headers/x-ms-file-name']", ".xlsx"] },
            { "startsWith": ["@triggerOutputs()?['headers/x-ms-file-name']", "Plannings"] },
            { "not": { "startsWith": ["@triggerOutputs()?['headers/x-ms-file-name']", "~$"] } }
          ]
        },
        "actions": {
          "Create_or_update_file_on_GitHub": {
            "type": "ApiConnection",
            "inputs": {
              "host": {
                "connection": { "name": "@parameters('$connections')['github']['connectionId']" }
              },
              "method": "put",
              "path": "/repos/OhLaPey/planning-urbansoccer/contents/@{triggerOutputs()?['headers/x-ms-file-name']}",
              "body": {
                "message": "📅 Sync planning @{triggerOutputs()?['headers/x-ms-file-name']}",
                "content": "@{base64(triggerBody())}"
              }
            }
          }
        }
      }
    }
  }
}
```

### Résultat

Chaque fois qu'un fichier `Plannings YYYY SXX.xlsx` est modifié dans SharePoint :
1. Power Automate détecte le changement (< 5 min)
2. Le fichier est poussé vers GitHub
3. Le workflow `generate.yml` se déclenche automatiquement
4. Les ICS + HTML sont régénérés et déployés sur GitHub Pages

---

## Option B : GitHub Actions + Microsoft Graph API

Pour un contrôle total, un workflow GitHub Actions interroge SharePoint
périodiquement via l'API Microsoft Graph.

### 1. Créer une App Registration Azure AD

1. Aller sur https://portal.azure.com → **Azure Active Directory** → **App registrations**
2. **New registration** :
   - Nom : `Planning UrbanSoccer Sync`
   - Type : `Accounts in this organizational directory only`
3. Aller dans **Certificates & secrets** → **New client secret**
   - Copier la valeur du secret
4. Aller dans **API permissions** → **Add a permission** :
   - **Microsoft Graph** → **Application permissions** :
     - `Sites.Read.All`
     - `Files.Read.All`
   - Cliquer sur **Grant admin consent**
5. Noter :
   - **Application (client) ID** → `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** → `AZURE_TENANT_ID`
   - **Client secret** → `AZURE_CLIENT_SECRET`

### 2. Trouver l'URL du site SharePoint

L'URL est du type : `https://votre-tenant.sharepoint.com/sites/USP-TOS`

### 3. Configurer les secrets GitHub

Aller dans **Settings > Secrets and variables > Actions** du repo et ajouter :

| Secret | Valeur |
|--------|--------|
| `AZURE_TENANT_ID` | ID du tenant Azure |
| `AZURE_CLIENT_ID` | ID de l'application |
| `AZURE_CLIENT_SECRET` | Secret de l'application |
| `SHAREPOINT_SITE_URL` | `https://xxx.sharepoint.com/sites/USP-TOS` |
| `SHAREPOINT_FOLDER_PATH` | `General/RH/Plannings` |

### 4. Le workflow s'exécute automatiquement

Le workflow `.github/workflows/sync-sharepoint.yml` :
- Tourne toutes les heures (lun-ven, 7h-19h)
- Télécharge les nouveaux/modifiés `Plannings YYYY SXX.xlsx`
- Régénère les ICS/HTML
- Commit et push automatiquement

Vous pouvez aussi le déclencher manuellement depuis l'onglet **Actions** du repo.

### Usage local du script

```bash
# Voir les fichiers disponibles
export AZURE_TENANT_ID=...
export AZURE_CLIENT_ID=...
export AZURE_CLIENT_SECRET=...
export SHAREPOINT_SITE_URL=https://xxx.sharepoint.com/sites/USP-TOS
python sync_sharepoint.py --dry-run

# Sync uniquement la semaine 11
python sync_sharepoint.py --pattern S11

# Sync tout
python sync_sharepoint.py
```

---

## Comparatif

| | Power Automate | GitHub Actions |
|---|---|---|
| **Facilité** | ⭐⭐⭐ Très simple | ⭐⭐ Config Azure AD |
| **Délai** | < 5 min | Toutes les heures |
| **Coût** | Inclus dans M365 | Gratuit (GitHub) |
| **Fiabilité** | Haute | Haute |
| **Contrôle** | Limité | Total |

**Recommandation** : Commencez par Power Automate (Option A). Si vous avez
besoin de plus de contrôle, passez à l'Option B.
