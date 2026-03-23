/**
 * Gmail → GitHub : Sync automatique des plannings UrbanSoccer
 *
 * Ce script surveille Gmail pour les emails contenant des plannings
 * en pièce jointe et les pousse automatiquement vers GitHub.
 *
 * Installation :
 *   1. Aller sur https://script.google.com → Nouveau projet
 *   2. Coller ce code
 *   3. Remplacer GITHUB_TOKEN par votre token
 *   4. Exécuter "setup" une fois (autorise les permissions)
 *   5. Le script tourne automatiquement toutes les 5 minutes
 */

// ══════════════════════════════════════════════════════════════════════════════
// CONFIGURATION — Modifiez ces valeurs
// ══════════════════════════════════════════════════════════════════════════════

const CONFIG = {
  // Token GitHub (Fine-grained, avec permission Contents: Read & Write)
  GITHUB_TOKEN: "VOTRE_TOKEN_GITHUB_ICI",

  // Repo GitHub
  GITHUB_OWNER: "OhLaPey",
  GITHUB_REPO: "planning-urbansoccer",
  GITHUB_BRANCH: "main",

  // Filtre email : objet contenant ce texte
  EMAIL_SUBJECT_FILTER: "Planning UrbanSoccer",

  // Label Gmail appliqué après traitement (pour ne pas retraiter)
  LABEL_PROCESSED: "Planning-Synced",

  // Pattern des fichiers de planning acceptés
  FILE_PATTERN: /^Plannings\s+\d{4}\s+S\d{1,2}.*\.xlsx$/i,
};

// ══════════════════════════════════════════════════════════════════════════════
// SETUP — Exécuter une seule fois
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Crée le label Gmail et installe le déclencheur automatique.
 * À exécuter UNE SEULE FOIS manuellement.
 */
function setup() {
  // Créer le label Gmail s'il n'existe pas
  let label = GmailApp.getUserLabelByName(CONFIG.LABEL_PROCESSED);
  if (!label) {
    label = GmailApp.createLabel(CONFIG.LABEL_PROCESSED);
    Logger.log("Label créé : " + CONFIG.LABEL_PROCESSED);
  }

  // Supprimer les anciens déclencheurs de ce script
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "checkEmails") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Créer un déclencheur toutes les 5 minutes
  ScriptApp.newTrigger("checkEmails")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("Déclencheur installé : checkEmails toutes les 5 minutes");
  Logger.log("Setup terminé !");
}

// ══════════════════════════════════════════════════════════════════════════════
// FONCTION PRINCIPALE — Vérification des emails
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Vérifie les nouveaux emails avec des plannings en pièce jointe
 * et les pousse vers GitHub.
 */
function checkEmails() {
  const label = GmailApp.getUserLabelByName(CONFIG.LABEL_PROCESSED);
  if (!label) {
    Logger.log("Label non trouvé. Exécutez setup() d'abord.");
    return;
  }

  // Chercher les emails non traités avec le bon objet
  const query = `subject:"${CONFIG.EMAIL_SUBJECT_FILTER}" has:attachment -label:${CONFIG.LABEL_PROCESSED}`;
  const threads = GmailApp.search(query, 0, 10);

  if (threads.length === 0) {
    Logger.log("Aucun nouveau planning à traiter.");
    return;
  }

  Logger.log(`${threads.length} email(s) à traiter.`);

  for (const thread of threads) {
    const messages = thread.getMessages();
    let allSuccess = true;

    for (const message of messages) {
      const attachments = message.getAttachments();

      for (const attachment of attachments) {
        const fileName = attachment.getName();

        // Vérifier que c'est un fichier de planning
        if (!CONFIG.FILE_PATTERN.test(fileName)) {
          Logger.log(`Ignoré (pas un planning) : ${fileName}`);
          continue;
        }

        Logger.log(`Traitement : ${fileName}`);

        // Récupérer le contenu en base64
        const content = Utilities.base64Encode(attachment.getBytes());

        // Pousser vers GitHub (avec retry)
        let success = pushToGitHub(fileName, content);
        if (!success) {
          // Retry une fois — le SHA a pu changer entre le GET et le PUT
          Logger.log(`Retry : ${fileName}`);
          Utilities.sleep(2000);
          success = pushToGitHub(fileName, content);
        }

        if (success) {
          Logger.log(`Poussé vers GitHub : ${fileName}`);
        } else {
          Logger.log(`ERREUR push GitHub : ${fileName}`);
          allSuccess = false;
        }
      }
    }

    // Marquer le thread comme traité SEULEMENT si tous les fichiers ont été poussés
    if (allSuccess) {
      thread.addLabel(label);
      thread.markRead();
    } else {
      Logger.log(`Thread non marqué (erreurs) — sera retraité au prochain cycle.`);
    }
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// GITHUB API — Upload de fichier
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Pousse un fichier vers GitHub via l'API Contents.
 * Gère la création ET la mise à jour (récupère le SHA si le fichier existe).
 */
function pushToGitHub(fileName, contentBase64) {
  const url = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/contents/${encodeURIComponent(fileName)}`;

  const headers = {
    Authorization: `Bearer ${CONFIG.GITHUB_TOKEN}`,
    Accept: "application/vnd.github+v3+json",
    "User-Agent": "UrbanSoccer-Planning-Sync",
  };

  // Vérifier si le fichier existe déjà (pour obtenir le SHA)
  let sha = null;
  try {
    const existing = UrlFetchApp.fetch(url + `?ref=${CONFIG.GITHUB_BRANCH}`, {
      method: "GET",
      headers: headers,
      muteHttpExceptions: true,
    });

    if (existing.getResponseCode() === 200) {
      const data = JSON.parse(existing.getContentText());
      sha = data.sha;
      Logger.log(`Fichier existant, SHA : ${sha}`);
    }
  } catch (e) {
    Logger.log(`Fichier nouveau : ${fileName}`);
  }

  // Extraire le numéro de semaine pour le message de commit
  const weekMatch = fileName.match(/S(\d{1,2})/);
  const weekInfo = weekMatch ? `S${weekMatch[1]}` : fileName;

  // Créer ou mettre à jour le fichier
  const payload = {
    message: `Sync planning SharePoint : ${weekInfo}`,
    content: contentBase64,
    branch: CONFIG.GITHUB_BRANCH,
  };

  if (sha) {
    payload.sha = sha;
  }

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "PUT",
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code = response.getResponseCode();
    if (code === 200 || code === 201) {
      return true;
    } else {
      Logger.log(`Erreur GitHub (${code}) : ${response.getContentText()}`);
      return false;
    }
  } catch (e) {
    Logger.log(`Exception GitHub : ${e.message}`);
    return false;
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// TEST — Pour vérifier que tout fonctionne
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Teste la connexion GitHub (sans envoyer de fichier).
 * Exécutez cette fonction pour vérifier que le token est valide.
 */
function testGitHubConnection() {
  const url = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}`;
  const headers = {
    Authorization: `Bearer ${CONFIG.GITHUB_TOKEN}`,
    Accept: "application/vnd.github+v3+json",
    "User-Agent": "UrbanSoccer-Planning-Sync",
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: headers,
      muteHttpExceptions: true,
    });

    const code = response.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      Logger.log(`Connexion OK ! Repo : ${data.full_name}`);
      Logger.log(`Branche par défaut : ${data.default_branch}`);
    } else {
      Logger.log(`Erreur (${code}) : vérifiez votre token GitHub.`);
      Logger.log(response.getContentText());
    }
  } catch (e) {
    Logger.log(`Exception : ${e.message}`);
  }
}
