# Service de mise en forme de mémoire (Word, arabe + latin)

Appelé par le scénario Make « Mémoire – mise en forme Word via Telegram ».

## Déploiement (ex. Render, Railway, Fly.io — tout hébergeur Docker)
1. Mettre ce dossier dans un dépôt GitHub (privé).
2. Créer un « Web Service » à partir du Dockerfile.
3. Variable d'environnement : `API_KEY` = une longue valeur secrète de votre choix.
4. Tester : `GET https://VOTRE-URL/health` doit répondre `{"status":"ok"}`.
5. Dans Make, ouvrir le scénario, module « Set multiple variables » : renseigner `base_url` (sans / final) et `api_key`.

## Endpoints (en-tête `Authorization: Bearer <API_KEY>`)
- `POST /spec-text`  file=cahier (.pdf/.docx) -> texte du cahier des charges
- `POST /format`     file=mémoire.docx, rules=JSON (facultatif) -> .docx formaté
- `POST /outline`    plan détecté (pour vérification)
- `GET  /default-rules`

## Ce que fait /format
Page de garde conservée (sans numéro), préliminaires (dédicace, remerciements, résumé/ملخص) détectés,
titres Chapitre/الفصل, Introduction/مقدمة, Conclusion/خاتمة, 1. / 1.1 / 1.1.1 mis en styles Titre 1-4,
paragraphes arabes en RTL (police arabe), latins en LTR, runs mixtes séparés, légendes de tableaux/figures,
vrai sommaire Word (champ TOC pré-rempli), sections : couverture / préliminaires (i, ii…) / corps (1, 2…).
