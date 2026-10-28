# Générateur de nomenclature et de dossiers de montage assisté par l'IA

## Démonstration
https://github.com/user-attachments/assets/29e2154b-422c-4f9d-8ab8-f981e7aa73b6

![Liste des pièces](captures/capture02)
![Formulaire de validation](captures/Capture04)
![Dossier de montage généré](captures/Capture03_)

Projet personnel : à partir d'une présentation de revue de conception (PowerPoint), produire la **nomenclature des pièces** (BOM), la faire **valider par l'équipe métier** dans une application **Power Apps**, puis générer automatiquement le **dossier de montage** d'un sous-ensemble avec **Power Automate**.

> Toutes les données sont **fictives** (véhicule concept « VX-1 »). Projet réalisé à titre personnel, sans lien avec une entreprise.

## Architecture

```

https://github.com/user-attachments/assets/0fd0d44c-d4d8-4b6a-aade-76a5a31a07a8


PowerPoint de revue de conception
        │
        ▼
Script Python + IA (Claude)  ──►  nomenclature.csv  (lignes douteuses marquées « à vérifier »)
        │
        ▼
Microsoft Lists « Pieces »
        │
        ▼
Application Power Apps
  • filtre par sous-ensemble, recherche, badge « À vérifier », compteur
  • formulaire de correction + bouton « Valider la ligne » (l'IA propose, le métier valide)
  • bouton « Générer le dossier de montage » (actif seulement si le sous-ensemble est entièrement validé)
        │  envoie le sous-ensemble + le contenu HTML
        ▼
Flux Power Automate « GenererDossierMontage »
  • crée le fichier dans OneDrive  • crée un lien de partage  • renvoie le lien à l'application
```

## Résultats mesurés (11 pièces attendues)

| Mode d'extraction | Résultat |
|---|---|
| Sans IA (lecture des tableaux uniquement) | **7/11** : ajouts, remplacement et changement de quantité écrits en texte libre ignorés |
| Avec IA | **11/11 pièces identifiées**, 10/11 quantités ; la quantité du joint, absente de la présentation, est **signalée au lieu d'être inventée** |

L'IA a aussi relevé des incohérences de conception : une hausse du nombre de vis « suite au renfort de charnière » sans pièce de renfort dans la nomenclature, et un saut de numérotation des références. Deux exécutions ne donnent pas toujours exactement le même résultat : c'est pour cela que la validation humaine fait partie de l'outil.

## Contenu du dépôt

| Fichier | Rôle |
|---|---|
| `donnees/presentation_conception_exemple.pptx` | Présentation fictive : tableau de nomenclature + évolutions en texte libre |
| `donnees/nomenclature_attendue.csv` | Résultat attendu, pour mesurer la qualité de l'extraction |
| `scripts/extraction_nomenclature.py` | Extraction (avec ou sans IA), contrôles automatiques, export CSV, comparaison |
| `captures/` | Captures d'écran de l'application et du flux |
| `export_power_platform/` | Exports de l'application Power Apps et du flux Power Automate |

## 1. Extraction de la nomenclature (Python)

```powershell
python -m pip install -r requirements.txt
cd scripts
# Sans IA
python extraction_nomenclature.py ..\donnees\presentation_conception_exemple.pptx --sans-ia --comparer ..\donnees\nomenclature_attendue.csv
# Avec IA (la clé n'est jamais écrite dans le code ni publiée)
$env:ANTHROPIC_API_KEY="votre_cle"
python extraction_nomenclature.py ..\donnees\presentation_conception_exemple.pptx --comparer ..\donnees\nomenclature_attendue.csv --sortie nomenclature_ia.csv
```

Contrôles automatiques : référence mal formée, quantité vide, nulle ou non numérique, champ manquant, doublon. Chaque ligne concernée reçoit `AVerifier = oui`.

## 2. Liste Microsoft Lists « Pieces »

Créée à partir de `nomenclature_ia.csv`. Colonnes : Reference (colonne Titre, appelée `Title` dans Power Apps), Designation, Quantite (nombre), SousEnsemble, Poste, Statut, AVerifier, Commentaire.

## 3. Application Power Apps (canevas)

Formules écrites pour Power Apps en français : `;` sépare les arguments et `;;` enchaîne les actions (en anglais : `,` et `;`).

**Écran `scrListe`**

| Contrôle | Propriété | Formule |
|---|---|---|
| scrListe | OnVisible | `ClearCollect(colSousEnsembles; {Value: "Tous"});; Collect(colSousEnsembles; Distinct(Pieces; SousEnsemble))` |
| drpSousEnsemble | Items / Default | `colSousEnsembles` / `"Tous"` |
| txtRecherche | Default / HintText | `""` / `"Rechercher une référence ou une désignation"` |
| lblCompteur | Text | `"Lignes à vérifier : " & CountRows(Filter(Pieces; AVerifier = "oui"))` |
| galPieces | TemplateFill | `If(ThisItem.AVerifier = "oui"; RGBA(255; 244; 229; 1); Color.White)` |
| galPieces | OnSelect | `EditForm(frmPiece);; Navigate(scrPiece)` |
| lblBadge | Visible | `ThisItem.AVerifier = "oui"` |

`galPieces.Items` :
```
Sort(
    Filter(Pieces;
        IsBlank(drpSousEnsemble.Selected.Value) || drpSousEnsemble.Selected.Value = "Tous" || SousEnsemble = drpSousEnsemble.Selected.Value;
        StartsWith(Title; txtRecherche.Text) || StartsWith(Designation; txtRecherche.Text)
    );
    Title
)
```

`btnGenerer.DisplayMode` (génération possible seulement pour un sous-ensemble entièrement validé) :
```
If(IsBlank(drpSousEnsemble.Selected.Value) || drpSousEnsemble.Selected.Value = "Tous" || CountRows(Filter(Pieces; SousEnsemble = drpSousEnsemble.Selected.Value; AVerifier = "oui")) > 0; DisplayMode.Disabled; DisplayMode.Edit)
```

`btnGenerer.OnSelect` : construit la page HTML du dossier (une ligne par pièce, triée par poste, avec `Concat`), appelle le flux et ouvre le fichier :
```
Set(varSE; drpSousEnsemble.Selected.Value);;
Set(varHTML; "<html>…<h1>Dossier de montage – " & varSE & "</h1>…" &
    Concat(Sort(Filter(Pieces; SousEnsemble = varSE); Poste);
        "<tr><td>" & Poste & "</td><td>" & Title & "</td><td>" & Designation & "</td><td>" & Quantite & "</td></tr>") &
    "</table></body></html>");;
Set(varLien; GenererDossierMontage.Run(varSE; varHTML).lien);;
Notify("Dossier de montage généré"; NotificationType.Success);;
Launch(varLien)
```

**Écran `scrPiece`**

| Contrôle | Propriété | Formule |
|---|---|---|
| frmPiece | DataSource / Item | `Pieces` / `galPieces.Selected` |
| frmPiece | OnSuccess | `Notify("Pièce enregistrée"; NotificationType.Success);; Back()` |
| btnRetour | OnSelect | `ResetForm(frmPiece);; Back()` |
| btnEnregistrer | OnSelect | `SubmitForm(frmPiece)` |
| btnValider | OnSelect | `Patch(Pieces; galPieces.Selected; frmPiece.Updates; {AVerifier: "non"});; Notify("Ligne validée"; NotificationType.Success);; Back()` |
| btnValider | DisplayMode | `If(frmPiece.Updates.Quantite > 0; DisplayMode.Edit; DisplayMode.Disabled)` (règle métier : pas de validation d'une quantité nulle) |
| btnValider | Visible | `galPieces.Selected.AVerifier = "oui"` |

## 4. Flux Power Automate « GenererDossierMontage »

1. **Power Apps (V2)** — entrées texte `SousEnsemble` et `ContenuHTML`.
2. **OneDrive Entreprise — Créer un fichier** — `/TP Nomenclature/DossierMontage<SousEnsemble>.html`, contenu `ContenuHTML`.
3. **OneDrive Entreprise — Créer un lien de partage** — lecture seule, étendue Organisation.
4. **Répondre à une application ou un flux PowerApp** — sortie texte `lien` = URL web du fichier.

## Choix techniques et enseignements

- **Validation humaine** : l'IA propose, l'utilisateur métier corrige et valide ; le dossier ne peut être généré que pour un sous-ensemble entièrement validé.
- **Délégation** : `CountRows` et `Distinct` ne sont pas délégués à SharePoint (calcul local sur les 500 premières lignes par défaut). Suffisant ici ; pour une vraie nomenclature de plusieurs milliers de pièces, il faudrait des requêtes déléguables ou Dataverse.
- **Liste dans « Mes listes »** : l'action SharePoint de Power Automate trouve mal les espaces OneDrive personnels ; l'application envoie donc elle-même le contenu au flux, qui ne fait que créer le fichier.
- **Colonne Titre** : la colonne « Reference » importée remplace la colonne Titre de la liste ; Power Apps la nomme `Title`.

## Pistes d'amélioration

Stockage dans Dataverse ou un site SharePoint d'équipe, historique des versions de la nomenclature, dossier de montage en PDF ou à partir d'un modèle Word, extraction IA directement dans Power Platform (AI Builder), ordre des opérations par poste.
