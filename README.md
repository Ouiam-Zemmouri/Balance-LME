# Balance LME — COFICAB

Dashboard Streamlit indépendant pour l'analyse de la balance LME (méthode FIFO) :
écart de valorisation entre les ventes et le stock + achats consommés, par fixation
LME, pour COFICAB Kenitra et COFICAB Maroc.

## Structure du repo

```
balance-lme/
├── balance_app.py        ← l'application Streamlit
├── requirements.txt      ← dépendances Python
├── balance_files/        ← tes fichiers Excel mensuels (déposés ici)
│   ├── COF KT - LME balance 06.2026.xlsx
│   ├── COF MA - LME balance 03.2026.xlsx
│   └── ...
└── README.md
```

## Ajouter un nouveau mois

1. Va dans le dossier `balance_files/` sur GitHub
2. "Add file" → "Upload files" → dépose ton/tes fichier(s) `.xlsx`
3. "Commit changes"
4. Ouvre le dashboard → clique **🔄 Refresh Data** dans la barre latérale

Aucune conversion, aucun script à lancer : le dashboard lit tous les `.xlsx` du
dossier `balance_files/` automatiquement à chaque ouverture (ou au clic sur
Refresh Data).

Le nom de fichier peut suivre n'importe lequel de ces formats — l'entité, le
mois et l'année sont détectés automatiquement :
- `COF KT - LME balance 06.2026.xlsx`
- `COF_KT_-_LME_balance_06_2026.xlsx`
- `COF-MR-LME-balance-03-2026.xlsx`

## Déploiement sur Streamlit Cloud

1. [share.streamlit.io](https://share.streamlit.io) → **New app**
2. Repository : ce repo
3. Branch : `main`
4. Main file path : `balance_app.py`
5. Deploy

## Mot de passe

Le mot de passe d'accès est défini dans `balance_app.py` (variable `PASSWORD`,
en haut du fichier).
