# Suivi Maintenance Bornes

Application web statique pour consolider des classeurs Excel de maintenance (2020 → 2026) et suivre les contrats.

## Démarrage rapide

1. Démarrez un serveur local (exemple) :

```bash
python -m http.server 8000
```

2. Ouvrez `http://localhost:8000` dans un navigateur.
3. Importez vos fichiers Excel via le bouton en haut de page.

## Rappels des colonnes utilisées

- R : Site avec maintenance (Oui / Non)
- O : Date contrat de maintenance (ou "Pas de contrat")
- N : Date facture (utilisée si O = Pas de contrat)
- L : Date maintenance 2ème année 1/2
- M : Date maintenance 3ème année 2/2
- I : Bornes
- H : Marque / série
- D : Site
- E : Adresse (avec lien Google Maps)

La colonne de région est attendue en T (modifiable dans le code si besoin).
