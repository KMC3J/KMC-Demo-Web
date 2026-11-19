# KMC – Démonstration web

Cette version est destinée à une démonstration technique avec un jeu de données **entièrement fictif**.

## Déploiement conseillé pour l'essai

1. Créer un dépôt GitHub **public de démonstration uniquement** (ne jamais y déposer les grands livres réels, exports Ciril, sauvegardes, identifiants ou données de la commune).
2. Ajouter uniquement `app.py`, `requirements.txt` et `README.md` à ce dépôt.
3. Sur Streamlit Community Cloud, créer une application depuis ce dépôt, choisir `app.py` comme fichier principal et lancer le déploiement.
4. Ouvrir l'URL fournie depuis les deux ordinateurs.

## Limites essentielles

- Données synthétiques, sans lien avec la comptabilité réelle de Saint-Aubin-lès-Elbeuf.
- Cette version ne doit recevoir aucun grand livre réel, fichier nominatif ou document interne.
- L'hébergement gratuit peut mettre l'application en veille et ne garantit pas un stockage permanent.
- La base locale à l'application peut être réinitialisée lors d'un redéploiement ou d'une remise en veille.
- L'application ne dispose pas encore d'un système complet de comptes utilisateurs et de contrôle d'accès multi-utilisateur.
- Ne pas utiliser pour une activité professionnelle réelle ni pour prendre des décisions budgétaires.

## Exécution locale de test

```bash
pip install -r requirements.txt
streamlit run app.py
```
