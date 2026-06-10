# 🥛 LaitTrack — Suivi automatisé des paquets de lait

Application web complète pour traiter et visualiser les données de prise de lait des employés.

## Structure du projet

```
lait-app/
├── backend/
│   ├── app.py              ← API Flask (Python)
│   └── requirements.txt    ← Dépendances Python
└── frontend/
    └── index.html          ← Application frontend (HTML/JS, zéro dépendance)
```

## 🚀 Lancement

### 1. Démarrer le backend (API)

```bash
cd backend
pip install -r requirements.txt
python app.py
```

Le serveur démarre sur **http://localhost:5000**

### 2. Ouvrir le frontend

Ouvrez simplement `frontend/index.html` dans votre navigateur.
> ⚠️ Le backend doit être lancé pour que l'app fonctionne.

---

## 📊 Format du fichier Excel attendu

| Employé         | 01/06 | 02/06 | 03/06 | … |
|-----------------|-------|-------|-------|---|
| Ahmed Ben Ali   | 1     | 0     | 1     | … |
| Fatima Khelil   | 0     | 1     | 1     | … |

- **Colonne A** : Noms des employés (ligne 1 = entête "Employé")
- **Colonnes suivantes** : Une colonne par jour (entête = date)
- **Valeurs** : `1` = a pris un paquet de lait, `0` = n'a pas pris

---

## ✨ Fonctionnalités

- **Drag & drop** du fichier Excel
- **Dashboard interactif** :
  - KPIs globaux (total paquets, taux, top employé)
  - Graphique à barres par employé
  - Donut chart de répartition
  - Carte de chaleur (heatmap) jour par jour
  - Tableau de classement détaillé
- **Export Excel** du rapport avec 2 feuilles :
  - Détail journalier (données colorisées + totaux automatiques)
  - Récapitulatif (résumé par employé avec pourcentages)

---

## 🔌 API Endpoints

| Méthode | Route | Description |
|---------|-------|-------------|
| `POST` | `/api/preview` | Analyse le fichier, retourne JSON |
| `POST` | `/api/download` | Génère et télécharge le rapport Excel |
| `GET` | `/api/health` | Vérification que l'API est active |
