# LaitTrack - Import et rapport Excel

Application web pour importer un fichier Excel de suivi des paquets de lait par employe et telecharger un rapport Excel genere automatiquement.

Le projet utilise :

- **Backend** : Flask + openpyxl
- **Frontend** : React
- **Format d'entree** : fichier Excel `.xlsx` ou `.xls`

## Structure du projet

```text
lait-app/
|-- backend/
|   |-- app.py
|   `-- requirements.txt
|-- frontend/
|   |-- package.json
|   |-- package-lock.json
|   |-- public/
|   |   `-- index.html
|   `-- src/
|       |-- App.js
|       |-- App.css
|       `-- index.js
`-- README.md
```

## Fonctionnalites actuelles

- Import d'un fichier Excel par clic ou drag and drop.
- Validation du format `.xlsx` / `.xls`.
- Envoi du fichier au backend Flask.
- Generation d'un rapport Excel.
- Telechargement automatique du rapport.

Le dashboard a ete retire pour le moment. L'interface garde seulement l'import du fichier et le bouton de telechargement.

## Format du fichier Excel attendu

Le fichier source doit avoir cette structure :

| Employe | 01/06 | 02/06 | 03/06 |
|---------|-------|-------|-------|
| Ahmed   | 1     | 0     | 1     |
| Fatima  | 0     | 1     | 1     |

Regles :

- Colonne A : noms des employes.
- Ligne 1 : dates.
- Valeur `1` : l'employe a pris un paquet de lait.
- Valeur `0` : l'employe n'a pas pris de paquet.

## Installation

### 1. Backend

Depuis la racine du projet :

```powershell
cd backend
pip install -r requirements.txt
```

### 2. Frontend

Dans un autre terminal :

```powershell
cd frontend
npm install
```

## Demarrage du projet

### 1. Lancer le backend Flask

```powershell
cd C:\Users\nbenj\Desktop\lait-app\lait-app\backend
python app.py
```

Le backend demarre sur :

```text
http://localhost:5000
```

### 2. Lancer le frontend React

Dans un deuxieme terminal :

```powershell
cd C:\Users\nbenj\Desktop\lait-app\lait-app\frontend
npm start
```

Le frontend demarre sur :

```text
http://localhost:3000
```

## API

### `GET /api/health`

Verifie que l'API est active.

### `POST /api/download`

Recoit un fichier Excel et retourne un rapport Excel telechargeable.

Champ attendu dans le formulaire :

```text
file
```

### `POST /api/preview`

Analyse le fichier et retourne les donnees JSON.

Cette route existe encore dans le backend, mais elle n'est plus utilisee par le frontend actuel car le dashboard est masque.

## Fichiers importants

- `backend/app.py` : API Flask, lecture Excel et generation du rapport.
- `frontend/src/App.js` : interface React et logique d'import/telechargement.
- `frontend/src/App.css` : styles du theme clair.
- `frontend/public/index.html` : fichier HTML racine utilise par React.

## Commandes utiles

Build React :

```powershell
cd frontend
npm run build
```

Verifier les dependances Python :

```powershell
cd backend
pip install -r requirements.txt
```
