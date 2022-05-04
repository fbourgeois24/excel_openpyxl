# Openpyxl lib
## Instancier la classe
`fichier_excel = excel_file(chemin/du/fichier.xlsx)`

## Ouvrir le classeur
`fichier_excel.open(data_only=False)`
data_only converti toutes les formules en données "brutes" dans le classeur

## Lire une ou plusieurs cellules dans le classeur
`fichier_excel.read(range="A1", sheet=0)`
- range : plage de cellules à lire
Peut être au format "A1" pour une cellule unique ou "A1:B2" pour une plage de cellules
Si cellule simple ex. "A1" la valeur est renvoyée directement
Si plage ex. "A1:B2" une liste à deux niveaux est renvoyée, le premier niveau contient les lignes donc ici le résultat renvoyé sera : `[[A1, B1], [A2, B2]]`
- sheet : feuille dans laquelle lire les données (par défaut la première)
On peut accéder aux feuilles par leur indice ou leur nom