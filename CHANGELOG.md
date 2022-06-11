#Changelog

## 11/06/2022
## Ajoute des procédures :
### VerifSaisieForm()
- Vérification des saisies de tous les contrôles avec champs requis.
- Utilisation : Appliquer sur l'évennement BeforeUpdate du form { `Private Sub Form_BeforeUpdate....` }
### VerifSaisieControl()
- Vérifie la saisie du contrôle actif, permet de remetre la couleur du label à l'origine si saisie faite.
- Utilisation : Appliquer sur l'évennement BeforeUpdate du contrôle  { `=VerifSaisieControl()` }
### VerifSaisieRestaureLbl()
- Restaure tous les labels modifiés(couleur texte), à utiliser sur l'évennement Cancel du form.
- Utilisation : Appliquer sur l'évennement Undo du form, {`=VerifSaisieRestaureLbl()`}cela restaure la couleur des texte label en cas d'annulation dans le form.
