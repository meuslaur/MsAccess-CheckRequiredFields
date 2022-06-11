# MsAccess-CheckRequiredFields
Utilitaire pour contrôler les saisies dans un formulaire
## Cet utilitaire contrôle les saisies dans un formulaire avant la MàJ de celui-ci
### Vérification des saisies des contrôles :
-  acCheckBox, acOptionButton, acToggleButton, acOptionGroup
-  acListBox, acComboBox 
-  acTextBox
### Contrôles éffectués :
- Vérifier la source du contrôle (`Not Null or <> "="`)
- Vérifier l'état du contrôle (`Enabled, Visible, Not Locked`)
- Vérifier si le champs source du contrôle et Required dans la table source.
- Vérifier si le contrôle contient une saisie (`<> VbNullString or Not IsNull`)
### Opérations effectuées :
- Stock les champs Required non saisis
- Modifie la couleur de texte du label du contrôle :
- - Sauvegrade la couler d'origine { `Function LblColorSauve()` }
- - Modifie la couleur du texte du label { `Sub LblColorApplique()` }
- - Restaure la couleur texte label si saisie correcte { `Sub LblColorRestaure()` }
- La couleur d'origine est enregistrée dans la prorpriété StatusBarText du contrôle
- Si StatusBarText contient du texte ils est restauré a l'origine.
