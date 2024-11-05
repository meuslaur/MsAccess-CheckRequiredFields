<img align="left" src="https://github.com/meuslaur/meuslaur/blob/main/Logo_MsAccess.png" width="64px">

# MsAccess-CheckRequiredFields
Utilitaire pour contrôler les saisies dans un formulaire
## Cet utilitaire contrôle les saisies dans un formulaire avant la MàJ de celui-ci
### Vérification des saisies des contrôles :
-  acCheckBox, acOptionButton, acToggleButton, acOptionGroup
-  acListBox, acComboBox 
-  acTextBox
### Contrôles effectués :
- Vérifier la source du contrôle (`Not Null or <> "="`)
- Vérifier l'état du contrôle (`Enabled, Visible, Not Locked`)
- Vérifier si le champs source du contrôle est Required dans la table source.
- Vérifier si le contrôle contient une saisie (`<> VbNullString or Not IsNull`)
### Opérations effectuées :
- Stock les champs Required non saisis
- Modifie la couleur de texte du label du contrôle :
- - Sauvegarde la couleur d'origine { `Function LblColorSauve()` }
- - Modifie la couleur du texte du label { `Sub LblColorApplique()` }
- - Restaure la couleur texte label si saisie correcte { `Sub LblColorRestaure()` }
- La couleur d'origine est enregistrée dans la propriété StatusBarText du contrôle
- Si StatusBarText contient du texte ils est restauré a l'origine.
### Utilisation :
- La fonction `VerifChampSaisieRequi()` retourne la liste des champs requis dans la table source qui n'ont pas était validés.
- Insèrez le code suivat sur l'évennement `BeforeUpdate` du formulaire :
```VBA
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo ERR_MajF
    Dim sRep As String

    sRep = VerifSaisieForm()
    If (sRep <> vbNullString) Then
        Cancel = True
        '// Your code here
        '// Your code here
        Exit Sub
    End If
'.....
End Sub
```

- Si le contrôle n'a pas de Label lié rien n'est modifié, dans ce cas vous pouvez utiliser la propriété `BorderColor`, 
en modifiant le code dans les procédures LblColorApplique(), LblColorSauve() et LblColorRestaure().

## Résumé

|   Créer le|   2022/06/11|
| - | - |
|   Auteur| [@meuslau](https://github.com/meuslaur)|
|   Catégorie|   MsAccess|
|   Type|   Utilitaire|
|   Langage|   VBA|
