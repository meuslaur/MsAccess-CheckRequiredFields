Attribute VB_Name = "MD_VerifSaisie"
'@Folder("Dev")
Option Compare Database
Option Explicit
' ------------------------------------------------------
' Name:    MD_VerifSaisie
' Kind:    Module
' Purpose: Fonctions de vérification pour le champs à saisie obligatoire.
' Author:  Laurent
' Date:    09/06/2022 - 14:08
' DateMod:
' ------------------------------------------------------

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
Private m_oFrm       As Form
Private m_oCtr       As Control
Private Const REPCOL As String = "#"
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' ----------------------------------------------------------------
'// Vérification de la saisie des champs obligatoire dans le form.
' ----------------------------------------------------------------
' Procedure Nom:            VerifChampSaisieRequi
' Sujet:                    Vérification saisie obligatoire des controles (Texte et Zone de liste) du formulaire
'                           La vérification est faite sur les champs la table source.
'                           les controles non null, non visible, non activés sont ignorés
' Procedure Kind:   Function
' Procedure Access: Public
'
'=== Paramètres ===
' oFrm (Form):
'==================
'
' Return Type: String Retourne la liste des champs à Saisir, ou une chaine vide si ok.
'
' Author:  Laurent
' Date:    15/04/2022
' DateMod: 10/06/2022 - 08:05
' ----------------------------------------------------------------
Public Function VerifChampSaisieRequi(sFrmName As String) As String
On Error GoTo ERR_VerifChampSaisieRequi

    Dim sSource  As String
    Dim sMsg     As String
    Dim bCheck   As Boolean
    Dim bRequis  As Boolean

    Set m_oFrm = Application.Forms(sFrmName)

    '// Parcourir les controles du form...
    For Each m_oCtr In m_oFrm.Controls

        sSource = vbNullString
'Debug.Print m_oCtr.Name 'TOGO: Test
        Select Case m_oCtr.ControlType

            Case acCheckBox, acOptionButton, acToggleButton         '106 Contrôle CheckBox (acOptionGroup) '105 Contrôle OptionButton (acOptionGroup)122 Contrôle ToggleButton (acOptionGroup) bouton bascule                         octr.parent.ControlType =107
                sSource = Nz(m_oCtr.ControlSource, vbNullString)

            Case acListBox, acComboBox                              '111 Contrôle ComboBox    RowSource RowSourceType   '110 Contrôle ListBox     RowSource RowSourceType
                sSource = Nz(m_oCtr.ControlSource, vbNullString)

            Case acOptionGroup                                      '107 Contrôle OptionGroup
                sSource = Nz(m_oCtr.ControlSource, vbNullString)

            Case acTextBox                                          '109 Contrôle TextBox
                sSource = Nz(m_oCtr.ControlSource, vbNullString)

        End Select

        bCheck = ((sSource <> vbNullString) And (Left$(sSource, 1) <> "="))     '// Source valide ?
        If bCheck Then bCheck = ControlValide()                                 '// Controle valide ?...
        If bCheck Then bRequis = ChampSaisieObligatoire(sSource)                '// Vérifie si la saisie est obligatoire...
        If bRequis Then bCheck = IIf(IsNull(m_oCtr.Value) Or (m_oCtr.Value = vbNullString), True, False)      '// Test si le ctr contient une valeur.

        If (bCheck And bRequis) Then
            sMsg = sMsg & sSource & ", "
            LblColorApplique    '// Met le label en rouge...
        ElseIf bRequis Then
            LblColorRestaure    '// Saisie faite, remettre sont label en l'état...
        End If

        bRequis = False
    Next

    VerifChampSaisieRequi = sMsg

SORTIE_VerifChampSaisieRequi:
    Set m_oCtr = Nothing
    Set m_oFrm = Nothing
    Exit Function

ERR_VerifChampSaisieRequi:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  VerifSaisie.MD_VerifSaisie.VerifChampSaisieRequi, ligne " & Erl & "."
    Resume SORTIE_VerifChampSaisieRequi
End Function
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ################################ PRIVATE SUB/FUNC ####################################

' ----------------------------------------------------------------
' Procedure Nom:                ChampSaisieObligatoire
' Sujet:                        Vérifier si la saisie du champ est obligatoire dans la table.0
' Procedure Kind:               Function
' Procedure Access:             Private
' Parameter oRst (Recordset):   Table a vérifier
' Return Type:                  Boolean, True si obligatoire, ou erreur
' Author:                       Laurent
' Date:                         15/04/2022
' ----------------------------------------------------------------
Private Function ChampSaisieObligatoire(sChamp As String) As Boolean
On Error GoTo ERR_ChampSaisieObligatoire

    ChampSaisieObligatoire = m_oFrm.Recordset.Fields(sChamp).Properties("Required")

SORTIE_ChampSaisieObligatoire:
    Exit Function

ERR_ChampSaisieObligatoire:
    MsgBox "Erreur " & Err.Number & vbCrLf & " (" & Err.Description & ")" & vbCrLf & "Dans la Function : 'ChampSaisieObligatoire', ligne " & Erl & "."
    Resume SORTIE_ChampSaisieObligatoire
End Function

' ----------------------------------------------------------------
' Procedure Nom:            ControlValide
' Sujet:                    Vérification si visible, activé et non vérouillé.
' Procedure Kind:           Function
' Procedure Access:         Private
' Return Type:              Boolean, TRUE si Visible activé non vérouillé.
' Author:                   Laurent
' Date:                     15/04/2022
' ----------------------------------------------------------------
Private Function ControlValide() As Boolean
    ControlValide = ((m_oCtr.Visible) And (m_oCtr.Enabled) And (Not m_oCtr.Locked))
End Function

Private Function LblColorSauve() As Boolean
'// Pour les contrôles avec source requis, et si le contrôle à un label lié,
'// sauve la couleur texte du label, dans StatusBarText du contrôle.
'// si StatusBarText contient déjà qq chose, on sauve la couleur à la fin du texte.

    If (m_oCtr.Controls.Count = 0) Then Exit Function               '// Pas de label, on sort.

    LblColorSauve = True

    If InStr(1, m_oCtr.StatusBarText, REPCOL) Then Exit Function    '// Déjà sauver, on sort.

    Dim sColor  As String
    Dim sTxt    As String

    sColor = Str$(m_oCtr.Controls(0).ForeColor) '// renvoi la couleur avec un espace devant ??
    sColor = LTrim$(sColor)
    sTxt = Nz(m_oCtr.StatusBarText, vbNullString)
    sTxt = sTxt & REPCOL & sColor & REPCOL
    m_oCtr.StatusBarText = sTxt

End Function

Private Sub LblColorApplique()
'// Pour les contrôles avec source requis, et si le contrôle à un label lié,
'// met le texte du label lié au contrôle en rouge.
    If (LblColorSauve() = False) Then Exit Sub   '// Pas de label, on sort.

    Dim sLblName As String
    sLblName = m_oCtr.Controls(0).Name
    m_oFrm.Controls(sLblName).ForeColor = RGB(255, 0, 0)

End Sub

Private Sub LblColorRestaure()
'// Pour les contrôles avec source requis, et si le contrôle à un label lié,
'// on restaure la couleur du texte d'origine du label lié au contrôle.
'// La couleur et stockée dans la prop StatusBarText du contrôle parent,
'// encradré par la Const REPCOL.

    If (m_oCtr.Controls.Count = 0) Then Exit Sub    '// Pas de label, on sort.

    Dim sBarTxt  As String
    Dim sLblName As String
    Dim sColor   As String
    Dim lLen     As Long
    Dim lPosD    As Long
    Dim lPosF    As Long

    sBarTxt = m_oCtr.StatusBarText

    '// Obtenir la position de la couleur.
    lPosD = InStr(1, sBarTxt, REPCOL)
    lPosF = InStr(lPosD + 2, sBarTxt, REPCOL)

    If (lPosD = 0) Then Exit Sub       '// Pas encore sauvegrader, on sort.

    '// Restaure la valeur de la prop StatusBarText.
    If (lPosD = 1) Then
        m_oCtr.StatusBarText = vbNullString
    Else
        m_oCtr.StatusBarText = Left$(sBarTxt, lPosD - 1)
    End If

    '// Extraction de la couleur.
    lLen = (lPosF - lPosD)
    sColor = Mid$(sBarTxt, lPosD + 1, lLen - 1)

    '// Restaure la couleur.
    sLblName = m_oCtr.Controls(0).Name
    m_oFrm.Controls(sLblName).ForeColor = Val(sColor)

End Sub
'// ################################# END PRIV. SUB/FUNC #################################
