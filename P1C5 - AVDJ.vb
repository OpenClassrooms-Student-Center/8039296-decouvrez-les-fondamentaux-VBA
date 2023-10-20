Sub Mise_a_jour()
'
' Mise_a_jour Macro
'
    'Création d'une nouvelle colonne
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K1").Select
    'Nom de la colonne
    ActiveCell.FormulaR1C1 = "Chiffre d'affaires1"
    Range("K2").Select
    Application.CutCopyMode = False
    'Formule pour calculer le chiffre d'affaires par produit
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K45")
    Range("K2:K45").Select
    'changement du format
    Selection.NumberFormat = "#,##0.00 $"
    Columns("K:K").Select
    Range("K2").Activate
    'Changement de la couleur et gras sur la colonne
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .bold = True
    End With
End Sub