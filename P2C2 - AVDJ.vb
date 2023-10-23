Sub Mise_a_jour_reporting()

Dim mois As Integer
Dim mois_titre As String
Dim total_CA As Double

'Création d'une nouvelle colonne
Columns("K:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("K1").Select
'Nom de la colonne
ActiveCell.FormulaR1C1 = "Chiffre d'affaires"
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
    .Bold = True
End With

'ajout des lignes
Range("a1:a2").EntireRow.Insert

'calcul de la date du jour
Range("A1").Value = "=TODAY()"

'calul du mois
Range("B1").Value = "=MONTH(RC[-1])"

'création de la somme pour le calcul du CA total
Range("K1").Value = "=SUM(R[3]C:R[46]C)"
    
'chargement de la variable mois
mois = Range("B1").Value
    
    'selection du mois pour le titre du reporting
    Select Case mois
        Case 1
        mois_titre = "Janvier"
        Case 2
        mois_titre = "Fevrier"
        Case 3
        mois_titre = "Mars"
        Case 4
        mois_titre = "Avril"
        Case 5
        mois_titre = "Mai"
        Case 6
        mois_titre = "Juin"
        Case 7
        mois_titre = "Juillet"
        Case 8
        mois_titre = "Aout"
        Case 9
        mois_titre = "Septembre"
        Case 10
        mois_titre = "Octobre"
        Case 11
        mois_titre = "Novembre"
        Case 12
        mois_titre = "Decembre"
    End Select

'Titre du reporting
Range("E1").Value = "Reporting du mois de " & mois_titre
    
total_CA = Range("K1").Value
Range("K1").Select
'Coloration du total CA
    If total_CA <= 89999 Then
        'En rouge et fond jaune
        With Selection.Interior
            .Color = 65535
        End With
        With Selection.Font
            .Color = -16776961
        End With
    ElseIf total_CA >= 100000 Then
        'en vert
        With Selection.Font
            .Color = -11489280
        End With
    Else
        'en orange
        With Selection.Font
            .Color = -16727809
        End With
    End If

End Sub
