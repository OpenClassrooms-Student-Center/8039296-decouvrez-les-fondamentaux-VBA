Sub Mise_a_jour_reporting()

Dim mois As Integer
Dim mois_titre As String
Dim test_CA As Double
Dim modif_var As String


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
    
test_CA = Range("K1").Value

'Mise en forme des cellules CA

For i = 4 To 100

test_CA = Range("K" & i).Value
Range("K" & i).Select
'En rouge et jaune
    If test_CA < 1000 Then
        'En rouge
        With Selection.Font
            .Color = -16776961
            .Bold = True
        End With
    Else
        'en vert
        With Selection.Font
            .Color = -11489280
            .Bold = True
        End With
    End If
Next i

'changement des données de la colonne Glutenfree
i = 4
Do While i < 100
    modif_var = Range("E" & i).Value
    If modif_var = "Oui" Then
    Range("E" & i).Value = "Gluten Free"
    Else
    End If
i = i + 1
Loop

'changement des données de la colonne Bio
Range("F4:F100").Select
Selection.Replace What:="Oui", Replacement:="Bio", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

End Sub



