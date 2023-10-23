'\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\ Module 1 /////////
'////////////////////////////

Sub demarrer_form()

UserForm1.Show

End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\ Userform 1 /////////
'//////////////////////////////

Option Explicit
Dim seuil As Double

Private Sub CommandButton1_Click()



'test si une feuille est déja présente
On Error GoTo GérerErreur
Sheets.Add(After:=ActiveSheet).Name = "Calcul"
Sheets("Data").Select



'Déclaration des variables
Dim ligne_date As Integer
Dim test_reporting As String
Dim date_du_jour As String
Dim test As Object

'chargement
seuil = TextBox2.Value
Unload UserForm1

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

'ajout des lignes
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

'création de la somme pour le calcul du CA total
Range("K1").Select
ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[46]C)"

'sous_programme pour le titre
UserForm1.titre_reporting
'sous programme pour la mise en forme du CA
UserForm1.mise_en_forme_CA (seuil)
'sous programme pour la mise en forme de Glutenfree
UserForm1.glutenfree

'changement des données de la colonne Bio
Range("F4:F100").Select
Selection.Replace What:="Oui", Replacement:="Bio", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

Sheets("Calcul").Range("A1").Value = "Calcul de la quantité moyenne de vente"
Sheets("Calcul").Range("A2").Value = "Calcul de la plus petite vente"
Sheets("Calcul").Range("A3").Value = "calcul de la plus grande vente"
Sheets("Calcul").Range("A4").Value = "nombre de vente à 0"
Sheets("Calcul").Range("A5").Value = "Somme du nombre de vente"
Sheets("Calcul").Range("B1").Value = "=AVERAGE(Data!R[1]C[8]:R[44]C[8])"
Sheets("Calcul").Range("B2").Value = "=MIN(Data!RC[8]:R[43]C[8])"
Sheets("Calcul").Range("B3").Value = "=MAX(Data!R[-1]C[8]:R[42]C[8])"
Sheets("Calcul").Range("B4").Value = "=COUNTIF(Data!R[-1]C[8]:R[43]C[8],0)"
Sheets("Calcul").Range("B5").Value = "=SUM(Data!R[-3]C[8]:R[40]C[8])"
Exit Sub

GérerErreur:
Application.DisplayAlerts = False
Unload UserForm1
ActiveSheet.Delete
Application.DisplayAlerts = True
MsgBox ("Mise à jour déjà faite")

End Sub

Sub glutenfree()

Dim modif_var As String
Dim i As Integer

i = 4
Do While i < 100
    modif_var = Range("E" & i).Value
    If modif_var = "Oui" Then
    Range("E" & i).Value = "Gluten Free"
    Else
    End If
i = i + 1
Loop

End Sub

Sub mise_en_forme_CA(seuil As Integer)

Dim test_CA As Double
Dim i As Integer
'Mise en forme des cellules CA
For i = 4 To 100
test_CA = Range("K" & i).Value
Range("K" & i).Select
'En rouge et jaune
    If test_CA < seuil Then
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

End Sub

Sub titre_reporting()

Dim mois As Integer
Dim mois_titre As String

    'calcul de la date du jour
    Range("A1").Value = "=TODAY()"
    
    'calul du mois
    Range("B1").Value = "=MONTH(RC[-1])"
    
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

End Sub

Private Sub CommandButton3_Click()

    ActiveWorkbook.SaveAs Filename:="D:\Export\test.xlsb", _
        FileFormat:=xlExcel12, CreateBackup:=False
        
End Sub

Private Sub CommandButton4_Click()
Unload UserForm1
Exit Sub
End Sub
