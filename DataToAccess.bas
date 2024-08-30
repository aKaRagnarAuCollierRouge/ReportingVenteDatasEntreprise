Attribute VB_Name = "DataToAccess"
Sub EnvoyerDataVersAccess()

Dim db As Object
Dim rs As Object
Dim appAccess As Object
Dim strSQL As String
Dim ws As Worksheet
Dim ligneMax As Integer
Dim cheminAccess As String

Range("B1").Select
Selection.End(xlDown).Select
ligne_max = ActiveCell.Row

'Acces a la base de données

cheminAccess = "C:\Users\Baptiste\Documents\PortfolioProject\VBA Project\Projet VBA 2\Database.accdb"

'Nom de la feuille contenant des données

Set ws = ThisWorkbook.Sheets("Data")

'Ouverture de la base de données

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase cheminAccess
Set db = appAccess.CurrentDb
'Nom de la table dans access
Set rs = db.OpenRecordset("Data")


'Boucle qui remplit les lignes une à une en choisissant les données


For i = 2 To ligne_max
    'Ajoute un nouvelle enregistrement
    
    rs.AddNew
    rs.Fields("Date").Value = ws.Cells(i, 1).Value
    rs.Fields("Référence").Value = ws.Cells(i, 2).Value
    rs.Fields("Poids").Value = ws.Cells(i, 3).Value
    rs.Fields("Désignation").Value = ws.Cells(i, 4).Value
    rs.Fields("GLUTEN FREE").Value = ws.Cells(i, 5).Value
    rs.Fields("BIO").Value = ws.Cells(i, 6).Value
    rs.Fields("Code Marque").Value = ws.Cells(i, 7).Value
    rs.Fields("Marque").Value = ws.Cells(i, 8).Value
    rs.Fields("TVA").Value = ws.Cells(i, 9).Value
    rs.Fields("Prix de Vente").Value = ws.Cells(i, 10).Value
    rs.Fields("Quantité mois").Value = ws.Cells(i, 11).Value
    rs.Fields("CA").Value = ws.Cells(i, 12).Value
    rs.Fields("Fournisseur").Value = ws.Cells(i, 13).Value
    'Mise à jour de la table
    rs.Update

Next i

'Fermeture de la BDD

app.Access.DoCmd.Quit acQuitSaveAll

'Libérez les objets

Set rs = Nothing
Set db = Nothing
Set appAccess = Nothing

End Sub
