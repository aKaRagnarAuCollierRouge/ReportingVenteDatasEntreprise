Attribute VB_Name = "Module1"


Sub AnalyseVariableQualitative()
Dim EndRange As Long
Dim RangeC As String

'Référence

EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("B2:B10000"))
RangeC = "B2:B" & CStr(EndRange)

Sheets("Reporting").Range("L14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))
AffichageValeursUniques RangeC, "L15"


'Glutenfree

EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("E2:E10000"))
RangeC = "E2:E" & CStr(EndRange)

Sheets("Reporting").Range("M14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))
AffichageValeursUniques RangeC, "M15"

'Bio
EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("F2:F10000"))
RangeC = "F2:F" & CStr(EndRange)

Sheets("Reporting").Range("N14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))
AffichageValeursUniques RangeC, "F15"

'Code marque

EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("G2:G10000"))
RangeC = "G2:G" & CStr(EndRange)

Sheets("Reporting").Range("O14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))
AffichageValeursUniques RangeC, "O15"

'Fournisseur
EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("L2:L10000"))
RangeC = "L2:L" & CStr(EndRange)

Sheets("Reporting").Range("P14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))
AffichageValeursUniques RangeC, "P15"
End Sub

Sub NbValeursUniques(ByVal zone_analyse As String, ByVal valeur_unique As String)
Dim C As Range
Dim Dico As Object

Set Dico = CreateObject("Scripting.Dictionary")
For Each C In Sheets("Data").Range(zone_analyse).SpecialCells(xlCellTypeVisible)
    If Not Dico.exists(C.Value) Then Dico.Add C.Value, C.Value
    
Sheets("Reporting").Range(valeur_unique).Resize(Dico.Count) = Application.Transpose(Dico)


End Sub


Sub AnalyseVariableQuantitative()

Dim EndRange As Long
Dim RangeC As String
'Prix de vente
EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("I2:I10000"))
RangeC = "I2:I" & CStr(EndRange)

Sheets("Reporting").Range("L4").Value = WorksheetFunction.Min(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("L5").Value = WorksheetFunction.Max(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("L6").Value = WorksheetFunction.Average(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("L7").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))


'TVA
EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("H2:H10000"))
RangeC = "H2:H" & CStr(EndRange)
Sheets("Reporting").Range("M4").Value = WorksheetFunction.Min(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("M5").Value = WorksheetFunction.Max(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("M6").Value = WorksheetFunction.Average(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("M7").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))


'Quantités
EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("J2:J10000"))
RangeC = "J2:J" & CStr(EndRange)

Sheets("Reporting").Range("N4").Value = WorksheetFunction.Min(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("N5").Value = WorksheetFunction.Max(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("N6").Value = WorksheetFunction.Average(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("N7").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))


'Poids

EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("C2:C10000"))
RangeC = "C2:C" & CStr(EndRange)


Sheets("Reporting").Range("O4").Value = WorksheetFunction.Min(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("O5").Value = WorksheetFunction.Max(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("O6").Value = WorksheetFunction.Average(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("O7").Value = WorksheetFunction.CountBlank(Sheets("Data").Range(RangeC))


'CA
EndRange = CountEmptyCellsFromBottom(Sheets("Data").Range("K2:K10000"))
RangeC = "K2:K" & CStr(EndRange)


Sheets("Reporting").Range("P4").Value = WorksheetFunction.Min(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("P5").Value = WorksheetFunction.Max(Sheets("Data").Range(RangeC))
Sheets("Reporting").Range("P6").Value = WorksheetFunction.Average(Sheets("Data").Range(RangeC))



End Sub

Sub AffichageValeursUniques(zone_analyse As String, ByVal valeur_unique As String)

    Dim C As Range
    Dim Dico As Object
    
    Set Dico = CreateObject("Scripting.Dictionary")
    
    For Each C In Sheets("Data").Range(zone_analyse).SpecialCells(xlCellTypeVisible)
        If Not Dico.exists(C.Value) Then Dico.Add C.Value, C.Value
    Next C
    Sheets("Reporting").Range(valeur_unique).Resize(Dico.Count) = Application.Transpose(Dico.Keys)
    

End Sub

Function CountEmptyCellsFromBottom(rng As Range) As Long
    Dim cell As Range
    Dim i As Long
    
    ' Boucler à travers les cellules de la plage en partant du bas
    For i = rng.Rows.Count To 1 Step -1
        If rng.Cells(i, 1).Value <> "" Then
            CountEmptyCellsFromBottom = rng.Cells(i, 1).Row
            Exit Function
        End If
    Next i
    
    ' Si toutes les cellules sont vides, retourner la première ligne de la plage
    CountEmptyCellsFromBottom = rng.Cells(1, 1).Row
End Function
