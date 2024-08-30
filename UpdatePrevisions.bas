Attribute VB_Name = "UpdatePrevisions"
Sub prevision()

Dim i As Integer
Dim mois_m1 As String

i = 32

'recherche de la plage de donn�es contenant les mois et les donn�es
Do While Range("C" & i).Value <> "Total g�n�ral"
i = i + 1
Loop

'On r�cup�re le dernier mois et on l'incr�mente d'un mois pour faire la pr�vision
mois_m1 = DateAdd("m", 1, WorksheetFunction.Max(Sheets("Analyses").Range("C32:C" & (i - 1))))

'Pr�vision pour BRESCIA
ActiveWorkbook.CreateForecastSheet Timeline:=Sheets("Analyses").Range( _
    "C32:C" & (i - 1)), Values:=Sheets("Analyses").Range("D32:D" & (i - 1)), ForecastEnd:= _
    mois_m1, ConfInt:=0.95, Seasonality:=1, ChartType:= _
    xlForecastChartTypeLine, Aggregation:=xlForecastAggregationAverage, _
    DataCompletion:=xlForecastDataCompletionInterpolate, ShowStatsTable:=False
ActiveSheet.Name = "Pr�vision BRESCIA"
    
'Pr�vision pour LECCE
ActiveWorkbook.CreateForecastSheet Timeline:=Sheets("Analyses").Range( _
    "C32:C" & (i - 1)), Values:=Sheets("Analyses").Range("E32:E" & (i - 1)), ForecastEnd:= _
    mois_m1, ConfInt:=0.95, Seasonality:=1, ChartType:= _
    xlForecastChartTypeLine, Aggregation:=xlForecastAggregationAverage, _
    DataCompletion:=xlForecastDataCompletionInterpolate, ShowStatsTable:=False
ActiveSheet.Name = "Pr�vision LECCE"

End Sub

Sub MAJ_Graphique()
    
ActiveWorkbook.RefreshAll

End Sub

