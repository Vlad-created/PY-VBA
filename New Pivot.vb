Sub CreatePivotTable()
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long

    Set DSheet = Worksheets("202505")

    LastRow = DSheet.Cells(DSheet.Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, DSheet.Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    Set PCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=PRange)

    Set PSheet = Worksheets.Add

    Set PTable = PCache.CreatePivotTable( _
        TableDestination:=PSheet.Cells(1, 1), _
        TableName:="PivotTableUn")

    With PTable
    End With

    With PTable
        .PivotFields("Gender").Orientation = xlRowField
        
    With .PivotFields("DAS ID C")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of Gender"
    End With
End With
End Sub




' Another Pivot'

Sub CreatePivotTable()
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long

    
    Set DSheet = Worksheets("202505")

    
    LastRow = DSheet.Cells(DSheet.Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, DSheet.Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)
    
    Set PCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=PRange)
    
    Set PSheet = Worksheets.Add
    
    Set PTable = PCache.CreatePivotTable( _
        TableDestination:=PSheet.Cells(1, 1), _
        TableName:="HC")
    
    With PTable
        On Error Resume Next 
        .PivotFields("Global Country Grouping").Orientation = xlRowField
        
        If Err.Number <> 0 Then
            MsgBox "Error setting 'Global Country Grouping' as Row Field: " & Err.Description
            Err.Clear
        End If
        
        With .PivotFields("DAS ID C")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of DAS ID C"
        End With
        On Error GoTo 0 
    End With
End Sub