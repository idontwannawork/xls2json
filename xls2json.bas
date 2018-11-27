Attribute VB_Name = "Module1"
Sub xls2json()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim tbl As Range
    Set tbl = ws.Cells(1, 1).CurrentRegion

    maxRow = tbl.Rows.Count
    maxCol = tbl.Columns.Count
    ReDim headerItem(maxCol - 1)
    ReDim recordItem(maxRow - 2)
    
    '1行目は項目名
    For c = 1 To maxCol
        headerItem(c - 1) = tbl(1, c)
    Next c
    
    'レコードの内容を取得
    For r = 2 To maxRow
    
        ReDim temps(maxCol - 1)
        
        For c = 1 To maxCol
            '各行の各セルを取得し見出しと組み合わせる。
            temps(c - 1) = """" & headerItem(c - 1) & """" & ":" & """" & tbl(r, c) & """"
        Next c
        
        recordItem(r - 2) = "{" & Join(temps, ",") & "}"
        
    Next r
    
    json = "[" & Join(recordItem, ",") & "]"
    
    '出力先の選択
    outputFilename = Application.GetSaveAsFilename("output.json", "JSON(*.json),*.json")
    
    If outputFilename = False Then
        Exit Sub
    End If
    
    Dim pre As Object
    Set pre = CreateObject("ADODB.Stream")
    
    'BOMなしUTF-8として出力
    With pre
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText json
        .Position = 0
        .Type = 1
        .Position = 3
        Dim buf As Variant
        buf = .Read()
        .Position = 0
        .Write buf
        .SetEOS
        .SaveToFile outputFilename, 2
        .Close
    End With
    
End Sub
