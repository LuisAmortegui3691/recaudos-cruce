Attribute VB_Name = "Módulo1"
Sub macroCruceRecaudos()

    Dim rutaLibros As String
    Dim libroInicial As String, libroFinal As String
    Dim rutaMacro As String
    Dim ultimaFilaDocFinal As Long, ultimaFilaDocInicial As Long
    Dim recaudoLF As String, polizaLF As String, remisionPF As String, placaLF As String, identificaLF As String
    Dim i As Long, j As Long
    Dim valorLibriInicial As String, valorLibroFinal As String
    Dim validationLibroInicial As String, rgbLibroInicial As String, numberDocLibroInicial As String, dateLibroInicial As String
    
    rutaMacro = ThisWorkbook.Sheets("main").Range("c4").Value
    
    If rutaMacro = "" Then
        MsgBox "La ruta esta vacia por favor diligenciela en la celda C4"
    ElseIf Right(rutaMacro, 1) <> "\" Then
        rutaMacro = rutaMacro & "\"
    End If
    

    libroInicial = ThisWorkbook.Sheets("main").Range("C2").Value
    libroFinal = ThisWorkbook.Sheets("main").Range("C3").Value
    
    ' Abrir documento Inicial y arreglar
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=rutaMacro & libroFinal
    Application.DisplayAlerts = True
    
    Workbooks(libroFinal).Sheets(1).Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Workbooks(libroFinal).Sheets(1).Columns("A:A").Select
    Workbooks(libroFinal).Sheets(1).Range("A3").Activate
    Selection.Insert Shift:=xlToRight
    Workbooks(libroFinal).Sheets(1).Range("AP1").Value = "validation"
    Workbooks(libroFinal).Sheets(1).Range("AQ1").Value = "rgb"
    Workbooks(libroFinal).Sheets(1).Range("AR1").Value = "numbert document"
    Workbooks(libroFinal).Sheets(1).Range("AS1").Value = "date"
    Workbooks(libroFinal).Sheets(1).Rows("1:1").Select
    Selection.AutoFilter
    
    Workbooks(libroFinal).Sheets(1).Range("$A$1:$AS$337").AutoFilter Field:=2, Criteria1:="=AN", _
        Operator:=xlOr, Criteria2:="=PD"
    Workbooks(libroFinal).Sheets(1).Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete Shift:=xlUp
    Workbooks(libroFinal).Sheets(1).Range("$A$1:$AS$151").AutoFilter Field:=2
    Range("C2").Select
    
    ultimaFilaDocFinal = Workbooks(libroFinal).Sheets(1).Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 2 To ultimaFilaDocFinal
        recaudoLF = Workbooks(libroFinal).Sheets(1).Range("C" & i).Value
        polizaLF = Workbooks(libroFinal).Sheets(1).Range("H" & i).Value
        remisionPF = Workbooks(libroFinal).Sheets(1).Range("I" & i).Value
        placaLF = Workbooks(libroFinal).Sheets(1).Range("L" & i).Value
        identificaLF = Workbooks(libroFinal).Sheets(1).Range("M" & i).Value
        
        Workbooks(libroFinal).Sheets(1).Range("A" & i).Value = recaudoLF & polizaLF & remisionPF & placaLF & identificaLF
    Next i
    
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=rutaMacro & libroInicial
    Application.DisplayAlerts = True
    
    Workbooks(libroInicial).Sheets(1).Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Workbooks(libroInicial).Sheets(1).Columns("A:A").Select
    Workbooks(libroInicial).Sheets(1).Range("A3").Activate
    Selection.Insert Shift:=xlToRight
    
    ultimaFilaDocFinal = Workbooks(libroInicial).Sheets(1).Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 2 To ultimaFilaDocFinal
        recaudoLF = Workbooks(libroInicial).Sheets(1).Range("C" & i).Value
        polizaLF = Workbooks(libroInicial).Sheets(1).Range("H" & i).Value
        remisionPF = Workbooks(libroInicial).Sheets(1).Range("I" & i).Value
        placaLF = Workbooks(libroInicial).Sheets(1).Range("L" & i).Value
        identificaLF = Workbooks(libroInicial).Sheets(1).Range("M" & i).Value
        
        Workbooks(libroInicial).Sheets(1).Range("A" & i).Value = recaudoLF & polizaLF & remisionPF & placaLF & identificaLF
    Next i
    
    
    For j = 2 To ultimaFilaDocFinal
        valorLibriInicial = Workbooks(libroInicial).Sheets(1).Range("A" & i).Value
        valorLibroFinal = Workbooks(libroFinal).Sheets(1).Range("A" & i).Value
        
        If valorLibriInicial = valorLibroFinal Then
            validationLibroInicial = Workbooks(libroInicial).Sheets(1).Range("AP" & j).Value
            rgbLibroInicial = Workbooks(libroInicial).Sheets(1).Range("AQ" & j).Value
            numberDocLibroInicial = Workbooks(libroInicial).Sheets(1).Range("AR" & j).Value
            dateLibroInicial = Workbooks(libroInicial).Sheets(1).Range("AS" & j).Value
            
            Workbooks(libroFinal).Sheets(1).Range("AP" & j).Value = validationLibroInicial
            Workbooks(libroFinal).Sheets(1).Range("AQ" & j).Value = rgbLibroInicial
            Workbooks(libroFinal).Sheets(1).Range("AR" & j).Value = numberDocLibroInicial
            Workbooks(libroFinal).Sheets(1).Range("AS" & j).Value = dateLibroInicial
        End If
    Next j
        

End Sub
