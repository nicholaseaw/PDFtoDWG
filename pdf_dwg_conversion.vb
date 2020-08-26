Option Explicit

Sub PdfToDwg()
    'Set background plot variable
    Dim backPlot As Integer
    
    'Declare variables
    Dim fileNames() As String, oPath As String, dPath As String, ePath As String, ExcelFilePath As String, i As Integer, filePDF As String, fName As String, dwgName As String, dwg As AcadDocument, pdfName As String
    Dim splitFileName() As String, brutName As String
    Dim DwgPath As String, DwgTemplate As String
    Dim inputDist As Variant, scalefactor As Integer
    
    'Declare distance variables
    Dim returnDist As Double
    Dim basePnt(0 To 2) As Double
    
    basePnt(0) = 0#: basePnt(1) = 0#: basePnt(2) = 0#
    
    'Declare Excel variables
    Dim oApp_Excel As Excel.Application
    Dim oBook As Excel.workbook
    
    'Create Excel
    Set oApp_Excel = CreateObject("EXCEL.APPLICATION")
    Set oBook = oApp_Excel.workbooks.Add
    
    backPlot = ThisDrawing.GetVariable("BACKGROUNDPLOT")
    ThisDrawing.SetVariable "BACKGROUNDPLOT", 0
    ThisDrawing.SetVariable "FILEDIA", 0
    
    'Original path of PDF file location
    oPath = "D:/Working Projects/VBA/AutoCAD/PDF/"
    
    'Destination path of DWG file location
    dPath = "D:/Working Projects/VBA/AutoCAD/DWG/"
    
    'Destination path of Excel file location
    ePath = "D:\Working Projects\VBA\AutoCAD\Excel\"
    
    'Destination path of template dwg
    DwgPath = Preferences.Files.TemplateDwgPath & "\acadiso.dwt"
    
    'Get DWG template
    DwgTemplate = Dir(DwgPath, vbNormal)
    
    'Get all XLXS file from path
    ExcelFilePath = Dir(ePath & "*.xlsx")
    
    'Get all PDF file from original path
    filePDF = Dir(oPath & "*.pdf")
    While filePDF <> ""
        ReDim Preserve fileNames(i)
        fileNames(i) = filePDF
        i = i + 1
        filePDF = Dir
    Wend
    
    'Check if dwg/bak file is in file path
    If Dir(dPath & "*.dwg") <> "" And Dir(dPath & "*.bak") <> "" Then
        Kill dPath & "/*.dwg"
        Kill dPath & "/*.bak"
    Else
    
    End If
    
    'Create sheet for dwg output in Excel
    oBook.Sheets(1).Cells(1, 1).Value = "File Name"
    oBook.Sheets(1).Cells(1, 2).Value = "Date created"
    oBook.Sheets(1).Cells(1, 3).Value = "Dwg Scale"
    
    'Convert PDF to DWG
    For i = 0 To UBound(fileNames)
        'Get individual file names
        fName = fileNames(i)
        splitFileName = Split(fName, ".")
        brutName = splitFileName(0)
        'Create path and name for PDF and DWG
        dwgName = dPath & brutName & ".dwg"
        pdfName = oPath & fName
        'Output files in DWG format to excel
        oBook.Sheets(1).Cells(i + 2, 1).Value = brutName & ".dwg"
        'Get path for DWG template, import PDF and convert to DWG at 1.0 scale
        ThisDrawing.Application.Documents.Add (DwgPath)
        ThisDrawing.SendCommand "(command ""-PDFIMPORT"" ""FILE"" """ & pdfName & """ ""1"" ""0,0"" ""1.0"" ""0.0"")" & vbCr
        
        'Get scale of PDF drawing
        
        
        
    '    MsgBox "Please measure distance between 2 points for scale otherwise input 1 if NTS (following distance will need to be 1)", vbOKOnly, "Scale correction"
    '    returnDist = ThisDrawing.Utility.GetDistance(, "Enter distance: ")
    '    inputDist = InputBox("Please enter dimension from drawing to correct scale:")
    '    Do
    '        If Not IsNumeric(inputDist) Then
    '            MsgBox "Please enter distance in numbers!"
    '            inputDist = InputBox("Please enter dimension from drawing to correct scale:")
    '        ElseIf inputDist < 0 Then
    '            MsgBox "Distance is negative. Please enter a positive value."
    '            inputDist = InputBox("Please enter dimension from drawing to correct scale:")
    '        ElseIf inputDist = "" Then
    '            MsgBox "Distance is blank. Please enter a value."
    '            inputDist = InputBox("Please enter dimension from drawing to correct scale:")
    '        Else
    '            Exit Do
    '        End If
    '    Loop
    '    scalefactor = inputDist / returnDist
    '    ThisDrawing.SendCommand ("SCALE" & vbCr & "ALL" & vbCr & vbCr & "0,0" & vbCr & scalefactor & vbCr)
    '    If scalefactor <> 1 Then
    '        oBook.Sheets(1).Cells(i + 2, 3).Value = scalefactor
    '    ElseIf scalefactor = 1 Then
    '        oBook.Sheets(1).Cells(i + 2, 3).Value = "NTS"
    '    End If
        
        'Save as DWG format
        ThisDrawing.SaveAs (dwgName)
        ThisDrawing.Close
        'oBook.Sheets(1).Cells(i + 2, 2).Value = FileDateTime(dwgName)
    Next i
    
    'Check if Excel file is in file path
    If ExcelFilePath <> "" Then
        Kill ePath & "\*.xlsx"
    End If
    
    oBook.SaveAs (ePath & "FileName.xlsx")
    oBook.Close
    oApp_Excel.Quit
    
    Set oBook = Nothing
End Sub
