Attribute VB_Name = "Módulo1"
Public planilhamacrocheckbookingerror As String
Public selectedservice As String
Public myPath As String
Public myFile As String

Sub CHECK_BOOKING_ERROR()
Attribute CHECK_BOOKING_ERROR.VB_ProcData.VB_Invoke_Func = "X\n14"
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb As Workbook
Dim myExtension As String
Dim FldrPicker As FileDialog

planilhamacrocheckbookingerror = ActiveWorkbook.Name

' limpa filtro de dia para poder inserir novos dados na database
    If ActiveSheet.AutoFilterMode Then
    'FILTRO ON
    Rows("1:1").Select
    Selection.AutoFilter
    End If

' DESCONGELANDO PAINEIS
    Range("D2").Select
    ActiveWindow.FreezePanes = False

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(fileName:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'Change First Worksheet's Background Fill Blue
      If Range("A2") <> "" Then
      Call FUN_CHECK_BOOKING_ERROR
      End If
    
    'Not Save and Close Workbook
      wb.Close SaveChanges:=False
Application.DisplayAlerts = True

    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

' inserindo filtro HOJE na data base para análise
    If ActiveSheet.AutoFilterMode Then
    'FILTRO ON
    Rows("1:1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=1, Criteria1:= _
        xlFilterToday, Operator:=xlFilterDynamic
    Else
    'FILTRO OFF
    Rows("1:1").Select
    Selection.AutoFilter Field:=1, Criteria1:= _
        xlFilterToday, Operator:=xlFilterDynamic
    End If
    
' COR FUNDO BRANCO
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
' CONGELANDO PAINEIS
    Range("D2").Select
    ActiveWindow.FreezePanes = True


'Message Box when tasks are completed
ActiveWorkbook.Save
MsgBox "Finalized!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Function FUN_CHECK_BOOKING_ERROR()

'
' LIMPANDO COLUNAS INUTEIS + LINHAS SEM BOOKING + ALTURA DAS LINHAS
'

'
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AA:AC").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AG:BQ").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:C").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("Z:Z").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Range("E35").Select
    Columns("AC:AC").Select
    Selection.Cut
    Columns("AE:AE").Select
    Selection.Insert Shift:=xlToRight
    
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Cells.Select
    Selection.RowHeight = 12.75


'
' INSERIR COLUNA SERVICO E DATA
'

'
      
'SERVICO
    
    
    Application.ScreenUpdating = True
    Windows(myFile).Activate
    SERVICEINPUT.Show
    Application.ScreenUpdating = False
    
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "SERVICE"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = selectedservice
    Range("A2").Select
    Selection.Copy
    Range("B1").Select
    Selection.End(xlDown).Select
    actualrow = Selection.Row()
    Range("A2:A" & actualrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    selectedservice = ""
    
'DATA
    
    dateofcreation = Date
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "DATE OF CREATION / AMEND"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = dateofcreation
    Range("A2").Select
    Selection.Copy
    Range("B1").Select
    Selection.End(xlDown).Select
    actualrow = Selection.Row()
    Range("A2:A" & actualrow).Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        
'
' CHECK CARRIER
'
'
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Check Carrier "
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]=""AM"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""AR""),AND(RC[-1]=""AM"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""UY""),AND(RC[-1]=""IA"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],5)=""USPEF""),AND(RC[-1]=""UC"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],5)=""USHOU""),AND(RC[-1]=""UC"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""CO"")," & _
        "AND(RC[-1]=""SA"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""IT""),AND(RC[-1]=""SX"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""ZA""),AND(RC[-1]=""AD"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""CL""),AND(RC[-1]=""AM"",RC[-10]=""ALCT"",RC[-6]=""S""),AND(RC[-1]=""IA"",RC[-10]=""ABUS""),AND(RC[-1]=""UC"",RC[-10]=""UCLA""),AND(RC[-1]=""AM"",RC[-10]=""SAEC"",RC[-6]=" & _
        """S"",LEFT(RC[-2],2)=""UY""),AND(RC[-1]=""AM"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""PY""),AND(RC[-1]=""AM"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""AR""),AND(RC[-1]=""IA"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""US""),AND(RC[-1]=""IA"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""CO""),AND(RC[-1]=""SA"",RC[-10]=""SAEC"",RC[-6]=""N""),AND(RC[-1]=""MM" & _
        """,RC[-10]=""ABAC""),AND(RC[-1]=""EP"",RC[-10]=""ABAC""),AND(RC[-1]=""AD"",RC[-10]=""ABAC""),AND(RC[-1]=""SX"",RC[-10]=""SAAF""),AND(RC[-1]=""KG"",RC[-10]=""ASIA"")),""OK"",""VERIFICAR"")"
    
    Range("L2").Select
    Selection.Copy
    Range("A1").Select
    Selection.End(xlDown).Select
    actualrow = Selection.Row()
    Range("L2:L" & actualrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Calculate
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False


'
' CHECK TS OR FEEDER VESSEL / PORT
'

'
    Columns("W:W").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Check TS Vessel"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(LEFT(RC[-14],2)=""BR"", OR(RC[-10]=""N"",RC[-9]="""", RC[-8]="""", RC[-7]="""", RC[-6]="""", RC[-5]="""", RC[-4]="""")),""VERIFICAR"",(IF(AND(OR(RC[-3]<>"""",RC[-2]<>"""",RC[-1]<>""""),OR(RC[-3]="""",RC[-2]="""",RC[-1]="""")),""VERIFICAR"",""OK"")))"
    Range("W2").Select
    Selection.Copy
    Range("A1").Select
    Selection.End(xlDown).Select
    actualrow = Selection.Row()
    Range("W2:W" & actualrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("W2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Calculate
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    
  
'
' CHECK SHIPPER + AGREEMENT PARTY
'

'
    Columns("AE:AE").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "Check Shipper / Ag.Party"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-7]="""",RC[-6]="""",RC[-3]="""",RC[-2]=""""),""VERIFICAR"",""OK"")"
    Range("AE2").Select
    Selection.Copy
    Range("A1").Select
    Selection.End(xlDown).Select
    actualrow = Selection.Row()
    Range("AE2:AE" & actualrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("AE2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Calculate
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False

  
  
'
' CHECK AGREEMENT NUMBER
'

'
    Columns("AH:AH").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Check Agree.No. + Product ID"
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-2]=""XXXX1234567 RE"",RC[-2]=""XXXX1234567 PE"",RC[-2]=""XXXX1234567 HA"",RC[-2]=""XXXX1234567 PU""),(IF(RC[-1]=99999,""OK"",""VERIFICAR"")),(IF(OR(RC[-2]="""", LEN(RC[-2])<>11),""VERIFICAR"",(IF(OR(RIGHT(LEFT(RC[-2],4),1)=""C"",RIGHT(LEFT(RC[-2],4),1)=""Q"",),""OK"",""VERIFICAR"")))))"
    Range("AH2").Select
    Selection.Copy
    Range("A1").Select
    Selection.End(xlDown).Select
    actualrow = Selection.Row()
    Range("AH2:AH" & actualrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Range("AH2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Calculate
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
        



'
' VERFICANDO SE HÁ ALGUM ERRO
'

'
    Range("AM1").Select
    ActiveCell.FormulaR1C1 = "Check Error"
    Range("AM2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-5]<>""OK"",RC[-8]<>""OK"",RC[-16]<>""OK"",RC[-27]<>""OK""),""VERIFICAR"",""OK"")"
    Range("AM2").Select
    Selection.Copy
    Range("A1").Select
    Selection.End(xlDown).Select
    actualrow = Selection.Row()
    Range("AM2:AM" & actualrow).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    Range("AM2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Calculate
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False





' EXCLUINDO LINHAS QUE NAO POSSUEM BOOKING ERROR
    If ActiveSheet.AutoFilterMode Then
    'FILTRO ON
    Rows("1:1").Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=39, Criteria1:="OK"
    Else
    'FILTRO OFF
    Rows("1:1").Select
    Selection.AutoFilter Field:=39, Criteria1:="OK"
    End If
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    Rows("1:1").Select
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete


'
' COLOCANDO EM DATABASE
'

'
    If Range("A2") <> "" Then

Application.DisplayAlerts = False
        If Range("A3") <> "" Then
        Range("A2:AR2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Else
        Range("A2:AR2").Select
        Selection.Copy
        End If
        
    Windows(planilhamacrocheckbookingerror).Activate
    Sheets("DATA_BASE").Activate
       
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    
    End If
    
        
End Function

'Sub Macro2()
''
'' LIMPANDO COLUNAS INUTEIS + LINHAS SEM BOOKING + ALTURA DAS LINHAS
''
'
''
'    Columns("A:A").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("AA:AC").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("AG:BQ").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("A:C").Select
'    Selection.Cut
'    Columns("E:E").Select
'    Selection.Insert Shift:=xlToRight
'    Columns("Z:Z").Select
'    Selection.Cut
'    Columns("I:I").Select
'    Selection.Insert Shift:=xlToRight
'    Range("E35").Select
'    Columns("AC:AC").Select
'    Selection.Cut
'    Columns("AE:AE").Select
'    Selection.Insert Shift:=xlToRight
'
'    Columns("A:A").Select
'    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Selection.EntireRow.Delete
'    Cells.Select
'    Selection.RowHeight = 12.75
'
'
''
'' INSERIR COLUNA SERVICO E DATA
''
'
''
'
''SERVICO
'
'
''''''    Application.ScreenUpdating = True
''''''    Windows(myFile).Activate
''''''    SERVICEINPUT.Show
''''''    Application.ScreenUpdating = False
'
'
'    selectedservice = "ASIA"
'
'    Columns("A:A").Select
'    Application.CutCopyMode = False
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("A1").Select
'    ActiveCell.FormulaR1C1 = "SERVICE"
'    Range("A2").Select
'    ActiveCell.FormulaR1C1 = selectedservice
'    Range("A2").Select
'    Selection.Copy
'    Range("B1").Select
'    Selection.End(xlDown).Select
'    actualrow = Selection.Row()
'    Range("A2:A" & actualrow).Select
'    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'
'    Range("A2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'    selectedservice = ""
'
''DATA
'
'    dateofcreation = Date
'    Columns("A:A").Select
'    Application.CutCopyMode = False
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("A1").Select
'    ActiveCell.FormulaR1C1 = "DATE OF CREATION / AMEND"
'    Range("A2").Select
'    ActiveCell.FormulaR1C1 = dateofcreation
'    Range("A2").Select
'    Selection.Copy
'    Range("B1").Select
'    Selection.End(xlDown).Select
'    actualrow = Selection.Row()
'    Range("A2:A" & actualrow).Select
'    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
'        xlNone, SkipBlanks:=False, Transpose:=False
'
''
'' CHECK CARRIER
''
''
'    Columns("L:L").Select
'    Application.CutCopyMode = False
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("L1").Select
'    ActiveCell.FormulaR1C1 = "Check Carrier "
'    Range("L2").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IF(OR(AND(RC[-1]=""AM"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""AR""),AND(RC[-1]=""AM"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""UY""),AND(RC[-1]=""IA"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],5)=""USPEF""),AND(RC[-1]=""UC"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],5)=""USHOU""),AND(RC[-1]=""UC"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""CO"")," & _
'        "AND(RC[-1]=""SA"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""IT""),AND(RC[-1]=""SX"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""ZA""),AND(RC[-1]=""AD"",RC[-10]=""ALCT"",RC[-6]=""N"",LEFT(RC[-2],2)=""CL""),AND(RC[-1]=""AM"",RC[-10]=""ALCT"",RC[-6]=""S""),AND(RC[-1]=""IA"",RC[-10]=""ABUS""),AND(RC[-1]=""UC"",RC[-10]=""UCLA""),AND(RC[-1]=""AM"",RC[-10]=""SAEC"",RC[-6]=" & _
'        """S"",LEFT(RC[-2],2)=""UY""),AND(RC[-1]=""AM"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""PY""),AND(RC[-1]=""AM"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""AR""),AND(RC[-1]=""IA"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""US""),AND(RC[-1]=""IA"",RC[-10]=""SAEC"",RC[-6]=""S"",LEFT(RC[-2],2)=""CO""),AND(RC[-1]=""SA"",RC[-10]=""SAEC"",RC[-6]=""N""),AND(RC[-1]=""MM" & _
'        """,RC[-10]=""ABAC""),AND(RC[-1]=""EP"",RC[-10]=""ABAC""),AND(RC[-1]=""AD"",RC[-10]=""ABAC""),AND(RC[-1]=""SX"",RC[-10]=""SAAF""),AND(RC[-1]=""KG"",RC[-10]=""ASIA"")),""OK"",""VERIFICAR"")"
'
'    Range("L2").Select
'    Selection.Copy
'    Range("A1").Select
'    Selection.End(xlDown).Select
'    actualrow = Selection.Row()
'    Range("L2:L" & actualrow).Select
'    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'
'    Range("L2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Calculate
''    Application.CutCopyMode = False
''    Selection.Copy
''    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
''        :=False, Transpose:=False
'
'
''
'' CHECK TS OR FEEDER VESSEL / PORT
''
'
''
'    Columns("W:W").Select
'    Application.CutCopyMode = False
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("W1").Select
'    ActiveCell.FormulaR1C1 = "Check TS Vessel"
'    Range("W2").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IF(AND(LEFT(RC[-14],2)=""BR"", OR(RC[-10]=""N"",RC[-9]="""", RC[-8]="""", RC[-7]="""", RC[-6]="""", RC[-5]="""", RC[-4]="""")),""VERIFICAR"",(IF(AND(OR(RC[-3]<>"""",RC[-2]<>"""",RC[-1]<>""""),OR(RC[-3]="""",RC[-2]="""",RC[-1]="""")),""VERIFICAR"",""OK"")))"
'    Range("W2").Select
'    Selection.Copy
'    Range("A1").Select
'    Selection.End(xlDown).Select
'    actualrow = Selection.Row()
'    Range("W2:W" & actualrow).Select
'    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'
'    Range("W2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Calculate
''    Application.CutCopyMode = False
''    Selection.Copy
''    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
''        :=False, Transpose:=False
'
'
''
'' CHECK SHIPPER + AGREEMENT PARTY
''
'
''
'    Columns("AE:AE").Select
'    Application.CutCopyMode = False
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("AE1").Select
'    ActiveCell.FormulaR1C1 = "Check Shipper / Ag.Party"
'    Range("AE2").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IF(OR(RC[-7]="""",RC[-6]="""",RC[-3]="""",RC[-2]=""""),""VERIFICAR"",""OK"")"
'    Range("AE2").Select
'    Selection.Copy
'    Range("A1").Select
'    Selection.End(xlDown).Select
'    actualrow = Selection.Row()
'    Range("AE2:AE" & actualrow).Select
'    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'
'    Range("AE2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Calculate
''    Application.CutCopyMode = False
''    Selection.Copy
''    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
''        :=False, Transpose:=False
'
'
'
''
'' CHECK AGREEMENT NUMBER
''
'
''
'    Columns("AH:AH").Select
'    Application.CutCopyMode = False
'    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("AH1").Select
'    ActiveCell.FormulaR1C1 = "Check Agree.No. + Product ID"
'    Range("AH2").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IF(OR(RC[-2]=""XXXX1234567 RE"",RC[-2]=""XXXX1234567 PE"",RC[-2]=""XXXX1234567 HA"",RC[-2]=""XXXX1234567 PU""),(IF(RC[-1]=99999,""OK"",""VERIFICAR"")),(IF(OR(RC[-2]="""", LEN(RC[-2])<>11),""VERIFICAR"",(IF(OR(RIGHT(LEFT(RC[-2],4),1)=""C"",RIGHT(LEFT(RC[-2],4),1)=""Q"",),""OK"",""VERIFICAR"")))))"
'    Range("AH2").Select
'    Selection.Copy
'    Range("A1").Select
'    Selection.End(xlDown).Select
'    actualrow = Selection.Row()
'    Range("AH2:AH" & actualrow).Select
'    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'
'    Range("AH2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Calculate
''    Application.CutCopyMode = False
''    Selection.Copy
''    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
''        :=False, Transpose:=False
'
'
'
'
''
'' VERFICANDO SE HÁ ALGUM ERRO
''
'
''
'    Range("AM1").Select
'    ActiveCell.FormulaR1C1 = "Check Error"
'    Range("AM2").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IF(OR(RC[-5]<>""OK"",RC[-8]<>""OK"",RC[-16]<>""OK"",RC[-27]<>""OK""),""VERIFICAR"",""OK"")"
'    Range("AM2").Select
'    Selection.Copy
'    Range("A1").Select
'    Selection.End(xlDown).Select
'    actualrow = Selection.Row()
'    Range("AM2:AM" & actualrow).Select
'    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'
'    Range("AM2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Calculate
''    Application.CutCopyMode = False
''    Selection.Copy
''    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
''        :=False, Transpose:=False
'
'
'
'
'
'' EXCLUINDO LINHAS QUE NAO POSSUEM BOOKING ERROR
'    If ActiveSheet.AutoFilterMode Then
'    'FILTRO ON
'    Rows("1:1").Select
'    Selection.AutoFilter
'    Selection.AutoFilter Field:=39, Criteria1:="OK"
'    Else
'    'FILTRO OFF
'    Rows("1:1").Select
'    Selection.AutoFilter Field:=39, Criteria1:="OK"
'    End If
'    Range("A2").Select
'    Range(Selection, Selection.End(xlDown)).ClearContents
'    Rows("1:1").Select
'    Selection.AutoFilter
'    Columns("A:A").Select
'    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Selection.EntireRow.Delete
'
' End Sub

