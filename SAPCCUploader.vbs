Sub CLear_Range()
'
' CLear_Range Macro
' Author: Wayan Wijesinghe

'
    Range("A5:Q5000").Clear
    Range("A5").Activate
    
End Sub




Sub Rounder()
    '
    ' CLear_Range Macro
    ' Author: Wayan Wijesinghe
    ' v.0
    ' Release data : 2020.07.15

    '
    Application.ScreenUpdating = False
    
    Range(Range("B5"), Range("B5000")).Clear
    Range(Range("D5"), Range("D5000")).Clear
    Range(Range("A5"), Range("Q5000")).ClearFormats
    
    For Each cell In Range("E5:Q5000")
        If IsEmpty(cell) = False Then
            If IsNumeric(cell) = True Then
                cell.Value = WorksheetFunction.Round(cell.Value, 2)
                cell.NumberFormat = "0.00"
            Else
                MsgBox ("Cell " & cell.Address & " contains non-numeric entries. Clean the data and try again")
                Application.ScreenUpdating = True
                Range("A5").Activate
                Exit Sub
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Range("A5").Activate
    
End Sub

Sub Tool_Saveastxtfile()
    ' Author: Wayan Wijesinghe
    'Save the file as a text file in the current folder. If the source file is not saved, saves a txt file in current users desktop


    ' Copy activesheet to the new workbook
    Dim pth As String, wbname As String

    pth = ActiveWorkbook.Path
    
    If pth = "" Then
        pth = Environ("USERPROFILE") & "\Desktop"
    End If
    
    wbname = InputBox("Please input a name for the textfile")
    If wbname = vbNullString Then
        MsgBox ("Either you didnt enter a name for the textfile or you cancelled. Existing...")
        Exit Sub
    End If
    
    
    ActiveSheet.Copy
    MsgBox "A Tab Delimited Text File will be saved in" & pth
    
    'Save new workbook as MyWb.xls(x) into the folder where ThisWorkbook is stored
    ActiveWorkbook.SaveAs pth & "\" & wbname, FileFormat:=xlText, CreateBackup:=False
    
    ' Close the saved copy
    ActiveWorkbook.Close False
    
End Sub
