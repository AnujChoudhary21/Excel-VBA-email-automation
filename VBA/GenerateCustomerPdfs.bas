Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateCustomerPDFs()

    Dim wdApp As Object
    Dim wdDoc As Object
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim templatePath As String
    Dim destinationFolder As String
    
    Dim employeeID As String
    Dim customerName As String
    Dim amountPending As String
    
    Dim outputPDF As String
    Dim uniquePDF As String
    
    Dim generatedCount As Long
    Dim skippedCount As Long
    
    ' Word constants for late binding
    Const wdFindContinue As Long = 1
    Const wdReplaceAll As Long = 2
    Const wdExportFormatPDF As Long = 17
    
    Set ws = ActiveSheet
    
    '==================================================
    ' STEP 1: Select Word template
    '==================================================
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        .Title = "Select Word Letter Template"
        .AllowMultiSelect = False
        
        .Filters.Clear
        .Filters.Add "Microsoft Word Documents", "*.docx;*.doc"
        
        If .Show <> -1 Then
            MsgBox "No Word template selected.", vbInformation
            Exit Sub
        End If
        
        templatePath = .SelectedItems(1)
        
    End With
    
    
    '==================================================
    ' STEP 2: Select destination folder
    '==================================================
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        
        .Title = "Select Destination Folder for PDFs"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then
            MsgBox "No destination folder selected.", vbInformation
            Exit Sub
        End If
        
        destinationFolder = .SelectedItems(1)
        
    End With
    
    
    '==================================================
    ' Find last row of Excel data
    '==================================================
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No customer data found in the worksheet.", _
               vbExclamation
        Exit Sub
    End If
    
    
    '==================================================
    ' Start Microsoft Word
    '==================================================
    
    On Error Resume Next
    
    Set wdApp = GetObject(, "Word.Application")
    
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "Microsoft Word could not be started.", _
               vbCritical
        Exit Sub
    End If
    
    wdApp.Visible = False
    
    
    '==================================================
    ' Process each Excel row
    '==================================================
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Generating customer PDFs..."
    
    For i = 2 To lastRow
        
        '----------------------------------------------
        ' Read Excel values
        '----------------------------------------------
        
        employeeID = Trim(CStr(ws.Cells(i, "A").Value))
        customerName = CStr(ws.Cells(i, "B").Value)
        amountPending = CStr(ws.Cells(i, "D").Value)
        
        
        '----------------------------------------------
        ' Skip rows without Employee ID
        '----------------------------------------------
        
        If employeeID <> "" Then
            
            On Error GoTo RowError
            
            '------------------------------------------
            ' Open a fresh copy of the Word template
            '------------------------------------------
            
            Set wdDoc = wdApp.Documents.Open( _
                            fileName:=templatePath, _
                            ReadOnly:=False, _
                            AddToRecentFiles:=False)
            
            
            '------------------------------------------
            ' Replace Customer Name
            '------------------------------------------
            
            ReplaceWordText wdDoc, _
                            "<<Customer Name>>", _
                            customerName
            
            
            '------------------------------------------
            ' Replace Amount Pending
            '------------------------------------------
            
            ReplaceWordText wdDoc, _
                            "<<Amount Pending>>", _
                            amountPending
            
            
            '------------------------------------------
            ' Create PDF filename
            '------------------------------------------
            
            outputPDF = destinationFolder & "\" & _
                        CleanFileName(employeeID) & ".pdf"
            
            ' Prevent overwriting an existing PDF
            uniquePDF = GetUniquePDFName(outputPDF)
            
            
            '------------------------------------------
            ' Export Word document as PDF
            '------------------------------------------
            
            wdDoc.ExportAsFixedFormat _
                    OutputFileName:=uniquePDF, _
                    ExportFormat:=wdExportFormatPDF
            
            
            '------------------------------------------
            ' Close Word document without saving
            '------------------------------------------
            
            wdDoc.Close SaveChanges:=False
            Set wdDoc = Nothing
            
            generatedCount = generatedCount + 1
            
        Else
            
            skippedCount = skippedCount + 1
            
        End If
        
        Application.StatusBar = _
            "Processing row " & i & " of " & lastRow
        
        GoTo ContinueLoop
        
        
RowError:
        
        skippedCount = skippedCount + 1
        
        If Not wdDoc Is Nothing Then
            On Error Resume Next
            wdDoc.Close SaveChanges:=False
            Set wdDoc = Nothing
            On Error GoTo 0
        End If
        
        Err.Clear
        
ContinueLoop:
        
        On Error GoTo 0
        
    Next i
    
    
    '==================================================
    ' Close Word
    '==================================================
    
    On Error Resume Next
    
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    On Error GoTo 0
    
    
    '==================================================
    ' Completion message
    '==================================================
    
    MsgBox "PDF generation completed!" & vbCrLf & vbCrLf & _
           "PDFs generated: " & generatedCount & vbCrLf & _
           "Rows skipped: " & skippedCount & vbCrLf & vbCrLf & _
           "Saved to:" & vbCrLf & destinationFolder, _
           vbInformation, "PDF Generation Complete"

    Exit Sub

End Sub


'=========================================================
' Replace text anywhere in the Word document
'=========================================================

Sub ReplaceWordText(ByVal wdDoc As Object, _
                    ByVal searchText As String, _
                    ByVal replacementText As String)

    Dim storyRange As Object
    Dim currentRange As Object
    
    For Each storyRange In wdDoc.StoryRanges
        
        Set currentRange = storyRange
        
        Do
            
            With currentRange.Find
                
                .ClearFormatting
                .Replacement.ClearFormatting
                
                .Text = searchText
                .Replacement.Text = replacementText
                
                .Forward = True
                .Wrap = 1
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                
                .Execute Replace:=2
                
            End With
            
            On Error Resume Next
            Set currentRange = currentRange.NextStoryRange
            On Error GoTo 0
            
        Loop Until currentRange Is Nothing
        
    Next storyRange

End Sub


'=========================================================
' Remove characters that are invalid in Windows filenames
'=========================================================

Function CleanFileName(ByVal fileName As String) As String

    Dim invalidCharacters As Variant
    Dim character As Variant
    
    invalidCharacters = Array( _
        "\", "/", ":", "*", "?", """", "<", ">", "|" _
    )
    
    For Each character In invalidCharacters
        fileName = Replace(fileName, character, "_")
    Next character
    
    CleanFileName = Trim(fileName)

End Function


'=========================================================
' Create unique PDF filename
'=========================================================

Function GetUniquePDFName(ByVal originalPath As String) As String

    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim counter As Long
    Dim newPath As String
    
    ' If PDF does not already exist
    If Dir(originalPath) = "" Then
        GetUniquePDFName = originalPath
        Exit Function
    End If
    
    ' Get folder path
    folderPath = Left(originalPath, _
                      InStrRev(originalPath, "\"))
    
    ' Get filename without .pdf
    fileName = Mid(originalPath, _
                   InStrRev(originalPath, "\") + 1)
    
    baseName = Left(fileName, _
                    Len(fileName) - 4)
    
    counter = 1
    
    Do
        
        newPath = folderPath & _
                  baseName & " (" & counter & ").pdf"
        
        counter = counter + 1
        
    Loop While Dir(newPath) <> ""
    
    GetUniquePDFName = newPath

End Function

