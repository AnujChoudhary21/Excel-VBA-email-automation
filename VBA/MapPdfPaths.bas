Attribute VB_Name = "Module2"
Option Explicit

Sub FindEmployeePDFs()

    Dim ws As Worksheet
    Dim folderPath As String
    Dim lastRow As Long
    Dim i As Long
    
    Dim employeeID As String
    Dim pdfPath As String
    
    Dim foundCount As Long
    Dim notFoundCount As Long
    
    Set ws = ActiveSheet
    
    '==================================================
    ' Select PDF folder
    '==================================================
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        
        .Title = "Select Folder Containing Employee PDFs"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then
            MsgBox "No folder selected.", vbInformation
            Exit Sub
        End If
        
        folderPath = .SelectedItems(1)
        
    End With
    
    
    '==================================================
    ' Find last row in Column A
    '==================================================
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No Employee IDs found in Column A.", _
               vbExclamation
        Exit Sub
    End If
    
    
    '==================================================
    ' Add header to Column E
    '==================================================
    
    ws.Range("E1").Value = "PDF File Path"
    ws.Range("E1").Font.Bold = True
    
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Searching for employee PDF files..."
    
    
    '==================================================
    ' Search PDF for each Employee ID
    '==================================================
    
    For i = 2 To lastRow
        
        employeeID = Trim(CStr(ws.Cells(i, "A").Value))
        
        If employeeID <> "" Then
            
            ' Expected filename:
            ' EmployeeID.pdf
            
            pdfPath = folderPath & "\" & employeeID & ".pdf"
            
            If Dir(pdfPath) <> "" Then
                
                '--------------------------------------
                ' PDF found
                '--------------------------------------
                
                ws.Cells(i, "E").Value = pdfPath
                
                ' Make the path clickable
                ws.Hyperlinks.Add _
                    Anchor:=ws.Cells(i, "E"), _
                    Address:=pdfPath, _
                    TextToDisplay:=pdfPath
                
                foundCount = foundCount + 1
                
            Else
                
                '--------------------------------------
                ' PDF not found
                '--------------------------------------
                
                ws.Cells(i, "E").Value = "PDF Not Found"
                
                notFoundCount = notFoundCount + 1
                
            End If
            
        Else
            
            ws.Cells(i, "E").Value = "No Employee ID"
            
        End If
        
        Application.StatusBar = _
            "Checking row " & i & " of " & lastRow
        
    Next i
    
    
    '==================================================
    ' Format Column E
    '==================================================
    
    ws.Columns("E").ColumnWidth = 60
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    
    '==================================================
    ' Completion message
    '==================================================
    
    MsgBox "PDF search completed!" & vbCrLf & vbCrLf & _
           "PDFs found: " & foundCount & vbCrLf & _
           "PDFs not found: " & notFoundCount, _
           vbInformation, "PDF Search Complete"

End Sub

