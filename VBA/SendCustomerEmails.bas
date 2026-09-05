Attribute VB_Name = "Module3"
Option Explicit

Sub PreviewAndSendCustomerEmails()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    
    Dim lastRow As Long
    Dim i As Long
    Dim firstValidRow As Long
    
    Dim customerName As String
    Dim emailID As String
    Dim amountPending As Variant
    Dim attachmentPath As String
    Dim emailBody As String
    
    Dim answer As VbMsgBoxResult
    Dim sentCount As Long
    Dim skippedCount As Long
    
    '====================================================
    ' Excel worksheet
    '====================================================
    
    Set ws = ActiveSheet
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No customer data found.", _
               vbExclamation, "No Data"
        Exit Sub
    End If
    
    
    '====================================================
    ' CONNECT TO OUTLOOK
    '====================================================
    
    On Error Resume Next
    
    'Try to connect to already-open Outlook
    Set OutlookApp = GetObject(, "Outlook.Application")
    
    On Error GoTo 0
    
    
    'If Outlook is not open, start it
    If OutlookApp Is Nothing Then
        
        On Error Resume Next
        
        Set OutlookApp = CreateObject("Outlook.Application")
        
        On Error GoTo 0
        
    End If
    
    
    '====================================================
    ' Check whether Outlook was successfully opened
    '====================================================
    
    If OutlookApp Is Nothing Then
        
        MsgBox _
            "Microsoft Outlook could not be started or automated." & _
            vbCrLf & vbCrLf & _
            "Please check the following:" & vbCrLf & _
            "1. Make sure Microsoft Outlook is installed." & vbCrLf & _
            "2. Open Outlook manually before running the macro." & vbCrLf & _
            "3. Make sure you are using Classic Outlook, not New Outlook." & _
            vbCrLf & vbCrLf & _
            "If you are using New Outlook, VBA COM automation is not supported.", _
            vbCritical, _
            "Outlook Connection Error"
        
        Exit Sub
        
    End If
    
    
    '====================================================
    ' Find first valid customer
    '====================================================
    
    firstValidRow = 0
    
    For i = 2 To lastRow
        
        emailID = Trim(CStr(ws.Cells(i, "C").Value))
        
        If emailID <> "" And InStr(1, emailID, "@") > 0 Then
            firstValidRow = i
            Exit For
        End If
        
    Next i
    
    
    If firstValidRow = 0 Then
        
        MsgBox "No valid email addresses were found in Column C.", _
               vbExclamation, "No Email Addresses"
        
        Exit Sub
        
    End If
    
    
    '====================================================
    ' Get first customer's information
    '====================================================
    
    customerName = Trim(CStr(ws.Cells(firstValidRow, "B").Value))
    emailID = Trim(CStr(ws.Cells(firstValidRow, "C").Value))
    amountPending = ws.Cells(firstValidRow, "D").Value
    attachmentPath = Trim(CStr(ws.Cells(firstValidRow, "E").Value))
    
    
    '====================================================
    ' Create first email for PREVIEW
    '====================================================
    
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    emailBody = _
        "Hello " & customerName & "," & vbCrLf & vbCrLf & _
        "Your Amount is Pending Rs." & _
        Format(amountPending, "#,##0.00") & vbCrLf & vbCrLf & _
        "Please check the Attachment File for more details" & _
        vbCrLf & vbCrLf & _
        "Regards"
    
    
    With OutlookMail
        
        .To = emailID
        
        .Subject = "Amount Pending - " & customerName
        
        .Body = emailBody
        
        'Attach PDF
        If attachmentPath <> "" Then
            
            If Dir(attachmentPath) <> "" Then
                .Attachments.Add attachmentPath
            Else
                MsgBox _
                    "Attachment file was not found:" & _
                    vbCrLf & vbCrLf & attachmentPath, _
                    vbExclamation, _
                    "Attachment Not Found"
            End If
            
        End If
        
        'Show preview
        .Display
        
    End With
    
    
    '====================================================
    ' Ask for confirmation
    '====================================================
    
    answer = MsgBox( _
        "The first email has been opened in Outlook." & _
        vbCrLf & vbCrLf & _
        "Please check:" & vbCrLf & _
        "• Customer Name" & vbCrLf & _
        "• Email Address" & vbCrLf & _
        "• Amount" & vbCrLf & _
        "• Attachment" & vbCrLf & vbCrLf & _
        "Do you want to SEND ALL emails?", _
        vbYesNo + vbQuestion, _
        "Confirm Bulk Email")
    
    
    '====================================================
    ' CANCEL
    '====================================================
    
    If answer = vbNo Then
        
        On Error Resume Next
        OutlookMail.Close 0
        On Error GoTo 0
        
        Set OutlookMail = Nothing
        Set OutlookApp = Nothing
        
        MsgBox "Email sending cancelled. No emails were sent.", _
               vbInformation, "Cancelled"
        
        Exit Sub
        
    End If
    
    
    '====================================================
    ' SEND ALL EMAILS
    '====================================================
    
    On Error Resume Next
    OutlookMail.Close 0
    Set OutlookMail = Nothing
    On Error GoTo 0
    
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Sending emails..."
    
    
    For i = 2 To lastRow
        
        customerName = Trim(CStr(ws.Cells(i, "B").Value))
        emailID = Trim(CStr(ws.Cells(i, "C").Value))
        amountPending = ws.Cells(i, "D").Value
        attachmentPath = Trim(CStr(ws.Cells(i, "E").Value))
        
        If emailID <> "" And InStr(1, emailID, "@") > 0 Then
            
            On Error GoTo EmailError
            
            Set OutlookMail = OutlookApp.CreateItem(0)
            
            emailBody = _
                "Hello " & customerName & "," & vbCrLf & vbCrLf & _
                "Your Amount is Pending Rs." & _
                Format(amountPending, "#,##0.00") & vbCrLf & vbCrLf & _
                "Please check the Attachment File for more details" & _
                vbCrLf & vbCrLf & _
                "Regards"
            
            
            With OutlookMail
                
                .To = emailID
                
                .Subject = "Amount Pending - " & customerName
                
                .Body = emailBody
                
                'Attach PDF
                If attachmentPath <> "" Then
                    
                    If Dir(attachmentPath) <> "" Then
                        .Attachments.Add attachmentPath
                    End If
                    
                End If
                
                'Send email
                .Send
                
            End With
            
            sentCount = sentCount + 1
            
            Set OutlookMail = Nothing
            
            GoTo NextCustomer
            
            
EmailError:
            
            skippedCount = skippedCount + 1
            
            If Not OutlookMail Is Nothing Then
                On Error Resume Next
                OutlookMail.Close 0
                Set OutlookMail = Nothing
                On Error GoTo 0
            End If
            
            Err.Clear
            
NextCustomer:
            
            On Error GoTo 0
            
        Else
            
            skippedCount = skippedCount + 1
            
        End If
        
        Application.StatusBar = _
            "Sending email " & (i - 1) & _
            " of " & (lastRow - 1)
        
    Next i
    
    
    '====================================================
    ' Finish
    '====================================================
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    
    MsgBox _
        "Email process completed!" & vbCrLf & vbCrLf & _
        "Emails sent: " & sentCount & vbCrLf & _
        "Rows skipped: " & skippedCount, _
        vbInformation, _
        "Completed"

End Sub

