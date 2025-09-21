# Automations-Tools-Statements-Automailer
Earlier, settlement reports had to be manually filtered in Excel, separate files were created for each merchant, and then emails were sent one by one. This manual process was not only time-consuming but also prone to human errors.


Sub SplitSaveAndMail_Gmail()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet, wsM As Worksheet
    Dim FSO As Object, FolderPath As String
    Dim uniqueCodes As Object, merchantCodes As Object
    Dim Code As Variant, cell As Range
    Dim LastRow As Long, mLastRow As Long
    
    ' Gmail account details
    Dim GmailID As String, GmailAppPassword As String
    GmailID = "XXXXgmail.com"
    GmailAppPassword = "XXXX XXXX XXXX XXXX"
    
    ' Set worksheets
    Set ws = ThisWorkbook.Sheets("Axis")
    Set wsM = ThisWorkbook.Sheets("Merchant")
    
    ' Folder path for split files
    FolderPath = ThisWorkbook.Path & "\Split Files"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(FolderPath) Then
        FSO.CreateFolder FolderPath
    End If
    
    
    LastRow = ws.Cells(ws.Rows.Count, "Z").End(xlUp).Row
    
    
    mLastRow = wsM.Cells(wsM.Rows.Count, "A").End(xlUp).Row
    
    ' Merchant dictionary
    Set merchantCodes = CreateObject("Scripting.Dictionary")
    For Each cell In wsM.Range("A2:A" & mLastRow)
        If Not merchantCodes.Exists(cell.Value) And cell.Value <> "" Then
            merchantCodes.Add cell.Value, 1
        End If
    Next cell
    
    ' Unique Axis codes
    Set uniqueCodes = CreateObject("Scripting.Dictionary")
    For Each cell In ws.Range("Z2:Z" & LastRow)
        If Not uniqueCodes.Exists(cell.Value) And cell.Value <> "" Then
            uniqueCodes.Add cell.Value, 1
        End If
    Next cell
    
    
    Dim EmailFiles As Object
    Set EmailFiles = CreateObject("Scripting.Dictionary")
    
    ' List of headers
    Dim headers() As Variant
    headers = Array("TERM_ID", "TRAN_DATE", "BATCH_NO", "CARD_NO", "GROSS_AMT", "MID", "RRN", "PROCESS_DATE", "Card Name", "Spay %", "Final Amount", "Code")
    
    
    For Each Code In uniqueCodes.Keys
        If merchantCodes.Exists(Code) Then
            Dim newWB As Workbook, Filename As String
            Dim SafeCode As String, MerchantName As String, ClientEmail As String
            Dim headerIndex As Long, sourceCol As Long, destCol As Long
            
            
            On Error Resume Next
            MerchantName = Application.WorksheetFunction.VLookup(Code, wsM.Range("A2:C" & mLastRow), 2, False)
            ClientEmail = Application.WorksheetFunction.VLookup(Code, wsM.Range("A2:C" & mLastRow), 3, False)
            On Error GoTo 0
            
            If MerchantName = "" Then MerchantName = "Unknown"
            If ClientEmail = "" Then GoTo SkipMail
            
            
            SafeCode = Replace(CStr(Code), "/", "_")
            SafeCode = Replace(SafeCode, "\", "_")
            SafeCode = Replace(SafeCode, ":", "_")
            SafeCode = Replace(SafeCode, "*", "_")
            SafeCode = Replace(SafeCode, "?", "_")
            SafeCode = Replace(SafeCode, "|", "_")
            
            MerchantName = Replace(MerchantName, "/", "_")
            MerchantName = Replace(MerchantName, "\", "_")
            MerchantName = Replace(MerchantName, ":", "_")
            MerchantName = Replace(MerchantName, "*", "_")
            MerchantName = Replace(MerchantName, "?", "_")
            MerchantName = Replace(MerchantName, "|", "_")
            
            ' Create new workbook
            Set newWB = Workbooks.Add
            
          
            ws.UsedRange.AutoFilter Field:=26, Criteria1:=Code
            
           
            destCol = 1
            For headerIndex = LBound(headers) To UBound(headers)
                On Error Resume Next
                sourceCol = ws.Rows(1).Find(What:=headers(headerIndex), LookIn:=xlValues, LookAt:=xlWhole).Column
                On Error GoTo 0
                
                If sourceCol > 0 Then
                    newWB.Sheets(1).Cells(1, destCol).Value = ws.Cells(1, sourceCol).Value
                    ws.Range(ws.Cells(2, sourceCol), ws.Cells(LastRow, sourceCol)).SpecialCells(xlCellTypeVisible).Copy
                    newWB.Sheets(1).Cells(2, destCol).PasteSpecial xlPasteValues
                    If headers(headerIndex) = "Spay %" Then
                        newWB.Sheets(1).Columns(destCol).NumberFormat = "0.00%"
                    End If
                    destCol = destCol + 1
                End If
            Next headerIndex
            
            newWB.Sheets(1).Columns.AutoFit
            newWB.Sheets(1).Name = SafeCode
            
            ' Save file
            Filename = FolderPath & "\" & SafeCode & " - " & MerchantName & ".xlsx"
            newWB.SaveAs Filename:=Filename, FileFormat:=xlOpenXMLWorkbook
            newWB.Close SaveChanges:=False
            
            
            If EmailFiles.Exists(ClientEmail) Then
                EmailFiles(ClientEmail) = EmailFiles(ClientEmail) & "|" & SafeCode & "|" & Filename
            Else
                EmailFiles.Add ClientEmail, SafeCode & "|" & Filename
            End If
        End If
SkipMail:
    Next Code
    
    ' === Send ONE mail per email ID with multiple attachments ===
    Dim emailKey As Variant, fileList As Variant
    Dim codesList As String
    
    For Each emailKey In EmailFiles.Keys
        codesList = ""
        fileList = Split(EmailFiles(emailKey), "|")
        
        ' Build Codes list for Subject
        For i = 0 To UBound(fileList) Step 2
            codesList = codesList & fileList(i) & ", "
        Next i
        If Right(codesList, 2) = ", " Then codesList = Left(codesList, Len(codesList) - 2)
        
        ' --- CDO Mail ---
        Dim CDO_Mail As Object, CDO_Config As Object, SMTP_Config As Object
        Set CDO_Mail = CreateObject("CDO.Message")
        Set CDO_Config = CreateObject("CDO.Configuration")
        CDO_Config.Load -1
        Set SMTP_Config = CDO_Config.Fields
        
        With SMTP_Config
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GmailID
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = GmailAppPassword
            .Update
        End With
        
        With CDO_Mail
            Set .Configuration = CDO_Config
            .From = GmailID
            .To = emailKey
            .Subject = "Axis POS Settlement Report (" & codesList & ")"
            .TextBody = "Dear Merchant," & vbNewLine & vbNewLine & _
                        "Please find attached your Axis POS Settlement Report for your reference." & vbNewLine & vbNewLine & _
                        "For any queries or clarifications, feel free to reach out to us." & vbNewLine & vbNewLine & _
                        "Best Regards," & vbNewLine & "Spay Team"
            
            
            For i = 1 To UBound(fileList) Step 2
                .AddAttachment fileList(i)
            Next i
            
            .Send
        End With
    Next emailKey
    
    ' Cleanup
    ws.AutoFilterMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Files created and grouped emails sent successfully!", vbInformation, "Task Complete"
End Sub


