Attribute VB_Name = "Module1"
Option Explicit

'=========================================================
' AP PAYMENT NOTIFICATION AUTOMATION
'
' False = create/display emails for review
' True  = actually send emails
'=========================================================
Private Const SEND_EMAILS As Boolean = False


Public Sub Run_AP_Payment_Notifications()

    Dim wsControl As Worksheet
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim paymentDate As Date
    
    Dim colEntity As Long
    Dim colPayDate As Long
    Dim colPayNo As Long
    Dim colVendor As Long
    Dim colPayMethod As Long
    Dim colPayAmount As Long
    Dim colInvoiceNo As Long
    Dim colInvoiceAmount As Long
    Dim colEmail1 As Long
    Dim colEmail2 As Long
    Dim colEmail3 As Long
    Dim colDear As Long
    
    Dim entities As Object
    Dim entityItem As Object
    Dim vendors As Object
    Dim vendorItem As Object
    
    Dim r As ListRow
    Dim key As Variant
    Dim vendorKey As String
    
    Dim entity As String
    Dim vendor As String
    Dim invoiceNo As String
    Dim email1 As String
    Dim email2 As String
    Dim email3 As String
    Dim dearText As String
    Dim valediction As String
    
    Dim invoiceAmount As Double
    Dim rowDate As Variant
    
    Dim olApp As Object
    Dim olMail As Object
    
    Dim emailCount As Long
    Dim htmlBody As String
    
    Dim ruleProblem As String
    
    On Error GoTo ErrHandler

    '=====================================================
    ' WORKBOOK OBJECTS
    '=====================================================
    
    Set wsControl = ThisWorkbook.Worksheets("AP Email Control")
    Set wsData = ThisWorkbook.Worksheets("AP Email Data")
    
    wsControl.Range("C4").ClearContents
    DoEvents
    
    wsControl.Range("C4").Value = "Starting..."
    DoEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'AP Email Data contains one table
    Set tbl = wsData.ListObjects(1)
    
    
    '=====================================================
    ' VALIDATE PAYMENT DATE
    '=====================================================
    
    If Not IsDate(wsControl.Range("C2").Value) Then
        
        MsgBox _
            "Please enter a valid Payment Date in cell C2.", _
            vbExclamation, _
            "AP Payment Notification"
        
        GoTo SafeExit
        
    End If
    
    paymentDate = DateValue(wsControl.Range("C2").Value)
    
    valediction = Trim(CStr(wsControl.Range("C6").Value))

    If Len(valediction) = 0 Then
        valediction = "Thank you,"
    End If
    
    
    '=====================================================
    ' STATUS
    '=====================================================
    
    wsControl.Range("C4").Value = _
        "Preparing AP payment notifications..."
    
    DoEvents
    
    
    '=====================================================
    ' POWER QUERY REFRESH
    '
    ' Intentionally disabled in the public/demo workbook.
    ' Business Central connection parameters must be
    ' configured before refresh can be enabled.
    '=====================================================
    
    'ThisWorkbook.RefreshAll
    'Application.CalculateUntilAsyncQueriesDone
    
    
    '=====================================================
    ' AP_Email_Final COLUMN POSITIONS
    '
    ' 1  Entity
    ' 2  Payment Date
    ' 3  Document No
    ' 4  Vendor No
    ' 5  Vendor
    ' 6  Payment Method
    ' 7  Payment Amount
    ' 8  Invoice No
    ' 9  Invoice Amount
    ' 10 Email 1
    ' 11 Email 2
    ' 12 Email 3
    ' 13 Dear
    ' 14 Name
    '=====================================================
    
    colEntity = 1
    colPayDate = 2
    colPayNo = 3
    colVendor = 5
    colPayMethod = 6
    colPayAmount = 7
    colInvoiceNo = 8
    colInvoiceAmount = 9
    colEmail1 = 10
    colEmail2 = 11
    colEmail3 = 12
    colDear = 13
    
    
    '=====================================================
    ' ENTITY DICTIONARY
    '
    ' One dictionary item = one entity email
    '=====================================================
    
    Set entities = CreateObject("Scripting.Dictionary")
    
    
    '=====================================================
    ' READ SELECTED PAYMENT DATE
    '=====================================================
    
    For Each r In tbl.ListRows
        
        rowDate = r.Range.Cells(1, colPayDate).Value
        
        If IsDate(rowDate) Then
            
            If DateValue(rowDate) = paymentDate Then
                
                entity = Trim(CStr( _
                    r.Range.Cells(1, colEntity).Value))
                
                vendor = Trim(CStr( _
                    r.Range.Cells(1, colVendor).Value))
                
                invoiceNo = Trim(CStr( _
                    r.Range.Cells(1, colInvoiceNo).Value))
                
                invoiceAmount = Val( _
                    r.Range.Cells(1, colInvoiceAmount).Value)
                
                
                'Skip incomplete / invalid output rows
                If Len(entity) = 0 _
                   Or Len(vendor) = 0 _
                   Or Len(invoiceNo) = 0 _
                   Or invoiceAmount = 0 Then
                
                    GoTo NextDataRow
                
                End If
                                
                
                email1 = Trim(CStr( _
                    r.Range.Cells(1, colEmail1).Value))
                
                email2 = Trim(CStr( _
                    r.Range.Cells(1, colEmail2).Value))
                
                email3 = Trim(CStr( _
                    r.Range.Cells(1, colEmail3).Value))
                
                dearText = Trim(CStr( _
                    r.Range.Cells(1, colDear).Value))
                
                
                '=========================================
                ' CREATE ENTITY IF FIRST TIME SEEN
                '=========================================
                
                If Not entities.Exists(entity) Then
                    
                    Set entityItem = _
                        CreateObject("Scripting.Dictionary")
                    
                    entityItem.Add "Email1", email1
                    entityItem.Add "Email2", email2
                    entityItem.Add "Email3", email3
                    entityItem.Add "Dear", dearText
                    
                    Set vendors = _
                        CreateObject("Scripting.Dictionary")
                    
                    entityItem.Add "Vendors", vendors
                    
                    entities.Add entity, entityItem
                    
                Else
                    
                    Set entityItem = entities(entity)
                    
                    
                    '=====================================
                    ' SAFETY CHECK
                    '
                    ' One entity email requires one
                    ' consistent recipient/greeting rule.
                    '=====================================
                    
                    If LCase(Trim(entityItem("Email1"))) <> _
                       LCase(email1) Then
                        
                        ruleProblem = _
                            "Email 1 differs within " & entity
                        
                        GoTo RuleError
                        
                    End If
                    
                    If LCase(Trim(entityItem("Email2"))) <> _
                       LCase(email2) Then
                        
                        ruleProblem = _
                            "Email 2 differs within " & entity
                        
                        GoTo RuleError
                        
                    End If
                    
                    If LCase(Trim(entityItem("Email3"))) <> _
                       LCase(email3) Then
                        
                        ruleProblem = _
                            "Email 3 differs within " & entity
                        
                        GoTo RuleError
                        
                    End If
                    
                    If Trim(entityItem("Dear")) <> dearText Then
                        
                        ruleProblem = _
                            "Dear greeting differs within " & entity
                        
                        GoTo RuleError
                        
                    End If
                    
                End If
                
                
                '=========================================
                ' GET VENDOR DICTIONARY
                '=========================================
                
                Set vendors = entityItem("Vendors")
                
                vendorKey = vendor
                
                
                '=========================================
                ' CREATE VENDOR SECTION
                '=========================================
                
                If Not vendors.Exists(vendorKey) Then
                    
                    Set vendorItem = _
                        CreateObject("Scripting.Dictionary")
                    
                    vendorItem.Add "Vendor", vendor
                    vendorItem.Add "InvoiceRows", ""
                    vendorItem.Add "VendorTotal", 0#
                    
                    vendors.Add vendorKey, vendorItem
                    
                End If
                
                
                Set vendorItem = vendors(vendorKey)
                
                
                '=========================================
                ' ADD INVOICE TO VENDOR
                '=========================================
                
                vendorItem("InvoiceRows") = _
                    vendorItem("InvoiceRows") & _
                    BuildInvoiceRow( _
                        paymentDate, _
                        invoiceNo, _
                        invoiceAmount)
                
                vendorItem("VendorTotal") = _
                    CDbl(vendorItem("VendorTotal")) + _
                    invoiceAmount
                
            End If
            
        End If
        
NextDataRow:
        
    Next r
    
    
    '=====================================================
    ' NOTHING FOUND
    '=====================================================
    
    If entities.Count = 0 Then
        
        wsControl.Range("C4").Value = _
            "No AP payment notifications found for " & _
            Format(paymentDate, "m/d/yyyy")
        
        MsgBox _
            "No AP payment notifications were found for " & _
            Format(paymentDate, "m/d/yyyy") & ".", _
            vbInformation, _
            "AP Payment Notification"
        
        GoTo SafeExit
        
    End If
    
    
    '=====================================================
    ' OPEN OUTLOOK
    '=====================================================
    
    Set olApp = CreateObject("Outlook.Application")
    
    
    '=====================================================
    ' ONE EMAIL PER ENTITY
    '=====================================================
    
    For Each key In entities.Keys
        
        entity = CStr(key)
        
        Set entityItem = entities(entity)
        Set vendors = entityItem("Vendors")
        
        
        htmlBody = BuildEntityEmail( _
            entity, _
            paymentDate, _
            entityItem("Dear"), _
            valediction, _
            vendors)
                
        Set olMail = olApp.CreateItem(0)
        
        
        With olMail
            
            .To = entityItem("Email1")
            
            .CC = BuildCC( _
                entityItem("Email2"), _
                entityItem("Email3"))
            
            
            .Subject = _
                entity & _
                " PAYMENTS " & _
                Format(paymentDate, "m/d/yy")
            
            
            '---------------------------------------------
            ' Display first so Outlook inserts signature
            '---------------------------------------------
            
            .Display
            
            
            '---------------------------------------------
            ' Put AP payment report above signature
            '---------------------------------------------
            
            .htmlBody = htmlBody & .htmlBody
            
            
            '---------------------------------------------
            ' TEST MODE / SEND MODE
            '---------------------------------------------
            
            If SEND_EMAILS Then
                .Send
            End If
            
        End With
        
        
        emailCount = emailCount + 1
        
    Next key
    
    
    '=====================================================
    ' CONFIRMATION
    '=====================================================
    
    If SEND_EMAILS Then
        
        wsControl.Range("C4").Value = _
            emailCount & _
            " entity email(s) sent for " & _
            Format(paymentDate, "m/d/yyyy")
        
        MsgBox _
            emailCount & _
            " AP payment notification(s) sent successfully.", _
            vbInformation, _
            "AP Payment Notifications Complete"
        
    Else
        
        wsControl.Range("C4").Value = _
            emailCount & _
            " entity email(s) created for review - " & _
            Format(paymentDate, "m/d/yyyy")
        
        MsgBox _
            emailCount & _
            " AP payment notification(s) created for review." & _
            vbCrLf & vbCrLf & _
            "Nothing was sent.", _
            vbInformation, _
            "AP Payment Notification Test"
        
    End If
    
    
SafeExit:
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Exit Sub
    
    
'=========================================================
' RECIPIENT/GREETING RULE ERROR
'=========================================================

RuleError:
    
    wsControl.Range("C4").Value = _
        "ERROR: " & ruleProblem
    
    MsgBox _
        "The AP payment notification process stopped because:" & _
        vbCrLf & vbCrLf & _
        ruleProblem & _
        vbCrLf & vbCrLf & _
        "Since we are creating one email per entity, " & _
        "all approved vendors within that entity must " & _
        "use the same Email 1, Email 2, Email 3 and Dear values.", _
        vbExclamation, _
        "AP Vendor Rule Conflict"
    
    GoTo SafeExit
    
    
'=========================================================
' GENERAL ERROR
'=========================================================

ErrHandler:
    
    Dim ErrorMessage As String
    
    ErrorMessage = Err.Description
    
    On Error Resume Next
    
    wsControl.Range("C4").Value = _
        "ERROR: " & ErrorMessage
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox _
        "AP payment notification process stopped:" & _
        vbCrLf & vbCrLf & _
        ErrorMessage, _
        vbCritical, _
        "AP Payment Notification Error"

End Sub


'=========================================================
' BUILD FULL ENTITY EMAIL
'=========================================================

Private Function BuildEntityEmail( _
    ByVal entity As String, _
    ByVal paymentDate As Date, _
    ByVal dearText As String, _
    ByVal valediction As String, _
    ByVal vendors As Object) As String

    Dim html As String
    Dim vendorKey As Variant
    Dim vendorItem As Object
    
    
    html = _
        "<html><body style='font-family:Calibri;" & _
        "font-size:11pt;'>"
    
    
    '=====================================================
    ' GREETING
    '
    ' Use Dear exactly as entered in AP_Vendors.
    '=====================================================
    
    If Len(Trim(dearText)) > 0 Then
        
        html = html & _
            "<p>" & _
            HtmlEncode(dearText) & _
            "</p>"
        
    End If
    
    
    '=====================================================
    ' INTRO
    '=====================================================
    
    html = html & _
        "<p>We made these " & _
        HtmlEncode(entity) & _
        " payments today.</p>"
    
    
    '=====================================================
    ' ENTITY HEADING
    '=====================================================
    
    html = html & _
        "<div style='text-align:left;" & _
        "font-weight:bold;font-size:13pt;" & _
        "margin-top:18px;margin-bottom:14px;'>"
    
    html = html & HtmlEncode(entity)
    
    html = html & "</div>"
    
    
    '=====================================================
    ' EACH VENDOR
    '=====================================================
    
    For Each vendorKey In vendors.Keys
        
        Set vendorItem = vendors(vendorKey)
        
        
        '-------------------------------------------------
        ' Vendor name
        '-------------------------------------------------
        
        html = html & _
            "<div style='font-weight:bold;" & _
            "margin-top:16px;margin-bottom:4px;'>"
        
        html = html & _
            HtmlEncode(vendorItem("Vendor"))
        
        html = html & "</div>"
        
        
        '-------------------------------------------------
        ' Vendor invoice table
        '-------------------------------------------------
        
        html = html & _
            "<table width='500' cellpadding='0' cellspacing='0' " & _
            "style='width:500px;border-collapse:collapse;" & _
            "table-layout:fixed;font-family:Calibri;" & _
            "font-size:11pt;margin-bottom:12px;'>"
        
        
        ' Header
        
        html = html & "<tr>"
        
        html = html & _
            "<th style='border:1px solid #808080;" & _
            "padding:4px 10px;text-align:left;" & _
            "min-width:90px;'>Date</th>"
        
        html = html & _
            "<th style='border:1px solid #808080;" & _
            "padding:4px 10px;text-align:left;" & _
            "min-width:260px;'>Invoice No.</th>"
        
        html = html & _
            "<th style='border:1px solid #808080;" & _
            "padding:4px 10px;text-align:right;" & _
            "min-width:100px;'>Amount</th>"
        
        html = html & "</tr>"
        
        
        ' Invoice rows
        
        html = html & vendorItem("InvoiceRows")
        
        
        '-------------------------------------------------
        ' Vendor total
        '-------------------------------------------------
        
        html = html & "<tr>"
        
        html = html & _
            "<td style='border:1px solid #808080;" & _
            "padding:4px 10px;'></td>"
        
        html = html & _
            "<td style='border:1px solid #808080;" & _
            "padding:4px 10px;'><b>Total</b></td>"
        
        html = html & _
            "<td style='border:1px solid #808080;" & _
            "padding:4px 10px;text-align:right;'><b>"
        
        html = html & _
            Format( _
                CDbl(vendorItem("VendorTotal")), _
                "#,##0.00")
        
        html = html & "</b></td>"
        
        html = html & "</tr>"
        
        html = html & "</table>"
        
    Next vendorKey
    
    
    html = html & "<br>"
    
    If Len(Trim(valediction)) > 0 Then
        html = html & _
            "<br><p>" & _
            HtmlEncode(valediction) & _
            "</p>"
    End If
    
    html = html & "</body></html>"
    
    
    BuildEntityEmail = html

End Function


'=========================================================
' BUILD ONE INVOICE ROW
'=========================================================

Private Function BuildInvoiceRow( _
    ByVal paymentDate As Date, _
    ByVal invoiceNo As String, _
    ByVal invoiceAmount As Double) As String

    Dim html As String
    
    
    html = "<tr>"
    
    
    'Date
    
    html = html & _
        "<td style='border:1px solid #808080;" & _
        "padding:4px 10px;'>"
    
    html = html & _
        Format(paymentDate, "m/d/yyyy")
    
    html = html & "</td>"
    
    
    'Invoice
    
    html = html & _
        "<td style='border:1px solid #808080;" & _
        "padding:4px 8px;width:260px;'>"
    
    html = html & HtmlEncode(invoiceNo)
    
    html = html & "</td>"
    
    
    'Amount
    
    html = html & _
        "<td style='border:1px solid #808080;" & _
        "padding:4px 10px;text-align:right;'>"
    
    html = html & _
        Format(invoiceAmount, "#,##0.00")
    
    html = html & "</td>"
    
    
    html = html & "</tr>"
    
    
    BuildInvoiceRow = html

End Function


'=========================================================
' BUILD CC
'=========================================================

Private Function BuildCC( _
    ByVal email2 As String, _
    ByVal email3 As String) As String

    Dim result As String
    
    
    If Len(Trim(email2)) > 0 Then
        result = Trim(email2)
    End If
    
    
    If Len(Trim(email3)) > 0 Then
        
        If Len(result) > 0 Then
            result = result & ";"
        End If
        
        result = result & Trim(email3)
        
    End If
    
    
    BuildCC = result

End Function


'=========================================================
' HTML ENCODING
'=========================================================

Private Function HtmlEncode( _
    ByVal txt As String) As String

    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    txt = Replace(txt, """", "&quot;")
    
    HtmlEncode = txt

End Function

