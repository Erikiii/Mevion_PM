Attribute VB_Name = "Module1"
Option Explicit

Public Sub SendPMAlerts()

    Const RECIPIENTS As String = _
        "xihao.han@mevion.com;yiqian.lv@mevion.com;bin.yu@mevion.com;peng.cui@mevion.com;wendong.tian@mevion.com"
    Const ALERT_DAYS As Long = 5

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim outlookApp As Object
    Dim email As Object
    Dim row As ListRow
    Dim dueDate As Variant
    Dim lastAlert As Variant
    Dim daysLeft As Long
    Dim body As String
    Dim alertRows As Collection
    Dim item As Variant

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("PM Schedule")
    Set tbl = ws.ListObjects("PM_Schedule")
    Set alertRows = New Collection

    'Refresh TODAY() and all schedule formulas before checking.
    Application.CalculateFull
    DoEvents

    body = "<html><body>" & _
           "<p>The following preventive maintenance tasks are due " & _
           "within " & ALERT_DAYS & " days:</p>" & _
           "<table border='1' cellpadding='5' cellspacing='0'>" & _
           "<tr style='background:#1F4E78;color:white;'>" & _
           "<th>Task ID</th><th>Equipment / Area</th>" & _
           "<th>Maintenance Task</th><th>Next Due</th>" & _
           "<th>Days Left</th></tr>"

    For Each row In tbl.ListRows

        dueDate = row.Range.Cells(1, _
            tbl.ListColumns("Next Due").Index).value

        lastAlert = row.Range.Cells(1, _
            tbl.ListColumns("Last Alert Sent").Index).value

        If IsDate(dueDate) Then
            daysLeft = DateDiff("d", Date, CDate(dueDate))

            'Include tasks due today through five days ahead.
            If daysLeft >= 0 And daysLeft <= ALERT_DAYS Then

                'Do not send the same task more than once today.
                If Not IsDate(lastAlert) _
                   Or DateValue(CDate(lastAlert)) <> Date Then

                    body = body & "<tr>" & _
                        "<td>" & HtmlEncode(row.Range.Cells(1, _
                            tbl.ListColumns("Task ID (WI)").Index).Text) & "</td>" & _
                        "<td>" & HtmlEncode(row.Range.Cells(1, _
                            tbl.ListColumns("Equipment / Area").Index).Text) & "</td>" & _
                        "<td>" & HtmlEncode(row.Range.Cells(1, _
                            tbl.ListColumns("Maintenance Task").Index).Text) & "</td>" & _
                        "<td>" & Format(CDate(dueDate), "yyyy-mm-dd") & "</td>" & _
                        "<td>" & daysLeft & "</td></tr>"

                    alertRows.Add row
                End If
            End If
        End If
    Next row

    If alertRows.Count = 0 Then Exit Sub

    body = body & "</table>" & _
           "<p>Please review and arrange the required maintenance.</p>" & _
           "</body></html>"

    'Connect to Classic Outlook.
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler

    If outlookApp Is Nothing Then
        Err.Raise vbObjectError + 1000, , _
            "Classic Outlook could not be started."
    End If

    Set email = outlookApp.CreateItem(0)

    With email
        .To = RECIPIENTS
        .Subject = "Preventive Maintenance Alert - " & _
                   Format(Date, "yyyy-mm-dd")
        .HTMLBody = body
        .Send
        'Use .Display instead of .Send while testing.
    End With
    
    'Adding delay for Outlook to wait a bit
    Application.Wait (Now + TimeValue("0:00:30"))

    'Record today's alert date only after Outlook accepts the email.
    For Each item In alertRows
        item.Range.Cells(1, _
            tbl.ListColumns("Last Alert Sent").Index).value = Date
    Next item

    ThisWorkbook.Save
    Exit Sub

ErrorHandler:
    Debug.Print "PM alert error: " & Err.Number & " - " & Err.Description

End Sub

Private Function HtmlEncode(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    HtmlEncode = value
End Function

