Sub CloseMailMerge()
'
' CloseMailMerge Macro
'
'
' clear mail merge association on this doc
 ActiveDocument.MailMerge.MainDocumentType = wdNotAMergeDocument
End Sub
Sub SelectedMergeFieldRemove()
'
' SelectedMergeFieldRemove Macro
'
'
    Dim fld As Field
'    Application.ScreenUpdating = False

'    If Selection.Fields.Count = 1 Then
    For Each fld In Selection.Fields
        If fld.Type = wdFieldMergeField Then
            fld.Unlink
        End If
    Next fld
End Sub
Sub FindBill()
'
' FindBill Macro
'
'
    Dim sBillNumber As String
    Dim lastActiveRecord As Integer

    On Error Resume Next

    ' Retrieve bill number from last search as default
    sBillNumber = ActiveDocument.Variables("sBillNum").Value
    If Err.Number = 5825 Then
        sBillNumber = ""
    End If

    sBillNumber = InputBox("Enter bill number", "Bill Number", sBillNumber)

    With ActiveDocument.MailMerge.DataSource
        If StrPtr(sBillNumber) = 0 Then
'            MsgBox ("canceled!")
            Exit Sub
        ElseIf sBillNumber = vbNullString Then
            MsgBox ("Resetting to beginning of data!")
            .ActiveRecord = wdFirstRecord
            Exit Sub
        End If

        lastActiveRecord = .ActiveRecord
        Application.ScreenUpdating = False

        ' Find the selected record (exact match).
        Do While .FindRecord(FindText:=sBillNumber, Field:="Bill_") = True
            If .DataFields("Bill_") = sBillNumber Then
                Exit Do
            End If
        Loop
        Application.ScreenUpdating = True

        If .DataFields("Bill_") <> sBillNumber Then
            .ActiveRecord = lastActiveRecord
            MsgBox ("Bill " & sBillNumber & " not found!")
        End If

        ActiveDocument.Variables("sBillNum") = sBillNumber
    End With
End Sub

Public Sub ListAllVariables()
    Dim V As Variable, S As String
    For Each V In ActiveDocument.Variables
        S = S & V.Name & vbTab & V.Value & vbNewLine
    Next V
    MsgBox S
End Sub

