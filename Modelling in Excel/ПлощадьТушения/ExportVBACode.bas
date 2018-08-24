Attribute VB_Name = "ExportVBACode"
'--------------������ ������� ��������� ��� �������� ���� VBA �� ������� ������-------------
'------------------����� ����� ���� ����������� ��������� ��� ����� ������------------------
Public Sub SaveVBACode()

    ExportVBA ThisWorkbook.Path & "\��������������\"

End Sub

Public Sub ExportVBA(sDestinationFolder As String)
'���������� ������� ����
    Dim oVBComponent As Object

    For Each oVBComponent In ThisWorkbook.VBProject.VBComponents
        If oVBComponent.Type = 1 Then
            ' Standard Module
            fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 2 Then
            ' Class
            fullName = sDestinationFolder & oVBComponent.Name & ".cls"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 3 Then
            ' Form
            fullName = sDestinationFolder & oVBComponent.Name & ".frm"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 100 Then
            ' Document
            fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            oVBComponent.Export fullName
        Else
            ' UNHANDLED/UNKNOWN COMPONENT TYPE
        End If
        Debug.Print "�������� " & fullName
    Next oVBComponent

End Sub

