Attribute VB_Name = "ExportVBACode"
'--------------������ ������� ��������� ��� �������� ���� VBA �� ������� ������-------------
'------------------����� ����� ���� ����������� ��������� ��� ����� ������------------------
Public Sub SaveVBACode()

    ExportVBA ThisWorkbook.Path & "\����\"

End Sub

Public Sub ExportVBA(sDestinationFolder As String)
'���������� ������� ����
    Dim oVBComponent As Object

    For Each oVBComponent In ThisWorkbook.VBProject.VBComponents
        If oVBComponent.Type = 1 Then
            ' Standard Module
            oVBComponent.Export sDestinationFolder & oVBComponent.Name & ".bas"
        ElseIf oVBComponent.Type = 2 Then
            ' Class
            oVBComponent.Export sDestinationFolder & oVBComponent.Name & ".cls"
        ElseIf oVBComponent.Type = 3 Then
            ' Form
            oVBComponent.Export sDestinationFolder & oVBComponent.Name & ".frm"
        ElseIf oVBComponent.Type = 100 Then
            ' Document
            oVBComponent.Export sDestinationFolder & oVBComponent.Name & ".bas"
        Else
            ' UNHANDLED/UNKNOWN COMPONENT TYPE
        End If
    Next oVBComponent

End Sub

