Attribute VB_Name = "export"
Sub ExportAndFormatExcel()
    Dim projectPath As String
    Dim projectName As String
    Dim exportFilePath As String
    Dim exportMapName As String
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Dim lastRow As Long
    Dim i As Long

    ' �������� ���� � ��� �������� ����� �������
    projectPath = ActiveProject.Path
    projectName = ActiveProject.Name
    
    ' ������������ ���� ��� ���������� ����������������� ����� Excel ����� � mpp ������
    exportFilePath = projectPath & "\" & Replace(projectName, ".mpp", "_�������.xlsx")
    
    ' �������� ����� ��������
    exportMapName = "������� ��� test"
    
    ' ������� ������������ ����, ���� �� ����
    On Error Resume Next
    Kill exportFilePath
    On Error GoTo 0
    
    ' �������������� ������ � �������������� ������������ ����� ��������
    FileSaveAs Name:=exportFilePath, _
               FormatID:="MSProject.ACE", _
               map:=exportMapName
               
    ' ������� ���� Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(exportFilePath)
    Set xlWorksheet = xlWorkbook.Sheets(1)
    
    ' �������� ��������� ����������� ������ � ������� C
    lastRow = xlWorksheet.Cells(xlWorksheet.Rows.Count, "C").End(-4162).Row ' xlUp
    
    ' �������������� �������� ���
    With xlWorksheet
        ' ���������� ������ ����� ��� �������� C, D, H, I
        .Range("C2:C" & lastRow).NumberFormat = "dd.mm.yyyy"
        .Range("D2:D" & lastRow).NumberFormat = "dd.mm.yyyy"
        .Range("H2:H" & lastRow).NumberFormat = "dd.mm.yyyy"
        .Range("I2:I" & lastRow).NumberFormat = "dd.mm.yyyy"
        
        ' �������� �������� �����
        For i = 2 To lastRow
            ' ��������� �������� ������ � ������������� ��� � ������ ������
            Dim oldValue As Date
            
            oldValue = .Cells(i, "C").Value
            .Cells(i, "C").Value = oldValue
            
            oldValue = .Cells(i, "D").Value
            .Cells(i, "D").Value = oldValue
            
            oldValue = .Cells(i, "H").Value
            .Cells(i, "H").Value = oldValue
            
            oldValue = .Cells(i, "I").Value
            .Cells(i, "I").Value = oldValue
        Next i
    End With
    
    ' ��������� � ������� ����
    xlWorkbook.Save
    xlWorkbook.Close
    
    ' �������� ������� Excel
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    ' ��������� �� �������� ��������
    MsgBox "������ ������������� � �������������� � " & exportFilePath
End Sub


