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

    ' Получить путь и имя текущего файла проекта
    projectPath = ActiveProject.Path
    projectName = ActiveProject.Name
    
    ' Формирование пути для сохранения экспортированного файла Excel рядом с mpp файлом
    exportFilePath = projectPath & "\" & Replace(projectName, ".mpp", "_экспорт.xlsx")
    
    ' Название схемы экспорта
    exportMapName = "Экспорт ГПР test"
    
    ' Удалить существующий файл, если он есть
    On Error Resume Next
    Kill exportFilePath
    On Error GoTo 0
    
    ' Экспортировать данные с использованием существующей схемы экспорта
    FileSaveAs Name:=exportFilePath, _
               FormatID:="MSProject.ACE", _
               map:=exportMapName
               
    ' Открыть файл Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(exportFilePath)
    Set xlWorksheet = xlWorkbook.Sheets(1)
    
    ' Получить последнюю заполненную строку в столбце C
    lastRow = xlWorksheet.Cells(xlWorksheet.Rows.Count, "C").End(-4162).Row ' xlUp
    
    ' Форматирование столбцов дат
    With xlWorksheet
        ' Установить формат ячеек для столбцов C, D, H, I
        .Range("C2:C" & lastRow).NumberFormat = "dd.mm.yyyy"
        .Range("D2:D" & lastRow).NumberFormat = "dd.mm.yyyy"
        .Range("H2:H" & lastRow).NumberFormat = "dd.mm.yyyy"
        .Range("I2:I" & lastRow).NumberFormat = "dd.mm.yyyy"
        
        ' Обновить значения ячеек
        For i = 2 To lastRow
            ' Прочитать значение ячейки и преобразовать его в нужный формат
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
    
    ' Сохранить и закрыть файл
    xlWorkbook.Save
    xlWorkbook.Close
    
    ' Очистить объекты Excel
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    ' Сообщение об успешном экспорте
    MsgBox "Проект экспортирован и отформатирован в " & exportFilePath
End Sub


