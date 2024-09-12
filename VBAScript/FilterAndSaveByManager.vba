'-------------------------------------
' Creation date : 2024/09/10
' Last update   : 2024/09/12
' Author        : F.W.Yue
' Tested on WPS Office 11.1.0.10162
' Description   :用于Excel报表的拆分
'-------------------------------------

Sub FilterAndSaveByManager()
    Dim manager As Variant
    Dim managers As String
    Dim managerList() As String
    Dim ws As Worksheet
    Dim originalFileName As String
    Dim newFilePath As String
    Dim desktopPath As String
    Dim wb As Workbook
    Dim success As Boolean
    Dim searchRowNumber As Long
    Dim fileExists As Boolean

    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    originalFileName = Replace(ThisWorkbook.Name, "市场部", "")
    managers = InputBox("请输入大区经理的名字，用逗号分隔：", "筛选大区经理", "刘波,胡冰雪")
    managerList = Split(managers, ",")

    For Each manager In managerList
        manager = Trim(manager)
        newFilePath = desktopPath & manager & "大区_" & originalFileName
        ThisWorkbook.SaveCopyAs (newFilePath)
        Set wb = Workbooks.Open(newFilePath)

        For Each ws In wb.Sheets
            If ws.PivotTables.Count = 0 Then
                success = FilterWorksheetByColumn(wb, ws, manager, 1, "大区经理")
                If Not success Then
                    success = FilterWorksheetByColumn(wb, ws, manager, 3, "大区经理")
                End If
            End If
        Next ws

        For Each ws In wb.Sheets
            If ws.PivotTables.Count > 0 Then
                Dim pt As PivotTable
                For Each pt In ws.PivotTables
                    On Error Resume Next
                    pt.PivotCache.Refresh
                    On Error GoTo 0
                Next pt
            End If
        Next ws

        wb.Close SaveChanges:=True
    Next manager
End Sub

Function FilterWorksheetByColumn(wb As Workbook, ws As Worksheet, filterValue As Variant, Optional searchRowNumber As Long = 1, Optional searchStr As String = "大区经理") As Boolean
    Dim searchRow As Range
    Dim targetCol As Range
    Dim lastRowNumber As Integer '最后一行名称
    Dim found As Boolean '是否找到
    Dim colNumber As Integer '筛选列号
    Dim tempWs As Worksheet '用于复制的临时表
    Dim usedRange As Range '复制范围
    Dim targetRange As Range '筛选范围
    Dim sheetName As String '表名
    Dim cellToCheck As Range
    Dim currentRow As Range
    Dim rowsToDelete As Range
    Dim numberOfRowsToDelete As Long
    Dim cellValue As Variant
    '配置信息打印
    Debug.Print "筛选的值: " & filterValue
    Debug.Print "查找的工作表: " & ws.Name
    Debug.Print "查找的行号: " & searchRowNumber
    '初始化
    Set rowsToDelete = Nothing
    Set rowsToDelete = Nothing
    numberOfRowsToDelete = 0
    FilterWorksheetByAddress = False
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    cellValue = ws.Cells(1, 1).Value
    Debug.Print "搜索行: " & cellValue
    
    ' 检查搜索行是否为空
    Set searchRow = ws.Rows(searchRowNumber)
    cellValue = searchRow.Cells(1, 1).Value
    Debug.Print "搜索行: " & cellValue
    If IsEmpty(cellValue) Then
        Debug.Print "搜索行是空的，无法继续"
        Exit Function
    End If

    ' 查找目标列
    Set targetCol = searchRow.Find(What:=searchStr, LookAt:=xlWhole)
    If targetCol Is Nothing Then Exit Function

    colNumber = targetCol.Column
    Debug.Print "查找到的列号: " & colNumber

    lastRowNumber = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Debug.Print "最后一列的行号: " & lastRowNumber

     ' 检查 searchRowNumber + 1 行 colNumber 列的单元格是否为空
     Set cellToCheck = ws.Cells(searchRowNumber + 1, colNumber)
    If IsEmpty(cellToCheck) Then
        startRowNumber = searchRowNumber + 1
    Else
        startRowNumber = searchRowNumber
    End If
    
    ' 遍历指定范围内的行,-1表示倒序
    For i = lastRowNumber To startRowNumber + 1 Step -1
        Set currentRow = ws.Cells(i, colNumber)
        If Not currentRow.Value = filterValue Then
            ' 如果当前行的值不等于过滤值，则添加到要删除的行范围
            If rowsToDelete Is Nothing Then
                Set rowsToDelete = currentRow.EntireRow
            Else
                Set rowsToDelete = Union(rowsToDelete, currentRow.EntireRow)
            End If
        End If
    numberOfRowsToDelete = numberOfRowsToDelete + 1
    Next i
    Debug.Print "需删除行数为: " & numberOfRowsToDelete
    
    
      '将超级表转换为区域
    If ws.ListObjects.Count > 0 Then
        For Each tbl In ws.ListObjects
            tbl.Unlist
        Next tbl
        Debug.Print "超级表已转换为区域"
    Else
        Debug.Print "无超级表"
    End If
    
    ' 删除所选行
    If Not rowsToDelete Is Nothing Then
        rowsToDelete.Delete Shift:=xlUp
    End If

End Function
