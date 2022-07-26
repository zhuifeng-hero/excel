# excel
VBA开发增加excel功能
## 使用说明
1. 导入自定义菜单
2. 添加自定义加载项
3. 调用宏文件

## 功能
### 合并多电子表
```
Sub MergeWorkbooks()

Dim usd As Integer
Dim wk, sht, wkn, awk, shtt As String

awk = ActiveWorkbook.Name
usd = Workbooks("VBA-Kevin.xlsm").Sheets("Sheets").Range("A1").CurrentRegion.Rows.Count
For i = 2 To usd
    wk = Workbooks("VBA-Kevin.xlsm").Sheets("Sheets").Range("A" & i).Value
    wkn = Workbooks("VBA-Kevin.xlsm").Sheets("Sheets").Range("C" & i).Value
    sht = Workbooks("VBA-Kevin.xlsm").Sheets("Sheets").Range("B" & i).Value

    shtt = Workbooks("VBA-Kevin.xlsm").Sheets("Sheets").Range("D" & i).Value
    Workbooks.Open Filename:=wk
If shtt = "" Then
    Workbooks(wkn).Sheets(1).Copy Before:=Workbooks(awk).Sheets(1)
    ActiveSheet.Name = sht
Else
    If shtt = "全" Then
        For j = 1 To Workbooks(wkn).Sheets.Count
         Workbooks(wkn).Sheets(j).Copy Before:=Workbooks(awk).Sheets(1)
         ActiveSheet.Name = sht & "!" & Workbooks(wkn).Sheets(j).Name
        Next
    Else
    On Error Resume Next
    Workbooks(wkn).Sheets(shtt).Name = sht
    
    Workbooks(wkn).Sheets(sht).Copy Before:=Workbooks(awk).Sheets(1)
   ' ActiveSheet.Name = sht
    End If
End If
Workbooks(wkn).Close SaveChanges:=False
Next
End Sub
```
### 合并Excel多小表
### 添加目录及超链接
### 自定义函数
- 拆分列
- 
