# EXCEL_VBA_數字變文字

EXCEL - 如何將數字轉變成文字



```
'EXCEL - 如何將數字轉變成文字
sub number2text
    'EXCEL_VBA_數字變文字  
    最大的行號 = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    最大的列號 = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column

    Do While ActiveCell.Row <= 最大的行號
      
       
        '在數字的前面加上單引號, 結果存到相鄰的空格內
        ActiveCell.Offset(0, 1).FormulaR1C1 = "'" & ActiveCell.Value
       
        '移到下一格
        ActiveCell.Offset(1, 0).Select
       
    '不斷的重複一上的動作, 直到最後一行
    Loop
end-sub

```


.  
.  
.  


# EXCEL VBA - 如何將文字轉變成數字

```
'EXCEL - 如何將文字轉變成數字
'簡單描述這個 EXCEL MACRO 的用途

'1) 文字格式的數目轉變成 EXCEL 認得, 可以加總的數字
'2) 去除文字首尾的多餘空格
'3) 去除多餘的 0
'…..還沒想到….



'copy the below——————

Sub TEXT_NUMBER()
'
' TEXT_NUMBER Macro
'  xiao_laba@yahoo.com.cn 在 3/6/2010 的巨集
'

Dim row As Integer
Dim col As Integer

Dim temp As Integer

    'EXCEL, 選整個 SHEET, 格式化, 等同[儲存格格式], [數值] [G/通用格式]
    Cells.Select
    Selection.NumberFormatLocal = "General"

    '計算 EXCEL SHEET 內包含資料的 行 x 列
    最大的行號 = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).row
    最大的列號 = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column


    row = 1
    col = 1   '設定從最左上角的地一個開始
   
   
    Do While col <= 最大的列號
         
   
        '在右邊插入空白列
        Columns(col + 1).Select
        Selection.NumberFormatLocal = "General" '必須設定格式為"通用", 否則插入的列可能將數字看成是文字
        Selection.Insert Shift:=xlToRight
       
        '禁止螢幕更新, VBA 處理資料時不閃爍, 可不用但速度變慢
        Application.ScreenUpdating = False
   
   
        Do While row <= 最大的行號
           
            '點選每列的第一格
            Cells(row, col).Select
                       
            If IsNull(ActiveCell) Then ActiveCell = ""  '某些CELL看似空白 ,但寫入公式會錯誤, 所以加入這行
            If IsEmpty(ActiveCell) Then ActiveCell = "" '某些CELL看似空白 ,但寫入公式會錯誤, 所以加入這行
            ActiveCell = Trim(ActiveCell.Value)         '不論格式, 先去除每格資料的頭尾的空格
           
            '檢查存儲格的內容是否包含純數字 (包括有單引號開頭的文本格式的數字)
            If IsNumeric(ActiveCell) Then
           
                '純數字的話, 右面的存儲格套入公式 X 1, 文字格式的數目就轉換成EXCEL認得的數字, 只存整數部份
                ActiveCell.Offset(0, 1) = "=RoundUp(RC[-1]*1,0)"
                ActiveCell.Offset(0, 1) = ActiveCell.Offset(0, 1).Value '去除公式, 只存結果
                ActiveCell.Offset(0, 1).NumberFormatLocal = "General"   ’
                If ActiveCell.Offset(0, 1) = 0 Then ActiveCell.Offset(0, 1) = "" '如果 = 0, 留空
               
            Else
               
                '如果非純數字, 直接依照文字格式存到右面的存儲格
                ActiveCell.Offset(0, 1) = ActiveCell.Value
           
            End If
           

            '移到下一格
            row = row + 1
          
        Loop    '不斷的重複以上的動作, 直到該列的最後一行
       
       
        '完成了這一列, 需要刪除此列, 並移到下一列, 第一行
        Columns(col).Select
        Selection.Delete Shift:=xlToRight   '刪除此列
        row = 1 '指向第一行
        col = col + 1   '指向下一列
        Cells(row, col).Select  '點選指向的 CELL
       
       
        '回覆螢幕更新, VBA 處理資料時會閃爍, 可不用但速度變慢
        Application.ScreenUpdating = True

   
    Loop        '不斷的重複一上的動作, 直到最後一列
   
   
    '完成處理整個 EXCEL SHEET 的資料後, 點選最左上角的 CELL
    Cells(1, 1).Select

End Sub



'copy the above——————

```

