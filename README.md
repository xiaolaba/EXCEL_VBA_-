# EXCEL_VBA_-
EXCEL_VBA_數字變文字



```
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
