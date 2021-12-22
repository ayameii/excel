Application.ScreenUpdating = False   '关闭屏幕更新
Cells(k + 1, 1).copy Cells(k , 1)       'k+1行1列的单元格复制到k行1列
Cells(k + 1, 1).ClearContents         '清除单元格内容
Cells(k + 1, 1).Resize(1, 13).Copy            '目标单元格之后的1行13列单元格全部复制
Sheets(i).Select                                  '第 i 个sheet被选中
For i =2 to 14 ---------Next                            'for 循环以2开始到14结束
Application.CutCopyMode = False                '清除剪切板内容
set ch = Nothing                                        '清除对象内容  
