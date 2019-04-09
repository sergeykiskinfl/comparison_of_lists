'  Copyright 2019 Sergey Kiskin

'  Licensed under the Apache License, Version 2.0 (the "License");
'  you may not use this file except in compliance with the License.
'  You may obtain a copy of the License at
'
'      http://www.apache.org/licenses/LICENSE-2.0
'
'  Unless required by applicable law or agreed to in writing, software
'  distributed under the License is distributed on an "AS IS" BASIS,
'  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or  implied.
'  See the License for the specific language governing permissions and
'  limitations under the License.

Option Explicit
Dim EA As Excel.Application
Dim WB As Excel.Workbook 'Обрабатываемая книга
Dim WS As Excel.Worksheet 'Обрабатываемый лист
Dim i&, n&
'i, n, - переменные - счетчики для циклов
Dim iLastRow&(0 To 1) 'Переменная для определения последней строки листа
Dim dictLists(0 To 3) As Object 'Массив словарей
Dim varKey 'Переменная для перебора ключей в словаре

Public Sub comparisonOfRange()

'Предварительная обработка
Set EA = Excel.Application

With EA

    .ScreenUpdating = False: .DisplayAlerts = False: .StatusBar = False

End With

Set WB = EA.Workbooks("Comparison_Of_Lists.xlsm")
Set WS = WB.Worksheets("Comparison_of_lists")

'Создание словарей
For i = 0 To 3
       
       Set dictLists(i) = CreateObject("Scripting.Dictionary")
              
Next i

iLastRow(0) = lastRow(WS)
iLastRow(1) = lastRow(WS, 2)

With WS

'Заполнение двух словарей данными диапазонов
For n = 0 To 1

    For i = 2 To iLastRow(n)
        
        If .Cells(i, n + 1).Value <> "" And Not dictLists(n).Exists(.Cells(i, n + 1).Value) Then
        
            dictLists(n).Add .Cells(i, n + 1).Value, 1
       
        End If
       
    Next i

Next n

End With

'Поиск различий списков и заполнение ими новых словарей
With dictLists(0)

    For Each varKey In .Keys
        
       If dictLists(1).Exists(varKey) Then
       
       Else:
       
        dictLists(2).Add varKey, 1
       
       End If
         
    Next varKey

End With

With dictLists(1)

    For Each varKey In .Keys
        
       If dictLists(0).Exists(varKey) Then
       
       Else:
       
        dictLists(3).Add varKey, 1
       
       End If
         
    Next varKey

End With

'Распечатка различий
With WS

    For i = 0 To dictLists(2).Count - 1
          
         .Cells(i + 2, 5).Value = dictLists(2).Keys()(i)
    
    Next i
    
    For i = 0 To dictLists(3).Count - 1
          
         .Cells(i + 2, 7).Value = dictLists(3).Keys()(i)
    
    Next i

End With

With EA

    .ScreenUpdating = True: .DisplayAlerts = True

End With

End Sub




