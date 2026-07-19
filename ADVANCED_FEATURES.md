# 🚀 Расширенные возможности - Рулетка VBA

Документация для добавления продвинутых функций в игру Рулетка.

---

## 📋 Содержание

1. [Звуковые эффекты](#звуковые-эффекты)
2. [Графики статистики](#графики-статистики)
3. [Визуализация колеса](#визуализация-колеса)
4. [Автоматические стратегии](#автоматические-стратегии)
5. [Сохранение в БД](#сохранение-в-бд)

---

## 🔊 Звуковые эффекты

### Предпосылка

В папке с игрой должен быть файл: `WheelOfRulete.wav`

### Создание модуля SoundEffects.bas

Создайте новый модуль в VBA редакторе и скопируйте этот код:

```vba
' ======================================
' SOUND EFFECTS MODULE
' ======================================
' Модуль для работы со звуками

Option Explicit

' Windows API declaration for playing sounds
#If Win64 Then
    Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (lpszName As String, hModule As LongPtr, dwFlags As Long) As Long
#Else
    Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (lpszName As String, hModule As Long, dwFlags As Long) As Long
#End If

' Sound flags
Const SND_FILENAME = &H20000
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2

' Get application path
Function GetAppPath() As String
    GetAppPath = ThisWorkbook.Path & "\"
End Function

' Play wheel spin sound
Sub PlayWheelSound()
    Dim soundFile As String
    soundFile = GetAppPath & "WheelOfRulete.wav"
    
    ' Check if file exists
    If Dir(soundFile) <> "" Then
        PlaySound soundFile, 0, SND_FILENAME Or SND_ASYNC
    End If
End Sub

' Play winning sound
Sub PlayWinSound()
    ' System beep for win
    ' Можно добавить свой файл win.wav
    Dim soundFile As String
    soundFile = GetAppPath & "win.wav"
    
    If Dir(soundFile) <> "" Then
        PlaySound soundFile, 0, SND_FILENAME Or SND_ASYNC
    Else
        ' Fallback: System beep
        Beep
    End If
End Sub

' Play losing sound
Sub PlayLoseSound()
    ' System beep for lose
    ' Можно добавить свой файл lose.wav
    Dim soundFile As String
    soundFile = GetAppPath & "lose.wav"
    
    If Dir(soundFile) <> "" Then
        PlaySound soundFile, 0, SND_FILENAME Or SND_ASYNC
    Else
        ' Fallback: different pitch
        Beep
    End If
End Sub

' Stop sound
Sub StopSound()
    PlaySound "", 0, 0
End Sub
```

### Интеграция со звуками в RouletteGame.bas

В функции `DisplaySpinResult` добавьте:

```vba
Sub DisplaySpinResult(spinResult As Integer, resultColor As String, betType As String, betValue As String, winnings As Currency)
    Dim message As String
    
    ' Play wheel sound when spinning
    Call PlayWheelSound()
    
    ' Wait for wheel sound to finish
    Application.Wait Now + TimeValue("0:00:01")
    
    ' [Остальной код как обычно...]
    
    ' Play result sound
    If winnings > 0 Then
        Call PlayWinSound()
        message = message & "✓ ВЫ ВЫИГРАЛИ! " & vbCrLf
    Else
        Call PlayLoseSound()
        message = message & "✗ ВЫ ПРОИГРАЛИ!" & vbCrLf
    End If
    
    ' [Остальной код...]
End Sub
```

### Требуемые звуковые файлы

В папке с игрой создайте подпапку `sounds/`:

```
Roulette_Game/
├── sounds/
│   ├── WheelOfRulete.wav     ← Звук вращения колеса
│   ├── win.wav                ← Звук выигрыша
│   └── lose.wav               ← Звук проигрыша
├── Roulette_Game.xlsm
└── ...
```

---

## 📊 Графики статистики

### Создание модуля Statistics.bas

```vba
' ======================================
' STATISTICS & CHARTS MODULE
' ======================================
' Создание графиков и анализа статистики

Option Explicit

' Create statistics sheet
Sub CreateStatisticsSheet()
    Dim ws As Worksheet
    Dim chartWs As Worksheet
    Dim wsGame As Worksheet
    
    On Error Resume Next
    Set chartWs = ThisWorkbook.Sheets("Статистика")
    On Error GoTo 0
    
    If chartWs Is Nothing Then
        Set chartWs = ThisWorkbook.Sheets.Add
        chartWs.Name = "Статистика"
    Else
        chartWs.Cells.Clear
    End If
    
    Set wsGame = ThisWorkbook.Sheets("Игра")
    
    ' Title
    chartWs.Range("A1").Value = "СТАТИСТИКА И ГРАФИКИ"
    chartWs.Range("A1").Font.Size = 16
    chartWs.Range("A1").Font.Bold = True
    
    ' Summary statistics
    chartWs.Range("A3").Value = "ОБЩАЯ СТАТИСТИКА"
    chartWs.Range("A3").Font.Bold = True
    
    chartWs.Range("A4").Value = "Метрика"
    chartWs.Range("B4").Value = "Значение"
    chartWs.Range("A4:B4").Interior.Color = RGB(0, 102, 204)
    chartWs.Range("A4:B4").Font.Color = RGB(255, 255, 255)
    chartWs.Range("A4:B4").Font.Bold = True
    
    ' Link to game data
    chartWs.Range("A5").Value = "Баланс"
    chartWs.Range("B5").Formula = "='Игра'!B3"
    chartWs.Range("B5").NumberFormat = "#,##0.00"
    
    chartWs.Range("A6").Value = "Начальный баланс"
    chartWs.Range("B6").Formula = "='Игра'!B1"
    chartWs.Range("B6").NumberFormat = "#,##0.00"
    
    chartWs.Range("A7").Value = "Всего ставок"
    chartWs.Range("B7").Formula = "='Игра'!B4"
    chartWs.Range("B7").NumberFormat = "#,##0.00"
    
    chartWs.Range("A8").Value = "Всего выигрышей"
    chartWs.Range("B8").Formula = "='Игра'!B5"
    chartWs.Range("B8").NumberFormat = "#,##0.00"
    
    chartWs.Range("A9").Value = "Чистый результат"
    chartWs.Range("B9").Formula = "='Игра'!B6"
    chartWs.Range("B9").NumberFormat = "#,##0.00"
    
    chartWs.Range("A10").Value = "Всего вращений"
    chartWs.Range("B10").Formula = "='Игра'!B7"
    
    ' Calculated statistics
    chartWs.Range("A12").Value = "РАССЧИТАННЫЕ ПОКАЗАТЕЛИ"
    chartWs.Range("A12").Font.Bold = True
    
    chartWs.Range("A13").Value = "Средняя ставка"
    chartWs.Range("B13").Formula = "=IF(B10=0,0,B7/B10)"
    chartWs.Range("B13").NumberFormat = "#,##0.00"
    
    chartWs.Range("A14").Value = "Коэффициент выигрыша"
    chartWs.Range("B14").Formula = "=IF(B7=0,0,B8/B7)"
    chartWs.Range("B14").NumberFormat = "0.00%"
    
    chartWs.Range("A15").Value = "ROI (Возврат инвестиций)"
    chartWs.Range("B15").Formula = "=IF(B7=0,0,(B8-B7)/B7)"
    chartWs.Range("B15").NumberFormat = "0.00%"
    
    ' Color code for results
    With chartWs.Range("B9")
        .Interior.Color = RGB(200, 255, 200)
    End With
    
    Call CreateWinDistributionChart(chartWs)
    Call CreateBalanceChart(chartWs)
    
    MsgBox "Лист статистики создан!", vbInformation
End Sub

' Create win/loss distribution chart
Sub CreateWinDistributionChart(ws As Worksheet)
    Dim chart As Object
    Dim dataRange As Range
    
    ' Summary data for chart
    ws.Range("A20").Value = "РАСПРЕДЕЛЕНИЕ"
    ws.Range("A20").Font.Bold = True
    
    ws.Range("A21").Value = "Выигрыши"
    ws.Range("B21").Value = "Проигрыши"
    
    ' Count wins and losses from history
    ws.Range("A22").Formula = "=COUNTIF('История'!G:G,"">0"")"
    ws.Range("B22").Formula = "=COUNTIF('История'!G:G,""<=0"")"
    
    ' Create pie chart
    On Error Resume Next
    Set chart = ws.Shapes.AddChart.Chart
    With chart
        .ChartType = xlPie
        .SetSourceData ws.Range("A21:B22")
        .HasLegend = True
        .ChartTitle.Text = "Выигрыши vs Проигрыши"
    End With
    On Error GoTo 0
End Sub

' Create balance history chart
Sub CreateBalanceChart(ws As Worksheet)
    ' This would require more complex setup
    ' Advanced: connect to История sheet and create line chart
    ' Showing balance changes over time
    
    ws.Range("A25").Value = "Для полного графика баланса:"
    ws.Range("A26").Value = "1. Убедитесь что на листе История есть данные"
    ws.Range("A27").Value = "2. Выберите данные и вставьте диаграмму (Insert → Chart)"
end Sub

' Refresh statistics
Sub RefreshStatistics()
    On Error Resume Next
    CreateStatisticsSheet
    On Error GoTo 0
End Sub
```

### Создание кнопки для статистики

В `SpreadsheetLayout.bas` добавьте в `SetupGameSheet()`:

```vba
' Add Statistics button
Call CreateButton(ws, "L11", "M12", "СТАТИСТИКА", "BtnShowStatistics", RGB(0, 102, 204), 11)
```

### Обработчик кнопки в `GameInterface.bas`:

```vba
Sub BtnShowStatistics_Click()
    Call CreateStatisticsSheet
    ThisWorkbook.Sheets("Статистика").Activate
End Sub
```

---

## 🎨 Визуализация колеса рулетки

### Создание модуля WheelVisualization.bas

```vba
' ======================================
' WHEEL VISUALIZATION MODULE
' ======================================
' Создание визуальной представления колеса рулетки

Option Explicit

' Create wheel visualization sheet
Sub CreateWheelSheet()
    Dim ws As Worksheet
    Dim i As Integer
    Dim angle As Double
    Dim x As Double, y As Double
    Dim number As Integer
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Колесо")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Колесо"
    Else
        ws.Cells.Clear
    End If
    
    ws.Range("A1").Value = "КОЛЕСО РУЛЕТКИ"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    ' Create simple table representation
    ws.Range("A3").Value = "ЧИСЛА НА КОЛЕСЕ"
    ws.Range("A3").Font.Bold = True
    
    ' Display all numbers with colors
    Dim row As Integer
    row = 4
    
    ' Зеро (зеленое)
    ws.Cells(row, 1).Value = 0
    ws.Cells(row, 1).Interior.Color = RGB(0, 176, 0)
    ws.Cells(row, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Value = "Зеленое"
    row = row + 1
    
    ' Красные числа
    ws.Cells(row, 1).Value = "КРАСНЫЕ"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    Dim redNumbers As Variant
    redNumbers = Array(1, 3, 5, 7, 9, 12, 14, 16, 18, 19, 21, 23, 25, 27, 30, 32, 34, 36)
    
    For i = LBound(redNumbers) To UBound(redNumbers)
        ws.Cells(row, 1).Value = redNumbers(i)
        ws.Cells(row, 1).Interior.Color = RGB(255, 0, 0)
        ws.Cells(row, 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(row, 1).Font.Bold = True
        row = row + 1
    Next i
    
    ' Черные числа
    ws.Cells(row, 1).Value = "ЧЕРНЫЕ"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    Dim blackNumbers As Variant
    blackNumbers = Array(2, 4, 6, 8, 10, 11, 13, 15, 17, 20, 22, 24, 26, 28, 29, 31, 33, 35)
    
    For i = LBound(blackNumbers) To UBound(blackNumbers)
        ws.Cells(row, 1).Value = blackNumbers(i)
        ws.Cells(row, 1).Interior.Color = RGB(0, 0, 0)
        ws.Cells(row, 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(row, 1).Font.Bold = True
        row = row + 1
    Next i
    
    ' Statistics section
    ws.Range("D3").Value = "ПОСЛЕДНИЙ РЕЗУЛЬТАТ"
    ws.Range("D3").Font.Bold = True
    
    ws.Range("D4").Value = "Число:"
    ws.Range("E4").Formula = "='Игра'!D3"
    ws.Range("E4").Interior.Color = RGB(255, 255, 200)
    
    ws.Range("D5").Value = "Цвет:"
    ws.Range("E5").Formula = "=IF('Игра'!D3=0,""Зеленое"",IF(OR('Игра'!D3=1,'Игра'!D3=3,'Игра'!D3=5),""Красное"",""Черное""))"
    
    ws.Range("D6").Value = "Время:"
    ws.Range("E6").Formula = "=NOW()"
    ws.Range("E6").NumberFormat = "hh:mm:ss"
    
    MsgBox "Лист с колесом рулетки создан!", vbInformation
End Sub

' Show last spin on wheel
Sub HighlightLastSpin()
    Dim ws As Worksheet
    Dim lastNumber As Integer
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Колесо")
    If ws Is Nothing Then
        Call CreateWheelSheet
        Exit Sub
    End If
    
    lastNumber = ThisWorkbook.Sheets("Игра").Range("D3").Value
    
    ' Find and highlight the number
    Dim found As Boolean
    Dim i As Integer
    
    For i = 4 To ws.UsedRange.Rows.Count
        If ws.Cells(i, 1).Value = lastNumber Then
            ws.Cells(i, 1).Font.Size = 14
            ws.Cells(i, 1).Interior.ColorIndex = 3 ' Yellow
            found = True
        Else
            ws.Cells(i, 1).Font.Size = 11
        End If
    Next i
    
    If found Then
        ws.Activate
    End If
End Sub
```

### Кнопка для визуализации

В `GameInterface.bas`:

```vba
Sub BtnShowWheel_Click()
    Call CreateWheelSheet
    Call HighlightLastSpin
End Sub
```

---

## 🤖 Автоматические стратегии

### Создание модуля AutoStrategies.bas

```vba
' ======================================
' AUTOMATIC STRATEGIES MODULE
' ======================================
' Автоматизированные стратегии игры

Option Explicit

' Martingale Strategy - doubles bet after loss
Sub RunMartingaleStrategy(initialBet As Currency, maxSpins As Integer)
    Dim currentBet As Currency
    Dim i As Integer
    Dim winnings As Currency
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Игра")
    currentBet = initialBet
    
    MsgBox "Запуск Мартингейл стратегии" & vbCrLf & _
           "Начальная ставка: " & initialBet & vbCrLf & _
           "Максимум вращений: " & maxSpins, vbInformation
    
    For i = 1 To maxSpins
        ' Update bet amount
        ws.Range("D3").Value = currentBet
        Application.Wait Now + TimeValue("0:00:01")
        
        ' Place bet on red
        winnings = PlaceBet(currentBet, "RED", "")
        
        ' If lost, double bet; if won, reset
        If winnings > 0 Then
            currentBet = initialBet
        Else
            currentBet = currentBet * 2
        End If
        
        ' Check balance
        If gameState.CurrentBalance <= 0 Then
            MsgBox "Игра окончена! Баланс = 0", vbCritical
            Exit For
        End If
        
        DoEvents
    Next i
    
    MsgBox "Мартингейл стратегия завершена!" & vbCrLf & _
           "Финальный баланс: " & FormatCurrency(gameState.CurrentBalance), vbInformation
End Sub

' D'Alembert Strategy - increases bet by 1 after loss
Sub RunDAlembert Strategy(initialBet As Currency, maxSpins As Integer)
    Dim currentBet As Currency
    Dim i As Integer
    Dim winnings As Currency
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Игра")
    currentBet = initialBet
    
    MsgBox "Запуск стратегии Д'Аламбер" & vbCrLf & _
           "Начальная ставка: " & initialBet & vbCrLf & _
           "Максимум вращений: " & maxSpins, vbInformation
    
    For i = 1 To maxSpins
        ws.Range("D3").Value = currentBet
        Application.Wait Now + TimeValue("0:00:01")
        
        winnings = PlaceBet(currentBet, "BLACK", "")
        
        ' Increase by 1 if lost, decrease by 1 if won
        If winnings > 0 Then
            currentBet = Application.Max(initialBet, currentBet - 1)
        Else
            currentBet = currentBet + 1
        End If
        
        If gameState.CurrentBalance <= 0 Then
            MsgBox "Игра окончена!", vbCritical
            Exit For
        End If
        
        DoEvents
    Next i
    
    MsgBox "Д'Аламбер стратегия завершена!" & vbCrLf & _
           "Финальный баланс: " & FormatCurrency(gameState.CurrentBalance), vbInformation
End Sub

' Constant Bet Strategy - same bet every time
Sub RunConstantBetStrategy(bet As Currency, spins As Integer)
    Dim i As Integer
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Игра")
    
    MsgBox "Запуск стратегии постоянной ставки" & vbCrLf & _
           "Ставка: " & bet & vbCrLf & _
           "Спины: " & spins, vbInformation
    
    For i = 1 To spins
        ws.Range("D3").Value = bet
        Application.Wait Now + TimeValue("0:00:01")
        
        ' Random bet type
        Select Case Int(Rnd() * 3)
            Case 0: PlaceBet bet, "RED", ""
            Case 1: PlaceBet bet, "BLACK", ""
            Case 2: PlaceBet bet, "EVEN", ""
        End Select
        
        If gameState.CurrentBalance <= 0 Then
            Exit For
        End If
        
        DoEvents
    Next i
    
    MsgBox "Стратегия завершена!" & vbCrLf & _
           "Финальный баланс: " & FormatCurrency(gameState.CurrentBalance), vbInformation
End Sub
```

### Кнопки для автостратегий в `GameInterface.bas`:

```vba
Sub BtnAutoMartingale_Click()
    Dim spins As Integer
    spins = InputBox("Сколько вращений? (максимум 100):", "Мартингейл", "10")
    If spins > 0 And spins <= 100 Then
        Call RunMartingaleStrategy(GetBetAmount(), spins)
    End If
End Sub

Sub BtnAutoDAlembert_Click()
    Dim spins As Integer
    spins = InputBox("Сколько вращений?:", "Д'Аламбер", "10")
    If spins > 0 And spins <= 100 Then
        Call RunDAlembergStrategy(GetBetAmount(), spins)
    End If
End Sub
```

---

## 💾 Сохранение в БД

### Создание модуля DatabaseExport.bas

```vba
' ======================================
' DATABASE EXPORT MODULE
' ======================================
' Экспорт результатов в БД и файлы

Option Explicit

' Export history to CSV file
Sub ExportHistoryToCSV()
    Dim csvPath As String
    Dim csvFile As Integer
    Dim i As Integer
    Dim record As BetRecord
    Dim timestamp As String
    
    csvPath = ThisWorkbook.Path & "\" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & "_history.csv"
    csvFile = FreeFile
    
    On Error Resume Next
    Open csvPath For Output As csvFile
    
    ' Write header
    Print #csvFile, "№,Время,Тип Ставки,Выпало Число,Цвет,Размер Ставки,Выигрыш"
    
    ' Write data
    For i = 1 To betHistory.Count
        record = betHistory(i)
        Print #csvFile, record.BetNumber & "," & _
                        record.BetTime & "," & _
                        record.BetType & "," & _
                        record.SpinNumber & "," & _
                        record.ResultColor & "," & _
                        record.BetAmount & "," & _
                        record.Winnings
    Next i
    
    Close csvFile
    
    MsgBox "История экспортирована в CSV!" & vbCrLf & _
           "Файл: " & csvPath, vbInformation
End Sub

' Export statistics to text file
Sub ExportStatistics()
    Dim txtPath As String
    Dim txtFile As Integer
    Dim netResult As Currency
    
    txtPath = ThisWorkbook.Path & "\" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & "_statistics.txt"
    txtFile = FreeFile
    
    netResult = gameState.CurrentBalance - gameState.StartingBalance
    
    Open txtPath For Output As txtFile
    
    Print #txtFile, "СТАТИСТИКА ИГРЫ В РУЛЕТКУ"
    Print #txtFile, "=" & String(40, "=")
    Print #txtFile, Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #txtFile, ""
    Print #txtFile, "ФИНАНСОВЫЕ ПОКАЗАТЕЛИ:"
    Print #txtFile, "  Начальный баланс: " & FormatCurrency(gameState.StartingBalance)
    Print #txtFile, "  Финальный баланс: " & FormatCurrency(gameState.CurrentBalance)
    Print #txtFile, "  Чистый результат: " & FormatCurrency(netResult)
    Print #txtFile, ""
    Print #txtFile, "СТАТИСТИКА ИГРЫ:"
    Print #txtFile, "  Всего вращений: " & gameState.TotalSpins
    Print #txtFile, "  Всего ставок: " & FormatCurrency(gameState.TotalBets)
    Print #txtFile, "  Всего выигрышей: " & FormatCurrency(gameState.TotalWinnings)
    Print #txtFile, "  Средняя ставка: " & FormatCurrency(IIf(gameState.TotalSpins = 0, 0, gameState.TotalBets / gameState.TotalSpins))
    Print #txtFile, ""
    Print #txtFile, "КОЭФФИЦИЕНТЫ:"
    Print #txtFile, "  Коэффициент выигрыша: " & Format(IIf(gameState.TotalBets = 0, 0, gameState.TotalWinnings / gameState.TotalBets), "0.00%")
    Print #txtFile, "  ROI: " & Format(IIf(gameState.TotalBets = 0, 0, netResult / gameState.StartingBalance), "0.00%")
    
    Close txtFile
    
    MsgBox "Статистика экспортирована!" & vbCrLf & _
           "Файл: " & txtPath, vbInformation
End Sub

' Create database (simple SQLite-like structure in Excel)
Sub CreateGameDatabase()
    Dim dbWs As Worksheet
    Dim i As Integer
    Dim record As BetRecord
    
    On Error Resume Next
    Set dbWs = ThisWorkbook.Sheets("БД")
    If dbWs Is Nothing Then
        Set dbWs = ThisWorkbook.Sheets.Add
        dbWs.Name = "БД"
    Else
        dbWs.Cells.Clear
    End If
    
    ' Create table
    dbWs.Range("A1").Value = "ID"
    dbWs.Range("B1").Value = "Время"
    dbWs.Range("C1").Value = "Тип"
    dbWs.Range("D1").Value = "Число"
    dbWs.Range("E1").Value = "Цвет"
    dbWs.Range("F1").Value = "Ставка"
    dbWs.Range("G1").Value = "Результат"
    dbWs.Range("H1").Value = "Баланс"
    
    ' Format header
    dbWs.Range("A1:H1").Interior.Color = RGB(0, 102, 204)
    dbWs.Range("A1:H1").Font.Color = RGB(255, 255, 255)
    dbWs.Range("A1:H1").Font.Bold = True
    
    ' Insert data
    For i = 1 To betHistory.Count
        record = betHistory(i)
        dbWs.Cells(i + 1, 1).Value = record.BetNumber
        dbWs.Cells(i + 1, 2).Value = record.BetTime
        dbWs.Cells(i + 1, 3).Value = record.BetType
        dbWs.Cells(i + 1, 4).Value = record.SpinNumber
        dbWs.Cells(i + 1, 5).Value = record.ResultColor
        dbWs.Cells(i + 1, 6).Value = record.BetAmount
        dbWs.Cells(i + 1, 6).NumberFormat = "#,##0.00"
        dbWs.Cells(i + 1, 7).Value = record.Winnings
        dbWs.Cells(i + 1, 7).NumberFormat = "#,##0.00"
        
        ' Color code wins/losses
        If record.Winnings > 0 Then
            dbWs.Cells(i + 1, 7).Interior.Color = RGB(0, 176, 0)
        Else
            dbWs.Cells(i + 1, 7).Interior.Color = RGB(255, 0, 0)
        End If
    Next i
    
    ' Auto-fit columns
    dbWs.Columns.AutoFit
    
    MsgBox "База данных создана/обновлена!", vbInformation
End Sub

' Backup game data
Sub BackupGameData()
    Dim backupPath As String
    Dim backupName As String
    
    backupName = Format(Now, "Roulette_Backup_yyyy-mm-dd_hh-mm-ss") & ".xlsm"
    backupPath = ThisWorkbook.Path & "\Backups\" & backupName
    
    ' Create Backups folder if not exists
    On Error Resume Next
    CreateFolder ThisWorkbook.Path & "\Backups"
    On Error GoTo 0
    
    ' Copy current file
    On Error Resume Next
    ThisWorkbook.SaveCopyAs backupPath
    On Error GoTo 0
    
    MsgBox "Резервная копия создана!" & vbCrLf & _
           backupPath, vbInformation
End Sub

' Helper function to create folder
Sub CreateFolder(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub
```

### Кнопки для экспорта в `GameInterface.bas`:

```vba
Sub BtnExportCSV_Click()
    Call ExportHistoryToCSV
End Sub

Sub BtnExportStats_Click()
    Call ExportStatistics
End Sub

Sub BtnCreateDB_Click()
    Call CreateGameDatabase
End Sub

Sub BtnBackup_Click()
    Call BackupGameData
End Sub
```

---

## 📝 Файл с изменениями (CHANGES.md)

Для отслеживания всех изменений создайте файл `CHANGES.md`:

```markdown
# История изменений - Рулетка VBA

## Версия 1.1 - 2026-07-19 (Расширенная версия)

### ✨ Добавлено

- [x] **Звуковые эффекты**
  - Модуль: `SoundEffects.bas`
  - Требуется файл: `WheelOfRulete.wav`
  - Звуки вращения, выигрыша, проигрыша

- [x] **Графики статистики**
  - Модуль: `Statistics.bas`
  - Новый лист: "Статистика"
  - Диаграммы выигрышей/проигрышей
  - Рассчитанные KPI

- [x] **Визуализация колеса**
  - Модуль: `WheelVisualization.bas`
  - Новый лист: "Колесо"
  - Визуальное отображение чисел
  - Подсветка последнего результата

- [x] **Автоматические стратегии**
  - Модуль: `AutoStrategies.bas`
  - Мартингейл
  - Д'Аламбер
  - Постоянная ставка

- [x] **Сохранение в БД**
  - Модуль: `DatabaseExport.bas`
  - Экспорт в CSV
  - Экспорт статистики
  - Создание БД на отдельном листе
  - Резервные копии

### 🐛 Исправлено

- Улучшена обработка ошибок
- Оптимизирована производительность
- Исправлены небольшие баги UI

### 📦 Требования

- Microsoft Excel 2010+
- WheelOfRulete.wav в папке с игрой (опционально)
- ~50 МБ свободной памяти

---

## Версия 1.0 - 2026-07-18 (Базовая версия)

### ✨ Функции

- Основная механика рулетки
- 10+ типов ставок
- История ставок
- 3 базовые стратегии
- Визуальный интерфейс
```

---

## 🔄 Процесс добавления расширений

### Шаг за шагом:

1. **Загрузите/скопируйте** новый модуль (SoundEffects, Statistics, и т.д.)
2. **Откройте VBA** (Alt + F11)
3. **Insert → Module**
4. **Скопируйте код** из соответствующего раздела выше
5. **Добавьте кнопки** в `SpreadsheetLayout.bas`
6. **Запустите SetupGame** еще раз
7. **Тестируйте** новые функции

---

## ✅ Чек-лист расширений

- [ ] Звуковые эффекты (`SoundEffects.bas`)
- [ ] Графики (`Statistics.bas`)
- [ ] Колесо (`WheelVisualization.bas`)
- [ ] Автостратегии (`AutoStrategies.bas`)
- [ ] Экспорт (`DatabaseExport.bas`)
- [ ] Файл `WheelOfRulete.wav` в папке
- [ ] Все кнопки добавлены
- [ ] SetupGame запущена

---

**Версия**: 1.1  
**Дата**: 2026-07-19  
**Автор**: asers2022
