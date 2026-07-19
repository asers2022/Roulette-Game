# 💻 Примеры кода и практические решения

Коллекция готовых кодов и примеров для расширения функционала Рулетки VBA.

---

## 📋 Содержание

1. [Базовые примеры](#базовые-примеры)
2. [Расширенные функции](#расширенные-функции)
3. [Интеграция с другими приложениями](#интеграция-с-другими-приложениями)
4. [Оптимизация производительности](#оптимизация-производительности)
5. [Решение реальных задач](#решение-реальных-задач)

---

## 🎯 Базовые примеры

### 1️⃣ Создание пользовательской формы для ввода

```vba
' Создать новую UserForm в VBA (Insert → UserForm)
' Добавить элементы управления:
' - Label: "Введите начальный бала��с"
' - TextBox: txBalance
' - CommandButton: cmdOK (Caption: "OK")
' - CommandButton: cmdCancel (Caption: "Отмена")

' Код в UserForm:
Private Sub cmdOK_Click()
    Dim balance As Currency
    
    If Not IsNumeric(txBalance.Value) Then
        MsgBox "Пожалуйста, введите число!", vbCritical
        Exit Sub
    End If
    
    balance = CDbl(txBalance.Value)
    
    If balance <= 0 Then
        MsgBox "Баланс должен быть больше нуля!", vbCritical
        Exit Sub
    End If
    
    ' Передать значение в главный модуль
    ThisWorkbook.Sheets("Игра").Range("B1").Value = balance
    InitializeGame balance
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function IsNumeric(val As Variant) As Boolean
    IsNumeric = Not IsError(CDbl(val))
End Function

' Вызов формы:
Sub ShowBalanceForm()
    UserFormBalance.Show
End Sub
```

---

### 2️⃣ Функция для проверки лицензии/версии

```vba
' Добавить в RouletteGame.bas
Function CheckVersion() As Boolean
    Const VERSION As String = "1.1"
    Const MIN_EXCEL_VERSION As Integer = 14 ' Excel 2010
    
    ' Проверка версии Excel
    If Application.Version < MIN_EXCEL_VERSION Then
        MsgBox "Требуется Excel 2010 или новее!", vbCritical
        CheckVersion = False
        Exit Function
    End If
    
    ' Проверка наличия всех модулей
    Dim project As Object
    Set project = ThisWorkbook.VBProject
    
    Dim requiredModules As Variant
    requiredModules = Array("RouletteGame", "GameInterface", "SpreadsheetLayout")
    
    Dim i As Integer
    For i = LBound(requiredModules) To UBound(requiredModules)
        If project.VBComponents(requiredModules(i)) Is Nothing Then
            MsgBox "Отсутствует модуль: " & requiredModules(i), vbCritical
            CheckVersion = False
            Exit Function
        End If
    Next i
    
    CheckVersion = True
End Function

' Вызывать в SetupGame:
Sub SetupGame()
    If Not CheckVersion Then Exit Sub
    Call CreateGameLayout
    Call InitializeGame(10000)
End Sub
```

---

### 3️⃣ Логирование всех действий

```vba
' Создать новый модуль Logger.bas
Option Explicit

Dim logFile As String

Sub InitializeLogger()
    logFile = ThisWorkbook.Path & "\game_log.txt"
    
    ' Создать файл если не существует
    If Dir(logFile) = "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateTextFile logFile
    End If
    
    LogEntry "====== ИГРА ЗАПУЩЕНА ======"
    LogEntry "Версия: 1.1"
    LogEntry "Время: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
End Sub

Sub LogEntry(message As String)
    On Error Resume Next
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFile For Append As fileNum
    Print #fileNum, Format(Now, "hh:mm:ss") & " - " & message
    Close fileNum
    
    On Error GoTo 0
End Sub

Sub LogBet(betType As String, amount As Currency, result As Integer, winnings As Currency)
    Dim message As String
    message = "BET: " & betType & " | Amount: " & amount & _
              " | Result: " & result & " | Winnings: " & winnings
    LogEntry message
End Sub

Sub LogError(errNum As Integer, errDesc As String)
    Dim message As String
    message = "ERROR #" & errNum & ": " & errDesc
    LogEntry message
End Sub

' Использование в RouletteGame.bas:
' После PlaceBet:
Call LogBet(betType, betAmount, gameState.SpinResult, gameState.LastWinnings)
```

---

## 🚀 Расширенные функции

### 1️⃣ Система уровней сложности

```vba
' Добавить в RouletteGame.bas

' Difficulty levels
Enum DifficultyLevel
    Easy = 1        ' Большой баланс, низкая ставка
    Normal = 2      ' Средний баланс, средняя ставка
    Hard = 3        ' Маленький баланс, высокая ставка
End Enum

Public currentDifficulty As DifficultyLevel

Sub SetDifficulty(difficulty As DifficultyLevel)
    currentDifficulty = difficulty
    
    Select Case difficulty
        Case Easy
            gameState.CurrentBalance = 100000
            MsgBox "Режим: ЛЕГКИЙ" & vbCrLf & "Баланс: 100,000", vbInformation
            
        Case Normal
            gameState.CurrentBalance = 10000
            MsgBox "Режим: НОРМАЛЬНЫЙ" & vbCrLf & "Баланс: 10,000", vbInformation
            
        Case Hard
            gameState.CurrentBalance = 1000
            MsgBox "Режим: СЛОЖНЫЙ" & vbCrLf & "Баланс: 1,000", vbInformation
    End Select
    
    UpdateGameDisplay
End Sub
```

---

### 2️⃣ Достижения и бейджи

```vba
' Новый модуль Achievements.bas
Option Explicit

' Achievement tracking
Type Achievement
    Name As String
    Description As String
    Unlocked As Boolean
    UnlockedDate As Date
End Type

Public achievements As Collection

Sub InitializeAchievements()
    Set achievements = New Collection
    
    ' Добавить достижения
    AddAchievement "Первая ставка", "Сделать первую ставку"
    AddAchievement "Миллионер", "Достичь баланса 1,000,000"
    AddAchievement "Без убытков", "Выиграть 10 ставок подряд"
    AddAchievement "Рискующий", "Поставить 100 за раз"
    AddAchievement "Везунчик", "Выиграть на 35:1"
    AddAchievement "Стратег", "Использовать все стратегии"
End Sub

Sub AddAchievement(name As String, description As String)
    Dim ach As Achievement
    ach.Name = name
    ach.Description = description
    ach.Unlocked = False
    achievements.Add ach
End Sub

Sub CheckAchievements()
    ' Проверка условий достижений
    
    If gameState.TotalSpins = 1 Then
        UnlockAchievement "Первая ставка"
    End If
    
    If gameState.CurrentBalance >= 1000000 Then
        UnlockAchievement "Миллионер"
    End If
    
    If gameState.LastWinnings > 3500 Then ' 35:1 выигрыш
        UnlockAchievement "Везунчик"
    End If
End Sub

Sub UnlockAchievement(achName As String)
    Dim i As Integer
    
    For i = 1 To achievements.Count
        If achievements(i).Name = achName Then
            If Not achievements(i).Unlocked Then
                achievements(i).Unlocked = True
                achievements(i).UnlockedDate = Now
                
                MsgBox "🏆 ДОСТИЖЕНИЕ РАЗБЛОКИРОВАНО!" & vbCrLf & _
                       achName & vbCrLf & _
                       "Описание: " & achievements(i).Description, vbInformation
            End If
            Exit For
        End If
    Next i
End Sub

Sub ShowAchievements()
    Dim message As String
    Dim i As Integer
    
    message = "═══════════════════════════" & vbCrLf
    message = message & "ВАШИ ДОСТИЖЕНИЯ" & vbCrLf
    message = message & "═══════════════════════════" & vbCrLf & vbCrLf
    
    For i = 1 To achievements.Count
        If achievements(i).Unlocked Then
            message = message & "✓ " & achievements(i).Name & vbCrLf
            message = message & "  (" & achievements(i).Description & ")" & vbCrLf & vbCrLf
        Else
            message = message & "✗ " & achievements(i).Name & vbCrLf
            message = message & "  (Заблокировано)" & vbCrLf & vbCrLf
        End If
    Next i
    
    MsgBox message, vbInformation
End Sub
```

---

### 3️⃣ Система рейтинга и лидербордов

```vba
' Новый модуль Leaderboard.bas
Option Explicit

Type PlayerScore
    PlayerName As String
    Score As Currency
    Spins As Integer
    WinRate As Double
    PlayDate As Date
End Type

Public leaderboard As Collection
Const MAX_SCORES = 10

Sub SaveScore(playerName As String)
    If leaderboard Is Nothing Then
        Set leaderboard = New Collection
    End If
    
    Dim score As PlayerScore
    score.PlayerName = playerName
    score.Score = gameState.CurrentBalance - gameState.StartingBalance
    score.Spins = gameState.TotalSpins
    score.WinRate = IIf(gameState.TotalBets = 0, 0, gameState.TotalWinnings / gameState.TotalBets)
    score.PlayDate = Now
    
    ' Найти позицию для вставки
    Dim i As Integer
    For i = 1 To leaderboard.Count
        If score.Score > leaderboard(i).Score Then
            leaderboard.Add score, , i
            Exit For
        End If
    Next i
    
    ' Если не найдена позиция, добавить в конец
    If i > leaderboard.Count Then
        leaderboard.Add score
    End If
    
    ' Сохранить только топ 10
    While leaderboard.Count > MAX_SCORES
        leaderboard.Remove leaderboard.Count
    Wend
    
    ' Экспортировать лидербоард
    ExportLeaderboard
    
    MsgBox "Ваш результат сохранен!" & vbCrLf & _
           "Позиция: " & i & " из " & leaderboard.Count, vbInformation
End Sub

Sub ShowLeaderboard()
    Dim message As String
    Dim i As Integer
    
    message = "═════════════════════════════════════" & vbCrLf
    message = message & "ТОП 10 ЛИДЕРОВ" & vbCrLf
    message = message & "═════════════════════════════════════" & vbCrLf & vbCrLf
    
    For i = 1 To leaderboard.Count
        message = message & i & ". " & leaderboard(i).PlayerName & vbCrLf
        message = message & "   Счет: " & FormatCurrency(leaderboard(i).Score) & vbCrLf
        message = message & "   Спины: " & leaderboard(i).Spins & vbCrLf
        message = message & "   Win Rate: " & Format(leaderboard(i).WinRate, "0.00%") & vbCrLf & vbCrLf
    Next i
    
    MsgBox message, vbInformation
End Sub

Sub ExportLeaderboard()
    Dim ws As Worksheet
    Dim i As Integer
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Лидербоард")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Лидербоард"
    Else
        ws.Cells.Clear
    End If
    
    ' Заголовки
    ws.Range("A1").Value = "Место"
    ws.Range("B1").Value = "Игрок"
    ws.Range("C1").Value = "Счет"
    ws.Range("D1").Value = "Спины"
    ws.Range("E1").Value = "Win Rate"
    ws.Range("F1").Value = "Дата"
    
    ws.Range("A1:F1").Interior.Color = RGB(0, 102, 204)
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Font.Color = RGB(255, 255, 255)
    
    ' Данные
    For i = 1 To leaderboard.Count
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = leaderboard(i).PlayerName
        ws.Cells(i + 1, 3).Value = leaderboard(i).Score
        ws.Cells(i + 1, 3).NumberFormat = "#,##0.00"
        ws.Cells(i + 1, 4).Value = leaderboard(i).Spins
        ws.Cells(i + 1, 5).Value = leaderboard(i).WinRate
        ws.Cells(i + 1, 5).NumberFormat = "0.00%"
        ws.Cells(i + 1, 6).Value = Format(leaderboard(i).PlayDate, "yyyy-mm-dd hh:mm")
    Next i
    
    ws.Columns.AutoFit
End Sub
```

---

## 🔗 Интеграция с другими приложениями

### 1️⃣ Отправка результатов по Email

```vba
' Требует Outlook
Sub SendGameResultsViaEmail()
    Dim outlook As Object
    Dim mailItem As Object
    Dim recipient As String
    Dim subject As String
    Dim body As String
    
    recipient = InputBox("Введите email адрес:", "Отправить результаты")
    
    If recipient = "" Then Exit Sub
    
    ' Создать объект Outlook
    On Error Resume Next
    Set outlook = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    If outlook Is Nothing Then
        MsgBox "Outlook не установлен!", vbCritical
        Exit Sub
    End If
    
    ' Создать письмо
    Set mailItem = outlook.CreateItem(0)
    
    subject = "Результаты игры в Рулетку - " & Format(Now, "yyyy-mm-dd hh:mm")
    
    body = "Привет!" & vbCrLf & vbCrLf & _
           "Мои результаты в игре Рулетка:" & vbCrLf & vbCrLf & _
           "Начальный баланс: " & FormatCurrency(gameState.StartingBalance) & vbCrLf & _
           "Финальный баланс: " & FormatCurrency(gameState.CurrentBalance) & vbCrLf & _
           "Чистый результат: " & FormatCurrency(gameState.CurrentBalance - gameState.StartingBalance) & vbCrLf & _
           "Всего ставок: " & gameState.TotalSpins & vbCrLf & _
           "Успехов!" & vbCrLf & vbCrLf & _
           "Рулетка VBA"
    
    With mailItem
        .To = recipient
        .Subject = subject
        .Body = body
        .Send
    End With
    
    MsgBox "Email отправлен!", vbInformation
End Sub
```

---

### 2️⃣ Загрузка результатов в облако (Google Sheets)

```vba
' Требует API ключ Google Sheets
Sub ExportToGoogleSheets()
    Dim i As Integer
    Dim record As BetRecord
    Dim jsonData As String
    
    ' Построить JSON данные
    jsonData = "{""data"": ["
    
    For i = 1 To betHistory.Count
        Set record = betHistory(i)
        
        If i > 1 Then jsonData = jsonData & ","
        
        jsonData = jsonData & "{" & _
                   """bet_number"":" & record.BetNumber & "," & _
                   """spin_number"":" & record.SpinNumber & "," & _
                   """bet_type"":""" & record.BetType & """," & _
                   """bet_amount"":" & record.BetAmount & "," & _
                   """winnings"":" & record.Winnings & _
                   "}"
    Next i
    
    jsonData = jsonData & "]}"
    
    ' Отправить запрос (требует метода HttpRequest)
    ' Это пример - требует дополнительной настройки
    
    MsgBox "Данные подготовлены для загрузки!", vbInformation
End Sub
```

---

## ⚡ Оптимизация производительности

### 1️⃣ Кэширование часто используемых значений

```vba
' Добавить в RouletteGame.bas

' Cache variables
Dim cachedRedNumbers As Collection
Dim cachedBlackNumbers As Collection
Dim redNumbersLoaded As Boolean

Sub InitializeCache()
    If Not redNumbersLoaded Then
        Set cachedRedNumbers = New Collection
        Set cachedBlackNumbers = New Collection
        
        ' Загрузить красные числа
        Dim redStr As String
        redStr = RED_NUMBERS
        
        ' Парсинг и кэширование
        Dim parts As Variant
        parts = Split(redStr, ",")
        
        Dim i As Integer
        For i = LBound(parts) To UBound(parts)
            cachedRedNumbers.Add CInt(parts(i))
        Next i
        
        redNumbersLoaded = True
    End If
End Sub

' Оптимизированная функция проверки цвета
Function IsRedFast(number As Integer) As Boolean
    If number = 0 Then
        IsRedFast = False
    Else
        On Error Resume Next
        cachedRedNumbers.Add number
        ' Если добавилось, значит это новое число
        ' Удалить для следующей проверки
        On Error GoTo 0
        IsRedFast = (InStr(RED_NUMBERS, CStr(number)) > 0)
    End If
End Function
```

---

### 2️⃣ Отключение обновления экрана при интенсивных операциях

```vba
' Оптимизированная функция истории
Sub UpdateHistoryDisplayOptimized()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim row As Integer
    Dim i As Integer
    Dim record As BetRecord
    
    Set ws = ThisWorkbook.Sheets("История")
    
    ws.Range("A2:G1000").ClearContents
    
    row = 2
    For i = 1 To betHistory.Count
        Set record = betHistory(i)
        
        With ws
            .Cells(row, 1).Value = record.BetNumber
            .Cells(row, 2).Value = record.BetTime
            .Cells(row, 3).Value = record.BetType
            .Cells(row, 4).Value = record.SpinNumber
            .Cells(row, 5).Value = record.ResultColor
            .Cells(row, 6).Value = record.BetAmount
            .Cells(row, 7).Value = record.Winnings
            
            If record.Winnings > 0 Then
                .Cells(row, 7).Interior.Color = RGB(0, 176, 0)
            Else
                .Cells(row, 7).Interior.Color = RGB(255, 0, 0)
            End If
        End With
        
        row = row + 1
    Next i
    
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

---

## 🎓 Решение реальных задач

### Задача 1️⃣: Защита от чрезмерной ставки

```vba
' Добавить в GameInterface.bas

Function GetBetAmountSafe() As Currency
    Dim ws As Worksheet
    Dim betValue As Variant
    Dim maxBet As Currency
    
    Set ws = ThisWorkbook.Sheets("Игра")
    
    betValue = ws.Range("D3").Value
    maxBet = Application.Min(100, gameState.CurrentBalance / 2) ' Не больше 50% баланса
    
    If betValue = "" Or betValue <= 0 Then
        MsgBox "Введите сумму ставки от 1 до " & maxBet, vbCritical
        GetBetAmountSafe = 0
        Exit Function
    End If
    
    If betValue > maxBet Then
        MsgBox "Максимальная ставка: " & FormatCurrency(maxBet), vbCritical
        GetBetAmountSafe = 0
        Exit Function
    End If
    
    GetBetAmountSafe = CDbl(betValue)
End Function
```

---

### Задача 2️⃣: Напоминание о лимите времени

```vba
' Новый модуль SessionTimer.bas
Option Explicit

Dim sessionStartTime As Date
Dim sessionLimitMinutes As Integer

Sub StartSessionTimer(limitMinutes As Integer)
    sessionStartTime = Now
    sessionLimitMinutes = limitMinutes
    MsgBox "Сессия запущена на " & limitMinutes & " минут", vbInformation
End Sub

Sub CheckSessionTime()
    Dim elapsed As Double
    Dim remaining As Integer
    
    elapsed = DateDiff("n", sessionStartTime, Now)
    remaining = sessionLimitMinutes - elapsed
    
    If remaining <= 5 And remaining > 0 Then
        MsgBox "Осталось " & remaining & " минут игры!", vbExclamation
    ElseIf remaining <= 0 Then
        MsgBox "Время сессии истекло!", vbCritical
        ' Рекомендация: остановить игру
    End If
End Sub

' Вызывать периодически
Sub MonitorSessionTime()
    On Error Resume Next
    
    If Not sessionStartTime = 0 Then
        Call CheckSessionTime
    End If
    
    ' Проверять каждую минуту
    Application.OnTime Now + TimeValue("0:01:00"), "MonitorSessionTime"
End Sub
```

---

### Задача 3️⃣: Анализ вероятности выигрыша

```vba
' Новый модуль ProbabilityAnalysis.bas
Option Explicit

Sub AnalyzeBetProbability()
    Dim message As String
    
    message = "════════════════════════════════════" & vbCrLf
    message = message & "АНАЛИЗ ВЕРОЯТНОСТИ ВЫИГРЫША" & vbCrLf
    message = message & "════════════════════════════════════" & vbCrLf & vbCrLf
    
    message = message & "1. КРАСНОЕ/ЧЕРНОЕ:" & vbCrLf
    message = message & "   Вероятность: 48.6% (18/37)" & vbCrLf
    message = message & "   Выплата: 1:1" & vbCrLf
    message = message & "   Ожидание: -2.7%" & vbCrLf & vbCrLf
    
    message = message & "2. НА ЧИСЛО:" & vbCrLf
    message = message & "   Вероятность: 2.7% (1/37)" & vbCrLf
    message = message & "   Выплата: 35:1" & vbCrLf
    message = message & "   Ожидание: -2.7%" & vbCrLf & vbCrLf
    
    message = message & "3. РЕКОМЕНДАЦИЯ:" & vbCrLf
    message = message & "   Ставьте на красное/черное для" & vbCrLf
    message = message & "   наибольших шансов на выигрыш!" & vbCrLf
    
    MsgBox message, vbInformation
End Sub

' Рассчитать математическое ожидание
Function CalculateExpectedValue(probability As Double, payout As Double) As Double
    ' EV = (вероятность выигрыша * выплата) - (вероятность проигрыша * ставка)
    Dim lossProb As Double
    lossProb = 1 - probability
    
    CalculateExpectedValue = (probability * payout) - (lossProb * 1)
End Function

Sub ShowBetAnalysis()
    Dim message As String
    Dim evRed As Double
    Dim evNumber As Double
    
    evRed = CalculateExpectedValue(0.486, 2)
    evNumber = CalculateExpectedValue(0.027, 36)
    
    message = "МАТЕМАТИЧЕСКОЕ ОЖИДАНИЕ:" & vbCrLf & vbCrLf
    message = message & "Красное/Черное: " & Format(evRed, "0.00%") & vbCrLf
    message = message & "На число: " & Format(evNumber, "0.00%") & vbCrLf & vbCrLf
    message = message & "Оба имеют отрицательное ожидание (-2.7%)" & vbCrLf
    message = message & "благодаря зеро!"
    
    MsgBox message, vbInformation
End Sub
```

---

### Задача 4️⃣: Экспорт отчета в Word

```vba
' Требует Microsoft Word
Sub ExportGameReportToWord()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim table As Object
    Dim i As Integer
    
    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If wordApp Is Nothing Then
        MsgBox "Microsoft Word не установлен!", vbCritical
        Exit Sub
    End If
    
    ' Создать новый документ
    Set wordDoc = wordApp.Documents.Add()
    
    ' Заголовок
    wordDoc.Range.InsertBefore "ОТЧЕТ О ИГРЕ В РУЛЕТКУ" & vbCrLf
    wordDoc.Range.Font.Size = 16
    wordDoc.Range.Font.Bold = True
    
    ' Добавить таблицу
    Set table = wordDoc.Tables.Add(wordDoc.Range, betHistory.Count + 1, 7)
    table.Style = "Table Grid"
    
    ' Заголовки таблицы
    With table.Rows(1).Range
        .Cells(1).Range.Text = "№"
        .Cells(2).Range.Text = "Время"
        .Cells(3).Range.Text = "Тип"
        .Cells(4).Range.Text = "Число"
        .Cells(5).Range.Text = "Ставка"
        .Cells(6).Range.Text = "Результат"
        .Cells(7).Range.Text = "Цвет"
    End With
    
    ' Данные
    For i = 1 To betHistory.Count
        With table.Rows(i + 1).Range
            .Cells(1).Range.Text = betHistory(i).BetNumber
            .Cells(2).Range.Text = betHistory(i).BetTime
            .Cells(3).Range.Text = betHistory(i).BetType
            .Cells(4).Range.Text = betHistory(i).SpinNumber
            .Cells(5).Range.Text = Format(betHistory(i).BetAmount, "#,##0.00")
            .Cells(6).Range.Text = Format(betHistory(i).Winnings, "#,##0.00")
            .Cells(7).Range.Text = betHistory(i).ResultColor
        End With
    Next i
    
    wordApp.Visible = True
    
    MsgBox "Отчет создан в Word!", vbInformation
End Sub
```

---

## 📊 Бизнес-примеры использования

### 1️⃣ Обучающий материал

```vba
' Режим обучения для новичков
Sub RunTrainingMode()
    Dim message As String
    
    message = "РЕЖИМ ОБУЧЕНИЯ" & vbCrLf & vbCrLf
    message = message & "Урок 1: Основы рулетки" & vbCrLf
    message = message & "- На колесе 37 чисел (0-36)" & vbCrLf
    message = message & "- 18 красных, 18 черных, 1 зеро" & vbCrLf & vbCrLf
    message = message & "Урок 2: Типы ставок" & vbCrLf
    message = message & "- Красное/Черное (1:1)" & vbCrLf
    message = message & "- На число (35:1)" & vbCrLf
    message = message & "- И многое другое!" & vbCrLf & vbCrLf
    message = message & "Нажмите OK для продолжения"
    
    MsgBox message, vbInformation
End Sub
```

---

### 2️⃣ Система рейтинга кзинов

```vba
' Для интеграции в казино систему
Sub CreateCasinoReport()
    Dim ws As Worksheet
    Dim row As Integer
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Отчет казино"
    
    ' Заголовок
    ws.Range("A1").Value = "ОТЧЕТ О РАБОТЕ КАЗИНО"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    
    row = 3
    
    ' Статистика по ставкам
    ws.Cells(row, 1).Value = "Всего ставок:"
    ws.Cells(row, 2).Value = betHistory.Count
    row = row + 1
    
    ws.Cells(row, 1).Value = "Всего деньгт в ставках:"
    ws.Cells(row, 2).Value = gameState.TotalBets
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "Всего выплачено:"
    ws.Cells(row, 2).Value = gameState.TotalWinnings
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "Прибыль казино:"
    ws.Cells(row, 2).Value = gameState.TotalBets - gameState.TotalWinnings
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Interior.Color = RGB(0, 176, 0)
    
    row = row + 2
    
    ' Популярные ставки
    ws.Cells(row, 1).Value = "ПОПУЛЯРНЫЕ СТАВКИ"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    ' Подсчет
    ws.Cells(row, 1).Value = "Ставка"
    ws.Cells(row, 2).Value = "Количество"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Font.Bold = True
    row = row + 1
    
    ' Вручную добавить статистику
    ws.Cells(row, 1).Value = "Красное"
    ws.Cells(row, 2).Formula = "=COUNTIF('История'!C:C,""RED"")"
End Sub
```

---

## ✅ Контрольный список реализации

### Базовые функции:
- [x] Основная механика рулетки
- [x] Типы ставок
- [x] История ставок
- [x] Статистика

### Расширенные функции:
- [ ] Звуковые эффекты
- [ ] Графики
- [ ] Визуализация колеса
- [ ] Автостратегии
- [ ] Сохранение в БД

### Опциональные:
- [ ] Достижения
- [ ] Лидербоард
- [ ] Email экспорт
- [ ] Защита от читеров
- [ ] Многопользовательский режим

---

## 🔧 Быстрый старт с примерами

```vba
' Скопируйте этот код в новый модуль для быстрого тестирования

Sub QuickStartWithExamples()
    ' 1. Инициализировать игру
    Call InitializeGame(10000)
    
    ' 2. Создать лидербоард
    Call InitializeAchievements
    
    ' 3. Запустить логирование
    Call InitializeLogger
    
    ' 4. Показать обучение
    Call RunTrainingMode
    
    MsgBox "Все системы инициализированы!", vbInformation
End Sub
```

---

**Версия**: 1.1  
**Дата**: 2026-07-19  
**Автор**: asers2022

---

## 📞 Поддержка

Если вам нужна помощь с кодом:

1. Проверьте комментарии в коде
2. Смотрите примеры на GitHub
3. Создавайте Issues для вопросов
4. Экспериментируйте и учитесь!

**Успехов в разработке!** 🚀
