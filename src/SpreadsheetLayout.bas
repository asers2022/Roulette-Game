' ======================================
' SPREADSHEET LAYOUT MODULE
' ======================================
' Создание и настройка листов с указанием расположения всех элементов

Option Explicit

' РАСПОЛОЖЕНИЕ ЭЛЕМЕНТОВ НА ЛИСТЕ "Игра":
'
' ЛЕВАЯ ЧАСТЬ (столбцы A-B) - СТАТИСТИКА И УПРАВЛЕНИЕ:
' ═════════════════════════════════════════════════════════
' Ячейка    │ Содержание
' ───────────┼────────────────────────────────────────────
' A1        │ "Начальный баланс:"
' B1        │ 10000 (редактируется пользователем)
' 
' A3        │ "Текущий баланс:"
' B3        │ [Обновляется автоматически - формула или код]
' 
' A4        │ "Всего ставок:"
' B4        │ [Сумма всех сделанных ставок]
' 
' A5        │ "Всего выигрышей:"
' B5        │ [Сумма всех выигрышей]
' 
' A6        │ "Чистый результат:"
' B6        │ [Баланс - Начальный баланс]
' 
' A7        │ "Всего вращений:"
' B7        │ [Количество спинов]
' 
' A9        │ "Текущая ставка:"
' B9        │ [Текущий результат последней ставки]
'
' ═════════════════════════════════════════════════════════
'
' СРЕДНЯЯ ЧАСТЬ (столбцы D-E) - УПРАВЛЕНИЕ СТАВКОЙ:
' ═════════════════════════════════════════════════════════
' D1        │ "Сумма ставки (1-100):"
' D3        │ [Текстовое поле/SpinButton для ввода суммы]
'           │ [По умолчанию: 10]
'
' ═════════════════════════════════════════════════════════
'
' ПРАВАЯ ЧАСТЬ (столбцы F-M) - КНОПКИ СТАВОК:
' ═════════════════════════════════════════════════════════
'
' РЯДЫ 1-2: УПРАВЛЕНИЕ ИГРОЙ
' ─────────────────────────────
' F1-G1     │ Кнопка "НАЧАТЬ ИГРУ" (BtnStartGame)
' H1-I1     │ Кнопка "СТАТИСТИКА" (BtnStatistics)
' J1-K1     │ Кнопка "ОЧИСТИТЬ ИСТОРИЮ" (BtnClearHistory)
' L1-M1     │ Кнопка "ВЫХОД" (BtnExit)
'
' ═════════════════════════════════════════════════════════
'
' РЯДЫ 3-4: ОСНОВНЫЕ СТАВКИ (КРАСНОЕ/ЧЕРНОЕ)
' ─────────────────────────────────────────────
' F3-G3     │ Кнопка "КРАСНОЕ" (BtnRed) - Красный фон
' H3-I3     │ Кнопка "ЧЕРНОЕ" (BtnBlack) - Черный фон, белый текст
' J3-K3     │ Кнопка "ЧЕТНОЕ" (BtnEven) - Синий фон
' L3-M3     │ Кнопка "НЕЧЕТНОЕ" (BtnOdd) - Оранжевый фон
'
' ═════════════════════════════════════════════════════════
'
' РЯДЫ 5-6: ВЫСОКИЕ/НИЗКИЕ И КОЛОННЫ
' ──────────────────────────────────
' F5-G5     │ Кнопка "HIGH (19-36)" (BtnHigh) - Светло-зеленый фон
' H5-I5     │ Кнопка "LOW (1-18)" (BtnLow) - Светло-красный фон
' J5-K5     │ Кнопка "КОЛОННА 1" (BtnColumn1)
' L5-M5     │ Кнопка "КОЛОННА 2" (BtnColumn2)
'
' ═════════════════════════════════════════════════════════
'
' РЯДЫ 7-8: ДЮЖИНЫ
' ───────────────
' F7-G7     │ Кнопка "ДЮЖИНА 1 (1-12)" (BtnDozen1)
' H7-I7     │ Кнопка "ДЮЖИНА 2 (13-24)" (BtnDozen2)
' J7-K7     │ Кнопка "ДЮЖИНА 3 (25-36)" (BtnDozen3)
' L7-M7     │ Кнопка "КОЛОННА 3" (BtnColumn3)
'
' ═════════════════════════════════════════════════════════
'
' РЯДЫ 9-10: СПЕЦИАЛЬНЫЕ СТАВКИ
' ──────────────────────────────
' F9-G9     │ Кнопка "НА ЧИСЛО (0-36)" (BtnSingleNumber)
' H9-I9     │ Кнопка "АВТО СПИН" (BtnAutoSpin)
' J9-K9     │ Кнопка "СБРОС ИГРЫ" (BtnResetGame)
' L9-M9     │ Кнопка "ИСТОРИЯ" (BtnViewHistory)
'
' ═════════════════════════════════════════════════════════
'
' РЯДЫ 11-12: СТРАТЕГИИ
' ──────────────────────
' F11-G11   │ Кнопка "КЛАССИЧЕСКАЯ" (BtnStrategyClassic)
' H11-I11   │ Кнопка "МАРТИНГЕЙЛ" (BtnStrategyMartingale)
' J11-K11   │ Кнопка "Д'АЛАМБЕР" (BtnStrategyDAlembert)
'
' ═════════════════════════════════════════════════════════

' Create Game Sheets Layout
Sub CreateGameLayout()
    Dim ws As Worksheet
    Dim historyWs As Worksheet
    
    On Error Resume Next
    
    ' Create or get "Игра" sheet
    Set ws = ThisWorkbook.Sheets("Игра")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Игра"
    End If
    
    ' Create or get "История" sheet
    Set historyWs = ThisWorkbook.Sheets("История")
    If historyWs Is Nothing Then
        Set historyWs = ThisWorkbook.Sheets.Add
        historyWs.Name = "История"
    End If
    
    ' Setup Game Sheet
    Call SetupGameSheet(ws)
    
    ' Setup History Sheet
    Call SetupHistorySheet(historyWs)
    
    MsgBox "Игровой макет создан успешно!", vbInformation
End Sub

' Setup Game Sheet with all UI elements
Sub SetupGameSheet(ws As Worksheet)
    Dim btn As Object
    
    ws.Activate
    ws.Cells.Clear
    
    ' ═══════════════════════════════════════════════════════════
    ' ЛЕВАЯ ЧАСТЬ - СТАТИСТИКА (A:B)
    ' ═══════════════════════════════════════════════════════════
    
    ' Начальный баланс
    ws.Range("A1").Value = "Начальный баланс:"
    ws.Range("A1").Font.Bold = True
    ws.Range("B1").Value = 10000
    ws.Range("B1").NumberFormat = "#,##0.00"
    
    ' Текущий баланс
    ws.Range("A3").Value = "Текущий баланс:"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").Value = 10000
    ws.Range("B3").NumberFormat = "#,##0.00"
    ws.Range("B3").Interior.Color = RGB(200, 255, 200)
    
    ' Всего ставок
    ws.Range("A4").Value = "Всего ставок:"
    ws.Range("A4").Font.Bold = True
    ws.Range("B4").Value = 0
    ws.Range("B4").NumberFormat = "#,##0.00"
    
    ' Всего выигрышей
    ws.Range("A5").Value = "Всего выигрышей:"
    ws.Range("A5").Font.Bold = True
    ws.Range("B5").Value = 0
    ws.Range("B5").NumberFormat = "#,##0.00"
    
    ' Чистый результат
    ws.Range("A6").Value = "Чистый результат:"
    ws.Range("A6").Font.Bold = True
    ws.Range("B6").Value = 0
    ws.Range("B6").NumberFormat = "#,##0.00"
    
    ' Всего вращений
    ws.Range("A7").Value = "Всего вращений:"
    ws.Range("A7").Font.Bold = True
    ws.Range("B7").Value = 0
    
    ' Текущий результат
    ws.Range("A9").Value = "Последний результат:"
    ws.Range("A9").Font.Bold = True
    ws.Range("B9").Value = 0
    ws.Range("B9").NumberFormat = "#,##0.00"
    ws.Range("B9").Interior.Color = RGB(255, 200, 200)
    
    ' ═══════════════════════════════════════════════════════════
    ' СРЕДНЯЯ ЧАСТЬ - ВВОД СТАВКИ (D:E)
    ' ═══════════════════════════════════════════════════════════
    
    ws.Range("D1").Value = "Сумма ставки (1-100):"
    ws.Range("D1").Font.Bold = True
    ws.Range("D3").Value = 10
    ws.Range("D3").NumberFormat = "0"
    ws.Range("D3").Interior.Color = RGB(255, 255, 200)
    ws.Range("D3").Font.Size = 14
    ws.Range("D3").Font.Bold = True
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 25
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("L").ColumnWidth = 12
    
    ' ═══════════════════════════════════════════════════════════
    ' ПРАВАЯ ЧАСТЬ - КНОПКИ СТАВОК (F:M)
    ' ═══════════════════════════════════════════════════════════
    
    ' РЯДЫ 1-2: УПРАВЛЕНИЕ ИГРОЙ
    Call CreateButton(ws, "F1", "G2", "НАЧАТЬ ИГРУ", "BtnStartGame", RGB(0, 128, 0), 12)
    Call CreateButton(ws, "H1", "I2", "СТАТИСТИКА", "BtnStatistics", RGB(0, 102, 204), 12)
    Call CreateButton(ws, "J1", "K2", "ОЧИСТИТЬ", "BtnClearHistory", RGB(204, 102, 0), 12)
    Call CreateButton(ws, "L1", "M2", "ВЫХОД", "BtnExit", RGB(204, 0, 0), 12)
    
    ' РЯДЫ 3-4: ОСНОВНЫЕ СТАВКИ
    Call CreateButton(ws, "F3", "G4", "КРАСНОЕ", "BtnRed", RGB(255, 0, 0), 14)
    Call CreateButton(ws, "H3", "I4", "ЧЕРНОЕ", "BtnBlack", RGB(0, 0, 0), 14)
    Call CreateButton(ws, "J3", "K4", "ЧЕТНОЕ", "BtnEven", RGB(0, 102, 255), 12)
    Call CreateButton(ws, "L3", "M4", "НЕЧЕТНОЕ", "BtnOdd", RGB(255, 102, 0), 12)
    
    ' РЯДЫ 5-6: HIGH/LOW И КОЛОННЫ
    Call CreateButton(ws, "F5", "G6", "HIGH (19-36)", "BtnHigh", RGB(0, 204, 0), 12)
    Call CreateButton(ws, "H5", "I6", "LOW (1-18)", "BtnLow", RGB(255, 102, 102), 12)
    Call CreateButton(ws, "J5", "K6", "КОЛОННА 1", "BtnColumn1", RGB(153, 153, 255), 10)
    Call CreateButton(ws, "L5", "M6", "КОЛОННА 2", "BtnColumn2", RGB(153, 153, 255), 10)
    
    ' РЯДЫ 7-8: ДЮЖИНЫ
    Call CreateButton(ws, "F7", "G8", "ДЮЖИНА 1\n(1-12)", "BtnDozen1", RGB(204, 204, 255), 10)
    Call CreateButton(ws, "H7", "I8", "ДЮЖИНА 2\n(13-24)", "BtnDozen2", RGB(204, 204, 255), 10)
    Call CreateButton(ws, "J7", "K8", "ДЮЖИНА 3\n(25-36)", "BtnDozen3", RGB(204, 204, 255), 10)
    Call CreateButton(ws, "L7", "M8", "КОЛОННА 3", "BtnColumn3", RGB(153, 153, 255), 10)
    
    ' РЯДЫ 9-10: СПЕЦИАЛЬНЫЕ СТАВКИ
    Call CreateButton(ws, "F9", "G10", "НА ЧИСЛО\n(0-36)", "BtnSingleNumber", RGB(255, 200, 124), 11)
    Call CreateButton(ws, "H9", "I10", "АВТО СПИН", "BtnAutoSpin", RGB(204, 255, 204), 11)
    Call CreateButton(ws, "J9", "K10", "СБРОС ИГРЫ", "BtnResetGame", RGB(255, 204, 0), 11)
    Call CreateButton(ws, "L9", "M10", "ИСТОРИЯ", "BtnViewHistory", RGB(204, 204, 204), 11)
    
    ' РЯДЫ 11-12: СТРАТЕГИИ
    Call CreateButton(ws, "F11", "G12", "КЛАССИЧЕСКАЯ", "BtnStrategyClassic", RGB(153, 255, 204), 10)
    Call CreateButton(ws, "H11", "I12", "МАРТИНГЕЙЛ", "BtnStrategyMartingale", RGB(255, 204, 153), 10)
    Call CreateButton(ws, "J11", "K12", "Д'АЛАМБЕР", "BtnStrategyDAlembert", RGB(204, 153, 255), 10)
    
    ' Добавить легенду колеса рулетки для визуализации
    Call DrawWheelVisualization(ws)
    
End Sub

' Create Button Helper Function
Sub CreateButton(ws As Worksheet, startCell As String, endCell As String, caption As String, macroName As String, color As Long, fontSize As Integer)
    Dim btn As Object
    Dim startRange As Range
    Dim endRange As Range
    
    Set startRange = ws.Range(startCell)
    Set endRange = ws.Range(endCell)
    
    ' Set row height for better button appearance
    ws.Rows(startRange.Row).RowHeight = 30
    
    ' Create button
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                 Left:=startRange.Left, _
                                 Top:=startRange.Top, _
                                 Width:=endRange.Left + endRange.Width - startRange.Left, _
                                 Height:=endRange.Top + endRange.Height - startRange.Top)
    
    With btn.Object
        .Caption = caption
        .Font.Size = fontSize
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = color
        .Name = macroName
    End With
    
    ' Assign macro to button
    ws.Shapes(btn.Name).OnAction = macroName & "_Click"
End Sub

' Draw Wheel Visualization (numbers 0-36)
Sub DrawWheelVisualization(ws As Worksheet)
    ' Add wheel numbers for reference
    ws.Range("A12").Value = "КОЛЕСО РУЛЕТКИ:"
    ws.Range("A12").Font.Bold = True
    
    ' Display red numbers
    ws.Range("A13").Value = "Красные:"
    ws.Range("B13").Value = RED_NUMBERS
    
    ' Display black numbers
    ws.Range("A14").Value = "Черные:"
    ws.Range("B14").Value = BLACK_NUMBERS
    
    ' Display zero
    ws.Range("A15").Value = "Зеро:"
    ws.Range("B15").Value = "0 (Зеленое)"
End Sub

' Setup History Sheet
Sub SetupHistorySheet(ws As Worksheet)
    ws.Activate
    ws.Cells.Clear
    
    ' Create headers
    ws.Range("A1").Value = "№"
    ws.Range("B1").Value = "Время"
    ws.Range("C1").Value = "Тип ставки"
    ws.Range("D1").Value = "Выпало число"
    ws.Range("E1").Value = "Цвет"
    ws.Range("F1").Value = "Размер ставки"
    ws.Range("G1").Value = "Выигрыш/Проигрыш"
    
    ' Format headers
    Dim headerRow As Range
    Set headerRow = ws.Range("A1:G1")
    headerRow.Font.Bold = True
    headerRow.Interior.Color = RGB(0, 102, 204)
    headerRow.Font.Color = RGB(255, 255, 255)
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 18
    
    ' Format number columns
    ws.Columns("F").NumberFormat = "#,##0.00"
    ws.Columns("G").NumberFormat = "#,##0.00"
    
    ws.Range("A1").Value = "ИСТОРИЯ СТАВОК"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    
    ws.Range("A2").Value = "№"
    ws.Range("B2").Value = "Время"
    ws.Range("C2").Value = "Тип ставки"
    ws.Range("D2").Value = "Выпало число"
    ws.Range("E2").Value = "Цвет"
    ws.Range("F2").Value = "Размер ставки"
    ws.Range("G2").Value = "Выигрыш/Проигрыш"
    
    Set headerRow = ws.Range("A2:G2")
    headerRow.Font.Bold = True
    headerRow.Interior.Color = RGB(0, 102, 204)
    headerRow.Font.Color = RGB(255, 255, 255)
End Sub

' Exit Game
Sub BtnExit_Click()
    If MsgBox("Вы уверены, что хотите выйти?", vbYesNo) = vbYes Then
        ThisWorkbook.Close
    End If
End Sub

' Main Setup Routine - Run this first!
Sub SetupGame()
    Call CreateGameLayout
    Call InitializeGame(10000)
    MsgBox "Игра установлена успешно! Используйте кнопки для начала игры.", vbInformation
End Sub
