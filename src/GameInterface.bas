' ======================================
' GAME INTERFACE MODULE
' ======================================
' Обработчики кнопок и элементов управления игрой

Option Explicit

' Button Click Handlers - Красное/Черное (RED/BLACK)
Sub BtnRed_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "RED", ""
    End If
End Sub

Sub BtnBlack_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "BLACK", ""
    End If
End Sub

' Button Click Handlers - Четные/Нечетные (EVEN/ODD)
Sub BtnEven_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "EVEN", ""
    End If
End Sub

Sub BtnOdd_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "ODD", ""
    End If
End Sub

' Button Click Handlers - Высокие/Низкие (HIGH/LOW 19-36 / 1-18)
Sub BtnHigh_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "HIGH", ""
    End If
End Sub

Sub BtnLow_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "LOW", ""
    End If
End Sub

' Button Click Handlers - Дюжины (DOZENS)
Sub BtnDozen1_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "DOZEN1", "1-12"
    End If
End Sub

Sub BtnDozen2_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "DOZEN2", "13-24"
    End If
End Sub

Sub BtnDozen3_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "DOZEN3", "25-36"
    End If
End Sub

' Button Click Handlers - Колонны (COLUMNS)
Sub BtnColumn1_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "COLUMN1", "1,4,7,10,13,16,19,22,25,28,31,34"
    End If
End Sub

Sub BtnColumn2_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "COLUMN2", "2,5,8,11,14,17,20,23,26,29,32,35"
    End If
End Sub

Sub BtnColumn3_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    
    If betAmount > 0 Then
        PlaceBet betAmount, "COLUMN3", "3,6,9,12,15,18,21,24,27,30,33,36"
    End If
End Sub

' Button Click Handlers - Одиночные числа (SINGLE NUMBERS)
Sub BtnSingleNumber_Click()
    Dim number As Integer
    Dim betAmount As Currency
    
    betAmount = GetBetAmount()
    
    Dim response As String
    response = InputBox("Введите номер (0-36):", "Ставка на число", "")
    
    If response <> "" Then
        On Error Resume Next
        number = CInt(response)
        On Error GoTo 0
        
        If number >= 0 And number <= 36 Then
            If betAmount > 0 Then
                PlaceBet betAmount, "SINGLE", CStr(number)
            End If
        Else
            MsgBox "Неверный номер! Используйте 0-36", vbCritical
        End If
    End If
End Sub

' Utility Functions

' Get bet amount from spinbox/textbox
Function GetBetAmount() As Currency
    Dim ws As Worksheet
    Dim betValue As Variant
    
    Set ws = ThisWorkbook.Sheets("Игра")
    On Error Resume Next
    
    betValue = ws.Range("D3").Value
    
    If betValue = "" Or betValue <= 0 Or betValue > 100 Then
        MsgBox "Введите сумму ставки от 1 до 100!", vbCritical
        GetBetAmount = 0
        Exit Function
    End If
    
    GetBetAmount = CDbl(betValue)
End Function

' Initialize Game Button
Sub BtnStartGame_Click()
    Dim startBalance As Currency
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Игра")
    
    startBalance = ws.Range("B1").Value
    If startBalance <= 0 Then
        startBalance = 10000
        ws.Range("B1").Value = 10000
    End If
    
    InitializeGame startBalance
    MsgBox "Игра инициализирована с начальным балансом: " & FormatCurrency(startBalance), vbInformation
End Sub

' Statistics Button
Sub BtnStatistics_Click()
    ShowGameStats
End Sub

' Reset Game Button
Sub BtnResetGame_Click()
    ResetGame
End Sub

' View History Button
Sub BtnViewHistory_Click()
    ThisWorkbook.Sheets("История").Activate
    MsgBox "Откройте лист 'История' для просмотра полной истории ставок", vbInformation
End Sub

' Strategy Selector - Classic
Sub BtnStrategyClassic_Click()
    LoadStrategy "Классическая"
    MsgBox "Выбрана стратегия: Классическая (обычные ставки)", vbInformation
End Sub

' Strategy Selector - Martingale
Sub BtnStrategyMartingale_Click()
    LoadStrategy "Мартингейл"
    MsgBox "Выбрана стратегия: Мартингейл (удвоение ставки после проигрыша)", vbInformation
End Sub

' Strategy Selector - D'Alembert
Sub BtnStrategyDAlembert_Click()
    LoadStrategy "Д'Аламбер"
    MsgBox "Выбрана стратегия: Д'Аламбер (постепенное увеличение ставки)", vbInformation
End Sub

' Auto Spin (continuous play)
Sub BtnAutoSpin_Click()
    Dim i As Integer
    Dim count As Integer
    Dim betAmount As Currency
    
    count = InputBox("Сколько спинов? (1-100):", "Автоматические вращения", "10")
    
    If count > 0 And count <= 100 Then
        betAmount = GetBetAmount()
        
        If betAmount > 0 Then
            For i = 1 To count
                DoEvents
                PlaceBet betAmount, "RED", ""
                Application.Wait Now + TimeValue("0:00:01")
            Next i
            
            MsgBox "Автоматические вращения завершены!", vbInformation
        End If
    End If
End Sub

' Clear History
Sub BtnClearHistory_Click()
    If MsgBox("Вы уверены? История будет удалена!", vbYesNo) = vbYes Then
        Set betHistory = New Collection
        ThisWorkbook.Sheets("История").Range("A2:G1000").ClearContents
        MsgBox "История очищена", vbInformation
    End If
End Sub
