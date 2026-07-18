Option Explicit

' ======================================
' ROULETTE GAME - Main Module
' ======================================
' Полная игра в рулетку с визуальным интерфейсом
' Версия: 1.0

' Constants for European Roulette (0-36)
Public Const WHEEL_NUMBERS As Integer = 37
Public Const RED_NUMBERS As String = "1,3,5,7,9,12,14,16,18,19,21,23,25,27,30,32,34,36"
Public Const BLACK_NUMBERS As String = "2,4,6,8,10,11,13,15,17,20,22,24,26,28,29,31,33,35"

' Game State Structure
Public Type GameState
    CurrentBalance As Currency
    StartingBalance As Currency
    BetAmount As Currency
    SpinResult As Integer
    LastWinnings As Currency
    TotalBets As Currency
    TotalWinnings As Currency
    TotalSpins As Integer
    GameActive As Boolean
    CurrentStrategy As String
End Type

' Bet Record Structure
Public Type BetRecord
    BetNumber As Integer
    SpinNumber As Integer
    BetType As String
    BetValue As String
    BetAmount As Currency
    Winnings As Currency
    ResultColor As String
    BetTime As String
End Type

' Global Game State
Public gameState As GameState
Public betHistory As Collection

' Initialize Game
Sub InitializeGame(Optional startingBalance As Currency = 10000)
    Set betHistory = New Collection
    
    gameState.CurrentBalance = startingBalance
    gameState.StartingBalance = startingBalance
    gameState.BetAmount = 0
    gameState.SpinResult = -1
    gameState.LastWinnings = 0
    gameState.TotalBets = 0
    gameState.TotalWinnings = 0
    gameState.TotalSpins = 0
    gameState.GameActive = True
    gameState.CurrentStrategy = "Classic"
    
    UpdateGameDisplay
    MsgBox "Игра инициализирована!" & vbCrLf & "Начальный баланс: " & FormatCurrency(startingBalance), vbInformation, "Рулетка"
End Sub

' Spin the Wheel
Function SpinWheel() As Integer
    Randomize
    SpinWheel = Int(Rnd() * WHEEL_NUMBERS)
    gameState.SpinResult = SpinWheel
    gameState.TotalSpins = gameState.TotalSpins + 1
End Function

' Determine if number is Red
Function IsRed(number As Integer) As Boolean
    If number = 0 Then IsRed = False: Exit Function
    IsRed = (InStr(RED_NUMBERS, CStr(number)) > 0)
End Function

' Determine if number is Black
Function IsBlack(number As Integer) As Boolean
    If number = 0 Then IsBlack = False: Exit Function
    IsBlack = (InStr(BLACK_NUMBERS, CStr(number)) > 0)
End Function

' Get Color Name
Function GetColorName(number As Integer) As String
    If number = 0 Then
        GetColorName = "Зеро"
    ElseIf IsRed(number) Then
        GetColorName = "Красное"
    Else
        GetColorName = "Черное"
    End If
End Function

' Place Bet and Spin
Function PlaceBet(betAmount As Currency, betType As String, betValue As String) As Currency
    Dim result As Integer
    Dim winnings As Currency
    Dim resultColor As String
    
    ' Validate bet
    If betAmount <= 0 Then
        MsgBox "Некорректная сумма ставки!", vbCritical
        PlaceBet = 0
        Exit Function
    End If
    
    If betAmount > gameState.CurrentBalance Then
        MsgBox "Недостаточно средств! Ваш баланс: " & FormatCurrency(gameState.CurrentBalance), vbCritical
        PlaceBet = 0
        Exit Function
    End If
    
    ' Deduct bet from balance
    gameState.CurrentBalance = gameState.CurrentBalance - betAmount
    gameState.BetAmount = betAmount
    gameState.TotalBets = gameState.TotalBets + betAmount
    
    ' Spin the wheel
    result = SpinWheel
    resultColor = GetColorName(result)
    
    ' Calculate winnings
    winnings = CalculateWinnings(result, betType, betValue, betAmount)
    
    ' Add winnings to balance
    gameState.CurrentBalance = gameState.CurrentBalance + winnings
    gameState.LastWinnings = winnings
    gameState.TotalWinnings = gameState.TotalWinnings + winnings
    
    ' Record the bet
    RecordBet betAmount, result, betType, betValue, winnings, resultColor
    
    ' Display result
    DisplaySpinResult result, resultColor, betType, betValue, winnings
    
    ' Update display
    UpdateGameDisplay
    
    PlaceBet = winnings
End Function

' Calculate Winnings Based on Bet Type
Function CalculateWinnings(spinResult As Integer, betType As String, betValue As String, betAmount As Currency) As Currency
    Dim payout As Double
    
    Select Case UCase(betType)
        Case "SINGLE"
            If CInt(betValue) = spinResult Then
                payout = 36 ' 35:1 + original bet
            Else
                payout = 0
            End If
            
        Case "RED"
            If IsRed(spinResult) Then
                payout = 2 ' 1:1 + original bet
            Else
                payout = 0
            End If
            
        Case "BLACK"
            If IsBlack(spinResult) Then
                payout = 2 ' 1:1 + original bet
            Else
                payout = 0
            End If
            
        Case "ODD"
            If spinResult > 0 And spinResult Mod 2 = 1 Then
                payout = 2
            Else
                payout = 0
            End If
            
        Case "EVEN"
            If spinResult > 0 And spinResult Mod 2 = 0 Then
                payout = 2
            Else
                payout = 0
            End If
            
        Case "HIGH"
            If spinResult >= 19 And spinResult <= 36 Then
                payout = 2
            Else
                payout = 0
            End If
            
        Case "LOW"
            If spinResult >= 1 And spinResult <= 18 Then
                payout = 2
            Else
                payout = 0
            End If
            
        Case "DOZEN1"
            If spinResult >= 1 And spinResult <= 12 Then
                payout = 3 ' 2:1 + original bet
            Else
                payout = 0
            End If
            
        Case "DOZEN2"
            If spinResult >= 13 And spinResult <= 24 Then
                payout = 3
            Else
                payout = 0
            End If
            
        Case "DOZEN3"
            If spinResult >= 25 And spinResult <= 36 Then
                payout = 3
            Else
                payout = 0
            End If
            
        Case "COLUMN1"
            If spinResult > 0 And spinResult Mod 3 = 1 Then
                payout = 3
            Else
                payout = 0
            End If
            
        Case "COLUMN2"
            If spinResult > 0 And spinResult Mod 3 = 2 Then
                payout = 3
            Else
                payout = 0
            End If
            
        Case "COLUMN3"
            If spinResult > 0 And spinResult Mod 3 = 0 Then
                payout = 3
            Else
                payout = 0
            End If
            
        Case Else
            payout = 0
    End Select
    
    CalculateWinnings = betAmount * payout
End Function

' Record Bet in History
Sub RecordBet(betAmount As Currency, spinResult As Integer, betType As String, betValue As String, winnings As Currency, resultColor As String)
    Dim record As BetRecord
    
    record.BetNumber = betHistory.Count + 1
    record.SpinNumber = spinResult
    record.BetType = betType
    record.BetValue = betValue
    record.BetAmount = betAmount
    record.Winnings = winnings
    record.ResultColor = resultColor
    record.BetTime = Format(Now, "hh:mm:ss")
    
    betHistory.Add record
End Sub

' Display Spin Result
Sub DisplaySpinResult(spinResult As Integer, resultColor As String, betType As String, betValue As String, winnings As Currency)
    Dim message As String
    
    message = "════════════════════════════════════" & vbCrLf
    message = message & "РЕЗУЛЬТАТ ВРАЩЕНИЯ" & vbCrLf
    message = message & "════════════════════════════════════" & vbCrLf & vbCrLf
    message = message & "Выпало число: " & spinResult & " (" & resultColor & ")" & vbCrLf & vbCrLf
    message = message & "Тип ставки: " & betType & vbCrLf
    message = message & "Значение ставки: " & betValue & vbCrLf
    message = message & "Сумма ставки: " & FormatCurrency(gameState.BetAmount) & vbCrLf & vbCrLf
    
    If winnings > 0 Then
        message = message & "✓ ВЫ ВЫИГРАЛИ! " & vbCrLf
        message = message & "Выигрыш: " & FormatCurrency(winnings) & vbCrLf
    Else
        message = message & "✗ ВЫ ПРОИГРАЛИ!" & vbCrLf
    End If
    
    message = message & vbCrLf & "════════════════════════════════════" & vbCrLf
    message = message & "Текущий баланс: " & FormatCurrency(gameState.CurrentBalance)
    
    MsgBox message, IIf(winnings > 0, vbInformation, vbExclamation), "Рулетка"
End Sub

' Update Game Display in Spreadsheet
Sub UpdateGameDisplay()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Игра")
    
    If ws Is Nothing Then Exit Sub
    
    ' Update balance
    ws.Range("B3").Value = gameState.CurrentBalance
    ws.Range("B4").Value = gameState.TotalBets
    ws.Range("B5").Value = gameState.TotalWinnings
    ws.Range("B6").Value = gameState.CurrentBalance - gameState.StartingBalance
    ws.Range("B7").Value = gameState.TotalSpins
    
    ' Update history display
    Call UpdateHistoryDisplay
End Sub

' Update History Display
Sub UpdateHistoryDisplay()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim row As Integer
    Dim i As Integer
    Dim record As BetRecord
    
    Set ws = ThisWorkbook.Sheets("История")
    If ws Is Nothing Then Exit Sub
    
    ' Clear previous data
    ws.Range("A2:G1000").ClearContents
    
    ' Add header
    row = 2
    For i = 1 To betHistory.Count
        Set record = betHistory(i)
        
        ws.Cells(row, 1).Value = record.BetNumber
        ws.Cells(row, 2).Value = record.BetTime
        ws.Cells(row, 3).Value = record.BetType
        ws.Cells(row, 4).Value = record.SpinNumber
        ws.Cells(row, 5).Value = record.ResultColor
        ws.Cells(row, 6).Value = record.BetAmount
        ws.Cells(row, 7).Value = record.Winnings
        
        ' Color code wins/losses
        If record.Winnings > 0 Then
            ws.Cells(row, 7).Interior.Color = RGB(0, 176, 0)
            ws.Cells(row, 7).Font.Color = RGB(255, 255, 255)
        Else
            ws.Cells(row, 7).Interior.Color = RGB(255, 0, 0)
            ws.Cells(row, 7).Font.Color = RGB(255, 255, 255)
        End If
        
        row = row + 1
    Next i
End Sub

' Show Game Statistics
Sub ShowGameStats()
    Dim message As String
    Dim netResult As Currency
    Dim winRate As Double
    
    netResult = gameState.CurrentBalance - gameState.StartingBalance
    winRate = IIf(gameState.TotalSpins = 0, 0, (gameState.TotalWinnings / gameState.TotalBets) * 100)
    
    message = "════════════════════════════════════" & vbCrLf
    message = message & "СТАТИСТИКА ИГРЫ" & vbCrLf
    message = message & "════════════════════════════════════" & vbCrLf & vbCrLf
    message = message & "Начальный баланс: " & FormatCurrency(gameState.StartingBalance) & vbCrLf
    message = message & "Текущий баланс: " & FormatCurrency(gameState.CurrentBalance) & vbCrLf
    message = message & "Чистый результат: " & FormatCurrency(netResult) & vbCrLf & vbCrLf
    message = message & "Всего ставок: " & gameState.TotalSpins & vbCrLf
    message = message & "Сумма ставок: " & FormatCurrency(gameState.TotalBets) & vbCrLf
    message = message & "Сумма выигрышей: " & FormatCurrency(gameState.TotalWinnings) & vbCrLf
    message = message & "Коэффициент: " & Format(winRate, "0.00") & "%" & vbCrLf & vbCrLf
    message = message & "Последний результат: " & FormatCurrency(gameState.LastWinnings)
    
    MsgBox message, vbInformation, "Статистика"
End Sub

' Reset Game
Sub ResetGame()
    If MsgBox("Вы уверены? Все данные будут удалены!", vbYesNo) = vbYes Then
        InitializeGame gameState.StartingBalance
    End If
End Sub

' Load betting strategy
Sub LoadStrategy(strategyName As String)
    gameState.CurrentStrategy = strategyName
    MsgBox "Стратегия загружена: " & strategyName, vbInformation
End Sub
