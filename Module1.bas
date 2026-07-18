Option Explicit

' Roulette Game Module
' A complete VBA implementation of a European Roulette game

' Constants for the game
Const WHEEL_NUMBERS As Integer = 37 ' 0-36 for European Roulette
Const RED_NUMBERS As String = "1,3,5,7,9,12,14,16,18,19,21,23,25,27,30,32,34,36"
Const BLACK_NUMBERS As String = "2,4,6,8,10,11,13,15,17,20,22,24,26,28,29,31,33,35"

' Game State Structure
Type GameState
    CurrentBalance As Currency
    BetAmount As Currency
    SpinResult As Integer
    LastWinnings As Currency
    TotalBets As Currency
    TotalWinnings As Currency
    GameHistory As Collection
End Type

Dim gameState As GameState

' Initialize the game
Sub InitializeGame(Optional startingBalance As Currency = 1000)
    Set gameState.GameHistory = New Collection
    gameState.CurrentBalance = startingBalance
    gameState.BetAmount = 0
    gameState.SpinResult = -1
    gameState.LastWinnings = 0
    gameState.TotalBets = 0
    gameState.TotalWinnings = 0
    MsgBox "Roulette Game Initialized!" & vbCrLf & "Starting Balance: $" & Format(startingBalance, "0.00"), vbInformation, "Game Start"
End Sub

' Spin the wheel and return a random number (0-36)
Function SpinWheel() As Integer
    Randomize
    SpinWheel = Int(Rnd() * WHEEL_NUMBERS)
    gameState.SpinResult = SpinWheel
End Function

' Determine if a number is red
Function IsRed(number As Integer) As Boolean
    If number = 0 Then IsRed = False: Exit Function
    IsRed = (InStr(RED_NUMBERS, CStr(number)) > 0)
End Function

' Determine if a number is black
Function IsBlack(number As Integer) As Boolean
    If number = 0 Then IsBlack = False: Exit Function
    IsBlack = (InStr(BLACK_NUMBERS, CStr(number)) > 0)
End Function

' Place a bet and spin
Function PlaceBet(betAmount As Currency, betType As String, betValue As String) As Currency
    Dim result As Integer
    Dim winnings As Currency
    
    ' Validate bet
    If betAmount <= 0 Then
        MsgBox "Invalid bet amount!", vbCritical
        PlaceBet = 0
        Exit Function
    End If
    
    If betAmount > gameState.CurrentBalance Then
        MsgBox "Insufficient balance! Your balance: $" & Format(gameState.CurrentBalance, "0.00"), vbCritical
        PlaceBet = 0
        Exit Function
    End If
    
    ' Deduct bet from balance
    gameState.CurrentBalance = gameState.CurrentBalance - betAmount
    gameState.BetAmount = betAmount
    gameState.TotalBets = gameState.TotalBets + betAmount
    
    ' Spin the wheel
    result = SpinWheel
    
    ' Calculate winnings
    winnings = CalculateWinnings(result, betType, betValue, betAmount)
    
    ' Add winnings to balance
    gameState.CurrentBalance = gameState.CurrentBalance + winnings
    gameState.LastWinnings = winnings
    gameState.TotalWinnings = gameState.TotalWinnings + winnings
    
    ' Record the bet
    RecordBet betAmount, betType, betValue, result, winnings
    
    ' Display result
    DisplaySpinResult result, betType, betValue, winnings
    
    PlaceBet = winnings
End Function

' Calculate winnings based on bet type
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
                payout = 2 ' 1:1 + original bet
            Else
                payout = 0
            End If
            
        Case "EVEN"
            If spinResult > 0 And spinResult Mod 2 = 0 Then
                payout = 2 ' 1:1 + original bet
            Else
                payout = 0
            End If
            
        Case "HIGH"
            If spinResult >= 19 And spinResult <= 36 Then
                payout = 2 ' 1:1 + original bet
            Else
                payout = 0
            End If
            
        Case "LOW"
            If spinResult >= 1 And spinResult <= 18 Then
                payout = 2 ' 1:1 + original bet
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
                payout = 3 ' 2:1 + original bet
            Else
                payout = 0
            End If
            
        Case "DOZEN3"
            If spinResult >= 25 And spinResult <= 36 Then
                payout = 3 ' 2:1 + original bet
            Else
                payout = 0
            End If
            
        Case Else
            payout = 0
    End Select
    
    CalculateWinnings = betAmount * payout
End Function

' Display spin result
Sub DisplaySpinResult(spinResult As Integer, betType As String, betValue As String, winnings As Currency)
    Dim message As String
    Dim resultColor As String
    
    resultColor = IIf(spinResult = 0, "Green", IIf(IsRed(spinResult), "Red", "Black"))
    
    message = "Wheel Spin Result: " & spinResult & " (" & resultColor & ")" & vbCrLf & vbCrLf
    message = message & "Bet Type: " & betType & vbCrLf
    message = message & "Bet Value: " & betValue & vbCrLf
    message = message & "Bet Amount: $" & Format(gameState.BetAmount, "0.00") & vbCrLf & vbCrLf
    
    If winnings > 0 Then
        message = message & "YOU WIN! $" & Format(winnings, "0.00") & vbCrLf
    Else
        message = message & "YOU LOSE!" & vbCrLf
    End If
    
    message = message & "Current Balance: $" & Format(gameState.CurrentBalance, "0.00")
    
    MsgBox message, IIf(winnings > 0, vbInformation, vbExclamation), "Spin Result"
End Sub

' Record bet in history
Sub RecordBet(betAmount As Currency, betType As String, betValue As String, spinResult As Integer, winnings As Currency)
    Dim record As String
    record = Now & " | Bet: $" & Format(betAmount, "0.00") & " | Type: " & betType & " | Result: " & spinResult & " | Winnings: $" & Format(winnings, "0.00")
    gameState.GameHistory.Add record
End Sub

' Display game statistics
Sub ShowGameStats()
    Dim message As String
    Dim totalBetsNet As Currency
    
    totalBetsNet = gameState.TotalWinnings - gameState.TotalBets
    
    message = "=== GAME STATISTICS ===" & vbCrLf & vbCrLf
    message = message & "Current Balance: $" & Format(gameState.CurrentBalance, "0.00") & vbCrLf
    message = message & "Total Bets Placed: $" & Format(gameState.TotalBets, "0.00") & vbCrLf
    message = message & "Total Winnings: $" & Format(gameState.TotalWinnings, "0.00") & vbCrLf
    message = message & "Net Result: $" & Format(totalBetsNet, "0.00") & vbCrLf & vbCrLf
    message = message & "Last Win/Loss: $" & Format(gameState.LastWinnings, "0.00")
    
    MsgBox message, vbInformation, "Game Statistics"
End Sub

' Display game history
Sub ShowGameHistory()
    Dim message As String
    Dim i As Integer
    Dim startIdx As Integer
    
    If gameState.GameHistory.Count = 0 Then
        MsgBox "No game history yet!", vbInformation
        Exit Sub
    End If
    
    ' Show last 10 bets
    startIdx = Application.Max(1, gameState.GameHistory.Count - 9)
    
    message = "=== LAST GAME HISTORY ===" & vbCrLf & vbCrLf
    For i = startIdx To gameState.GameHistory.Count
        message = message & gameState.GameHistory(i) & vbCrLf
    Next i
    
    MsgBox message, vbInformation, "Game History"
End Sub

' Reset the game
Sub ResetGame(Optional newBalance As Currency = 1000)
    InitializeGame newBalance
End Sub

' Example usage - Main game loop
Sub PlayRoulette()
    Dim choice As String
    Dim betAmount As Currency
    Dim betType As String
    Dim betValue As String
    Dim continuePlay As Boolean
    
    InitializeGame 1000
    
    continuePlay = True
    Do While continuePlay And gameState.CurrentBalance > 0
        choice = InputBox("Choose action:" & vbCrLf & _
                         "1 - Place Bet" & vbCrLf & _
                         "2 - View Stats" & vbCrLf & _
                         "3 - View History" & vbCrLf & _
                         "4 - Quit" & vbCrLf & vbCrLf & _
                         "Current Balance: $" & Format(gameState.CurrentBalance, "0.00"), "Roulette Game")
        
        Select Case choice
            Case "1"
                betAmount = CDbl(InputBox("Enter bet amount:", "Roulette Game"))
                betType = InputBox("Bet type (SINGLE, RED, BLACK, ODD, EVEN, HIGH, LOW, DOZEN1, DOZEN2, DOZEN3):", "Roulette Game")
                betValue = InputBox("Bet value (number for SINGLE, or leave blank for color/odd/even):", "Roulette Game", "")
                PlaceBet betAmount, betType, IIf(betValue = "", "0", betValue)
                
            Case "2"
                ShowGameStats
                
            Case "3"
                ShowGameHistory
                
            Case "4"
                continuePlay = False
                ShowGameStats
                MsgBox "Thanks for playing!", vbInformation, "Game End"
        End Select
    Loop
    
    If gameState.CurrentBalance = 0 Then
        MsgBox "Game Over! You're out of balance!", vbCritical, "Bust"
    End If
End Sub
