Attribute VB_Name = "Franz_Export_Position"
Option Explicit

Sub exportPositionToCSV()
'' Creates a .csv containing the position for each account.
''
'' Positions Type:  0 - Error (unrecognized)
''                  1 - Call
''                  2 - Put
''                  3 - SPY
''                  4 - ES Future
''
'' Columns: 1 - Expiration Date     10 - Theta
''          2 - Strike              11 - Charm
''          3 - Type                12 - OEV
''          4 - Quantity            13 - Bid
''          5 - ImpliedVol          14 - Ask
''          6 - Theoretical Price   15 - TTE
''          7 - Delta               16 - DTE
''          8 - Gamma               17 - Forward
''          9 - Vega                18 - Multiplier
''                                  19 - Spot Price

    Call clearImmediateWindow
    Call xlUpdates(False)
    
    'Save the output here
    Dim path As String
    path = "A:\Matt Franz\Positions\"
    
    'First position row in PosSumTotal1 and PosSumTotal2
    Dim fr As Integer
    fr = 12
    
    'Define column layout
    Dim expC As Integer, kC As Integer, qC As Integer, priceC As Integer
    Dim deltaC As Integer, gammaC As Integer, vegaC As Integer, thetaC As Integer
    Dim charmC As Integer, oevC As Integer, ivC As Integer, oTypeC As Integer
    Dim bidC As Integer, askC As Integer, dteC As Integer, fwdC As Integer
    expC = 2
    kC = 3
    oTypeC = 4
    qC = 5
    ivC = 9
    priceC = 10
    deltaC = 11
    gammaC = 12
    vegaC = 13
    thetaC = 14
    charmC = 15
    oevC = 16
    fwdC = 25
    bidC = 29
    askC = 30
    dteC = 33
    
    Dim wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook
    
    Dim spyPrice As Double
    spyPrice = wb.Sheets("Summary").Range("B3").value
    
    Dim spotPrice As Double
    spotPrice = wb.Sheets("Summary").Range("B8").value
    
    Dim nAcct As Integer, iAcct As Integer
    nAcct = 2
    
    For iAcct = 1 To nAcct
        Dim sheetName As String
        sheetName = getSheetName(iAcct)
        Set ws = wb.Sheets(sheetName)
        
        'So we can read IV
        ws.Cells(11, 9).value = "IVol"
        ws.Calculate
        DoEvents
        
        Dim lr As Integer
        lr = countPositions(ws, fr)
        
        'Read data
        Dim expirationDate As Variant
        expirationDate = ws.Range(ws.Cells(fr, expC), ws.Cells(lr, expC)).value
        
        Dim strike As Variant
        strike = ws.Range(ws.Cells(fr, kC), ws.Cells(lr, kC)).value
        
        Dim oType As Variant
        oType = ws.Range(ws.Cells(fr, oTypeC), ws.Cells(lr, oTypeC)).value
        
        Dim quantity As Variant
        quantity = ws.Range(ws.Cells(fr, qC), ws.Cells(lr, qC)).value
        
        Dim impliedVol As Variant
        impliedVol = ws.Range(ws.Cells(fr, ivC), ws.Cells(lr, ivC)).value
        
        Dim price As Variant
        price = ws.Range(ws.Cells(fr, priceC), ws.Cells(lr, priceC)).value
        
        Dim delta As Variant
        delta = ws.Range(ws.Cells(fr, deltaC), ws.Cells(lr, deltaC)).value
        
        Dim gamma As Variant
        gamma = ws.Range(ws.Cells(fr, gammaC), ws.Cells(lr, gammaC)).value
        
        Dim vega As Variant
        vega = ws.Range(ws.Cells(fr, vegaC), ws.Cells(lr, vegaC)).value
        
        Dim theta As Variant
        theta = ws.Range(ws.Cells(fr, thetaC), ws.Cells(lr, thetaC)).value
        
        Dim charm As Variant
        charm = ws.Range(ws.Cells(fr, charmC), ws.Cells(lr, charmC)).value
        
        Dim oev As Variant
        oev = ws.Range(ws.Cells(fr, oevC), ws.Cells(lr, oevC)).value
        
        Dim bid As Variant
        bid = ws.Range(ws.Cells(fr, bidC), ws.Cells(lr, bidC)).value
        
        Dim ask As Variant
        ask = ws.Range(ws.Cells(fr, askC), ws.Cells(lr, askC)).value
        
        Dim tte As Variant
        tte = ws.Range(ws.Cells(fr, dteC), ws.Cells(lr, dteC)).value
        
        Dim expDate As Variant, jPos As Integer, dte As Variant, row As Integer
        Dim nPos As Integer
        nPos = lr - fr
        ReDim dte(0 To nPos, 1)
        For jPos = 0 To nPos
            row = jPos + fr
            expDate = ws.Range(ws.Cells(row, expC), ws.Cells(row, expC)).value
            dte(jPos, 1) = expDate - Now()
        Next jPos
        
        Dim fwd As Variant
        fwd = ws.Range(ws.Cells(fr, fwdC), ws.Cells(lr, fwdC)).value
        
        'Export data to a .csv file
        Dim acctName As String
        If iAcct = 1 Then
            acctName = "s12"
        ElseIf iAcct = 2 Then
            acctName = "mio"
        Else
            acctName = "ERROR"
        End If
        
        Dim fileName As String
        fileName = "position_" & acctName & "_" & Format(Now, "YYYYMMDD") & "_" & Format(Now, "HHMM") & ".csv"
        Debug.Print fileName
        
        Dim OutputFileNum As Integer
        OutputFileNum = FreeFile
        
        Open path & fileName For Output Lock Write As #OutputFileNum
        
        'Create file headers
        Print #OutputFileNum, "Exp Date" & "," & _
                              "Strike" & "," & _
                              "C/P" & "," & _
                              "Quantity" & "," & _
                              "Implied Vol" & "," & _
                              "Theo Price" & "," & _
                              "Delta" & "," & _
                              "Gamma" & "," & _
                              "Vega" & "," & _
                              "Theta" & "," & _
                              "Charm" & "," & _
                              "OEV" & "," & _
                              "Bid" & "," & _
                              "Ask" & "," & _
                              "TTE" & "," & _
                              "DTE" & "," & _
                              "Forward" & "," & _
                              "Multiplier"; "," & _
                              "Spot"
        
        Dim r As Integer
        For r = LBound(expirationDate, 1) To UBound(expirationDate, 1)
            Dim position As String
            
            If r = 1 Or r = 2 Then
            'SPY Stock
                If r = 1 Then
                    position = "" & "," & _
                               "" & "," & _
                               "3" & "," & _
                               100 * quantity(r, 1) & "," & _
                               "" & "," & _
                               spyPrice & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "" & "," & _
                               "1" & "," & _
                               spotPrice
                    Print #OutputFileNum, position
                Else
                    'Skip the second leg of the combo
                End If
            ElseIf strike(r, 1) = 0 Then
            'Future
                position = expirationDate(r, 1) & "," & _
                           "" & "," & _
                           "4" & "," & _
                           quantity(r, 1) & "," & _
                           "" & "," & _
                           price(r, 1) & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "" & "," & _
                           "50" & "," & _
                           spotPrice
                Print #OutputFileNum, position
            Else
            'Options
                ' Conver C/P to 1/2
                If oType(r, 1) = "C" Then
                    oType(r, 1) = 1
                ElseIf oType(r, 1) = "P" Then
                    oType(r, 1) = 2
                Else
                    oType(r, 1) = 0
                End If
                
                'Convert date string to serial
                expirationDate(r, 1) = CDbl(expirationDate(r, 1))
            
                position = expirationDate(r, 1) & "," & _
                           strike(r, 1) & "," & _
                           oType(r, 1) & "," & _
                           quantity(r, 1) & "," & _
                           impliedVol(r, 1) & "," & _
                           price(r, 1) & "," & _
                           delta(r, 1) & "," & _
                           gamma(r, 1) & "," & _
                           vega(r, 1) & "," & _
                           theta(r, 1) & "," & _
                           charm(r, 1) & "," & _
                           oev(r, 1) & "," & _
                           bid(r, 1) & "," & _
                           ask(r, 1) & "," & _
                           tte(r, 1) & "," & _
                           dte(r - 1, 1) & "," & _
                           fwd(r, 1) & "," & _
                           "100" & "," & _
                           spotPrice
                Print #OutputFileNum, position
            End If
        Next r
                              
        Close OutputFileNum
    Next iAcct
    
    Call xlUpdates(True)
    Debug.Print "Done."
End Sub

Private Function countPositions(ByVal ws As Worksheet, ByVal row As Integer) As Integer
    Dim pos As Variant
    pos = ws.Range("A" & row).value
    
    Do While pos <> ""
        row = row + 1
        pos = ws.Range("A" & row).value
    Loop
    countPositions = row - 1
End Function

Private Function getSheetName(ByVal acctNum As Integer) As String
    If acctNum = 1 Then
        getSheetName = "PosSumTotal1"
    ElseIf acctNum = 2 Then
        getSheetName = "PosSumTotal2"
    Else
        Debug.Print "Account number not recognized. 1 and 2 are valid."
        getSheetName = "ERROR"
    End If
End Function

Private Sub xlUpdates(turnOn As Boolean)
'Turns excel display updates on/off so code runs faster
    If turnOn = True Then
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.Calculate
    Else
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If
End Sub

Private Sub clearImmediateWindow()
    Dim i As Integer
    For i = 1 To 100
        Debug.Print vbNewLine
    Next i
    Debug.Print "*******************"
    Debug.Print "Running..."
End Sub
