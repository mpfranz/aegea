Attribute VB_Name = "Franz_MAC_Risk"
Option Explicit
Option Private Module

Sub test()
    Call clearImmediateWindow
    Call xlUpdates(False)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Test")
    
    Dim rowArr As Variant
    rowArr = Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22)
    
    'Inputs
    Dim row As Variant
    For Each row In rowArr
        Debug.Print "row = " & row
    
        Dim spotPrice As Double
        spotPrice = 2126.41
    
        Dim posType As String
        posType = ws.Range("D" & row).Value
        
        Dim k As Double
        k = ws.Range("C" & row).Value
        
        Dim expirationDate As Date
        expirationDate = ws.Range("B" & row).Value
    
        Dim quantity As Double
        quantity = ws.Range("E" & row).Value
        
        Dim y As Double
        y = ws.Range("G" & row).Value
        
        Dim r As Double
        r = ws.Range("H" & row).Value
        
        Dim calendarDTE As Double
        'calendarDTE = expirationDate - Date
        calendarDTE = expirationDate - DateValue("10/28/2016")
        
        Dim curPrice As Double
        curPrice = ws.Range("F" & row).Value
    
    'Convert the market price to an implied vol
        Dim forward As Double
        forward = calcForwardPrice(spotPrice, r, y, calendarDTE, False)
    
        Dim impliedVol As Double
        impliedVol = priceToImpliedVol(curPrice, posType, k, forward, calendarDTE, False, r)
        ws.Range("I" & row).Value = impliedVol
        'impliedVol = ws.Range("O" & row).Value
    
    'Price Risk
        Dim priceShockArr As Variant
        priceShockArr = Array(-0.5, -0.45, -0.4, -0.35, -0.3, -0.25, -0.2, -0.175, -0.15, -0.125, -0.1, _
                              -0.08, -0.06, -0.04, -0.03, -0.02, -0.01, 0, 0.01, 0.02, 0.03, 0.04, 0.06, _
                              0.08, 0.1, 0.125, 0.15, 0.175, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5)

        Dim i As Integer
        For i = 0 To UBound(priceShockArr)
            Dim priceShock As Double
            priceShock = priceShockArr(i)
            
            Dim mlPriceRisk
            mlPriceRisk = mlMarginPriceRisk(posType, k, calendarDTE, quantity, curPrice, _
                forward, impliedVol, priceShock, r)

            If priceShock = 0 Then ws.Range("K" & row).Value = mlPriceRisk
        Next i

    'Vol Risk
        Dim mlVolRisk As Double
        mlVolRisk = mlMarginVolRisk(posType, k, calendarDTE, quantity, forward, impliedVol)
        Debug.Print "mlVolRisk = " & mlVolRisk & vbNewLine

    'Liquidity Risk
        Dim mlGrossLiqRisk As Double, mlNetLiqRisk As Double
        mlGrossLiqRisk = mlMarginVolLiquidityRisk(posType, k, calendarDTE, quantity, forward, impliedVol, True)
        mlNetLiqRisk = mlMarginVolLiquidityRisk(posType, k, calendarDTE, quantity, forward, impliedVol, False)
        Debug.Print "mlGrossLiqRisk = " & mlGrossLiqRisk
        Debug.Print "mlNetLiqRisk = " & mlNetLiqRisk & vbNewLine
    Next row
    
    Call xlUpdates(True)
    Debug.Print "Done."
End Sub

Function mlMarginPriceRisk(ByVal posType As String, ByVal k As Double, ByVal calendarDTE As Double, _
    ByVal q As Double, ByVal curPrice As Double, ByVal fwd As Double, ByVal iv As Double, _
    ByVal priceShock As Double, Optional ByVal r As Double = 0) As Double
    
    Dim multiplier As Integer
    multiplier = getMultiplier(posType)
    
    Dim newPrice As Double
    If posType = "SPY" Or posType = "F" Then
        newPrice = (1 + priceShock) * curPrice
    Else
        Dim newFwd As Double
        newFwd = (1 + priceShock) * fwd
        newPrice = priceBS(newFwd, k, calendarDTE, iv, posType, False, r)
    End If

    Dim pnl As Double
    pnl = multiplier * q * (newPrice - curPrice)
    
    mlMarginPriceRisk = pnl
End Function

Function mlMarginVolRisk(ByVal posType As String, ByVal k As Double, ByVal calendarDTE As Double, _
    ByVal q As Double, ByVal f As Double, ByVal iv As Double) As Double
    
    Dim pnl As Double
    If posType = "SPY" Or posType = "F" Then
        pnl = 0
    Else
        Dim vega As Double
        vega = 100 * q * vegaBS(f, k, iv, calendarDTE, False)
        
        Dim volShock As Double
        volShock = calcVolShock(calendarDTE)
        
        pnl = vega * volShock * iv * 100
    End If

    mlMarginVolRisk = pnl
End Function

Function mlMarginVolLiquidityRisk(ByVal posType As String, ByVal k As Double, ByVal calendarDTE As Double, _
    ByVal q As Double, ByVal forward As Double, ByVal iv As Double, ByVal useGross As Boolean) As Double
                                    
    Dim pnl As Double
    If posType = "SPY" Or posType = "F" Then
        pnl = 0
    Else
        Dim volLiqCharge As Double
        volLiqCharge = calcVolLiqCharge(forward, k)
        
        Dim volShock As Double
        volShock = calcVolShock(calendarDTE)
        
        Dim liqVolShift As Double
        liqVolShift = calcLiqVolShift(volLiqCharge, volShock, iv)
        
        Dim vega As Double
        vega = 100 * q * vegaBS(forward, k, iv, calendarDTE, False)
        
        Dim netVegaLiqRisk As Double, grossVegaLiqRisk As Double
        netVegaLiqRisk = vega * liqVolShift
        grossVegaLiqRisk = Abs(netVegaLiqRisk)
        
        If useGross = True Then
            pnl = grossVegaLiqRisk
        Else
            pnl = netVegaLiqRisk
        End If
    End If
    
    mlMarginVolLiquidityRisk = pnl
End Function

Function getMultiplier(posType As String) As Integer
    If posType = "SPY" Then
        getMultiplier = 1
    ElseIf posType = "F" Then
        getMultiplier = 50
    ElseIf posType = "C" Or posType = "P" Then
        getMultiplier = 100
    Else
        getMultiplier = 0
        Debug.Print "Error: Postion type not recognized: " & posType
    End If
End Function

Function calcLiqVolShift(ByVal volLiqCharge As Double, ByVal volShock As Double, impliedVol As Double) As Double
    Dim liqVolShift As Double
    liqVolShift = volLiqCharge * volShock * impliedVol * 100
    liqVolShift = Application.Min(2, liqVolShift)
    calcLiqVolShift = liqVolShift
End Function

Function calcVolLiqCharge(ByVal forward As Double, strike As Double) As Double
    Dim moneyness As Double
    moneyness = Abs((forward - strike) / forward)

    If moneyness < 0.3 Then
        calcVolLiqCharge = 0.1
    ElseIf moneyness < 0.45 Then
        calcVolLiqCharge = 0.2
    Else
        calcVolLiqCharge = 0.3
    End If
End Function

Function calcVolShock(ByVal calendarDTE As Double) As Double
    Dim monthsToExp As Double
    monthsToExp = 12 * calendarDTE / 365
    
    Dim threeMonthVolShock As Double
    threeMonthVolShock = 0.26
    
    Dim volShock As Double
    If monthsToExp <= 0 Then
        volShock = 0
    Else
        volShock = threeMonthVolShock * Sqr(3 / monthsToExp)
    End If
    
    calcVolShock = volShock
End Function
