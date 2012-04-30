Attribute VB_Name = "Module1"
Option Explicit
Private Function CopyArrayDimensions(inArray As Variant) As Variant
    Dim vntResult() As Variant
    ReDim vntResult(UBound(inArray, 1) - LBound(inArray, 1), _
        UBound(inArray, 2) - LBound(inArray, 2))
    CopyArrayDimensions = vntResult
End Function

Private Function CalculateSMA(inArray As Variant, nPeriod As Integer) As Variant

    Dim vntResult() As Variant
    vntResult = CopyArrayDimensions(inArray)
     
    Dim i As Integer
    For i = LBound(inArray, 2) To UBound(inArray, 2)
        Dim j As Integer
        ' First nPeriod array elements intentionally left empty.
        For j = LBound(inArray, 1) + (nPeriod - 1) To UBound(inArray, 1)
          Dim startPnt As Integer
          startPnt = j - (nPeriod - 1)
          Dim avg As Double
          avg = 0
          Dim k As Integer
          For k = 0 To nPeriod - 1
              avg = avg + inArray(startPnt + k, i)
          Next k
          avg = avg / nPeriod
          vntResult(j - LBound(inArray, 1), i - LBound(inArray, 2)) = avg
        Next j
    Next i
    CalculateSMA = vntResult
End Function

'Simple Moving Average
Public Function SMA(inRange As Range, nPeriod As Integer) As Variant
    SMA = CalculateSMA(inRange.Value, nPeriod)
End Function

Private Function CalculateEMA(inArray As Variant, nPeriod As Integer) As Variant
    
    Dim alpha As Double
    alpha = 2 / (nPeriod + 1)
    
    Dim vntResult() As Variant
    vntResult = CopyArrayDimensions(inArray)
    
    Dim i As Integer
    For i = LBound(inArray, 2) To UBound(inArray, 2)
        vntResult(nPeriod - LBound(inArray, 1), i - LBound(inArray, 2)) = inArray(nPeriod, i)
        
        Dim j As Integer
        ' First nPeriod array elements intentionally left empty.
        For j = LBound(inArray, 1) + (nPeriod) To UBound(inArray, 1)
          vntResult(j - LBound(inArray, 1), i - LBound(inArray, 2)) = _
            inArray(j, i) * alpha + _
            vntResult((j - 1) - LBound(inArray, 1), i - LBound(inArray, 2)) * (1 - alpha)
        Next j
    Next i
    CalculateEMA = vntResult
End Function
'Exponential Moving Average
Public Function EMA(inRange As Range, nPeriod As Integer)
    EMA = CalculateEMA(inRange.Value, nPeriod)
End Function

Private Function CalculateMACD(inArray As Variant, nShortPeriod As Integer, nLongPeriod As Integer, nSignalPeriod As Integer)
    Dim shortEma() As Variant
    Dim longEma() As Variant
    
    shortEma = CalculateEMA(inArray, nShortPeriod)
    longEma = CalculateEMA(inArray, nLongPeriod)
    
    Dim macdResult() As Variant
    macdResult = CopyArrayDimensions(shortEma)
    
    Dim i As Integer
    For i = LBound(shortEma, 2) To UBound(shortEma, 2)
        Dim j As Integer
        For j = WorksheetFunction.Max(nLongPeriod - 1, nShortPeriod - 1) + LBound(shortEma, 1) To UBound(shortEma, 1)
            macdResult(j - LBound(shortEma, 1), i - LBound(shortEma, 2)) = shortEma(j, i) - longEma(j, i)
        Next j
    Next i
    
    CalculateMACD = macdResult
    
End Function

Public Function MACD(inRange As Range, nShortPeriod As Integer, nLongPeriod As Integer, nSignalPeriod As Integer)
    Dim calcMacd() As Variant
    calcMacd = CalculateMACD(inRange.Value, nShortPeriod, nLongPeriod, nSignalPeriod)
    
    Dim calcSig() As Variant
    calcSig = CalculateSMA(calcMacd, nSignalPeriod)
  
    Dim macdCols As Integer
    macdCols = UBound(calcMacd, 2) + 1
    ReDim Preserve calcMacd(UBound(calcMacd, 1) - LBound(calcMacd, 1), (macdCols + UBound(calcSig, 2)) - LBound(calcSig, 2))
    
    Dim i As Integer
    For i = LBound(calcSig, 2) To UBound(calcSig, 2)
        Dim j As Integer
        
        For j = LBound(calcSig, 1) To UBound(calcSig, 1)
            calcMacd(j, i + macdCols) = calcSig(j, i)
        Next j
    Next i
    MACD = calcMacd
End Function

'Bollinger Bands
Public Function BOLLINGER(inRange As Range, nPeriod As Integer, nK As Double) As Variant
    Dim calculatedSma() As Variant
    Dim inRangeValue() As Variant
    inRangeValue = inRange.Value
    
    calculatedSma = CalculateSMA(inRangeValue, nPeriod)
    
    Dim calculatedStdDev() As Variant
    calculatedStdDev = CopyArrayDimensions(calculatedSma)
    
    Dim i As Integer
    For i = LBound(calculatedSma, 1) + nPeriod To UBound(calculatedSma, 1)
        Dim j As Integer
        Dim sumDist As Variant
        sumDist = 0
        For j = 0 To nPeriod
            Dim dist As Variant
            dist = inRangeValue((i - j) + LBound(inRangeValue, 1), LBound(inRangeValue, 2)) - calculatedSma(i, LBound(calculatedSma, 2))
            dist = dist * dist
            sumDist = sumDist + dist
        Next j
        sumDist = sumDist / nPeriod
        sumDist = Sqr(sumDist)
        calculatedStdDev(i, LBound(calculatedStdDev, 2)) = sumDist
    Next i
    
    '4 = 4 extra columns: Upper band, Lower band, %b and BandWidth
    ReDim Preserve calculatedSma(UBound(calculatedSma, 1) - LBound(calculatedSma, 1), (4 + UBound(calculatedStdDev, 2)) - LBound(calculatedStdDev, 2))
    
    For j = LBound(calculatedStdDev, 1) To UBound(calculatedStdDev, 1)
        ' Upper Band
        calculatedSma(j, 1) = calculatedSma(j, 0) + calculatedStdDev(j, 0) * nK
        ' Lower Band
        calculatedSma(j, 2) = calculatedSma(j, 0) - calculatedStdDev(j, 0) * nK
        '%b
        Dim bandRange As Variant
        bandRange = calculatedSma(j, 1) - calculatedSma(j, 2) ' Upper - Lower
        If (bandRange <> 0) Then
            calculatedSma(j, 3) = (inRangeValue(j + LBound(inRangeValue, 1), LBound(inRangeValue, 2)) - calculatedSma(j, 2)) / bandRange
        Else
            calculatedSma(j, 3) = 0
        End If
        'BandWidth
        If (calculatedSma(j, 0) <> 0) Then
            calculatedSma(j, 4) = bandRange / calculatedSma(j, 0)
        Else
            calculatedSma(j, 4) = 0
        End If
    Next j
    
    BOLLINGER = calculatedSma
End Function

'Relative Strength Index
Public Function RSI(inRange As Range, nPeriod As Integer) As Variant

    Dim avgGain() As Variant
    Dim avgLoss() As Variant
    Dim inRangeValue() As Variant

    inRangeValue = inRange.Value
    avgGain = CopyArrayDimensions(inRangeValue)
    avgLoss = CopyArrayDimensions(inRangeValue)
     
    Dim row As Integer
    For row = LBound(avgGain, 1) + (nPeriod) To UBound(avgGain, 1)
        
        Dim i As Integer
        Dim thisAvgGain As Variant
        Dim thisAvgLoss As Variant
        thisAvgGain = 0
        thisAvgLoss = 0
        For i = 0 To nPeriod - 1
            Dim change As Variant
            change = inRangeValue(LBound(inRangeValue, 1) + (row - i), LBound(inRangeValue, 2)) - inRangeValue(LBound(inRangeValue, 1) + (row - (i + 1)), LBound(inRangeValue, 2))
            
            If (change > 0) Then
                thisAvgGain = thisAvgGain + change
            ElseIf (change < 0) Then
                thisAvgLoss = thisAvgLoss - change  'avgLoss should be +ve, hence subtraction of negative value
            End If
            ' if change == 0 then both gain and loss = 0
            
        Next i
        
        thisAvgGain = thisAvgGain / nPeriod
        thisAvgLoss = thisAvgLoss / nPeriod
        
        avgGain(row, LBound(avgGain, 2)) = thisAvgGain
        avgLoss(row, LBound(avgLoss, 2)) = thisAvgLoss
    Next row
    
    Dim emaGain() As Variant
    Dim emaLoss() As Variant
    emaGain = CopyArrayDimensions(avgGain)
    emaLoss = CopyArrayDimensions(avgLoss)
    emaGain = CalculateEMA(avgGain, nPeriod)
    emaLoss = CalculateEMA(avgLoss, nPeriod)
    
    Dim resultArray() As Variant
    resultArray = CopyArrayDimensions(emaGain)
    
    For i = LBound(resultArray, 1) + nPeriod + 1 To UBound(resultArray, 1)
        If (emaLoss(i, LBound(emaLoss, 2)) = 0) Then
            resultArray(i, LBound(resultArray, 2)) = 100
        Else
            resultArray(i, LBound(resultArray, 2)) = 100 - (100 / (1 + (emaGain(i, LBound(emaGain, 2)) / emaLoss(i, LBound(emaLoss, 2)))))
        End If
    Next i
    
    RSI = resultArray
End Function


