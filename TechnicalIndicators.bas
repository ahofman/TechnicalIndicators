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
    ReDim vntResult(UBound(inArray, 1), UBound(inArray, 2))
     
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
