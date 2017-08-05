Option Strict On

Public NotInheritable Class Validation

#Region "IsNumberic メソッド (+1) "

    ''' <summary>
    ''' 文字列が数値であるかどうかを返します。
    ''' </summary>
    ''' <param name="strTarget">検査対象となる文字列。</param>
    ''' <returns>指定した文字列が数値であれば　True。それ以外は　False。</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function IsNumberic(ByVal strTarget As String) As Boolean
        Return Double.TryParse(strTarget,
                               System.Globalization.NumberStyles.Any,
                               Nothing,
                               0.0#)
    End Function

    ''' <summary>
    ''' オブジェクトが数値であるかどうかを返します。
    ''' </summary>
    ''' <param name="oTarget">検査対象となるオブジェクト</param>
    ''' <returns>指定した文字列が数値であれば　True。それ以外は　False。</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function IsNumberic(ByVal oTarget As Object) As Boolean
        Return IsNumberic(oTarget.ToString())
    End Function
#End Region

#Region " IsDate　メソッド　"

    ''' <summary>
    ''' 指定した年ー月ー日が正しい日付であるかどうかを返します。
    ''' </summary>
    ''' <param name="iYear">検査対象となる年</param>
    ''' <param name="iMonth">検査対象となる月</param>
    ''' <param name="iDay">検査対象となる日</param>
    ''' <returns>指定した年ー月ー日が正しい日付であれば　True。それ以外　False。</returns>
    ''' <remarks></remarks>
    Public Shared Function IsDate(ByVal iYear As Integer,
                                  ByVal iMonth As Integer,
                                  ByVal iDay As Integer) As Boolean
        If (Date.MinValue.Year > iYear OrElse Date.MaxValue.Year < iYear) Then
            Return False
        End If

        If (Date.MinValue.Month > iMonth OrElse iMonth > Date.MaxValue.Month) Then
            Return False
        End If

        Dim iLastDay As Integer = Date.DaysInMonth(iYear, iMonth)

        If (Date.MinValue.Day > iDay OrElse iDay > iLastDay) Then
            Return False
        End If

        Return True
    End Function
#End Region

#Region " Val　メソッド（＋２）"
    Public Overloads Shared Function Val(ByVal stTarget As String) As Double

        'Null　値の場合は　０　を返す
        If (stTarget Is Nothing) Then
            Return 0
        End If

        Dim iCurrent As Integer
        Dim iLength As Integer = stTarget.Length

        '評価対象外の文字列をスキップする。
        For iCurrent = 0 To iLength
            Select Case stTarget.Chars(iCurrent)
                Case " "c, " "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf
                Case Else
                    Exit For
            End Select
        Next

        '終端までに有効な文字が見つからなかった場合は　0 を返す
        If (iCurrent >= iLength) Then
            Return 0.0#
        End If

        Dim bMinus As Boolean

        '先頭にある符号を判断する
        Select Case stTarget.Chars(iCurrent)
            Case "-"c
                bMinus = True
                iCurrent += 1
            Case "+"c
                iCurrent += 1
        End Select

        Dim iValidLength As Integer
        Dim iPriod As Integer
        Dim dReturn As Double
        Dim bDecimal As Boolean
        Dim bShisuMark As Boolean

        '１　文字ずつ有効な文字かどうか判断する
        While (iCurrent < iLength)
            Dim chCurrent As Char = stTarget.Chars(iCurrent)
            Select Case chCurrent
                Case " "c, "　"c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf
                    iCurrent += 1
                Case "0"c
                    iCurrent += 1
                    If (iValidLength <> 0 OrElse bDecimal) Then
                        iValidLength += 1
                        dReturn = (dReturn * 10) + Double.Parse(chCurrent.ToString())
                    End If
                Case "1"c To "9"c
                    iCurrent += 1
                    iValidLength += 1
                    dReturn = (dReturn * 10) + Double.Parse(chCurrent.ToString())
                Case "."c
                    iCurrent += 1

                    If bDecimal Then
                        Exit While
                    End If

                    bDecimal = True
                    iPriod = iValidLength
                Case "e"c, "E"c, "d"c, "D"c
                    bShisuMark = True
                    iCurrent += 1
                    Exit While
                Case Else
                    Exit While
            End Select
        End While

        Dim iDecimal As Integer = 0

        '少数点が判定された場合
        If bDecimal Then
            iDecimal = iValidLength - iPriod
        End If

        '指数が判定された場合
        If bShisuMark Then
            Dim bShisuValid As Boolean
            Dim bShisuMinus As Boolean
            Dim dCoef As Double

            '指数を検証する
            While (iCurrent < iLength)
                Dim chCurrent As Char = stTarget.Chars(iCurrent)

                If (chCurrent = " "c OrElse chCurrent = "　"c _
                   OrElse chCurrent = ControlChars.Cr _
                   OrElse chCurrent = ControlChars.Lf _
                   OrElse chCurrent = ControlChars.Tab) Then
                    iCurrent += 1
                ElseIf (chCurrent >= "0"c AndAlso chCurrent <= "9"c) Then
                    dCoef = (dCoef * 10) + Double.Parse(chCurrent.ToString())
                    iCurrent += 1
                ElseIf chCurrent = "+"c Then
                    If (bShisuValid) Then
                        Exit While
                    End If

                    bShisuValid = True
                    iCurrent += 1
                ElseIf (chCurrent <> "-"c OrElse bShisuValid) Then
                    Exit While
                Else
                    bShisuValid = True
                    bShisuMinus = True
                    iCurrent += 1
                End If
            End While

            '指数の符号に応じて累乗する
            If bShisuMinus Then
                dCoef += iDecimal
                dReturn *= System.Math.Pow(10, -dCoef)
            Else
                dCoef -= iDecimal
                dReturn *= System.Math.Pow(10, dCoef)
            End If
        ElseIf (bDecimal AndAlso iDecimal <> 0) Then
            dReturn /= System.Math.Pow(10, iDecimal)
        End If

        '無限大の場合は　０　を返す。
        If (Double.IsInfinity(dReturn)) Then
            Return 0.0#
        End If

        'マイナス判定の場合はマイナスで返す
        If (bMinus) Then
            Return -dReturn
        End If

        Return dReturn
    End Function


    Public Overloads Shared Function Val(ByVal chTarget As Char) As Integer
        Return CType(Val(chTarget.ToString()), Integer)
    End Function

    Public Overloads Shared Function Val(ByVal oTarget As Object) As Double
        If (Not oTarget Is Nothing) Then
            Return Val(oTarget.ToString())
        End If
        Return 0.0#
    End Function
#End Region
End Class
