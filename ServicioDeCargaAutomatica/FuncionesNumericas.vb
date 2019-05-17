Module FuncionesNumericas

    ''' <summary>
    ''' Agrega 000 a un numero para que sea de tre caracteres su conversión
    ''' Por ejemplo
    ''' 0   ->  000
    ''' 1   ->  001
    ''' 12  ->  012
    ''' 918 ->  918
    ''' </summary>
    ''' <param name="Numero"></param>
    ''' <returns></returns>
    Function CompletaNumero(ByVal Numero As Integer)
        Dim Completado As String
        If Numero <= 0 Then
            Completado = "000"
        ElseIf Numero < 10 Then
            Completado = "00" + Numero.ToString
        ElseIf Numero < 100 Then
            Completado = "0" + Numero.ToString
        Else
            Completado = Numero.ToString
        End If
        Return Completado
    End Function

    ''' <summary>
    ''' Redondea un decimal a la cantidad de decimales especificados
    ''' </summary>
    ''' <param name="NumeroARedondear"></param>
    ''' <param name="NumeroDeDecimales"></param>
    ''' <returns></returns>
    Function RedondeaDecimal(ByVal NumeroARedondear As Decimal, ByVal NumeroDeDecimales As Integer)
        Dim Redondeado As Decimal = 0
        Dim ConverNume As Decimal = Decimal.Parse(10 ^ NumeroDeDecimales)
        Redondeado = Decimal.Parse(Integer.Parse(NumeroARedondear * ConverNume)) * 1 / ConverNume
        Return Redondeado
    End Function

    Function Text2DecimalRounded(ByVal NumeroARedondear As String, ByVal NumeroDeDecimales As Integer)
        Dim MyDecimal As Decimal = 0
        Dim Retorno As String = ""
        If Decimal.TryParse(NumeroARedondear, MyDecimal) Then
            Retorno = RedondeaDecimal(MyDecimal, NumeroDeDecimales).ToString
        End If
        Return Retorno
    End Function

End Module