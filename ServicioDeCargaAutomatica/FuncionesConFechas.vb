Imports System.Windows.Forms
Imports System.Windows.Forms.Integration
Module FuncionesConFechas

    ''' <summary>
    ''' Permite pasar de un string de fecha a uno de consulta SQL
    ''' </summary>
    ''' <param name="MyS"></param>
    ''' <returns></returns>
    Public Function DateString2stringSQL(ByRef MyS As String)
        Dim MyTime As DateTime
        Dim Str = ""
        If (DateTime.TryParse(MyS.Trim, MyTime)) Then
            Str = MyTime.Year.ToString + "-" + Num2strCero(MyTime.Month) + "-" + Num2strCero(MyTime.Day.ToString) _
                + " " + Num2strCero(MyTime.Hour.ToString) + ":" + Num2strCero(MyTime.Minute.ToString) + ":" + Num2strCero(MyTime.Second.ToString) + "." + MyTime.Millisecond.ToString
            Str = Str.Trim
        Else
            MsgBox("Error al convertir la fecha",, "Convertidor de fecha")
        End If
        Return Str
    End Function

    ''' <summary>
    ''' Permite pasa de una variable DateTime a una variable de busqueda SQL
    ''' </summary>
    ''' <param name="MyTime"></param>
    ''' <returns></returns>
    Public Function Date2stringSQL(ByRef MyTime As DateTime)
        Return MyTime.Year.ToString + "-" + Num2strCero(MyTime.Month) + "-" + Num2strCero(MyTime.Day.ToString) _
            + " " + Num2strCero(MyTime.Hour.ToString) + ":" + Num2strCero(MyTime.Minute.ToString) + ":" + Num2strCero(MyTime.Second.ToString) + ".000" '+ MyTime.Millisecond.ToString
    End Function

    ''' <summary>
    ''' Permite pasa de una variable DateTime a una variable de busqueda SQL, con únicamente la Date
    ''' </summary>
    ''' <param name="MyTime"></param>
    ''' <returns></returns>
    Public Function DateOnlyDate2stringSQL(ByRef MyTime As DateTime)
        Return MyTime.Year.ToString + "-" + Num2strCero(MyTime.Month) + "-" + Num2strCero(MyTime.Day.ToString)
    End Function

    ''' <summary>
    ''' Permite pasa de una variable DateTime a una variable de busqueda SQL, con únicamente la Date
    ''' </summary>
    ''' <param name="MyTime"></param>
    ''' <returns></returns>
    Public Function DateOnlyTime2stringSQL(ByRef MyTime As DateTime)
        Return Num2strCero(MyTime.Hour.ToString) + ":" + Num2strCero(MyTime.Minute.ToString) + ":" + Num2strCero(MyTime.Second.ToString) + ".000" '+ MyTime.Millisecond.ToString
    End Function

    ''' <summary>
    ''' Permite agregar el 0 faltante a los números que son menores a 10
    ''' </summary>
    ''' <param name="Num"></param>
    ''' <returns></returns>
    Function Num2strCero(ByVal Num As Integer)
        If (Num > -1 AndAlso Num < 10) Then
            Return "0" + Num.ToString
        Else
            Return Num.ToString
        End If
    End Function

    ''' <summary>
    ''' Permite pasar de una variable Date time a un string que tiene yyyy-MM-dd
    ''' </summary>
    ''' <param name="MyTime"></param>
    ''' <returns></returns>
    Function Date2strFecha(ByRef MyTime As DateTime)
        Return MyTime.Year.ToString + "-" + Num2strCero(MyTime.Month) + "-" + Num2strCero(MyTime.Day.ToString)
    End Function

    ''' <summary>
    ''' Permite pasar de una variable Date time a un string que tiene hh-mm-ss.uuu
    ''' Horas, minutos, segundos y milisegundos
    ''' </summary>
    ''' <param name="MyTime"></param>
    ''' <returns></returns>
    Function Date2strHMS(ByRef MyTime As DateTime)
        Return Num2strCero(MyTime.Hour.ToString) + ":" + Num2strCero(MyTime.Minute.ToString) + ":" + Num2strCero(MyTime.Second.ToString) + "." + MyTime.Millisecond.ToString
    End Function

    ''' <summary>
    ''' Tomando los parametros devuelve el tiempo como dateTime
    ''' </summary>
    ''' <param name="año"></param>
    ''' <param name="mes"></param>
    ''' <param name="dia"></param>
    ''' <param name="hora"></param>
    ''' <param name="minuto"></param>
    ''' <param name="segundo"></param>
    ''' <returns></returns>
    Function Num2date(ByVal año As Integer, ByVal mes As Integer, ByVal dia As Integer, ByVal hora As Integer, ByVal minuto As Integer, ByVal segundo As Integer)
        Dim MyTime As DateTime
        Dim Str = año.ToString + "-" + Num2strCero(mes) + "-" + Num2strCero(dia) _
            + " " + Num2strCero(hora) + ":" + Num2strCero(minuto) + ":" + Num2strCero(segundo) + "." + "0" 'Now.Millisecond.ToString
        If (Not DateTime.TryParse(Str, MyTime)) Then
            MsgBox("Error al convertir",, "Módulo num2date")
        End If
        Return MyTime
    End Function

    ''' <summary>
    ''' Permite pasar de Date time a una lista con los parametros
    ''' Año, Mes, Día, Hora, Minuto, Segundo
    ''' </summary>
    ''' <param name="MyTime"></param>
    ''' <returns></returns>
    Function Date2List(ByVal MyTime As DateTime)
        Try
            Dim listI As New List(Of Integer)
            listI.Add(MyTime.Year)
            listI.Add(MyTime.Month)
            listI.Add(MyTime.Day)
            listI.Add(MyTime.Hour)
            listI.Add(MyTime.Minute)
            listI.Add(MyTime.Second)
            Return listI
        Catch er As System.Exception
            MsgBox("Error al pasar a una lista",, "Módulo date2List")
            Return ListDeO(6)
        End Try

    End Function

    ''' <summary>
    ''' Pasa de una lista a una DateTime
    ''' </summary>
    ''' <param name="LisDate"></param>
    ''' <returns></returns>
    Function List2date(ByVal LisDate As List(Of Integer))
        Return Num2date(LisDate(0), LisDate(1), LisDate(2), LisDate(3), LisDate(4), LisDate(5))
    End Function

    ''' <summary>
    ''' Combinando los valores
    ''' </summary>
    ''' <param name="Thoras"></param>
    ''' <param name="Tmin"></param>
    ''' <param name="FechaYMD"></param>
    ''' <returns></returns>
    Function DateMixComboPicker(ByVal Thoras As String, ByVal Tmin As String, ByVal FechaYMD As DateTime)
        Dim Lista As New List(Of Integer)
        Dim Horas = 0
        Dim Min = 0
        Dim Fecha As DateTime = Now
        'FechaYMD = Now
        Dim Cob0 = Integer.TryParse(Thoras, Horas)
        Dim Cob1 = Integer.TryParse(Tmin, Min)
        If (Cob0 AndAlso Cob1) Then
            Try
                Fecha = Num2date(FechaYMD.Year, FechaYMD.Month, FechaYMD.Day, Horas, Min, Now.Second())
            Catch ex As System.Exception
                MsgBox("Error en la conversión",, "Módulo DateMixComboPicker")
            End Try
        Else
            MsgBox("Error en la conversión, en la conversión explicita" + Cob0.ToString + Cob1.ToString,, "Módulo DateMixComboPicker")
        End If
        Return Fecha
    End Function

    ''' <summary>
    ''' Devulve true si el texto ingresado como numerico, se encuentra entre un rango especifico
    ''' </summary>
    ''' <param name="tex"></param>
    ''' <param name="Lmin"></param>
    ''' <param name="lmax"></param>
    ''' <returns></returns>
    Function BuscaEntreTextInNum(ByVal tex As String, ByVal Lmin As Integer, ByVal lmax As Integer)
        Dim numero As Integer
        If (Integer.TryParse(tex, numero)) Then
            If (numero >= Lmin AndAlso numero <= lmax) Then
                Return True
            Else
                Return False
            End If
        Else
            MsgBox("Por favor checa los campos de entrada",, "Módulo buscaEntreTextInNum")
            Return False
        End If
    End Function

    ''' <summary>
    ''' Permite referenciar los combos a unos valores, donde los indices corresponden a un valor numerico.
    ''' </summary>
    ''' <param name="MyCombo"></param>
    ''' <param name="Lmini"></param>
    ''' <param name="Lsuperior"></param>
    ''' <param name="Remplazo"></param>
    ''' <returns></returns>
    Function ChecaCombo(ByRef MyCombo As ComboBox, ByVal Lmini As Integer, ByVal Lsuperior As Integer, ByVal Remplazo As Integer)
        If (Not MyCombo.Text = "") Then
            If (Not BuscaEntreTextInNum(MyCombo.Text, Lmini, Lsuperior)) Then
                MyCombo.SelectedIndex = Remplazo
                Return False
            Else
                Return True
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Permite realizar la operacion de FechaF-FechaI
    ''' Si el último valor es positivo indica que el resultado es positivo o -1 si el resultado era negativo
    ''' </summary>
    ''' <param name="FechaI"></param>
    ''' <param name="FechaF"></param>
    ''' <returns></returns>
    Function RestaFechas(ByVal FechaF As List(Of Integer), ByVal FechaI As List(Of Integer))
        Dim Resultado = ListDeO(FechaI.Count + 1)
        Dim IndexMayor As Integer = 0
        Dim EsNecesarioC As Boolean = False
        Try
            Dim MyIndex As Integer = FechaI.Count - 1
            While MyIndex >= 0
                Resultado(MyIndex) = FechaF(MyIndex) - FechaI(MyIndex)
                If (Resultado(MyIndex) <> 0) Then
                    IndexMayor = MyIndex
                End If
                If (Resultado(MyIndex) < 0) Then
                    EsNecesarioC = True
                End If
                MyIndex = MyIndex - 1
            End While
            '-----Interpretando el resultado
            'If (EsNecesarioC) Then
            If (Resultado(IndexMayor) > 0) Then
                'Calling the function
                Resultado(6) = 1
            Else
                Resultado(IndexMayor) = Math.Abs(Resultado(IndexMayor))
                Resultado(6) = -1
                'Calling the function
            End If
            'Here we need to call to that function
            'MsgBox(LiF.MyList2Text(Resultado, 0, 5))
            Resultado = CorrigeFehcha(Resultado)
            If (Not EstaVacia(Resultado)) Then
                Return Resultado
            Else
                MsgBox("Error al corregir la fecha",, "Comprobación fase II en el módulo RestaFechas")
            End If
            'Else
            ' MsgBox("Hurra")
            'End If
        Catch er As System.Exception
            MsgBox(er.Message,, "Módulo RestaFechas")
        End Try
        Return ListDeO(FechaI.Count + 1)
    End Function

    ''' <summary>
    ''' Permite corregir los signos negativos obtenidos durante la operación.
    ''' </summary>
    ''' <param name="ResuFecha"></param>
    ''' <returns></returns>
    Function CorrigeFehcha(ByVal ResuFecha As List(Of Integer))
        Dim Corregida As List(Of Integer)
        Corregida = ResuFecha
        Try
            ' 0-> Años.
            ' 1-> Meses.
            ' 2-> Dias.
            ' 3-> Horas.
            ' 4-> Minutos.
            ' 5-> Segundos.
            ' 6-> Valor de dirección.
            If (ResuFecha.Capacity = 8) Then
                If (Corregida(0) < 0) Then
                    MsgBox("Que extraño")
                End If
                If (Corregida(1) < 0) Then  '-------------------------------------Corrigiendo meses
                    Corregida(0) = Corregida(0) - 1     'Entonces bajamos un año
                    Corregida(1) = Corregida(1) + 12        ' Y sumamos 12 meses
                End If
                If (Corregida(2) < 0) Then  '------------------------------------Corrigiendo días
                    'Fase de corrección
                    Corregida(1) = Corregida(1) - 1     'Entonces me bajo un mes
                    If (Corregida(1) = 0) Then              'No existe el mes 0
                        Corregida(0) = Corregida(0) - 1         'Entonces bajamos un año
                        Corregida(1) = 12                           'Y nos quedamos en el mes 12
                    End If
                    'Aplicando
                    Corregida(2) = Corregida(2) + DateTime.DaysInMonth(Now.Year, Now.Month)  'Ahora si sumamos los días

                End If
                If (Corregida(3) < 0) Then  '------------------------------------Corrigiendo las horas
                    Corregida(2) = Corregida(2) - 1 'Entonces nos bajamos un día
                    If (Corregida(2) = 0) Then          'No existe el día cero
                        Corregida(1) = Corregida(1) - 1     'Entonces me bajo un mes
                        If (Corregida(1) = 0) Then              'No existe el mes cero
                            Corregida(0) = Corregida(0) - 1         'Entonces me bajo un año
                            Corregida(1) = 12                            'Y nos quedamos en el mes 12
                        End If
                        Corregida(2) = DateTime.DaysInMonth(Now.Year, Now.Month)    'Nos quedamos en el último del mes
                    End If
                    'Aplicando
                    Corregida(3) = Corregida(3) + 24                                                'Y añadimos las 24 horas
                End If

                If (Corregida(4) < 0) Then  '------------------------------------Corrigiendo los minutos
                    Corregida(3) = Corregida(3) - 1 'Entonces nos bajamos una hora
                    If (Corregida(3) < 0) Then          'Como no hay horas negativas
                        Corregida(2) = Corregida(2) - 1     'Entonces nos bajamos un día
                        If (Corregida(2) = 0) Then              'No existe el día cero
                            Corregida(1) = Corregida(1) - 1         'Entonces me bajo un mes
                            If (Corregida(1) = 0) Then                  'No existe el mes cero
                                Corregida(0) = Corregida(0) - 1             'Entonces me bajo un año
                                Corregida(1) = 12                               'Y nos quedamos en el mes 12
                            End If
                            Corregida(2) = DateTime.DaysInMonth(Now.Year, Now.Month)     'Nos quedamos en el último del mes
                        End If
                        Corregida(3) = Corregida(3) + 24                                                'Y añadimos las 24 horas
                    End If
                    'Aplicando
                    Corregida(4) = Corregida(4) + 60                                                        'Y añadimos 60 minutos.
                End If
                If (Corregida(5) < 0) Then
                    Corregida(4) = Corregida(4) - 1
                    If (Corregida(4) < 0) Then  '------------------------------------Corrigiendo los minutos
                        Corregida(3) = Corregida(3) - 1 'Entonces nos bajamos una hora
                        If (Corregida(3) < 0) Then          'Como no hay horas negativas
                            Corregida(2) = Corregida(2) - 1     'Entonces nos bajamos un día
                            If (Corregida(2) = 0) Then              'No existe el día cero
                                Corregida(1) = Corregida(1) - 1         'Entonces me bajo un mes
                                If (Corregida(1) = 0) Then                  'No existe el mes cero
                                    Corregida(0) = Corregida(0) - 1             'Entonces me bajo un año
                                    Corregida(1) = 12                               'Y nos quedamos en el mes 12
                                End If
                                Corregida(2) = DateTime.DaysInMonth(Now.Year, Now.Month)     'Nos quedamos en el último del mes
                            End If
                            Corregida(3) = Corregida(3) + 24                                                'Y añadimos las 24 horas
                        End If
                        'Aplicando
                        Corregida(4) = Corregida(4) + 60                                                        'Y añadimos 60 minutos.
                    End If
                    Corregida(5) = Corregida(5) + 60                                                                    'Y añadimos 60 segundos
                End If
                Return Corregida
            End If
        Catch er As System.Exception
            MsgBox(er.Message,, "Módulo corrigue fecha")
        End Try
        Return ListDeO(7)
    End Function

    ''' <summary>
    ''' Permite crear una lista de integers que simplemente esta llena de ceros según el número especificado.
    ''' </summary>
    ''' <param name="NumCeros"></param>
    ''' <returns></returns>
    Function ListDeO(ByVal NumCeros As Integer)
        Dim Dev As New List(Of Integer)
        Dim NuRep As Integer = 0
        While NuRep < NumCeros
            NuRep = NuRep + 1
            Dev.Add(0)
        End While
        Return Dev
    End Function

    Function EstaVacia(ByVal Ls As List(Of Integer))
        Dim EstaVacio As Boolean = True
        For Each Entero As Integer In Ls
            If (Entero <> 0) Then
                EstaVacio = False
            End If
        Next
        Return EstaVacio
    End Function

    Function res2Format(ByVal res As List(Of Integer))
        Return res(0).ToString + "-" + Num2strCero(res(1).ToString) + "-" + Num2strCero(res(2).ToString) _
           + " " + Num2strCero(res(3).ToString) + ":" + Num2strCero(res(4).ToString) + ":" + Num2strCero(res(5).ToString) + "." + "0" 'Now.Millisecond.ToString
    End Function
    Function Seconds2Milis(ByVal Seconds As Integer)
        Return Seconds * 1000
    End Function
    Function Minutes2Milis(ByVal Minutes As Integer)
        Return Seconds2Milis(Minutes * 60)
    End Function
    Function Hour2Miles(ByVal Horas As Integer)
        Return Minutes2Milis(Horas * 60)
    End Function
End Module