Module FuncionesParaListas
    ''' <summary>
    ''' Permite añadir un lista de strins a otra lista
    ''' </summary>
    ''' <param name="ToAdd"></param>
    ''' <param name="MyList"></param>
    Sub AddList2List(ByVal ToAdd As List(Of String), ByRef MyList As List(Of String))
        For Each str As String In ToAdd
            MyList.Add(str)
        Next
    End Sub

    ''' <summary>
    ''' Función para los cortes de listas de Strings en base a un index
    ''' inicial y uno final.
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <param name="Pini"></param>
    ''' <param name="Pfin"></param>
    ''' <returns></returns>
    Function Cut2List(ByVal MyList As List(Of String), ByVal Pini As Integer, Pfin As Integer)
        Dim MyListC As List(Of String) = New List(Of String)
        For Index As Integer = Pini To Pfin
            MyListC.Add(MyList(Index))
        Next
        Return MyListC
    End Function

    ''' <summary>
    ''' Remplaza el caracter especificado por un espacio, y si se encontraba al incio o al final
    ''' desaparecera, como
    ''' " 'Comida' " ->"Comida"
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <param name="MyChar"></param>
    ''' <returns></returns>
    Function KillChar(ByRef MyList As List(Of String), ByVal MyChar As Char)
        Dim MyRet = New List(Of String)
        For Each str As String In MyList
            MyRet.Add(str.Replace(MyChar, " ").Trim)
        Next
        Return MyRet
    End Function

    ''' <summary>
    ''' Sustituye una cadena por un "" en cada elemento de la lista
    ''' "NULL Hamburgesa "-> "Hamburgesa"
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <param name="MyStr"></param>
    ''' <returns></returns>
    Function ReplaceStr(ByRef MyList As List(Of String), ByVal MyStr As String)
        Dim MyRet = New List(Of String)
        For Each str As String In MyList
            MyRet.Add(str.Replace(MyStr, "").Trim)
        Next
        Return MyRet
    End Function

    Function CreateListBooleanWithOnly(ByVal Longitud As Integer, ByVal MyValues As Boolean)
        Dim MyRet As List(Of Boolean) = New List(Of Boolean)
        For Repeticion As Integer = 0 To Longitud - 1
            MyRet.Add(MyValues)
        Next
        Return MyRet
    End Function

    ''' <summary>
    ''' Permite buscar un String en una lista y devolver el Índice de donde fue encotrado
    ''' Retorna un -1 en caso de no ser encontrado
    ''' </summary>
    ''' <param name="MyStr"></param>
    ''' <param name="MyList"></param>
    ''' <returns></returns>
    Function FindIndexStrInList(ByVal MyStr As String, ByVal MyList As List(Of String))
        Dim Index As Integer = 0
        For Each Str As String In MyList
            If String.Compare(Str, MyStr) = 0 Then
                Return Index
                Exit Function
            End If
        Next
        Return -1
    End Function
    ''' <summary>
    ''' Como se supone que vienen del menor al mayor, siendo el index 0 el menor
    ''' Esta funcion permite retornar el valor más proximo continuo
    ''' </summary>
    ''' <param name="MyActual"></param>
    ''' <param name="MyList"></param>
    ''' <returns></returns>
    Function FindTheNextNumber2Next(ByVal MyActual As Integer, ByVal MyList As List(Of Integer))
        Dim Index As Integer = 0
        Dim LongOfList As Integer = MyList.Count()
        Dim Finded As Boolean = False
        Try
            If LongOfList > 0 Then
                For Index = 0 To LongOfList - 1
                    If MyActual < MyList(Index) Then
                        Finded = True
                        Return MyList(Index)
                        Exit Function
                    End If
                Next
                If Finded = False Then
                    Return MyList(0)
                Else
                    Return MyList(Index)
                End If
            Else
                Informe("No se han logrado obtener los datos de los intervalos")
                Return -1
            End If
        Catch er As System.Exception
            Informe("Error al buscar el siguiente minuto para iniciar el servicio, datos: " + vbNewLine _
                 + "Valor buscado: " + MyActual.ToString() _
                 + List2Secciones(ListNum2Str(MyList)) + vbNewLine _
                 + "Error: " + er.Message.ToString _
                 + "Index: " + Index.ToString _
                 + "Longitud: " + LongOfList.ToString _
                 + "Finded: " + Finded.ToString)
            Return -2
        End Try
    End Function
    Function ListStr2Num(ByVal Lista As List(Of String))
        Dim MyListConverted As List(Of Integer) = New List(Of Integer)
        Dim Temporal As Integer = 0
        For Each Value As Integer In Lista
            If Integer.TryParse(Value, Temporal) Then
                MyListConverted.Add(Temporal)
            Else
                Informe("Error al convertir la lista de tiempos a valores númericos: " + vbNewLine _
                    + List2Secciones(Lista))
                MyListConverted.Clear()
                Return MyListConverted
            End If
        Next
        Return MyListConverted
    End Function
    Function ListNum2Str(ByVal Lista As List(Of Integer))
        Dim MyListConverted As List(Of String) = New List(Of String)
        For Each Value As Integer In Lista
            MyListConverted.Add(Value.ToString)
        Next
        Return MyListConverted
    End Function
    Function AreNumberOnList(ByVal MyNumber As Integer, ByVal OnList As List(Of Integer))
        For Each MyInt As Integer In OnList
            If MyInt = MyNumber Then
                Return 1
                Exit Function
            End If
        Next
        Return 0
    End Function
End Module