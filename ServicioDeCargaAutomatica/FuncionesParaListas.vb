Module FuncionesParaListas

    'Function List2List(ByVal MyList As List(Of String), ByVal ToMyList As List(Of String), ByVal MyStr As String)
    '    Dim MyError As Boolean = True
    '    If MyList.Count = ToMyList.Count Then
    '        For Index As Integer = 0 To MyList.Count - 1
    '            If String.Compare(MyList(Index), MyStr) Then
    '                Return ToMyList(Index)
    '                Exit Function
    '            End If
    '        Next
    '    End If
    '    Return "e" 'e de error
    'End Function
#If Not DEBUG Then
    ''' <summary>
    ''' Permite pasar de una lista de datos a una checked list box
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <param name="MyCheckedL"></param>
    Sub List2CheckedList(ByVal MyList As List(Of String), ByVal MyStart As Integer, ByVal MyEnd As Integer, ByRef MyCheckedL As CheckedListBox)
        If (MyEnd - MyStart) + 1 = MyCheckedL.Items.Count Then
            Dim Index As Integer
            For Index = 0 To MyCheckedL.Items.Count - 1
                If MyList(Index + MyStart) = True.ToString Then
                    MyCheckedL.SetItemChecked(Index, True)
                ElseIf MyList(Index + MyStart) = False.ToString() Then
                    MyCheckedL.SetItemChecked(Index, False)
                Else
                    MsgBox("Error favor de notificar a sistemas",, "List2checkedList error")
                End If
            Next
        End If
    End Sub
#End If
#If Not DEBUG Then
    ''' <summary>
    ''' Permite obtener una lista de Strings con los valores del CheckedListBox
    ''' </summary>
    ''' <param name="MyCheckList"></param>
    ''' <returns>MyList that is a list</returns>
    Function CheckedList2List(ByRef MyCheckList As CheckedListBox)
        Dim MyList As List(Of String) = New List(Of String)
        MyList.Capacity = MyCheckList.Items.Count
        Dim Index As Integer = 0
        For Index = 0 To MyCheckList.Items.Count - 1
            'MyList.Add(MyCheckList.GetItemCheckState(Index).ToString)
            If String.Compare("Checked", MyCheckList.GetItemCheckState(Index).ToString) = 0 Then
                MyList.Add(True.ToString)
            ElseIf String.Compare("Unchecked", MyCheckList.GetItemCheckState(Index).ToString) = 0 Then
                MyList.Add(False.ToString)
            Else
                MsgBox("Error en CheckedList2List")
            End If
        Next
        Return MyList
    End Function
#End If

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

End Module