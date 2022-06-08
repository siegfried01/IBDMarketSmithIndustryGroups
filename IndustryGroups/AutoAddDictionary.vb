
Public Class AutoAddDictionary(Of Key As IComparable(Of Key), Value As {New})
    Inherits Dictionary(Of Key, Value)
    Default Public Overloads Property Item(key As Key) As Value
        Get
            If MyBase.Keys.Contains(key) Then
                Return MyBase.Item(key)
            Else
                MyBase.Item(key) = New Value
                Return MyBase.Item(key)
            End If
        End Get
        Set(value As Value)
            If Not MyBase.ContainsKey(key) Then
                MyBase.Add(key, value)
            Else
                MyBase.Item(key) = value
            End If
        End Set
    End Property
End Class
Public Class AutoAddMinMaxDictionary(Of Key As IComparable(Of Key))
    Inherits AutoAddDictionary(Of Key, Int32)

    Public minValue As Int32 = Int32.MaxValue
    Public maxValue As Int32 = Int32.MinValue

    Default Public Overloads Property Item(key As Key) As Int32
        Get
            If MyBase.Keys.Contains(key) Then
                Return MyBase.Item(key)
            Else
                MyBase.Item(key) = New Int32
                Return MyBase.Item(key)
            End If
        End Get
        Set(value As Int32)
            If value > maxValue Then
                maxValue = value
            End If
            If value < minValue Then
                minValue = value
            End If
            If Not MyBase.ContainsKey(key) Then
                MyBase.Add(key, value)
            Else
                MyBase.Item(key) = value
            End If
        End Set
    End Property

End Class