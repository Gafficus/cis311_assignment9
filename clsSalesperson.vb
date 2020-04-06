Public Class clsSalesperson
    Private strFirstName As String
    Private strLastName As String
    Private intOrderID As Integer
    Private intID As Integer
    Private sngGamesSales As Single
    Private intGamesQuantity As Integer
    Private sngDollsSales As Single
    Private intDollsQuantity As Integer
    Private sngBuildingSales As Single
    Private intBuildingQuantity As Integer
    Private sngModelSales As Single
    Private intModelQuantity As Single
    'Private Property sngTotalSales As Single
    Public Function getStrFirstName()
        Return strFirstName
    End Function
    Public Function getStrLastName()
        Return strLastName
    End Function
    Public Function getIntOrderID()
        Return intOrderID
    End Function
    Public Function getIntID()
        Return intID
    End Function
    Public Function getSngGamesSales()
        Return sngGamesSales
    End Function
    Public Function getIntGamesQuantity()
        Return intGamesQuantity
    End Function
    Public Function getSngDollsSales()
        Return sngDollsSales
    End Function
    Public Function getIntDollsQuantity()
        Return intDollsQuantity
    End Function
    Public Function getSngBuildingSales()
        Return sngBuildingSales
    End Function
    Public Function getIntBuildingQuantity()
        Return intBuildingQuantity
    End Function
    Public Function getSngModelSales()
        Return sngModelSales
    End Function
    Public Function getIntModelQuantity()
        Return intModelQuantity
    End Function
    '--- End Getters

    Public Sub setStrFirstName(ByVal inValue)
        strFirstName = inValue
    End Sub
    Public Sub setStrLastName(ByVal inValue)
        strLastName = inValue
    End Sub
    Public Sub setIntOrderID(ByVal inValue)
        intOrderID = inValue
    End Sub
    Public Sub setIntID(ByVal inValue)
        intID = inValue
    End Sub
    Public Sub setSngGamesSales(ByVal inValue)
        sngGamesSales = inValue
    End Sub
    Public Sub setIntGamesQuantity(ByVal inValue)
        intGamesQuantity = inValue
    End Sub
    Public Sub setSngDollsSales(ByVal inValue)
        sngDollsSales = inValue
    End Sub
    Public Sub setIntDollsQuantity(ByVal inValue)
        intDollsQuantity = inValue
    End Sub
    Public Sub setSngBuildingSales(ByVal inValue)
        sngBuildingSales = inValue
    End Sub
    Public Sub setIntBuildingQuantity(ByVal inValue)
        intBuildingQuantity = inValue
    End Sub
    Public Sub setSngModelSales(ByVal inValue)
        sngModelSales = inValue
    End Sub
    Public Sub setIntModelQuantity(ByVal inValue)
        intModelQuantity = inValue
    End Sub
    Public Sub New(strFirstName As String, strLastName As String, intOrderID As Integer,
                   intID As Integer, sngGamesSales As Single, intGamesQuantity As Integer,
                   sngDollsSales As Single, intDollsQuantity As Integer, sngBuildingSales As Single,
                   intBuildingQuantity As Integer, sngModelSales As Single, intModelQuantity As Single)
        If strFirstName Is Nothing Then
            Throw New ArgumentNullException(NameOf(strFirstName))
        End If

        If strLastName Is Nothing Then
            Throw New ArgumentNullException(NameOf(strLastName))
        End If

        Me.strFirstName = strFirstName
        Me.strLastName = strLastName
        Me.intOrderID = intOrderID
        Me.intID = intID
        Me.sngGamesSales = sngGamesSales
        Me.intGamesQuantity = intGamesQuantity
        Me.sngDollsSales = sngDollsSales
        Me.intDollsQuantity = intDollsQuantity
        Me.sngBuildingSales = sngBuildingSales
        Me.intBuildingQuantity = intBuildingQuantity
        Me.sngModelSales = sngModelSales
        Me.intModelQuantity = intModelQuantity
        'Me.sngTotalSales = sngTotalSales
    End Sub

    Public Overrides Function ToString() As String
        Return strFirstName & " " & strLastName & " " &
            intOrderID.ToString() & " " & intID & " "
    End Function
End Class
