Public Class clsSalesperson
    Private Property strFirstName As String
    Private Property strLastName As String
    Private Property intOrderID As Integer
    Private Property intID As Integer
    Private Property sngGamesSales As Single
    Private Property intGamesQuantity As Integer
    Private Property sngDollsSales As Single
    Private Property intDollsQuantity As Integer
    Private Property sngBuildingSales As Single
    Private Property intBuildingQuantity As Integer
    Private Property sngModelSales As Single
    Private Property intModelQuantity As Single
    'Private Property sngTotalSales As Single

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
