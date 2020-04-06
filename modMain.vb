Imports Microsoft.Office.Interop
Module modMain
    Dim anExcel As Excel.Application

    Dim mySalesForce As New List(Of clsSalesperson)
    Sub Main()
        Console.WriteLine("Loading Excel...")
        anExcel = New Excel.Application()
        createSalesForce()
        anExcel.Workbooks.Add()
        anExcel.Sheets.Add()
        Console.WriteLine("Writting data to Excel...")
        createHeader()
        Dim intLastRow As Integer = createMainData()
        createFormulaHeaders(intLastRow)
        createSalesForceSales(intLastRow)
        createSalesForceQuantity()
        Console.WriteLine("Opening Excel...")
        anExcel.Visible = True
        Console.WriteLine("Press Any Key To Exit.")
        Console.ReadKey()
        anExcel.Quit()
        anExcel = Nothing
    End Sub


    Private Sub createSalesForceSales(ByVal intLastRow)

    End Sub
    Private Sub createSalesForceQuantity()

    End Sub
    Private Function createMainData()
        Dim intRow = 1
        Dim intCol = 1
        Dim intColumOfData1 As Integer
        Dim intColumOfData2 As Integer
        intRow += 1
        For Each member In mySalesForce
            With anExcel
                .Cells(intRow, intCol) = member.getStrFirstName()
                intCol += 1
                .Cells(intRow, intCol) = member.getStrLastName()
                intCol += 1
                .Cells(intRow, intCol) = member.getIntOrderID()
                intCol += 1
                .Cells(intRow, intCol) = member.getIntID()
                intCol += 1
                intCol += 1

                intColumOfData1 = intCol
                .Cells(intRow, intCol) = member.getSngGamesSales()
                intCol += 1
                .Cells(intRow, intCol) = member.getSngDollsSales()
                intCol += 1
                .Cells(intRow, intCol) = member.getSngBuildingSales()
                intCol += 1
                .Cells(intRow, intCol) = member.getSngModelSales()
                intCol += 1
                .Cells(intRow, intCol) = "=sum(" & getColumnLetter(intColumOfData1) & intRow & ".." & getColumnLetter(intColumOfData1 + 3) & intRow & ")"
                intCol += 1
                .Cells(intRow, intCol) = "=min(" & getColumnLetter(intColumOfData1) & intRow & ".." & getColumnLetter(intColumOfData1 + 3) & intRow & ")"
                intCol += 1
                .Cells(intRow, intCol) = "=average(" & getColumnLetter(intColumOfData1) & intRow & ".." & getColumnLetter(intColumOfData1 + 3) & intRow & ")"
                intCol += 1
                .Cells(intRow, intCol) = "=max(" & getColumnLetter(intColumOfData1) & intRow & ".." & getColumnLetter(intColumOfData1 + 3) & intRow & ")"
                intCol += 1
                intCol += 1
                intColumOfData2 = intCol
                .Cells(intRow, intCol) = member.getIntGamesQuantity()
                intCol += 1
                .Cells(intRow, intCol) = member.getIntDollsQuantity()
                intCol += 1
                .Cells(intRow, intCol) = member.getIntBuildingQuantity()
                intCol += 1
                .Cells(intRow, intCol) = member.getIntModelQuantity()
                intCol += 1
                .Cells(intRow, intCol) = "=sum(" & getColumnLetter(intColumOfData2) & intRow & ".." & getColumnLetter(intColumOfData2 + 3) & intRow & ")"
                intCol += 1
                .Cells(intRow, intCol) = "=min(" & getColumnLetter(intColumOfData2) & intRow & ".." & getColumnLetter(intColumOfData2 + 3) & intRow & ")"
                intCol += 1
                .Cells(intRow, intCol) = "=average(" & getColumnLetter(intColumOfData2) & intRow & ".." & getColumnLetter(intColumOfData2 + 3) & intRow & ")"
                intCol += 1
                .Cells(intRow, intCol) = "=max(" & getColumnLetter(intColumOfData2) & intRow & ".." & getColumnLetter(intColumOfData2 + 3) & intRow & ")"
                intCol += 1
                intRow += 1
                intCol = 1
            End With
        Next
        Return intRow
    End Function
    Private Sub createHeader()
        Dim intRow As Integer = 1
        Dim intCol As Integer = 1
        Dim strHeaders As New List(Of String)
        strHeaders.Add("First Name")
        strHeaders.Add("Last Name")
        strHeaders.Add("Order ID")
        strHeaders.Add("Employee ID")
        strHeaders.Add(" ")
        strHeaders.Add("Games Sales")
        strHeaders.Add("Dolls Sales")
        strHeaders.Add("Build Sales")
        strHeaders.Add("Model Sales")
        strHeaders.Add("Total Sale")
        strHeaders.Add("Min Sales")
        strHeaders.Add("Avg Sales")
        strHeaders.Add("Max Sales")
        strHeaders.Add(" ")
        strHeaders.Add("Games Qty.")
        strHeaders.Add("Dolls Qty.")
        strHeaders.Add("Build Qty.")
        strHeaders.Add("Model Qty.")
        strHeaders.Add("Total Qty.")
        strHeaders.Add("Max Qty.")
        strHeaders.Add("Avg  Qty.")
        strHeaders.Add("Min Qty.")

        For Each header In strHeaders
            anExcel.Cells(intRow, intCol) = header.ToString
            intCol += 1
            Next
    End Sub
    Private Sub createFormulaHeaders(ByVal intLastRow)
        Dim intLastRowOfData As Integer = intLastRow '123123124124124124124124
        Dim intCol = 5
        Dim intCounter As Integer = 6

        intLastRow += 1
        With anExcel
            .Cells(intLastRow, intCol) = "Total:"
            While intCounter <= 22
                If intCounter <> 14 Then
                    .Cells(intLastRow, intCounter) = "=sum(" & getColumnLetter(intCounter) & 1 & ".." & getColumnLetter(intCounter) & intLastRowOfData & ")"
                End If
                intCounter += 1
            End While
            intCounter = 6
            intLastRow += 1
            .Cells(intLastRow, intCol) = "Max:"
            While intCounter <= 22
                If intCounter <> 14 Then
                    .Cells(intLastRow, intCounter) = "=max(" & getColumnLetter(intCounter) & 1 & ".." & getColumnLetter(intCounter) & intLastRowOfData & ")"
                End If
                intCounter += 1
            End While
            intCounter = 6
            intLastRow += 1
            .Cells(intLastRow, intCol) = "Avg:"
            While intCounter <= 22
                If intCounter <> 14 Then
                    .Cells(intLastRow, intCounter) = "=average(" & getColumnLetter(intCounter) & 1 & ".." & getColumnLetter(intCounter) & intLastRowOfData & ")"
                End If
                intCounter += 1
            End While
            intCounter = 6
            intLastRow += 1
            .Cells(intLastRow, intCol) = "Min:"
            While intCounter <= 22
                If intCounter <> 14 Then
                    .Cells(intLastRow, intCounter) = "=min(" & getColumnLetter(intCounter) & 1 & ".." & getColumnLetter(intCounter) & intLastRowOfData & ")"
                End If
                intCounter += 1
            End While
        End With
    End Sub
    Private Sub createSalesForce()
        mySalesForce.Add(New clsSalesperson("Robert", "Phillips", 103, 1015, 115.54, 4, 108.15, 3, 102.15, 1, 107.19, 5))
        mySalesForce.Add(New clsSalesperson("Susan", "Ricardo", 98, 1016, 174.15, 6, 132.14, 4, 181.54, 4, 185.67, 5))
        mySalesForce.Add(New clsSalesperson("William", "Acerba", 203, 1017, 165.34, 4, 193.43, 2, 154.65, 3, 192.23, 4))
        mySalesForce.Add(New clsSalesperson("Jill", "Quercas", 102, 1018, 186.85, 3, 196.65, 3, 324.44, 5, 175.34, 7))
        mySalesForce.Add(New clsSalesperson("Anthony", "Stallman", 104, 1019, 175.54, 4, 283.43, 6, 293.23, 4, 192.54, 2))
        mySalesForce.Add(New clsSalesperson("Scott", "Jarod", 36, 1020, 293.43, 5, 349.34, 3, 345.64, 3, 418.23, 2))
        mySalesForce.Add(New clsSalesperson("Fred", "Nostrandt", 12, 1021, 482.23, 4, 384.23, 2, 384.45, 4, 934.53, 4))
        mySalesForce.Add(New clsSalesperson("Leanne", "McCulloch", 215, 1022, 239.34, 2, 594.23, 4, 495.23, 5, 394.39, 9))
        mySalesForce.Add(New clsSalesperson("Valina", "Farland", 220, 1023, 394.54, 5, 495.45, 4, 594.23, 9, 293.43, 4))
        mySalesForce.Add(New clsSalesperson("Ashton", "Blasdell", 221, 1024, 473.99, 9, 293.98, 2, 485.38, 8, 384.95, 3))
        mySalesForce.Add(New clsSalesperson("Cullen", "Italski", 123, 1025, 494.53, 5, 340.89, 2, 830.0, 8, 348.53, 9))
        mySalesForce.Add(New clsSalesperson("Haleigh", "Turner", 144, 1026, 847.23, 9, 837.83, 4, 849.87, 7, 837.44, 8))
        mySalesForce.Add(New clsSalesperson("John", "Egland", 212, 1027, 282.29, 8, 101.87, 2, 192.82, 7, 172.33, 2))
        mySalesForce.Add(New clsSalesperson("Debbie", "Young", 133, 1028, 283.34, 8, 211.18, 2, 321.28, 2, 392.87, 7))
        mySalesForce.Add(New clsSalesperson("Larry", "Hon", 135, 1029, 293.45, 8, 374.54, 8, 847.34, 7, 283.43, 8))
        mySalesForce.Add(New clsSalesperson("Doug", "Ulysses", 132, 1030, 238.45, 2, 283.34, 2, 485.22, 2, 382.12, 8))
        mySalesForce.Add(New clsSalesperson("Bea", "Conrad", 201, 1031, 283.43, 2, 234.45, 5, 583.45, 4, 734.73, 8))
        mySalesForce.Add(New clsSalesperson("Ed", "Klute", 134, 1032, 293.43, 5, 837.45, 8, 934.98, 7, 938.28, 5))
        mySalesForce.Add(New clsSalesperson("Brian", "Larton", 143, 1033, 193.45, 5, 985.34, 3, 349.59, 9, 934.34, 2))
        mySalesForce.Add(New clsSalesperson("Cory", "Gerard", 200, 1034, 194.9, 9, 180.03, 4, 293.92, 3, 234.2, 9))
        mySalesForce.Add(New clsSalesperson("Aubrey", "Vander", 185, 1035, 102.32, 4, 293.04, 3, 203.98, 2, 203.0, 4))
        mySalesForce.Add(New clsSalesperson("Ted", "Xerxes", 181, 1036, 103.43, 2, 103.45, 2, 394.28, 4, 425.23, 6))
        mySalesForce.Add(New clsSalesperson("DeAnn", "Davis", 202, 1037, 192.23, 3, 283.43, 3, 384.23, 2, 384.98, 8))
        mySalesForce.Add(New clsSalesperson("Ron", "Zening", 76, 1038, 102.23, 3, 493.34, 3, 495.45, 4, 450.3, 9))
        mySalesForce.Add(New clsSalesperson("Peggy", "Wallis", 199, 1039, 103.43, 3, 394.04, 9, 493.23, 2, 940.2, 2))
        mySalesForce.Add(New clsSalesperson("Amy", "Oloff", 187, 1040, 102.3, 2, 184.03, 4, 103.45, 2, 394.34, 8))
    End Sub
End Module
