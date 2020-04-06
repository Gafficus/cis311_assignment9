'------------------------------------------------------------
'-                File Name : modGlobal.vb                     - 
'-                Part of Project: Assign9                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 5 Apr 2020                    -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the global variable used by the program-                                     - 
'------------------------------------------------------------	
Module modGlobals
    Public Enum column
        A = 1
        B
        C
        D
        E
        F
        G
        H
        I
        J
        K
        L
        M
        N
        O
        P
        Q
        R
        S
        T
        U
        V
        W
        X
        Y
        Z
    End Enum
    Public Enum employee
        strFirstName = 0
        strLastName
        intOrderID
        intID
        sngGamesSales
        intGamesQuantity
        sngDollsSales
        intDollsQuantity
        sngBuildingSales
        intBuildingQuantity
        sngModelSales
        intModelQuantity
    End Enum
    '------------------------------------------------------------
    '-                Subprogram Name: getColumnLetter          -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: The date you wrote it         -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is will take in an iteger and retun a    –
    '- character as a string                                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- intLetter - the integer coming into the program          -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- letter – holds the cahracter to be returned              -
    '------------------------------------------------------------
    Public Function getColumnLetter(ByVal intLetter)
        Dim letter As String = "z"
        Select Case intLetter
            Case 1
                letter = "A"
            Case 2
                letter = "B"
            Case 3
                letter = "C"
            Case 4
                letter = "D"
            Case 5
                letter = "E"
            Case 6
                letter = "F"
            Case 7
                letter = "G"
            Case 8
                letter = "H"
            Case 9
                letter = "I"
            Case 10
                letter = "J"
            Case 11
                letter = "K"
            Case 12
                letter = "L"
            Case 13
                letter = "M"
            Case 14
                letter = "N"
            Case 15
                letter = "O"
            Case 16
                letter = "P"
            Case 17
                letter = "Q"
            Case 18
                letter = "R"
            Case 19
                letter = "S"
            Case 20
                letter = "T"
            Case 21
                letter = "U"
            Case 22
                letter = "V"
            Case 23
                letter = "W"
            Case 24
                letter = "X"
            Case 25
                letter = "Y"
            Case 26
                letter = "Z"
        End Select
        Return letter
    End Function
End Module
