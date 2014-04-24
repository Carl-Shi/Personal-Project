Sub GeneratePlan()

Dim nIndex1 As Integer
Dim nIndex2 As Integer
Dim nTotal As Integer
Dim strName(150) As String
Dim strOriginalSeat(150) As String
Dim strNewSeat(150) As String
Dim nRound(150) As Integer

' Import all data
nTotal = 0

For nIndex1 = 7 To 56 Step 7
    For nIndex2 = 2 To 13
        If Sheets("7F").Cells(nIndex1, nIndex2) <> "" Then
            nTotal = nTotal + 1
            strName(nTotal) = Trim(Sheets("7F").Cells(nIndex1, nIndex2))
            strOriginalSeat(nTotal) = Trim(Sheets("7F").Cells(nIndex1 + 1, nIndex2))
            strNewSeat(nTotal) = Trim(Sheets("7F").Cells(nIndex1 + 2, nIndex2))
        End If
    Next
Next

For nIndex1 = 7 To 56 Step 7
    For nIndex2 = 2 To 13
        If Sheets("8F").Cells(nIndex1, nIndex2) <> "" Then
            nTotal = nTotal + 1
            strName(nTotal) = Trim(Sheets("8F").Cells(nIndex1, nIndex2))
            strOriginalSeat(nTotal) = Trim(Sheets("8F").Cells(nIndex1 + 1, nIndex2))
            strNewSeat(nTotal) = Trim(Sheets("8F").Cells(nIndex1 + 2, nIndex2))
        End If
    Next
Next

For nIndex1 = 1 To nTotal
    nRound(nIndex1) = 0
Next

'First group, need not move
For nIndex1 = 1 To nTotal
    If strNewSeat(nIndex1) = strOriginalSeat(nIndex1) Then
        Sheets("Plan").Cells(nIndex1, 1) = strName(nIndex1) + " " + strNewSeat(nIndex1)
        nRound(nIndex1) = 1
    End If
Next

'Second group, new seat is available
For nIndex1 = 1 To nTotal
    If nRound(nIndex1) = 0 Then
        For nIndex2 = 1 To nTotal
            If nRound(nIndex2) <> 1 And strNewSeat(nIndex1) = strOriginalSeat(nIndex2) Then
                Exit For
            End If
        Next
    
        If nIndex2 > nTotal Then
            Sheets("Plan").Cells(nIndex1, 2) = strName(nIndex1) + " " + strOriginalSeat(nIndex1) + "->" + strNewSeat(nIndex1)
            nRound(nIndex1) = 2
        End If
    End If
Next

'Third group, new seat is available after second group's move
For nIndex1 = 1 To nTotal
    If nRound(nIndex1) = 0 Then
        For nIndex2 = 1 To nTotal
            If nRound(nIndex2) <> 1 And nRound(nIndex2) <> 2 And strNewSeat(nIndex1) = strOriginalSeat(nIndex2) Then
                Exit For
            End If
        Next
    
        If nIndex2 > nTotal Then
            Sheets("Plan").Cells(nIndex1, 3) = strName(nIndex1) + " " + strOriginalSeat(nIndex1) + "->" + strNewSeat(nIndex1)
            nRound(nIndex1) = 3
        End If
    End If
Next

'Four group, the remaining people
For nIndex1 = 1 To nTotal
    If nRound(nIndex1) = 0 Then
        ' If someone depends on me, it means I need to move to temp office first
        For nIndex2 = 1 To nTotal
            If nRound(nIndex2) = 0 And strOriginalSeat(nIndex1) = strNewSeat(nIndex2) Then
                Sheets("Plan").Cells(nIndex1, 2) = strName(nIndex1) + " " + strOriginalSeat(nIndex1) + "->Temp"
                Sheets("Plan").Cells(nIndex1, 4) = strName(nIndex1) + " Temp->" + strNewSeat(nIndex1)
                nRound(nIndex1) = 4
            
                Sheets("Plan").Cells(nIndex2, 3) = strName(nIndex2) + " " + strOriginalSeat(nIndex2) + "->" + strNewSeat(nIndex2)
                nRound(nIndex2) = 3
                Exit For
            End If
        Next
        
        ' If no one depends on me, it means I am the last one
        If nIndex2 > nTotal Then
            Sheets("Plan").Cells(nIndex1, 4) = strName(nIndex1) + " " + strOriginalSeat(nIndex1) + "->" + strNewSeat(nIndex1)
            nRound(nIndex1) = 4
        End If
    End If
Next

End Sub
