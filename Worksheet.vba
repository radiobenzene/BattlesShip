Private Sub Worksheet_SelectionChange(ByVal Target As Range)

'0 is Horizontal direction
'1 is Vertical direction
    
CountSteps
PlaceFlag

DeleteSingleShip

DeleteDoubleShip 0
DeleteDoubleShip 1

DeleteTripleShip 1
DeleteTripleShip 0

DeleteFourShip 1
DeleteFourShip 0

If (Cells(10, 19).Value = 0 And Cells(11, 19).Value = 0 And Cells(12, 19).Value = 0 And Cells(13, 19).Value = 0) Then
    ClearCells
    ColorCells
        'EndGame form 
    EndGame.Show
End If
End Sub

Function CountSteps()
    Dim counter As Integer 'step counter
    counter = 0
    If ((ActiveCell.Interior.ColorIndex = 5) And (ActiveCell.Value <> 1)) Then
    'PlaceFlag
    counter = counter + 1
    Cells(15, 19).Value = Cells(15, 19).Value + 1
    End If
End Function

Public Function PlaceFlag()
'If clicked placing flag
If ((ActiveCell.Interior.ColorIndex <> 15) And ((ActiveCell.Value = 1) Or (ActiveCell.Value = 2) Or (ActiveCell.Value = 3) Or (ActiveCell.Value = 4))) Then
    ActiveCell.Interior.ColorIndex = 3
    ActiveCell.Font.ColorIndex = ActiveCell.Interior.ColorIndex
End If
If (ActiveCell.Interior.ColorIndex <> 15 And (ActiveCell.Value = " " Or ActiveCell.Value = 0) And ActiveCell.Interior.ColorIndex = 5) Then
    ActiveCell.Interior.ColorIndex = 50
    ActiveCell.Font.ColorIndex = ActiveCell.Interior.ColorIndex
End If
    
End Function

Function ReduceCounter(row As Integer, col As Integer)
   Cells(row, col).Value = Cells(row, col).Value - 1
   If (Cells(row, col).Value < 0) Then
   Cells(row, col).Value = Cells(row, col).Value + 1
   End If
End Function

Function DeleteSingleShip()
        Dim isColored As Boolean
        isColored = False
    If (ActiveCell.Interior.ColorIndex = 3 And ActiveCell.Value = 1 And isColored = False) Then
        isColored = True
        Cells(13, 19).Value = Cells(13, 19) - 1
        OpenSingleBoundary ActiveCell.row, ActiveCell.Column
    End If
    If (Cells(13, 19).Value < 0) Then
        Cells(13, 19).Value = Cells(13, 19).Value + 1
    End If
End Function

Function DeleteDoubleShip(dir As Integer)
Dim isColored As Boolean
Select Case dir
Case 0
    'Vertical Placement
    
    isColored = False
     If ((Cells(ActiveCell.row, ActiveCell.Column).Value = 2 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 1, ActiveCell.Column).Value = 2 And Cells(ActiveCell.row + 1, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
         isColored = True
         OpenDoubleBoundary ActiveCell.row, ActiveCell.Column, 0
         Cells(12, 19).Value = Cells(12, 19).Value - 1
    End If
    If ((Cells(ActiveCell.row, ActiveCell.Column).Value = 2 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 1, ActiveCell.Column).Value = 2 And Cells(ActiveCell.row - 1, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
         OpenDoubleBoundary ActiveCell.row - 1, ActiveCell.Column, 0
         isColored = True
         Cells(12, 19).Value = Cells(12, 19).Value - 1
        
    End If
    
Case 1
    'Horizontal Placement
    isColored = False
    If ((Cells(ActiveCell.row, ActiveCell.Column).Value = 2 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 1).Value = 2 And Cells(ActiveCell.row, ActiveCell.Column + 1).Interior.ColorIndex = 3) And isColored = False) Then 'On Left cell
        isColored = True
        OpenDoubleBoundary ActiveCell.row, ActiveCell.Column, 1
        Cells(12, 19).Value = Cells(12, 19).Value - 1
    End If
    If ((Cells(ActiveCell.row, ActiveCell.Column).Value = 2 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 1).Value = 2 And Cells(ActiveCell.row, ActiveCell.Column - 1).Interior.ColorIndex = 3) And isColored = False) Then 'On right Cell
        isColored = True
        OpenDoubleBoundary ActiveCell.row, ActiveCell.Column - 1, 1
        Cells(12, 19).Value = Cells(12, 19).Value - 1
    End If
    
         If (Cells(12, 19).Value < 0) Then
            Cells(12, 19).Value = Cells(12, 19).Value + 1
         End If
    End Select
         
    
    
End Function

Function DeleteTripleShip(dir As Integer)
    Dim isColored As Boolean
    
Select Case dir
    Case 1
        'Horizontal ship
        isColored = False
        If ((Cells(ActiveCell.row, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 1).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column + 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 2).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column + 2).Interior.ColorIndex = 3) And isColored = False) Then
            isColored = True
            OpenTripleBoundary ActiveCell.row, ActiveCell.Column, 1
            Cells(11, 19).Value = Cells(11, 19).Value - 1
        ElseIf (((Cells(ActiveCell.row, ActiveCell.Column).Value = 3) And (Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3)) And ((Cells(ActiveCell.row, ActiveCell.Column - 1).Value = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 1).Interior.ColorIndex = 3)) And ((Cells(ActiveCell.row, ActiveCell.Column - 2).Value = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 2).Interior.ColorIndex = 3)) And isColored = False) Then
            isColored = True
            OpenTripleBoundary ActiveCell.row, ActiveCell.Column - 2, 1
            Cells(11, 19).Value = Cells(11, 19).Value - 1
        ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 1).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column - 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 1).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column + 1).Interior.ColorIndex = 3) And isColored = False) Then
            isColored = True
            OpenTripleBoundary ActiveCell.row, ActiveCell.Column - 1, 1
            Cells(11, 19).Value = Cells(11, 19).Value - 1
        End If
        
    Case 0
        'Vertical Ship
        isColored = False
        If (((Cells(ActiveCell.row, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3)) And (Cells(ActiveCell.row + 1, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row + 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 2, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row + 2, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
            isColored = True
            OpenTripleBoundary ActiveCell.row, ActiveCell.Column, 0
            Cells(11, 19).Value = Cells(11, 19).Value - 1
        ElseIf ((((Cells(ActiveCell.row, ActiveCell.Column).Value = 3) And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3)) And (Cells(ActiveCell.row - 1, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row - 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 2, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row - 2, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
            isColored = True
            OpenTripleBoundary ActiveCell.row - 2, ActiveCell.Column, 0
            Cells(11, 19).Value = Cells(11, 19).Value - 1
        ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 1, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row - 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 1, ActiveCell.Column).Value = 3 And Cells(ActiveCell.row + 1, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
            isColored = True
            OpenTripleBoundary ActiveCell.row - 1, ActiveCell.Column, 0
            Cells(11, 19).Value = Cells(11, 19).Value - 1
        End If
        
End Select
         If (Cells(11, 19).Value < 0) Then
         Cells(11, 19).Value = Cells(11, 19).Value + 1
        End If
End Function

Function DeleteFourShip(dir As Integer)
    Dim isColored As Boolean
Select Case dir
Case 1
    'Horizontal Ship
    isColored = False
    If ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 1).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column + 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 2).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column + 2).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 3).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column + 3).Interior.ColorIndex = 3) And isColored = False) Then 'Left cell
         isColored = True
         OpenFourBoundary ActiveCell.row, ActiveCell.Column, 1
         Cells(10, 19).Value = Cells(10, 19).Value - 1
    ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 1).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column - 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 2).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column - 2).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 3).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column - 3).Interior.ColorIndex = 3) And isColored = False) Then 'Right cell
        isColored = True
        OpenFourBoundary ActiveCell.row, ActiveCell.Column - 3, 1
        Cells(10, 19).Value = Cells(10, 19).Value - 1
    ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 1).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column + 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 1).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column - 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 2).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column - 2).Interior.ColorIndex = 3) And isColored = False) Then
        isColored = True
        OpenFourBoundary ActiveCell.row, ActiveCell.Column - 2, 1
        Cells(10, 19).Value = Cells(10, 19).Value - 1
    ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column - 1).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column - 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 1).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column + 1).Interior.ColorIndex = 3) And (Cells(ActiveCell.row, ActiveCell.Column + 2).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column + 2).Interior.ColorIndex = 3) And isColored = False) Then
        isColored = True
        OpenFourBoundary ActiveCell.row, ActiveCell.Column - 1, 1
        Cells(10, 19).Value = Cells(10, 19).Value - 1
    End If
    
Case 0
    'Vertical ship
    isColored = False
    If ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 1, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row + 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 2, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row + 2, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 3, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row + 3, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then 'Top cell
        isColored = True
        OpenFourBoundary ActiveCell.row, ActiveCell.Column, 0
        Cells(10, 19).Value = Cells(10, 19).Value - 1
    ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 1, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row - 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 2, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row - 2, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 3, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row - 3, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
        isColored = True
        OpenFourBoundary ActiveCell.row - 3, ActiveCell.Column, 0
        Cells(10, 19).Value = Cells(10, 19).Value - 1
    ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 1, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row - 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 1, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row + 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 2, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row + 2, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
        isColored = True
        OpenFourBoundary ActiveCell.row - 1, ActiveCell.Column, 0
        Cells(10, 19).Value = Cells(10, 19).Value - 1
    ElseIf ((Cells(ActiveCell.row, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row + 1, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row + 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 1, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row - 1, ActiveCell.Column).Interior.ColorIndex = 3) And (Cells(ActiveCell.row - 2, ActiveCell.Column).Value = 4 And Cells(ActiveCell.row - 2, ActiveCell.Column).Interior.ColorIndex = 3) And isColored = False) Then
        isColored = True
        OpenFourBoundary ActiveCell.row - 2, ActiveCell.Column, 0
        Cells(10, 19).Value = Cells(10, 19).Value - 1
    
    End If
        
End Select
         If (Cells(10, 19).Value < 0) Then
        Cells(10, 19).Value = Cells(10, 19).Value + 1
        End If
End Function

Function OpenSingleBoundary(row As Integer, col As Integer)
    'Opening cell if ship is opened
    
    For i = row - 1 To row + 1
        For j = col - 1 To col + 1
            If (Cells(i, j).Interior.ColorIndex <> 15) Then
            Cells(row, col).Interior.ColorIndex = 3
             Cells(i, j).Interior.ColorIndex = 4
             Cells(i, j).Font.ColorIndex = 4 'ActiveCell.Interior.ColorIndex
             'Cells(i, j).Value = X
            End If
        Next j
    Next i
End Function


Function OpenDoubleBoundary(row As Integer, col As Integer, dir As Integer)
Select Case dir
Case 0
    'Vertical Ship
    For i = row - 1 To row + 2
        For j = col - 1 To col + 1
            If (Cells(i, j).Interior.ColorIndex <> 15 And Cells(row, col).Value = 2 And (Cells(row + 1, col).Value = 2 Or Cells(row - 1, col).Value = 2)) Then
                Cells(row, col).Interior.ColorIndex = 3
                Cells(row + 1, col).Interior.ColorIndex = 3
                Cells(i, j).Interior.ColorIndex = 4
                Cells(i, j).Font.ColorIndex = 4
            End If
        Next j
    Next i
    
Case 1
    'Horizontal Ship
    For i = row - 1 To row + 1
        For j = col - 1 To col + 2
            If (Cells(i, j).Interior.ColorIndex <> 15 And Cells(row, col).Value = 2 And (Cells(row, col + 1).Value = 2 Or Cells(row, col - 1).Value = 2)) Then
                Cells(row, col).Interior.ColorIndex = 3
                Cells(row, col + 1).Interior.ColorIndex = 3
                Cells(i, j).Interior.ColorIndex = 4
                Cells(i, j).Font.ColorIndex = 4
            End If
        Next j
       Next i
        
    End Select
End Function

Function OpenTripleBoundary(row As Integer, col As Integer, dir As Integer)
Select Case dir
Case 1
    'Horizontal case
    For i = row - 1 To row + 1
        For j = col - 1 To col + 3
            If (Cells(i, j).Interior.ColorIndex <> 15) Then
                Cells(row, col).Interior.ColorIndex = 3
                Cells(row, col + 1).Interior.ColorIndex = 3
                Cells(row, col + 2).Interior.ColorIndex = 3
                Cells(i, j).Interior.ColorIndex = 4
                Cells(i, j).Font.ColorIndex = 4
            End If
        Next j
    Next i

Case 0
    'Vertical ship
    For i = row - 1 To row + 3
        For j = col - 1 To col + 1
            If (Cells(i, j).Interior.ColorIndex <> 15) Then
                Cells(row, col).Interior.ColorIndex = 3
                Cells(row + 1, col).Interior.ColorIndex = 3
                Cells(row + 2, col).Interior.ColorIndex = 3
                Cells(i, j).Interior.ColorIndex = 4
                Cells(i, j).Font.ColorIndex = 4
            End If
        Next j
    Next i
    End Select
End Function

Function OpenFourBoundary(row As Integer, col As Integer, dir As Integer)
Select Case dir
Case 1
    'Horizontal ship
    For i = row - 1 To row + 1
        For j = col - 1 To col + 4
            If (Cells(i, j).Interior.ColorIndex <> 15) Then
                Cells(row, col).Interior.ColorIndex = 3
                Cells(row, col + 1).Interior.ColorIndex = 3
                Cells(row, col + 2).Interior.ColorIndex = 3
                Cells(row, col + 3).Interior.ColorIndex = 3
                Cells(i, j).Interior.ColorIndex = 4
                Cells(i, j).Font.ColorIndex = 4
            End If
        Next j
    Next i
Case 0
    'Vertical ship
    For i = row - 1 To row + 4
        For j = col - 1 To col + 1
            If (Cells(i, j).Interior.ColorIndex <> 15) Then
                Cells(row, col).Interior.ColorIndex = 3
                Cells(row + 1, col).Interior.ColorIndex = 3
                Cells(row + 2, col).Interior.ColorIndex = 3
                Cells(row + 3, col).Interior.ColorIndex = 3
                Cells(i, j).Interior.ColorIndex = 4
                Cells(i, j).Font.ColorIndex = 4
            End If
        Next j
    Next i
End Select
End Function

Function UserMessage()
    'MessageBox at the end of the game
    Dim Message As String
    Dim Win As String
    Message = "You won!"
    Win = "WINNER"
    'demo = "HELP.HLP"
    Context = 1000
     'MessageUser = MsgBox(Message, vbOKOnly, Win, demo, Context)
    'If (MessageUser = vbOK) Then
        Workbooks("ThisWorkbook").Close 'SaveChanges = True
    End If
End Function

Function FinEnd()
    If (Cells(4, 15).Value = 0 And Cells(5, 15).Value = 0 And Cells(6, 15).Value = 0 And Cells(7, 15).Value = 0) Then
        UserMessage
    End If
End Function

Function Equate(row As Integer, col As Integer, num As Integer, index_col As Integer)
    Dim equal As Boolean
    equal = False
    If ((Cells(row, col).Value = num) And (Cells(row, col).Interior.ColorIndex = index_col)) Then
        equal = True
    Else
        equal = False
    End If
    Equate = equal
End Function

