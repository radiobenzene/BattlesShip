'Original author - Uditangshu Aurangabadkar

Sub Start_Game()
    InitGame
    ColorCells 'Coloring all cells
    ClearCells 'Clearing all cells

    'Placing only single ships
    direction = Int((1 - 0 + 1 * Rnd) + 0) 'Random number between 0 and 1

    PlaceFourShip direction 'Placing four block ship with the direction as a random integer 

    'The other ships can be placed using numbers or even the random direction
    PlaceTripleShip 1 
    PlaceTripleShip 0

    PlaceDoubleShip 0
    PlaceDoubleShip 1
    PlaceDoubleShip 1

    PlaceSingleShip 0
    PlaceSingleShip 0
    PlaceSingleShip 0
    PlaceSingleShip 0
End Sub

Function ColorCells()
    Dim selected_color As Integer
    selected_color = 5
    'Coloring cells with the selected_color
    Range("E5:N14").Interior.ColorIndex = selected_color
End Function

Function InitGame()
    InitStepCounter
    InitShipCount
End Function

'Timer function is disabled
Function InitTimer()
    Range("R6").Value = "0:05:00"
End Function

Function InitShipCount()
    Cells(10, 19).Value = 1
    Cells(11, 19).Value = 2
    Cells(12, 19).Value = 3
    Cells(13, 19).Value = 4
End Function

Function InitStepCounter()
    Cells(15, 19).Value = 0
End Function

Function ClearCells()
    Range("E5:N14").Value = " "
End Function

Sub CreateRandomParameters(upper_lim As Integer, lower_lim As Integer)
    row = Int((upper_lim - lower_lim + 1) * Rnd + lower_lim) 'Random number between 2 and 11
    col = Int((upper_lim - lower_lim + 1) * Rnd + lower_lim) 'Random number between 2 and 11   
End Sub



Function PlaceSingleShip(dir As Integer)
tag2:    CreateRandomParameters 14, 5
    If ((isEmpty(row, col) = True) And (Cells(row, col).Interior.ColorIndex <> 15) And (Cells(row, col).Value <> 1) And (canPaint(row, col, 1, 0) = True)) Then
        Cells(row, col).Value = 1
        Cells(row, col).Font.ColorIndex = 5 'Coloring the text in cell
        Paint row, col, 1, dir
    Else
      GoTo tag2
       
    End If
End Function

Function PlaceDoubleShip(dir As Integer)
'Placing two ships
tag:    CreateRandomParameters 14, 5
Select Case dir
Case 0
    'If vertical ship
    If (((isEmpty(row, col) = True) And (isEmpty(row + 1, col) = True)) And (Cells(row, col).Interior.ColorIndex <> 15) And (Cells(row, col).Value <> 1) And (Cells(row, col).Value <> 2) And (Cells(row + 1, col).Value <> 2) And (canPaint(row, col, 2, dir) = True)) Then
        Cells(row, col).Value = 2
        Cells(row + 1, col).Value = 2
        Cells(row, col).Font.ColorIndex = 5
        Cells(row + 1, col).Font.ColorIndex = 5
        Paint row, col, 2, dir 'change here too
        
    Else
        'CreateRandomParameters 12, 2
        GoTo tag
    End If
    
Case 1
    If ((isEmpty(row, col) = True) And (isEmpty(row, col + 1) = True) And (Cells(row, col).Interior.ColorIndex <> 15) And canPaint(row, col, 2, dir)) Then
        Cells(row, col).Value = 2
        Cells(row, col + 1).Value = 2
        Cells(row, col).Font.ColorIndex = 5
        Cells(row, col + 1).Font.ColorIndex = 5
        Paint row, col, 2, dir
    Else
        GoTo tag
    End If
    
End Select
End Function

Function PlaceTripleShip(dir As Integer)
'Placing triple-ships
tag3:   CreateRandomParameters 14, 5
Select Case dir
Case 1
    If ((isEmpty(row, col) = True) And (isEmpty(row, col + 1) = True) And (isEmpty(row, col + 2) = True) And ((Cells(row, col + 1).Value <> 2) Or (Cells(row, col + 2).Value <> 2)) And (canPaint(row, col, 3, dir) = True)) Then
        Cells(row, col).Value = 3
        Cells(row, col + 1).Value = 3
        Cells(row, col + 2).Value = 3
        
        Paint row, col, 3, dir
        Cells(row, col).Font.ColorIndex = 5
        Cells(row, col + 1).Font.ColorIndex = 5
        Cells(row, col + 2).Font.ColorIndex = 5
        
    Else
        GoTo tag3
    End If
Case 0
    'Vertical Ship Placement
    If (isEmpty(row, col) = True And (isEmpty(row + 1, col) = True) And (isEmpty(row + 2, col) = True) And (canPaint(row, col, 3, dir) = True)) Then
        Cells(row, col).Value = 3
        Cells(row + 1, col).Value = 3
        Cells(row + 2, col).Value = 3
        Paint row, col, 3, dir
        
        Cells(row, col).Font.ColorIndex = 5
        Cells(row + 1, col).Font.ColorIndex = 5
        Cells(row + 2, col).Font.ColorIndex = 5
    Else
        GoTo tag3
    End If
    
End Select
End Function

Function PlaceFourShip(dir As Integer)
'Placing four-deck ships
tag4: CreateRandomParameters 14, 5
Select Case dir
Case 0
    If ((isEmpty(row, col) = True) And (isEmpty(row + 1, col) = True) And (isEmpty(row + 2, col) = True) And (isEmpty(row + 3, col) = True) And ((Cells(row + 1, col).Value <> 3) Or (Cells(row + 2, col).Value <> 3) Or (Cells(row + 3, col).Value <> 3)) And canPaint(row, col, 4, dir) = True) Then
        Cells(row, col).Value = 4
        Cells(row + 1, col).Value = 4
        Cells(row + 2, col).Value = 4
        Cells(row + 3, col).Value = 4
        
        Paint row, col, 4, dir
        Cells(row, col).Font.ColorIndex = 5
        Cells(row + 1, col).Font.ColorIndex = 5
        Cells(row + 2, col).Font.ColorIndex = 5
        Cells(row + 3, col).Font.ColorIndex = 5
    Else
        GoTo tag4
    End If
Case 1
    If (isEmpty(row, col) = True And (isEmpty(row, col + 1) = True) And (isEmpty(row, col + 2) = True) And (isEmpty(row, col + 3) = True) And (canPaint(row, col, 4, dir) = True)) Then
        Cells(row, col).Value = 4
        Cells(row, col + 1).Value = 4
        Cells(row, col + 2).Value = 4
        Cells(row, col + 3).Value = 4
        
        Paint row, col, 4, dir
        
        Cells(row, col).Font.ColorIndex = 5
        Cells(row, col + 1).Font.ColorIndex = 5
        Cells(row, col + 2).Font.ColorIndex = 5
        Cells(row, col + 3).Font.ColorIndex = 5
    Else
        GoTo tag4
    End If
    
End Select
End Function

Function isEmpty(row As Integer, col As Integer)
    Dim checkPlacement As Boolean
    If ((Cells(row, col).Value = " ") And (Cells(row, col).Interior.ColorIndex <> 15) And (Cells(row, col).Interior.ColorIndex = 5)) Then
        checkPlacement = True 'Cell is Empty
    Else
        checkPlacement = False 'Cell is not Empty
    End If
    isEmpty = checkPlacement
End Function



 
Function StartTimer()
    interval = Now + TimeValue("00:00:01")
    If Range("R6").Value = 0 Then Exit Function
    Range("R6") = Range("R6") - TimeValue("00:00:01")
    Application.OnTime interval, "Game_Timer"
End Function

Function StopTimer()
    Application.OnTime earliestTime:=interval, Procedure:=Game_Timer, Schedule:=False
End Function

Function canPaint(row As Integer, col As Integer, ship_size As Integer, dir As Integer)
'Can paint surroundings
    Dim isAvailable As Boolean
    Select Case ship_size
        Case 1
            'One-decked ship
            For i = row - 1 To row + 1
                For j = col - 1 To col + 1
                    If ((Cells(i, j).Interior.ColorIndex <> 15)) Then 'And (Cells(i, j).Value = " ")) Then
                        isAvailable = True
                    Else
                        isAvailable = False
                        End If
                Next j
            Next i
            
        Case 2
            Select Case dir
            Case 0
                    'Vertical
                    For i = row - 1 To row + 2
                        For j = col - 1 To col + 1
                            If ((Cells(i, j).Interior.ColorIndex <> 15) And (Cells(i, j).Value = " ")) Then
                                isAvailable = True
                            Else
                                isAvailable = False
                            End If
                        Next j
                    Next i
            Case 1
                'Horizontal
                For i = row - 1 To row + 1
                    For j = col - 1 To col + 2
                        If ((Cells(i, j).Interior.ColorIndex <> 15) And (Cells(i, j).Value = " ")) Then
                                isAvailable = True
                            Else
                                isAvailable = False
                            End If
                        Next j
                    Next i

            End Select
        
        Case 3
            Select Case dir
            Case 1
                For i = row - 1 To row + 1
                    For j = col - 1 To col + 3
                        If ((Cells(i, j).Interior.ColorIndex <> 15)) Then
                            isAvailable = True
                        Else
                            isAvailable = False
                        End If
                    Next j
                Next i
                
            Case 0
                For i = row - 1 To row + 3
                    For j = col - 1 To col + 1
                        If (Cells(i, j).Interior.ColorIndex <> 15) Then
                            isAvailable = True
                        Else
                            isAvaialble = False
                        End If
                    Next j
                Next i
            End Select
        
        Case 4
            Select Case dir
            Case 0
                For i = row - 1 To row + 4
                    For j = col - 1 To col + 1
                        If (Cells(i, j).Interior.ColorIndex <> 15) Then
                            isAvailable = True
                        Else
                            isAvailable = False
                        End If
                    Next j
                Next i
            Case 1
                For i = row - 1 To row + 1
                    For j = col - 1 To col + 4
                        If (Cells(i, j).Interior.ColorIndex <> 15) Then
                            isAvailable = True
                        Else
                            isAvailable = False
                        End If
                    Next j
                Next i
                
            End Select
            
        End Select
        canPaint = isAvailable
End Function

Sub Paint(row As Integer, col As Integer, ship_size As Integer, dir As Integer)
    Select Case ship_size
        Case 1
            Cells(row, col).Value = 1
            For i = row - 1 To row + 1
                For j = col - 1 To col + 1
                If (Cells(i, j).Interior.ColorIndex <> 15) Then ' And Cells(i, j).Value = "") Then
                    Cells(i, j).Value = 0
                    Cells(row, col).Value = 1
                    Cells(i, j).Font.ColorIndex = 5
                End If
                Next j
            Next i
        
        Case 2
                Select Case dir
                Case 0
                    Cells(row, col).Value = 2
                    Cells(row + 1, col).Value = 2
                    For i = row - 1 To row + 2
                        For j = col - 1 To col + 1
                        If (Cells(i, j).Interior.ColorIndex <> 15) Then
                            Cells(i, j).Value = 0
                            Cells(i, j).Font.ColorIndex = ActiveCell.Interior.ColorIndex
                            
                            Cells(row, col).Value = 2
                            Cells(row + 1, col).Value = 2
                            Cells(i, j).Font.ColorIndex = 5
                        End If
                        Next j
                    Next i
                    
                 Case 1
                    Cells(row, col).Value = 2
                    Cells(row, col + 1).Value = 2
                    For i = row - 1 To row + 1
                        For j = col - 1 To col + 2
                            If (Cells(i, j).Interior.ColorIndex <> 15) Then
                                Cells(i, j).Value = 0
                                Cells(i, j).Font.ColorIndex = ActiveCell.Interior.ColorIndex
                                Cells(row, col).Value = 2
                                Cells(row, col + 1).Value = 2
                                Cells(i, j).Font.ColorIndex = 5
                            End If
                        Next j
                    Next i
                 End Select
                
        Case 3
                Select Case dir
                Case 1
                    Cells(row, col).Value = 3
                    Cells(row, col + 1).Value = 3
                    Cells(row, col + 2).Value = 3
                    For i = row - 1 To row + 1
                        For j = col - 1 To col + 3
                            If (Cells(i, j).Interior.ColorIndex <> 15) Then
                            Cells(i, j).Value = 0
                            Cells(row, col).Value = 3
                            Cells(row, col + 1).Value = 3
                            Cells(row, col + 2).Value = 3
                            Cells(i, j).Font.ColorIndex = 5
                            End If
                        Next j
                    Next i
                    
                Case 0
                    Cells(row, col).Value = 3
                    Cells(row + 1, col).Value = 3
                    Cells(row + 2, col).Value = 3
                    
                    For i = row - 1 To row + 3
                        For j = col - 1 To col + 1
                            If (Cells(i, j).Interior.ColorIndex <> 15) Then
                                Cells(i, j).Value = 0
                                
                                Cells(row, col).Value = 3
                                Cells(row + 1, col).Value = 3
                                Cells(row + 2, col).Value = 3
                                Cells(i, j).Font.ColorIndex = 5
                            End If
                        Next j
                    Next i
        End Select
            
        Case 4
            Select Case dir
            Case 0
                Cells(row, col).Value = 4
                Cells(row + 1, col).Value = 4
                Cells(row + 2, col).Value = 4
                Cells(row + 3, col).Value = 4
                
                For i = row - 1 To row + 4
                    For j = col - 1 To col + 1
                        If (Cells(i, j).Interior.ColorIndex <> 15) Then
                        
                            Cells(i, j).Value = 0
                            Cells(row, col).Value = 4
                            Cells(row + 1, col).Value = 4
                            Cells(row + 2, col).Value = 4
                            Cells(row + 3, col).Value = 4
                            Cells(i, j).Font.ColorIndex = 5
                        End If
                    Next j
                Next i
                
            Case 1
                Cells(row, col).Value = 4
                Cells(row, col + 1).Value = 4
                Cells(row, col + 2).Value = 4
                Cells(row, col + 3).Value = 4
                
                For i = row - 1 To row + 1
                    For j = col - 1 To col + 4
                        If (Cells(i, j).Interior.ColorIndex <> 15) Then
                        
                            Cells(i, j).Value = 0
                            Cells(row, col).Value = 4
                            Cells(row, col + 1).Value = 4
                            Cells(row, col + 2).Value = 4
                            Cells(row, col + 3).Value = 4
                            Cells(i, j).Font.ColorIndex = 5
                        End If
                    Next j
                Next i
                        
            End Select
            
    End Select
End Sub


