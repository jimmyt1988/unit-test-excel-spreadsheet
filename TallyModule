Function Tally(refCell As Range, decayValueCell As Range, recoverValueCell As Range) As Single
    
    Dim previousCellValue As Single
    Dim outputValue As Single
    
    ' Get the value from the cell to the left of refCell.
    previousCellValue = Application.Caller.Offset(0, -1).value
    refCellValue = refCell.value
    
    ' Check if the decay value is exactly equal to 1 (this might be a bit inflexible if you expect a range of values).
    If refCellValue >= 1 Then
        ' If so, subtract the decay value from the previous cell value and ensure it doesn't go below zero.
        outputValue = Application.Max(previousCellValue - decayValueCell.value, 0)
    Else
        ' If decay value is not 1, check if the previous cell value is at least 1.
        If decayValueCell >= 1 Then
            ' If so, the output remains the same as the previous cell value.
            outputValue = previousCellValue
        Else
            ' If the previous cell value is less than 1, add the recovery value.
            ' Also ensure the result does not exceed 1.
            outputValue = Application.Min(previousCellValue + recoverValueCell.value, 1)
        End If
    End If

    Tally = outputValue

End Function

Sub Tally_Should_Decay_Value_By_Decay_Value_Cell_Test(debugService As DebugServiceClass)
    ' Arrange
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TalliesTest")
    
    ' Spreadsheet variables
    Dim decayValueCell As Range: Set decayValueCell = ws.Range("A1")
    Dim recoverValueCell As Range: Set recoverValueCell = ws.Range("A2")
    
    ' A decay value for the method
    decayValueCell.value = 0.2
    
    ' Not used in this test
    recoverValueCell.value = 0
    
    ' Cell range variables
    Dim state As Single: state = 1
    Dim refCell As Range: Set refCell = ws.Range("D1")
    Dim secondRefCell As Range: Set secondRefCell = ws.Range("E1")
    Dim firstActiveCell As Range: Set firstActiveCell = ws.Range("D2")
    Dim secondActiveCell As Range: Set secondActiveCell = ws.Range("E2")
    Dim firstValue As Range: Set firstValue = ws.Range("C2")
    
    ' Formulas that use Tally in the worksheet
    firstActiveCell.Formula = "=Tally(D1, $A$1, $A$2)"
    secondActiveCell.Formula = "=Tally(E1, $A$1, $A$2)"
    
    ' Sets the starting value
    firstValue.value = 1
    
    ' Denotes that the element in the game is on at t(n).
    refCell.value = state
    
    ' Denotes that the element in the game is on t(n+1).
    secondRefCell.value = state
    
    ' Force Excel to recalculate to ensure formulas are evaluated
    ws.Calculate
    
    ' Test results variables
    Dim firstCellResult As Single: firstCellResult = 0
    Dim secondCellResult As Single: secondCellResult = 0
    Dim expectedFirstCellResult As Single: expectedFirstCellResult = 0.8
    Dim expectedSecondCellResult As Single: expectedSecondCellResult = 0.6
    
    ' Act
    
    firstCellResult = firstActiveCell.value
    secondCellResult = secondActiveCell.value
     
    'Assert
    
    debugService.Assert "Tally_Should_Decay_Value_By_Decay_Value_Cell_Test - 1", firstCellResult = expectedFirstCellResult
    debugService.Assert "Tally_Should_Decay_Value_By_Decay_Value_Cell_Test - 2", secondCellResult = expectedSecondCellResult
End Sub

Sub Tally_Should_Not_Decay_Below_Zero_Test(debugService As DebugServiceClass)
    ' Arrange
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TalliesTest")
    
    ' Spreadsheet variables
    Dim decayValueCell As Range: Set decayValueCell = ws.Range("A1")
    Dim recoverValueCell As Range: Set recoverValueCell = ws.Range("A2")
    
    ' A decay value for the method
    decayValueCell.value = 0.2
    
    ' Not used in this test
    recoverValueCell.value = 0
    
    ' Cell range variables
    Dim state As Single: state = 1
    Dim refCell As Range: Set refCell = ws.Range("D1")
    Dim activeCell As Range: Set activeCell = ws.Range("D2")
    Dim firstValue As Range: Set firstValue = ws.Range("C2")
    
    ' Formulas that use Tally in the worksheet
    activeCell.Formula = "=Tally(D1, $A$1, $A$2)"
    
    ' Sets the starting value
    firstValue.value = 0
    
    ' Denotes that the element in the game is on at t(n).
    refCell.value = state
    
    ' Force Excel to recalculate to ensure formulas are evaluated
    ws.Calculate
    
    ' Test results variables
    Dim cellResult As Single: cellResult = 0
    Dim expectedCellResult As Single: expectedCellResult = 0
    
    ' Act
    
    cellResult = activeCell.value
     
    'Assert
    
    debugService.Assert "Tally_Should_Not_Decay_Below_Zero_Test", cellResult = expectedCellResult
End Sub

Sub Tally_Should_Not_Recover_Above_One_Test(debugService As DebugServiceClass)
    ' Arrange
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TalliesTest")
    
    ' Spreadsheet variables
    Dim decayValueCell As Range: Set decayValueCell = ws.Range("A1")
    Dim recoverValueCell As Range: Set recoverValueCell = ws.Range("A2")
    
    ' Not used in this test
    decayValueCell.value = 0
    
    ' The value to recover by
    recoverValueCell.value = 0.2
    
    ' Cell range variables
    Dim state As Single: state = 0
    Dim refCell As Range: Set refCell = ws.Range("D1")
    Dim activeCell As Range: Set activeCell = ws.Range("D2")
    Dim firstValue As Range: Set firstValue = ws.Range("C2")
    
    ' Formulas that use Tally in the worksheet
    activeCell.Formula = "=Tally(D1, $A$1, $A$2)"
    
    ' Sets the starting value
    firstValue.value = 1
    
    ' Denotes that the element in the game is on at t(n).
    refCell.value = state
    
    ' Force Excel to recalculate to ensure formulas are evaluated
    ws.Calculate
    
    ' Test results variables
    Dim cellResult As Single: cellResult = 0
    Dim expectedCellResult As Single: expectedCellResult = 1
    
    ' Act
    
    cellResult = activeCell.value
     
    'Assert
    
    debugService.Assert "Tally_Should_Not_Recover_Above_One_Test", cellResult = expectedCellResult
End Sub

