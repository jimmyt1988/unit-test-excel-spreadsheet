Option Explicit

Private debugService As DebugServiceClass

Private Sub Class_Initialize()
    Set debugService = New DebugServiceClass
End Sub

Public Sub RunAllTests()

    Tally_Should_Decay_Value_By_Decay_Value_Cell_Test debugService
    Tally_Should_Not_Decay_Below_Zero_Test debugService
    Tally_Should_Not_Recover_Above_One_Test debugService
    
    debugService.DisplayAssertions
    
End Sub
