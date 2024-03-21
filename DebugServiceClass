Option Explicit

Private assertionsMade As Long
Private assertionsPassed As Long
Private passedTests As Collection
Private failedTests As Collection

Private Sub Class_Initialize()
    Set passedTests = New Collection
    Set failedTests = New Collection
End Sub

Public Sub Assert(nameOfTest As String, condition As Boolean)
    assertionsMade = assertionsMade + 1
    If condition Then
        assertionsPassed = assertionsPassed + 1
        passedTests.Add nameOfTest
    Else
        failedTests.Add nameOfTest
    End If
End Sub

Public Sub DisplayAssertions()
    Dim message As String
    message = "Total Assertions Made: " & assertionsMade & vbCrLf & _
              "Assertions Passed: " & assertionsPassed & vbCrLf & vbCrLf

    If failedTests.Count > 0 Then
        message = message + "Failed Tests:" & vbCrLf & "- " & Join(CollectionToArray(failedTests), vbCrLf & "- ")
    End If

    MsgBox message, vbInformation, "Assertion Summary"
End Sub

Public Function TotalAssertions() As Long
    TotalAssertions = assertionsMade
End Function

Public Function PassedAssertions() As Long
    PassedAssertions = assertionsPassed
End Function

' Helper function to convert a collection to an array
Private Function CollectionToArray(coll As Collection) As String()
    If coll.Count = 0 Then
        CollectionToArray = Split("") ' Return an empty array if the collection is empty.
    Else
        Dim arr() As String
        ReDim arr(1 To coll.Count)
        Dim i As Long
        For i = 1 To coll.Count
            arr(i) = coll(i)
        Next i
        CollectionToArray = arr
    End If
End Function


