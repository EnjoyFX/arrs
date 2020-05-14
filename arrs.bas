Attribute VB_Name = "arrs"
Option Explicit

Function add_to_arr(arr As Variant, element As Variant, Optional ByVal base As Integer = 0) As Variant

    Dim the_max As Long

    If IsEmpty(arr) Then
        ReDim arr(base To base)
        arr(base) = element
    Else
        the_max = UBound(arr) + 1
        ReDim Preserve arr(LBound(arr) To the_max)
        arr(the_max) = element
    End If

    add_to_arr = arr

End Function

