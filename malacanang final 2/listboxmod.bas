Attribute VB_Name = "Listboxmod"
Public Sub xListKillDupes(listbox As listbox)
'Kills duplicate items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.list(Search1&) = listbox.list(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub

