Attribute VB_Name = "ModCollections"
Option Explicit

Private vItem As Variant
Private UniqueNum&

Public ColUndo As New Collection
Public ColRedo As New Collection
Public Sub DeleteCollections()

    For Each vItem In ColUndo
        ColUndo.Remove 1
    Next

    For Each vItem In ColRedo
        ColRedo.Remove 1
    Next

    UniqueNum = 0

End Sub
Public Sub UpdateUndo()
    UniqueNum = UniqueNum + 1
        Form1.picReal.Picture = Form1.picReal.Image
        ColUndo.Add Item:=Form1.picReal.Picture, Key:=CStr(UniqueNum)
    Form1.cmdUndo.Visible = ColUndo.Count > 1
    Form1.cmdRedo.Visible = ColRedo.Count > 0
    If Dirty = False Then Exit Sub
    Dirty = True
End Sub
Public Sub DoUnDo()
    ColRedo.Add ColUndo.Item(ColUndo.Count)
    ColUndo.Remove ColUndo.Count
    Form1.picReal.Picture = ColUndo.Item(ColUndo.Count)
    Form1.picReal.Refresh
    With Form1
        .PaintDown
        .Refresh
   End With
    Form1.cmdUndo.Visible = ColUndo.Count > 1
    Form1.cmdRedo.Visible = ColRedo.Count > 0
End Sub
Public Sub DoReDo()
Form1.cmdRedo.Visible = ColRedo.Count > 0
    ColUndo.Add ColRedo.Item(ColRedo.Count)
    ColRedo.Remove ColRedo.Count
        Form1.picReal.Picture = ColUndo.Item(ColUndo.Count)
    Form1.picReal.Refresh
    With Form1
        .PaintDown
        .Refresh
    End With
 Form1.cmdRedo.Visible = ColRedo.Count > 0
 Form1.cmdUndo.Visible = ColUndo.Count > 1
End Sub
Public Sub ClearRedo()
    For Each vItem In ColRedo
        ColRedo.Remove 1
    Next
End Sub
