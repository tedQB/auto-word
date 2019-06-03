VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "chk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents obj As MSForms.CheckBox
Attribute obj.VB_VarHelpID = -1

Private Sub obj_Change()
    Dim n
    n = Val(Replace(obj.Name, "CheckBox", ""))
    Sheets(1).Range("B" & n).Interior.ColorIndex = IIf(obj.value, 13, 0)
End Sub


