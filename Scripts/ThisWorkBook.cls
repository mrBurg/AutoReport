VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    Set ws = Sheets("AutoReport")
    
    ws.Activate
    
    ModuleControls.FillAll ws
End Sub
