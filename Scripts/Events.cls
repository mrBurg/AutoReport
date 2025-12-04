VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЦяКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_OLEObjectEvents(ByVal Obj As OLEObject, ByVal EventName As String)
    If EventName = "Click" Then
        If TypeName(Obj.Object) = "CheckBox" Then

            Dim cb As CheckBox
            Set cb = Obj.Object

            Dim r As Range
            Set r = Obj.TopLeftCell

            If cb.Value = True Then
                r.Value = "Показати"
            Else
                r.Value = "Приховати"
            End If
        End If
    End If
End Sub

