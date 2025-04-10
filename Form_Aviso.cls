VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Aviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Dim n As Integer
Const tt As Integer = 5000

Private Sub Form_Timer()
n = n + 1
If n = 2 Then
    Me.Etiqueta9.Visible = False
    Me.TimerInterval = 0
End If
End Sub

Private Sub Imagen4_Click()
Me.Etiqueta9.Visible = True
Me.TimerInterval = tt

End Sub
