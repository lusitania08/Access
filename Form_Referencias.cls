VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Referencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Me.Lista0.RowSource = ""
Dim IntRef As Integer
Dim StrRef As String
Dim F As String
Dim ff As Boolean
For IntRef = 1 To References.Count
     If References(IntRef).IsBroken Then
        F = " (** FALTA **)"
        ff = True
        Else
        F = " "
        ff = False
     End If
    StrRef = References(IntRef).Name & F & " - " & References(IntRef).FullPath
    Me.Lista0.AddItem StrRef
Next
End Sub

