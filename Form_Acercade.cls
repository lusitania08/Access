VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Acercade"
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


Private Sub Form_Load()
Dim sql As String
sql = "SELECT TOP 1 Revisiones.Id, Revisiones.Revision FROM Revisiones ORDER BY Revisiones.Id DESC;"
Dim rst1 As DAO.Recordset
Set rst1 = Me.Application.CurrentDb.OpenRecordset(sql)
Do While Not rst1.EOF
    Me.Etiqueta10.Caption = "Version " & rst1!Revision
    rst1.MoveNext
Loop
Me.Etiqueta11.Caption = "Version de Microsoft Access " & Me.Application.SysCmd(acSysCmdAccessVer)
Me.Etiqueta9.Visible = True
Me.TimerInterval = tt
End Sub

Private Sub Form_Timer()
n = n + 1
If n >= 2 Then
    Me.Etiqueta9.Visible = False
    Me.TimerInterval = 0
    n = 0
End If
End Sub

Private Sub Imagen4_Click()
Me.Etiqueta9.Visible = True
Me.Etiqueta9.Caption = "Logistica, es la ciencia del planeamiento, ejecucion del movimiento y sostenimiento de las fuerzas"
Me.TimerInterval = tt

End Sub


Private Sub Imagen4_DblClick(Cancel As Integer)
Me.Etiqueta9.Visible = True
Me.Etiqueta9.Caption = "Et siluit terra, in conspectu eius." & vbNewLine & "'Y la tierra enmudecio en su presencia.'"
Me.TimerInterval = tt
End Sub
