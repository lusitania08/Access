VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_OpcionesConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SecuenciaSql As String
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Private Sub Comando2_Click()
MarcarTodo
End Sub

Private Sub Comando3_Click()
DesMarcarTodo
End Sub

Private Sub Comando4_Click()
On Error GoTo Salerr
DoCmd.OpenReport "ConfReal", acViewPreview
Exit Sub
Salerr:
If Err.Number = 2501 Then
    Exit Sub
    Else
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error no controlado"
End If

End Sub

Private Sub Comando5_Click()
Dim sql As String
Dim n As String
If IsNull(Me.Lista0.Value) Or Me.Lista0.Value = "" Then Exit Sub
n = Me.Lista0.Value
sql = "SELECT InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.Nivel, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel, InfConfReal.Imprimir FROM InfConfReal WHERE  (((InfConfReal.Padre) Like '" & n & "*'));"
Form_Panel.SecuenciaSql = sql
DoCmd.OpenReport "ConfReal", acViewPreview
End Sub

Private Sub Lista0_Click()
Me.Comando5.Enabled = True
End Sub

Private Sub Lista0_DblClick(Cancel As Integer)
Dim rst As DAO.Recordset
Dim sql As String
Dim a
For Each a In Me.Lista0.ItemsSelected
    'comprobar si se imprime o no
    Dim imPri As Boolean
    imPri = Not Me.Lista0.Column(3, a)
    Dim Imprimir As String
    If imPri Then
        Imprimir = "True"
        Else
        Imprimir = "False"
    End If
    sql = "UPDATE InfConfReal SET InfConfReal.Imprimir =" & Imprimir & " WHERE  (((InfConfReal.Padre) Like'" & Me.Lista0.ItemData(a) & "*'));"
    Debug.Print sql
    Dim db As Database
    Set db = Me.Application.CurrentDb
    db.Execute sql
    Me.Lista0.Requery
Next
End Sub


Sub MarcarTodo()
Dim sql As String
    'comprobar si se imprime o no
    sql = "UPDATE InfConfReal SET InfConfReal.Imprimir = True;"
    Dim db As Database
    Set db = Me.Application.CurrentDb
    db.Execute sql
    Me.Lista0.Requery
End Sub

Sub DesMarcarTodo()
Dim sql As String
    'comprobar si se imprime o no
    sql = "UPDATE InfConfReal SET InfConfReal.Imprimir = False;"
    Dim db As Database
    Set db = Me.Application.CurrentDb
    db.Execute sql
    Me.Lista0.Requery

End Sub

Sub prueba()
    'Debug.Print sql
    Set rst = Me.Application.CurrentDb.OpenRecordset(sql)
    Do While Not rst.EOF
        rst.Edit
        If imPri Then
            rst!Imprimir = False
            Else
            rst!Imprimir = True
        End If
        rst.Update
        rst.MoveNext
    Loop

End Sub
