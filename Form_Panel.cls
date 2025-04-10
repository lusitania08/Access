VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Public SecuenciaSql As String

Private Sub Form_Load()
Dim rst As DAO.Recordset
Set rst = Me.Application.CurrentDb.OpenRecordset("ConfReal")
If rst.RecordCount = 0 Then
    MsgBox "No existen datos en la aplicacion. Debe extraer primero el fichero de 'Configuracion Real Breve de SIGLE' y posteriormente importarlo.", vbInformation + vbOKOnly, "No existen datos"
    Marco9.Value = 3
End If
Marco9_AfterUpdate
Dim sql As String
sql = "SELECT TOP 1 Revisiones.Id, Revisiones.Revision FROM Revisiones ORDER BY Revisiones.Id DESC;"
Dim rst1 As DAO.Recordset
Set rst1 = Me.Application.CurrentDb.OpenRecordset(sql)
Do While Not rst1.EOF
    Me.Etiqueta20.Caption = "Version " & rst1!Revision
    rst1.MoveNext
Loop

End Sub

Private Sub Marco9_AfterUpdate()
Dim b As Integer
b = Me.Marco9.Value
Select Case b
    Case 1
    'configuracion real
    Me.Secundario0.SourceObject = "ConfReal"
    Case 2
    'Opciones de configuracion
    Me.Secundario0.SourceObject = "OpcionesConf"
    Case 3
    'Importar datos
    Me.Secundario0.SourceObject = "Importar"
    Case 4
    'Comparar configuraciones
    Me.Secundario0.SourceObject = "Comparar"
    Case 5
    'Ayuda
    Me.Secundario0.SourceObject = "Help"
    Case 6
    'Ayuda
    Me.Secundario0.SourceObject = "Acercade"
End Select

End Sub
