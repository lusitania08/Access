VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Historial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Me.Lista0.RowSource = ""
Dim rst As DAO.Recordset
Dim RowsourceLista As String
Set rst = Me.Application.CurrentDb.OpenRecordset("Local")
If Not rst.BOF Then rst.MoveFirst
RowsourceLista = "Archivo;Fecha de Creacion;Fecha Importacion;Unidad;"
Do While Not rst.EOF
    RowsourceLista = RowsourceLista & rst!Archivo & ";" & rst!fechacreacion & ";" & rst!fechaImportacion & ";" & rst!Unidad & ";"
    rst.MoveNext
Loop
    Me.Lista0.RowSource = RowsourceLista
End Sub

