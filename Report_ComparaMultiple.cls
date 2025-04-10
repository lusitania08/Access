VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_ComparaMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Public D As String
Dim k As Integer
Public nn As String
Public Aus As Integer
Private Sub Detalle_Format(Cancel As Integer, FormatCount As Integer)
If Me.Falta.Value Then
    Me.NSerie.Visible = False
    Else
    Me.NSerie.Visible = True
End If
If Me.Alternativo.Value Then
    Me.Etiqueta29.Visible = True
    Else
    Me.Etiqueta29.Visible = False
End If
If Not Me.Falta.Value Then
    If Not IsNull(Me.Niv.Value) Then
        NV = Int(Me.Niv.Value)
        l = NV * 150
        Else
        NV = 0
        l = 0
    End If
    If NV = 1 Then
        Me.NOCt.FontBold = True
       Else
        Me.NOCt.FontBold = False
    End If
    Me.NOCt.Left = 0 + l
    Else
    Me.NOCt.FontBold = False
End If
End Sub

Private Sub EncabezadoDelGrupo2_Format(Cancel As Integer, FormatCount As Integer)
If Not Me.Falta.Value Then
    Me.Etiqueta9.Caption = "Articulo patron. Componentes cargados"
    Me.NSerie.Visible = True
    Else
    Me.Etiqueta9.Caption = "Resultado de comparacion con Nº de Serie Patron. Componentes ausentes."
    Me.NSerie.Visible = False
End If

End Sub

Private Sub EncabezadoDelGrupo4_Format(Cancel As Integer, FormatCount As Integer)
    If k = 0 Then
        D = Me.Descripcion.Value
    End If
    k = k + 1

Dim n, s As String
n = Left(Me.Padre, 13)
nn = n
s = Right(Me.Padre.Value, Len(Padre) - 14)
Me.Etiqueta11.Caption = D & " (" & n & ") - " & s
End Sub




Private Sub PieDelInforme_Format(Cancel As Integer, FormatCount As Integer)
Dim sql As String
sql = "SELECT InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.CodSistema FROM InfConfReal WHERE (((InfConfReal.Padre) Like '" & nn & "*'));"
'Debug.Print sql
Dim nreg, Nfalta As Integer
Dim r As DAO.Recordset
Set r = Me.Application.CurrentDb.OpenRecordset(sql)
If Not r.EOF Then r.MoveLast
nreg = r.RecordCount
sql = "SELECT ComparaNOC.Padre, ComparaNOC.Falta, ComparaNOC.Alternativo FROM ComparaNOC WHERE (((ComparaNOC.Padre) Like '" & nn & "*') AND ((ComparaNOC.Falta)=True) AND ((ComparaNOC.Alternativo)=False));"
Set r = Me.Application.CurrentDb.OpenRecordset(sql)
'Debug.Print sql
If Not r.EOF Then r.MoveLast
Nfalta = r.RecordCount
Dim res As Double
res = 1 - (Nfalta / nreg)
Me.Etiqueta31.Caption = "La comparacion nos indica que tiene un porcentaje de carga de " & Format(res, "0%")
End Sub

Private Sub Report_NoData(Cancel As Integer)
Cancel = True
End Sub

