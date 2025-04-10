VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_ConfReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************


Private Sub Detalle_Format(Cancel As Integer, FormatCount As Integer)

    Dim l As Integer
    Dim NV As Integer
    Dim plus As String
    plus = ""
    If Me.CaComp.Value = 0 And Me.NSerie.Value <> "" Then
        Me.CaComp.Visible = False
        Else
        Me.CaComp.Visible = True
    End If
    If Not IsNull(Me.Nivel.Value) Then
        NV = Int(Me.Nivel.Value)
        l = NV * 100
        
        If Nivel > 1 Then
            plus = "|"
        Else
            plus = ""
        End If
        
        If NV > 0 Then
            Me.txtNOC.Value = plus & String(NV - 1, Chr(133)) & Me.NOC.Value
            Else
            Me.txtNOC.Value = Me.NOC.Value
        End If
        
        Else
        Me.txtNOC.Value = Me.NOC.Value
        NV = 0
        l = 0
    End If
        If NV = 1 Or NV = 0 Then
            Me.txtNOC.FontBold = True
            'Me.Etiqueta13.Visible = False
           Else
            Me.txtNOC.FontBold = False
            'Me.Etiqueta13.Visible = True
            'Me.Etiqueta13.Left = 0 + l - 150
    
            'Me.txtNOC.Value = "+" & Me.NOC.Value
        End If
            Me.txtNOC.Left = 0 + l
     'Me.Descripcion.Left = 1701 + l
     'Me.NSerie.Left = 6522 + l
     'Me.CodSistema.Left = 8223 + l
End Sub

Private Sub Report_Close()
Form_Panel.SecuenciaSql = "SELECT  InfConfReal.Orden,InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel, InfConfReal.Imprimir FROM InfConfReal WHERE (((InfConfReal.Imprimir)=True)); "
End Sub

Private Sub Report_NoData(Cancel As Integer)
    Cancel = True

End Sub

Private Sub Report_Open(Cancel As Integer)
If Form_Panel.SecuenciaSql <> "" Then
    Me.RecordSource = Form_Panel.SecuenciaSql
    SecuenciaSql = ""
    NoVerRegistro = False
End If

End Sub
