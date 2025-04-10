VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Compara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Public CargNOC1, CargNOC2 As Integer
Public AusNOC1, AusNOC2 As Integer
Public NOC1, NOC2 As String
Public Carg, Aus, n As Integer

Private Sub Detalle_Format(Cancel As Integer, FormatCount As Integer)
Dim l As Integer
Dim NV As Integer
If Not Me.Falta.Value Then
    Carg = Carg + 1
    Else
    Aus = Aus + 1
End If
If Me.Alternativo.Value Then
    Carg = Carg + 1
    Aus = Aus - 1
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
    Me.Etiqueta9.Caption = "Resultado de la comparacion. Componentes cargados"
    Me.NSerie.Visible = True
    Else
    Me.Etiqueta9.Caption = "Resultado de la comparacion. Componentes ausentes"
    Me.NSerie.Visible = False
End If

End Sub




Private Sub EncabezadoDelGrupo4_Format(Cancel As Integer, FormatCount As Integer)
    If n = 0 Then
        NOC1 = Me.Etiqueta11.Value
        Else
        NOC2 = Me.Etiqueta11.Value
    End If
End Sub

Private Sub PieDelGrupo1_Format(Cancel As Integer, FormatCount As Integer)
    Dim porc1, porc2 As Double
If n = 0 Then
    CargNOC1 = Carg
    AusNOC1 = Aus
    Else
    CargNOC2 = Carg
    AusNOC2 = Aus
    If (CargNOC2 + AusNOC2) > 0 Then
        porc2 = CargNOC2 / (CargNOC2 + AusNOC2)
        Else
        porc2 = 1
    End If
End If
n = n + 1
    Carg = 0
    Aus = 0
    porc1 = CargNOC1 / (CargNOC1 + AusNOC1)
    Me.Etiqueta17.Caption = "NOC 1: " & NOC1
    Me.Etiqueta18.Caption = CargNOC1 & " Componentes cargados"
    Me.Etiqueta19.Caption = AusNOC1 & " Componentes ausentes con respecto al NOC 2"
    Me.Etiqueta26.Caption = "Porcentaje probable de carga de configuracion real: " & Format(porc1, "0%")
    Me.Etiqueta20.Caption = "NOC 2: " & NOC2
    Me.Etiqueta21.Caption = CargNOC2 & " Componentes cargados"
    Me.Etiqueta22.Caption = AusNOC2 & " Componentes ausentes con respecto al NOC 1"
    Me.Etiqueta27.Caption = "Porcentaje probable de carga de configuracion real: " & Format(porc2, "0%")
End Sub

Private Sub Report_NoData(Cancel As Integer)
Cancel = True
End Sub
