VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Public Linea, Caracter As Integer
Public StrLinea As String
Dim nd As String

Private Sub Form_Load()
Dim ndx As Node
Set ndx = Me.TreeView0.Nodes.Add(, , "raiz", "Ayuda Conf Real")
nd = ndx.Key
Set ndx = Me.TreeView0.Nodes.Add("raiz", tvwChild, "Importar", "Importar Configuracion")
Set ndx = Me.TreeView0.Nodes.Add("raiz", tvwChild, "ConfReal", "Arbol de Configuracion Real")
Set ndx = Me.TreeView0.Nodes.Add("raiz", tvwChild, "Impconf", "Imprimir Configuracion")
Set ndx = Me.TreeView0.Nodes.Add("raiz", tvwChild, "Comparar", "Comparar configuraciones")
ndx.EnsureVisible
End Sub

Private Sub TreeView0_NodeClick(ByVal Node As Object)
If Node.Key = nd Then
    Exit Sub
    Else
    Me.Txt.Value = ""
    Caracter = 0
End If
    Dim rst As DAO.Recordset
    Set rst = Me.Application.CurrentDb.OpenRecordset("TblHelp")
    If Not rst.BOF Then rst.MoveFirst
    Do While Not rst.EOF
        If Node.Key = rst!tipoAyuda Then
            StrLinea = rst!T_linea
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
'If Me.Verificacion15.Value Then
    'Me.TimerInterval = 50
    'Else
    Me.TimerInterval = 0
    Me.Txt.Value = StrLinea
'End If
nd = Node.Key
End Sub



Private Sub Form_Open(Cancel As Integer)
Me.TimerInterval = 0
Caracter = 1
'Me.Verificacion15.Value = False
End Sub

Private Sub Form_Timer()
EscribirBP

End Sub

Sub EscribirBP()
Dim StrTemp As String
Dim I, k As Integer
Const Inicio = 1000
I = Len(StrLinea)
    StrTemp = Left(StrLinea, Caracter)
k = Len(StrTemp)

If k < I Then
    Me.Txt.Value = StrTemp
    Caracter = Caracter + 5
    Else:
    Me.TimerInterval = 0
End If
End Sub




