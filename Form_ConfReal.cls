VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ConfReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Public NoCSel As String
Public NOCPadre As String
Public TxtNOCPadre As String
Sub ArticulosTRW()
'Cargar articulos en treeview
On Error GoTo ErrGo

Dim seisdigitos As String

Dim rst As DAO.Recordset
Dim ndx As Node
Dim ndx1 As Node
Set ndx = Me.Trw.Nodes.Add(, , "raiz", "Configuracion Real")
ndx.Bold = True
Dim Strsql As String
Dim cd, ser, Strkey, StrPadre As String
Dim ExisteLaKey As Boolean
'Cargar articulos
Strsql = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.dsArtSuperior FROM ConfReal ORDER BY ConfReal.dsArtSuperior;"
Set rst = Me.Application.CurrentDb.OpenRecordset(Strsql)
If Not rst.BOF Then rst.MoveFirst
Me.Application.SysCmd acSysCmdClearStatus
Me.Application.SysCmd acSysCmdInitMeter, "Cargando articulos...", rst.RecordCount
Do While Not rst.EOF
    If rst!cdArtSuperior = "" Or rst!cdArtSuperior = ";" Then
        GoTo nuevoRec
    End If
    Strkey = "Articulo+" & rst!cdArtSuperior
    Set ndx = Me.Trw.Nodes.Add("raiz", tvwChild, Strkey, rst!dsArtSuperior & " (" & rst!cdArtSuperior & ")")
        ndx.Bold = True
nuevoRec:
    rst.MoveNext
Loop
'Cargar articulos seriados de primer nivel
Strsql = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.dsArtSuperior, ConfReal.cdSistemaSuperior, ConfReal.vaNroSerieSuperior FROM ConfReal ORDER BY ConfReal.cdArtSuperior;"
Set rst = Me.Application.CurrentDb.OpenRecordset(Strsql)
If Not rst.BOF Then rst.MoveFirst
Me.Application.SysCmd acSysCmdClearStatus
Me.Application.SysCmd acSysCmdInitMeter, "Cargando articulos seriados de primer nivel...", rst.RecordCount
Do While Not rst.EOF
        If rst!cdArtSuperior = "" Or rst!cdArtSuperior = ";" Then
            GoTo nuevoRec1
        End If
        
       
        
    Strkey = rst!cdArtSuperior & "+" & rst!vaNroSerieSuperior & "+" & Devolver6Digitos(rst!cdSistemaSuperior)
    StrPadre = "Articulo+" & rst!cdArtSuperior
    Set ndx = Me.Trw.Nodes.Add(StrPadre, tvwChild, Strkey, Devolver6Digitos(rst!cdSistemaSuperior) & " >> " & rst!dsArtSuperior & " (" & rst!cdArtSuperior & ")" & " - " & rst!vaNroSerieSuperior)
        'ndx.Bold = True
        ndx.Tag = StrPadre
        
    Me.Application.SysCmd acSysCmdUpdateMeter, rst.AbsolutePosition
nuevoRec1:
rst.MoveNext
Loop
'Cargar resto de articulos
Strsql = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.vaNroSerieSuperior, ConfReal.cdSistemaSuperior, ConfReal.cdArt, ConfReal.dsArt, ConfReal.cdSistema, ConfReal.vaNroSerie, ConfReal.cdArtBajo, ConfReal.dsArtBajo, ConfReal.cdSistemaBajo, ConfReal.vaNroSerieBajo, ConfReal.caComp, ConfReal.cdEsNroSerie FROM ConfReal WHERE (((ConfReal.cdArtBajo)<>'')) ORDER BY ConfReal.cdArtSuperior, ConfReal.vaNroSerieSuperior, ConfReal.cdSistemaBajo;"
'Debug.Print Strsql
Set rst = Me.Application.CurrentDb.OpenRecordset(Strsql)
If Not rst.BOF Then rst.MoveFirst
Me.Application.SysCmd acSysCmdClearStatus
Me.Application.SysCmd acSysCmdInitMeter, "Cargando articulos seriados de segundo nivel...", rst.RecordCount
Do While Not rst.EOF
    If rst!vaNroSerieBajo <> "" Then
        Strkey = rst!cdArtBajo & "+" & rst!vaNroSerieBajo & "+" & Devolver6Digitos(rst!cdSistemaBajo)
        StrPadre = rst!cdArt & "+" & rst!vaNroSerie & "+" & Devolver6Digitos(rst!cdSistema)

        Set ndx = Me.Trw.Nodes.Add(StrPadre, tvwChild, Strkey, Devolver6Digitos(rst!cdSistemaBajo) & " >> " & rst!dsArtBajo & " (" & rst!cdArtBajo & ")" & " - " & rst!vaNroSerieBajo)
        ndx.Tag = StrPadre
    Else
Vuelta:
        Strkey = "Baj+" & Str(CDbl(Now)) & Rnd() 'rst!cdArtBajo & "+" & rst!vaNroSerieBajo & "+" & rst!cdSistemaBajo)
        
        StrPadre = rst!cdArt & "+" & rst!vaNroSerie & "+" & Devolver6Digitos(rst!cdSistema)
        
        'Debug.Print rst!cdSistemaBajo

        Set ndx = Me.Trw.Nodes.Add(StrPadre, tvwChild, Strkey, Devolver6Digitos(rst!cdSistemaBajo) & " >> " & rst!dsArtBajo & " (" & rst!cdArtBajo & ")" & " - Cant: " & rst!CaComp)
        ndx.Tag = StrPadre
    End If
NuevoReg:
    Me.Application.SysCmd acSysCmdUpdateMeter, rst.AbsolutePosition
    rst.MoveNext
Loop
Me.Trw.Nodes(1).Expanded = True
Me.Application.SysCmd acSysCmdRemoveMeter
'ndx.EnsureVisible
rst.Close
Exit Sub
ErrGo:
Dim er As Single
er = Err.Number
Select Case er
    Case 35601
    StrPadre = rst!cdArtSuperior & "+" & rst!vaNroSerieSuperior & "+" & Devolver6Digitos(rst!cdSistemaSuperior)
    Strkey = rst!cdArt & "+" & rst!vaNroSerie & "+" & Devolver6Digitos(rst!cdSistema)
    Set ndx = Me.Trw.Nodes.Add(StrPadre, tvwChild, Strkey, rst!dsArtBajo & " (" & rst!cdArtBajo & ")" & " - " & rst!vaNroSerie)
    Resume Vuelta
    'Debug.Print Err.
    Case 35602
    Resume NuevoReg
    
    Case Else
    MsgBox Err.Description & " - " & Err.Number
    Resume NuevoReg

End Select
End Sub

Function Devolver6Digitos(codigoSistema As String) As String
' Devolver 6 digitos del codigo de sisterma
Dim codigo As String
codigo = codigoSistema
If Len(codigo) < 6 Then

    Dim numeroceros As Integer
    numeroceros = 6 - Len(codigo)
    Devolver6Digitos = Format(codigo, "000000")
    Exit Function
    Else
    Devolver6Digitos = codigoSistema
    Exit Function

End If

End Function

Private Sub Comando1_Click()
Me.Trw.Nodes.Clear
ArticulosTRW
End Sub

Private Sub Comando14_Click()
'If NoCSel = "" Then
    'MsgBox "Debe seleccionar un CUF,s para poder imprimir el listado de componentes.", vbInformation + vbOKOnly, "Seleccion de CUF"
    'Exit Sub
'End If
Dim sql As String
sql = "SELECT  InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.Nivel, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel, InfConfReal.Imprimir, InfConfReal.Padre1 FROM InfConfReal WHERE  (((InfConfReal.Padre) ='" & NoCSel & "'));"
'Debug.Print sql
If Right(sql, 5) = "+'));" Then
    sql = Left(sql, Len(sql) - 5) + "'));"
End If

Dim rst As DAO.Recordset
Set rst = Me.Application.CurrentDb.OpenRecordset(sql)

'Debug.Print sql

'Dim n As Integer
'For n = 0 To 255
'Debug.Print "n=" & n & " > " & Chr(n)
'Next

If rst.RecordCount = 0 Then
    rst.Close
    MsgBox "Debe seleccionar un CUF,s para poder imprimir el listado de componentes.", vbInformation + vbOKOnly, "Seleccion de CUF"
    Exit Sub
End If
Form_Panel.SecuenciaSql = sql
'Debug.Print sql
DoCmd.OpenReport "ConfReal", acViewPreview
End Sub


Private Sub Comando20_Click()
If NOCPadre = "" Then Exit Sub
Dim sql As String
Dim nreg, Nreg2 As Integer
Dim Mensaje As String
sql = "SELECT InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.Nivel, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel, InfConfReal.Imprimir FROM InfConfReal WHERE (((InfConfReal.Padre) Like '" & NOCPadre & "*'));"
'Debug.Print sql
Dim rst As DAO.Recordset
Set rst = Me.Application.CurrentDb.OpenRecordset(sql)
If rst.RecordCount = 0 Then
    rst.Close
End If
rst.MoveLast
nreg = rst.RecordCount
rst.Close
sql = "SELECT InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.Nivel, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel, InfConfReal.Imprimir FROM InfConfReal WHERE (((InfConfReal.Padre) Like '" & NOCPadre & "*'));"
'Debug.Print sql
Set rst = Me.Application.CurrentDb.OpenRecordset(sql)
rst.MoveLast
Nreg2 = rst.RecordCount
Mensaje = "Cantidad del articulo " & TxtNOCPadre & ": " & nreg & " registros" & vbNewLine & "Existen un total de " & Nreg2 & " componentes cargados. " & vbNewLine & "Una media de " & Format((Nreg2 - nreg) / nreg, "0.0") & " componentes por articulo principal."

MsgBox Mensaje, vbInformation + vbOKOnly, "Informe del articulo"
End Sub

Private Sub Form_Load()
Me.Trw.Nodes.Clear
ArticulosTRW
End Sub

Private Sub Texto12_AfterUpdate()
If Me.Texto12.Value = "" Or IsNull(Me.Texto12.Value) Then Exit Sub
Buscar (Me.Texto12.Value)
Me.Texto12.Value = ""
End Sub
Sub Buscar(busqueda As String)
'Me.trw.Nodes.nodeselected.Expanded = False
Dim ndx As Node
Dim n As Integer
Dim longitud As Integer
Dim LongBus As Integer
LongBus = Len(busqueda)
Dim Txt, txtFragmento As String
For Each ndx In Me.Trw.Nodes
    'If ndx.Expanded Then ndx.Expanded = False
    Txt = ndx.Text
    longitud = Len(Txt)
    For n = 1 To longitud
        txtFragmento = Mid(Txt, n, LongBus)
        If txtFragmento = busqueda Then
            ndx.Selected = True
            ndx.EnsureVisible
            Dim Respuesta
            Respuesta = MsgBox("¿Desea continuar con la busqueda?", vbQuestion + vbYesNo, "Buscar")
            If Respuesta = vbNo Then
                Exit Sub
            End If
            
        End If
    Next
Next
End Sub




Private Sub trw_NodeClick(ByVal Node As Object)
If Node.Index = 1 Then Exit Sub
Dim ndx As Node
Set ndx = Me.Trw.Nodes(Node.Index)
If Left(Node.Key, 3) <> "Art" Then
    'NoCSel = Left(Node.Key, Len(Node.Key) - 6)
    DeterminaPadre Node
    'Debug.Print NoCSel
    Me.Comando14.Enabled = True
    Else
    NoCSel = ""
    NOCPadre = Right(Node.Key, 13)
    TxtNOCPadre = Node.Text
    Me.Comando14.Enabled = False
End If
'If ndx.Bold = False Then
'    Else
'End If
End Sub


Sub DeterminaPadre(ByVal nd As Node)
'Determina recursivamente el nodo padre
If nd Is Nothing Then
    Exit Sub
End If
Dim Padre As String
Padre = Left(nd.Parent.Key, 8)

'Debug.Print "Recursivo " & nd.Key, codSist
If Padre = "Articulo" Then
    If Len(nd.Key) > 6 Then
        NoCSel = Left(nd.Key, Len(nd.Key) - 6)
        NOCPadre = Right(nd.Parent.Key, 13)
        TxtNOCPadre = nd.Parent.Text
        Exit Sub
        Else
        DeterminaPadre nd.Parent
    End If
End If
End Sub




