VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Comparar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************

Private Sub Comando15_Click()
Comando7_Click
DoCmd.OpenReport "Compara", acViewPreview

End Sub

Private Sub Comando16_Click()
On Error GoTo ErrSal
Dim Art1 As String
Art1 = Me.Lista0.Value & "+" & Me.Lista2.Value
Dim rst As DAO.Recordset
Dim sql As String
Dim NOCComp As String
'NOC,s de la comparacion
sql = "SELECT DISTINCT InfConfReal.Padre FROM InfConfReal WHERE (((InfConfReal.Padre)<>'" & Art1 & "'))  AND (((InfConfReal.Padre) Like '" & Left(Art1, 13) & "*'));"
Set rst = Me.Application.CurrentDb.OpenRecordset(sql)
Dim Primero As Boolean
Do While Not rst.EOF
    NOCComp = rst!Padre
    If Not Primero Then
        ComprobarConfMultiple3 Art1, NOCComp, True
        Primero = True
        Else
        ComprobarConfMultiple3 Art1, NOCComp, False
    End If
    rst.MoveNext
Loop
DoCmd.OpenReport "ComparaMultiple", acViewPreview
Exit Sub
ErrSal:
If Err.Number = 2501 Then
    Exit Sub
    Else
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error no controlado"
End If
End Sub

Private Sub Comando7_Click()
Dim Art1, Art2 As String
Art1 = Me.Lista0.Value & "+" & Me.Lista2.Value
Art2 = Me.Lista0.Value & "+" & Me.Lista5.Value
Me.Lista8.RowSource = ""
Me.Lista10.RowSource = ""
ComprobarConf3 Art1, Art2, Lista10, True
ComprobarConf3 Art2, Art1, Lista8, False
If Me.Lista10.ListCount = 0 Then
    Me.Lista10.AddItem ("Todos los articulos de " & Me.Lista2.Value & ", estan presentes en " & Me.Lista5.Value)
End If
If Me.Lista8.ListCount = 0 Then
    Me.Lista8.AddItem ("Todos los articulos de " & Me.Lista5.Value & ", estan presentes en " & Me.Lista2.Value)
End If
Me.Comando15.Enabled = True
Me.Etiqueta9.Caption = "Componentes de " & Me.Lista5.Value & " que faltan en " & Me.Lista2.Value
Me.Etiqueta11.Caption = "Componentes de " & Me.Lista2.Value & " que faltan en " & Me.Lista5.Value
End Sub

Private Sub Lista0_AfterUpdate()
Dim Sql1 As String
Dim n As String
If IsNull(Me.Lista0.Value) Then Exit Sub
n = Me.Lista0.Value
Sql1 = "SELECT DISTINCT ConfReal.vaNroSerieSuperior, ConfReal.cdArtSuperior  FROM ConfReal WHERE (((ConfReal.cdArtSuperior)='" & n & "'));"
Me.Lista2.RowSource = Sql1
Me.Lista2.Value = ""
Me.Lista5.Value = ""
Me.Lista8.RowSource = ""
Me.Lista10.RowSource = ""
If Me.Lista2.ListCount = 1 Then
    Me.Lista8.AddItem ("No es posible la comparacion. Solo existe un Nº de serie.")
    Me.Comando7.Enabled = False
    Me.Comando16.Enabled = False
End If
Me.Lista2.SetFocus
End Sub


Private Sub Lista2_AfterUpdate()
Dim Sql1 As String
Dim n As String
If IsNull(Me.Lista0.Value) Then Exit Sub
n = Me.Lista0.Value
'SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.vaNroSerieSuperior FROM ConfReal WHERE (((ConfReal.cdArtSuperior)="1005123505888") AND ((ConfReal.vaNroSerieSuperior)<>"ET 004060HK"));
Sql1 = "SELECT DISTINCT ConfReal.vaNroSerieSuperior, ConfReal.cdArtSuperior FROM ConfReal WHERE (((ConfReal.cdArtSuperior)='" & n & "') AND ((ConfReal.vaNroSerieSuperior)<>'" & Me.Lista2.Value & "'));"
Me.Lista5.RowSource = Sql1
Me.Lista5.Value = ""
If Me.Lista2.Value <> "" And Me.Lista2.ListCount > 1 Then
    Me.Comando16.Enabled = True
    Else
    Me.Comando16.Enabled = False
End If

If Me.Lista2.Value <> "" And Me.Lista5.Value <> "" Then
    Me.Comando7.Enabled = True
    Else
    Me.Comando7.Enabled = False
    Me.Comando15.Enabled = False
    Me.Lista8.RowSource = ""
    Me.Lista10.RowSource = ""
    Me.Etiqueta9.Caption = "Componentes de NOC 2 que faltan en NOC 1"
    Me.Etiqueta11.Caption = "Componentes de NOC 1 que faltan en NOC 2"
End If
Me.Lista5.SetFocus

End Sub

Private Sub Lista5_AfterUpdate()
If Me.Lista5.Value <> "" Then
    Me.Comando7.Enabled = True
    Me.Comando7.SetFocus
    Else
    Me.Lista8.RowSource = ""
    Me.Lista10.RowSource = ""
    Me.Comando7.Enabled = False
    Me.Comando15.Enabled = False
    Me.Etiqueta9.Caption = "Componentes de NOC 2 que faltan en NOC 1"
    Me.Etiqueta11.Caption = "Componentes de NOC 1 que faltan en NOC 2"
End If
End Sub
Sub ComprobarConfMultiple3(ByVal A1 As String, ByVal A2 As String, ByVal BorrarReg As Boolean)
'comparar un articulo patron con los demas articulos efectuando los siguientes parametros
'COMPONENTES AUSENTES EN LOS ARTICULOS COMPARADOS
'1.- El codigo de sistema del componente del articulo comparado NO ESTA en el articulo patron
'2.- El codigo de sistema del componente del articulo comparado ESTA en el articulo patron, pero la cantidad no es la misma
'COMPONENTES ALTERNATIVOS EN LOS ARTICULOS COMPARADOS
'1.- El codigo de sistema del componente del articulo comparado ESTA en el articulo patron, pero el NOC es distinto. Puede haber varios
'articulos con mismo codigo de sistema y distinto NOC.
'COMPONENTES AUSENTES EN EL ARTICULO PATRON
'1.- El codigo de sistema del componente del articulo comparado, NO ESTA en el articulo patron
Dim Sql1, Sql2, SqlU As String
Sql1 = "SELECT InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel FROM InfConfReal WHERE (((InfConfReal.Padre)='" & A1 & "')) ORDER BY InfConfReal.CodSistema;"
Sql2 = "SELECT InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel FROM InfConfReal WHERE (((InfConfReal.Padre)='" & A2 & "')) ORDER BY InfConfReal.CodSistema;"
'borrar los registros de la tabla de comparacion
Dim db As Database
Set db = Me.Application.CurrentDb
If BorrarReg Then
    SqlU = "DELETE ComparaNOC.* FROM ComparaNOC;"
    db.Execute SqlU
End If
Dim rst1, rst2 As DAO.Recordset
Set rst1 = Me.Application.CurrentDb.OpenRecordset(Sql1)
Set rst2 = Me.Application.CurrentDb.OpenRecordset(Sql2)
If Not rst1.BOF Then rst1.MoveFirst
Dim NOC1, NOC2, Cod1, Cod2 As String
Dim Cant1, Cant2 As Integer
Dim EstaElArticulo As Boolean
Dim Alternativ As Boolean
Dim Flg As Boolean
Dim NAlt, DAlt, NV1, NV2 As String
Do While Not rst1.EOF
    NOC1 = rst1!NOC
    Cod1 = rst1!CodSistema
    Cant1 = rst1!CaComp
    NV1 = rst1!Nivel
    If BorrarReg Then
        'Insertar en la tabla los componentes del articulo patron
        SqlU = "INSERT INTO ComparaNOC ( Orden, Padre, NOC, Descripcion, NSerie, CaComp, CodSistema, Niv) VALUES ('" & rst1!Orden & "', '" & A1 & "', '" & NOC1 & "', '" & rst1!Descripcion & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "', '" & NV1 & "');"
        db.Execute SqlU
    End If
    Alternativ = False
    EstaElArticulo = False
    
    If Not rst2.BOF Then rst2.MoveFirst
    Do While Not rst2.EOF
        NOC2 = rst2!NOC
        Cod2 = rst2!CodSistema
        NV2 = rst2!Nivel
        If Not IsNull(rst2!CaComp) And rst2!CaComp <> "" Then
            Cant2 = rst2!CaComp
        End If
        
        If Cod1 = Cod2 Then
            If NOC1 = NOC2 Then
                Alternativ = False
                EstaElArticulo = True
                Exit Do
                Else
                NAlt = NOC2
                DAlt = rst2!Descripcion
                Alternativ = True
            End If
            If Cant1 > Cant2 Then
                Cant1 = Cant1 - Cant2
            End If
            EstaElArticulo = True
        End If
        rst2.MoveNext
    Loop
    If Not EstaElArticulo Or Alternativ Then
        'ausente
            Dim w As String
            If Alternativ Then
                w = "True"
               Else
                w = "False"
            End If
            If Alternativ Then
                SqlU = "INSERT INTO ComparaNOC ( Orden, Padre, NOC, Descripcion, NSerie, CaComp, CodSistema, Falta, Alternativo, Niv ) VALUES ('" & rst1!Orden & "', '" & A2 & "', '" & NAlt & "', '" & DAlt & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "', " & "True, " & w & ", '" & NV2 & "');"
                Else
                SqlU = "INSERT INTO ComparaNOC ( Orden, Padre, NOC, Descripcion, NSerie, CaComp, CodSistema, Falta, Alternativo, Niv ) VALUES ('" & rst1!Orden & "', '" & A2 & "', '" & NOC1 & "', '" & rst1!Descripcion & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "', " & "True, " & w & ", '" & NV2 & "');"
            End If
            db.Execute SqlU
        Alternativ = False
        Else
        EstaElArticulo = False
    End If
rst1.MoveNext
Loop
End Sub

Sub ComprobarConf3(ByVal A1 As String, ByVal A2 As String, ByVal Lista As ListBox, ByVal BorrarReg As Boolean)
'Comprobar las configuraciones de los NOC seleccionados
Dim Sql1, Sql2, SqlU As String
Dim Altern, NV1, NV2 As String
Sql1 = "SELECT  InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel FROM InfConfReal WHERE (((InfConfReal.Padre)='" & A1 & "')) ORDER BY InfConfReal.CodSistema;"
Sql2 = "SELECT  InfConfReal.Orden, InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.CaComp, InfConfReal.CodSistema, InfConfReal.Nivel FROM InfConfReal WHERE (((InfConfReal.Padre)='" & A2 & "')) ORDER BY InfConfReal.CodSistema, InfConfReal.NOC;"
            'Debug.Print Sql2
'borrar los registros de la tabla de comparacion
Dim db As Database
Set db = Me.Application.CurrentDb
If BorrarReg Then
    SqlU = "DELETE ComparaNOC.* FROM ComparaNOC;"
    db.Execute SqlU
End If
Dim rst1, rst2 As DAO.Recordset
Set rst1 = Me.Application.CurrentDb.OpenRecordset(Sql1)
Set rst2 = Me.Application.CurrentDb.OpenRecordset(Sql2)
If Not rst1.BOF Then rst1.MoveFirst
'listar los articulos del NOC1. Buscar articulos que no esten en NOC 2
'Numero de serie<>"". Si tienen mismo noc, mismo codsistema
'Numero de serie<>"". Si tienen mismo noc, mismo codsistema. Comprobar cantidad
Dim NOC1, NOC2, Cod1, Cod2 As String
Dim Cant1, Cant2 As Integer
Dim NOEstaElArticulo As Boolean
Do While Not rst1.EOF
    NOC1 = rst1!NOC
    Cod1 = rst1!CodSistema
    Cant1 = rst1!CaComp
    NV1 = rst1!Nivel
    'Insertar en la tabla compara los datos del articulo
    SqlU = "INSERT INTO ComparaNOC ( Orden, Padre, NOC, Descripcion, NSerie, CaComp, CodSistema, Niv ) VALUES ('" & rst1!Orden & "', '" & A1 & "', '" & NOC1 & "', '" & rst1!Descripcion & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "', '" & rst1!Nivel & "');"
    db.Execute SqlU
    If Not rst2.BOF Then rst2.MoveFirst
    Do While Not rst2.EOF
        NOC2 = rst2!NOC
        Cod2 = rst2!CodSistema
        NV2 = rst2!Nivel
        If Not IsNull(rst2!CaComp) And rst2!CaComp <> "" Then
            Cant2 = rst2!CaComp
        End If
        If Cod1 = Cod2 Then
            If NOC1 = NOC2 Then
                Alternativ = False
                EstaElArticulo = True
                Exit Do
                Else
                NAlt = NOC2
                DAlt = rst2!Descripcion
                Alternativ = True
            End If
            If Cant1 > Cant2 Then
                Cant1 = Cant1 - Cant2
            End If
            EstaElArticulo = True
        End If
        rst2.MoveNext
    Loop
    If Not EstaElArticulo Or Alternativ Then
        'ausente
            Dim w As String
            If Alternativ Then
                w = "True"
                Altern = "Alternativo"
               Else
                w = "False"
                Altern = ""
            End If
            If Alternativ Then
                SqlU = "INSERT INTO ComparaNOC ( Orden, Padre, NOC, Descripcion, NSerie, CaComp, CodSistema, Falta, Alternativo, Niv ) VALUES ('" & rst1!Orden & "', '" & A2 & "', '" & NOC1 & "', '" & DAlt & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "', " & "True, " & w & ", '" & NV2 & "');"
                Else
                SqlU = "INSERT INTO ComparaNOC ( Orden, Padre, NOC, Descripcion, NSerie, CaComp, CodSistema, Falta, Alternativo, Niv ) VALUES ('" & rst1!Orden & "', '" & A2 & "', '" & NOC1 & "', '" & rst1!Descripcion & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "', " & "True, " & w & ", '" & NV2 & "');"
            End If
            db.Execute SqlU
            If rst1!NSerie <> "" Then
                Lista.AddItem (NOC1 & " - " & rst1!Descripcion & " - " & Cod1 & " - " & Altern)
                Else
                Dim s As String
                If Cant1 > 1 Then
                    s = "Faltan "
                    Else
                    s = "Falta "
                End If
                Lista.AddItem (NOC1 & " - " & rst1!Descripcion & " - " & Cod1 & " (" & s & Cant1 & ")" & " - " & Altern)
            End If
        Alternativ = False
        Else
        EstaElArticulo = False
    End If
    rst1.MoveNext
Loop

End Sub

Sub ComprobarConf2(ByVal A1 As String, ByVal A2 As String, ByVal Lista As ListBox, ByVal BorrarReg As Boolean)
'Comprobar las configuraciones de los NOC seleccionados
Dim Sql1, Sql2, SqlU As String
Dim Altern As Boolean
Sql1 = "SELECT InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.CaComp, InfConfReal.CodSistema FROM InfConfReal WHERE (((InfConfReal.Padre)='" & A1 & "')) ORDER BY InfConfReal.CodSistema;"
Sql2 = "SELECT InfConfReal.Padre, InfConfReal.NOC, InfConfReal.Descripcion, InfConfReal.NSerie, InfConfReal.CaComp, InfConfReal.CodSistema FROM InfConfReal WHERE (((InfConfReal.Padre)='" & A2 & "')) ORDER BY InfConfReal.CodSistema, InfConfReal.NOC;"
            'Debug.Print Sql2
'borrar los registros de la tabla de comparacion
Dim db As Database
Set db = Me.Application.CurrentDb
If BorrarReg Then
    SqlU = "DELETE ComparaNOC.* FROM ComparaNOC;"
    db.Execute SqlU
End If
Dim rst1, rst2 As DAO.Recordset
Set rst1 = Me.Application.CurrentDb.OpenRecordset(Sql1)
Set rst2 = Me.Application.CurrentDb.OpenRecordset(Sql2)
If Not rst1.BOF Then rst1.MoveFirst
'listar los articulos del NOC1. Buscar articulos que no esten en NOC 2
'Numero de serie<>"". Si tienen mismo noc, mismo codsistema
'Numero de serie<>"". Si tienen mismo noc, mismo codsistema. Comprobar cantidad
Dim NOC1, NOC2, Cod1, Cod2 As String
Dim Cant1, Cant2 As Integer
Dim NOEstaElArticulo As Boolean
Do While Not rst1.EOF
    NOC1 = rst1!NOC
    Cod1 = rst1!CodSistema
    Cant1 = rst1!CaComp
    'Insertar en la tabla compara los datos del articulo
    SqlU = "INSERT INTO ComparaNOC ( Padre, NOC, Descripcion, NSerie, CaComp, CodSistema ) VALUES ('" & A1 & "', '" & NOC1 & "', '" & rst1!Descripcion & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "');"
    db.Execute SqlU
    If Not rst2.BOF Then rst2.MoveFirst
    Do While Not rst2.EOF
        NOC2 = rst2!NOC
        Cod2 = rst2!CodSistema
        If Not IsNull(rst2!CaComp) Then
            Cant2 = rst2!CaComp
        End If
            If Cod1 = Cod2 Then
                If NOC1 <> NOC2 Then
                    Altern = True
                    NOEstaElArticulo = True
                    If Cant1 = Cant2 Then
                        Exit Do
                    End If
                    Else
                    Altern = False
                    NOEstaElArticulo = False
                End If
                
                If Cant1 > Cant2 Then
                    NOEstaElArticulo = True
                    Cant1 = Cant1 - Cant2
                    Else
                    NOEstaElArticulo = False
                    Exit Do
                End If
                
                Else
                NOEstaElArticulo = True
                
            End If
        rst2.MoveNext
    Loop
    If NOEstaElArticulo Then
            SqlU = "INSERT INTO ComparaNOC ( Padre, NOC, Descripcion, NSerie, CaComp, CodSistema, Falta ) VALUES ('" & A2 & "', '" & NOC1 & "', '" & rst1!Descripcion & "', '" & rst1!NSerie & "', '" & Cant1 & "', '" & Cod1 & "', " & "True);"
            db.Execute SqlU
        If rst1!NSerie <> "" Then
            Lista.AddItem (NOC1 & " - " & rst1!Descripcion & " - " & Cod1 & Altern)
            Else
            Dim s As String
            If Cant1 > 1 Then
                s = "Faltan "
                Else
                s = "Falta "
            End If
            Lista.AddItem (NOC1 & " - " & rst1!Descripcion & " - " & Cod1 & " (" & s & Cant1 & ")" & Altern)
        End If
        NOEstaElArticulo = False
    End If
    rst1.MoveNext
Loop

End Sub

Private Sub Lista8_AfterUpdate()
Me.Lista8.ControlTipText = Me.Lista8.Value
End Sub
Private Sub Lista10_AfterUpdate()
Me.Lista10.ControlTipText = Me.Lista10.Value
End Sub


