VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Importar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'**********************************
'*© Bgda Joaquin Martinez Rus 2011*
'**********************************
Public contador As Integer
Const Tiempo As Integer = 3000
Public CargarInicio As Boolean
Public Cancelacion As Boolean
Public OrdenInf As Long
Public Nivel As Integer
Public PadreCargado As Boolean

Private Sub Comando0_Click()
Dim F
F = ap_OpenFile2
If F = "" Then Exit Sub
'Me.SetFocus
Me.Comando1.SetFocus
Me.Comando0.Enabled = False
'*********************
'IMPORTAR DATOS
'*********************
Me.ProgressBar4.Visible = True
importar (F)

'*********************
'CARGAR DATOS PARA INFORME
'*********************
InfConfReal2
Me.ProgressBar4.Visible = False

'CargarconfRealInforme
If Not Cancelacion Then
    Me.Etiqueta2.ForeColor = RGB(0, 0, 0)
    Me.Etiqueta2.Caption = "La importacion de datos se ha realizado con exito"
    CargarInicio = True
    Else
    Me.Etiqueta2.Caption = "ATENCION! EL ARCHIVO NO ES VALIDO"
    Me.Etiqueta2.ForeColor = RGB(255, 0, 0)
    Me.Etiqueta5.Caption = "El fichero debe ser CONFIGURACION REAL BREVE"
    Me.Application.SysCmd acSysCmdClearStatus
    Me.TimerInterval = Tiempo
    Cancelacion = False
End If

Me.Etiqueta2.Visible = True
Me.TimerInterval = Tiempo
Me.Comando0.Enabled = True
End Sub

Private Sub Comando1_Click()
Dim Respuesta
Respuesta = MsgBox("Se va a proceder al borrado de todos los datos de la aplicacion. Una vez borrados, puede importar de nuevo los datos del fichero de SIGLE 2010. ¿Desea continuar?", vbCritical + vbYesNo, "Borrar registros")
If Respuesta = vbNo Then
    Exit Sub
End If
Dim sql As String
sql = "DELETE ConfReal.* FROM ConfReal;"
Set db = Me.Application.CurrentDb
db.Execute sql
sql = "DELETE InfConfReal.* FROM InfConfReal;"
db.Execute sql
sql = "DELETE ComparaNOC.* FROM ComparaNOC;"
db.Execute sql
'MsgBox "Se han borrado todos los registros", vbInformation + vbOKOnly, "Borrar registros"
Me.Etiqueta2.Visible = True
Me.Etiqueta2.Caption = "Se han borrado todos los registros"
Me.Etiqueta13.Caption = ""
Me.TimerInterval = Tiempo
End Sub

Private Sub Comando7_Click()
Dim F
F = ap_OpenFile2
If F = "" Then Exit Sub
importarAusentes (F)
End Sub

Sub importar(File As String)
Me.Etiqueta13.Caption = ""
Dim fso, fil1, ts, ts2
Set fso = CreateObject("Scripting.FileSystemObject")
Set fil1 = fso.GetFile(File)
Set ts = fil1.OpenAsTextStream(1)
Set ts2 = fil1.OpenAsTextStream(1)
Me.Application.SysCmd acSysCmdClearStatus

Dim db As Database
Set db = Me.Application.CurrentDb
Dim sqlA As String
sqlA = "DELETE ConfReal.* FROM ConfReal;"
db.Execute sqlA
sqlA = "DELETE InfConfReal.* FROM InfConfReal;"
db.Execute sqlA
sqlA = "DELETE ComparaNOC.* FROM ComparaNOC;"
db.Execute sqlA

Dim rr As DAO.Recordset
Set rr = Me.Application.CurrentDb.OpenRecordset("Local")
rr.AddNew
rr!Archivo = fil1.Name
rr!fechacreacion = fil1.DateCreated

'Set ts = fso.OpenTextFile(strPath2, ForReading, True) Tambien he
Dim Fila As String
Dim n, k, j, Caracter, contador, UltimoCar As Integer
Dim totalLineas As Long
Dim StrCaracter As String
Dim Campo(19)
Dim rst As DAO.Recordset
Set rst = Me.Application.CurrentDb.OpenRecordset("ConfReal")
n = rst.Fields.Count
While Not ts2.AtEndOfStream
    ts2.ReadLine
    totalLineas = totalLineas + 1
Wend
ts2.Close
Me.Application.SysCmd acSysCmdInitMeter, "Importando datos...", totalLineas
Me.Etiqueta5.Caption = "Importando datos..."
Me.ProgressBar4.Max = totalLineas
Me.ProgressBar4.Min = 0.001

Me.Repaint
While Not ts.AtEndOfStream
    Fila = ts.ReadLine
    Me.Application.SysCmd acSysCmdUpdateMeter, j
    'Debug.Print Fila
    'RELLENAR LA MATRIZ CON LOS CAMPOS DE LA LINEA
    UltimoCar = 1
    For Caracter = 1 To Len(Fila)
        StrCaracter = Mid(Fila, Caracter, 1)
        If StrCaracter = ";" Then
            If UltimoCar > 1 Then
                Campo(contador) = Mid(Fila, UltimoCar + 1, Caracter - UltimoCar - 1)
                Else
                Campo(contador) = Mid(Fila, UltimoCar, Caracter - UltimoCar)
            End If
                UltimoCar = Caracter
            contador = contador + 1
        End If
    Next
    If j = 0 Then
        Dim CamposBuenos As Integer
        'Nombres de los campos. Si los nombres son verdaderos, el archivo es bueno.
        For k = 1 To 15
            If rst.Fields(k).Name = Campo(k - 1) Then
                CamposBuenos = CamposBuenos + 1
            End If
        Next
        If CamposBuenos < 14 Then
                Cancelacion = True
                Exit Sub
        End If
        contador = 0
        GoTo sal
    End If
    If j = 1 Then
        'determinar la unidad
        rr!Unidad = Campo(0)
    End If
    
    contador = 0
    rst.AddNew
    For k = 1 To 15
        'Debug.Print rst.Fields(k).Name & " = " & Campo(k - 1)
        rst.Fields(k) = Campo(k - 1)
    Next
    rst.Update
sal:
    j = j + 1
    Me.ProgressBar4.Value = j
Wend
Me.Application.SysCmd acSysCmdClearStatus
        rr!fechaImportacion = Now
        rr.Update
        rr.Close

End Sub


Private Sub CargaInfConfRecursiva(NOC As String, CodSistema As String, NSerie As String)

'rotar en modo recursivo por todos los nodos y asignar se mismo orden
Dim Sql1 As String
Dim rst1 As DAO.Recordset
Dim rst As DAO.Recordset
Dim StrPadre As String

Nivel = Nivel + 1

Sql1 = "SELECT DISTINCT ConfReal.cdUco, ConfReal.cdArtSuperior, ConfReal.dsArtSuperior, " & _
    "ConfReal.cdSistemaSuperior, ConfReal.vaNroSerieSuperior, ConfReal.cdArt, ConfReal.dsArt, " & _
    "ConfReal.cdSistema, ConfReal.vaNroSerie, ConfReal.cdArtBajo, ConfReal.dsArtBajo, " & _
    "ConfReal.cdSistemaBajo, ConfReal.vaNroSerieBajo, ConfReal.caComp, ConfReal.cdEsNroSerie " & _
    "FROM ConfReal " & _
    "WHERE (((ConfReal.cdArt)='" & NOC & "') AND " & _
    "((ConfReal.cdSistema)='" & CodSistema & "') AND " & _
    "((ConfReal.vaNroSerie)='" & NSerie & "')) " & _
    "ORDER BY ConfReal.cdSistemaBajo;"
    'Debug.Print Sql1
Set rst1 = Me.Application.CurrentDb.OpenRecordset(Sql1)
Set rst = Me.Application.CurrentDb.OpenRecordset("InfConfReal")


If rst1.RecordCount > 0 Then
    

'If rst1!vaNroSerieSuperior = "ET 078274-VE" Then
'      Debug.Print Sql1
  
'End If
    If Not rst1.BOF Then rst1.MoveFirst
    Do While Not rst1.EOF
        
        Dim FatherOne As String
        StrPadre = rst1!cdArtSuperior & "+" & rst1!vaNroSerieSuperior
                            
        If Not PadreCargado Then
            rst.AddNew
            OrdenInf = 10
            rst!Orden = OrdenInf
            rst!Padre = StrPadre
            rst!NOC = rst1!cdArtSuperior
            rst!Descripcion = rst1!dsArtSuperior
            rst!NSerie = rst1!vaNroSerieSuperior
            rst!CodSistema = rst1!cdSistemaSuperior
            rst!Padre1 = StrPadre
            rst!CaComp = rst1!CaComp
            rst!Nivel = 0
            rst!Imprimir = True
            rst.Update
            PadreCargado = True
        ElseIf rst1!cdArtBajo <> "" Then
                rst.AddNew
                rst!Orden = OrdenInf
                rst!Padre = StrPadre
                rst!NOC = rst1!cdArtBajo
                rst!Descripcion = rst1!dsArtBajo
                rst!NSerie = rst1!vaNroSerieBajo
                rst!CodSistema = rst1!cdSistemaBajo
                FatherOne = rst1!cdArt & "+" & rst1!vaNroSerie
                
                rst!Padre1 = FatherOne
                rst!CaComp = rst1!CaComp
                OrdenInf = OrdenInf + 10
                rst!Nivel = Nivel
                rst!Imprimir = True
                rst.Update
            
        End If
        If rst1!cdArtBajo <> "" Then
            CargaInfConfRecursiva rst1!cdArtBajo, rst1!cdSistemaBajo, rst1!vaNroSerieBajo
        End If
        rst1.MoveNext
    Loop
End If
Nivel = Nivel - 1
End Sub


Private Sub InfConfReal2()

Me.Etiqueta5.Caption = "Validando datos..." & " >>> 0%"
Me.Repaint
Dim ini As Date
Dim contador As Integer
ini = Now
'Rellena la tabla de informe de conf real con los datos de la tabla ConfReal

Dim rstN1 As DAO.Recordset
'Dim rstN2 As DAO.Recordset
Dim Norden As Long
'Dim rstN3 As DAO.Recordset
'Dim rstN4 As DAO.Recordset
'Dim rstN5 As DAO.Recordset

Dim rst As DAO.Recordset
Dim Sql1 As String
'*************************
'ARTICULOS SUPERIORES
'*************************
Sql1 = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.dsArtSuperior, ConfReal.cdSistemaSuperior, ConfReal.vaNroSerieSuperior FROM ConfReal"
Set rstN1 = Me.Application.CurrentDb.OpenRecordset(Sql1)

Set rst = Me.Application.CurrentDb.OpenRecordset("InfConfReal")
Dim StrPadre As String
'Nivel 1
Me.Application.SysCmd acSysCmdClearStatus
Dim tot As Integer
tot = rstN1.RecordCount
Me.ProgressBar4.Max = tot

Me.ProgressBar4.Min = 0.1
Me.ProgressBar4.Value = 0.1

Me.Application.SysCmd acSysCmdInitMeter, "Validando datos de configuracion real...", tot

Do While Not rstN1.EOF

        If rstN1!cdArtSuperior = "" Then
            GoTo Nuevo
        End If
        
        contador = contador + 1
         If Not IsNull(rstN1!dsArtSuperior) And rstN1!dsArtSuperior <> "" Then
            'CALCULAR LA VELOCIDAD DE DATOS CADA 10 REGISTROS
            If contador Mod 20 = 0 Then
                ' Calcular tiempo que queda
                CuantoQueda ini, rstN1.AbsolutePosition, tot
            End If
            
            Dim porc As Double
            porc = rstN1.AbsolutePosition / tot
            Me.Etiqueta5.Caption = "Validando datos de " & rstN1!dsArtSuperior & " >>> " & Format(porc, "0%")
            Me.ProgressBar4.Value = rstN1.AbsolutePosition + 1
            Me.Repaint
        End If
        'StrPadre = rstN1!cdArtSuperior & "+" & rstN1!vaNroSerieSuperior
            
            
        CargaInfConfRecursiva rstN1!cdArtSuperior, rstN1!cdSistemaSuperior, rstN1!vaNroSerieSuperior
        
Nuevo:
    Norden = 0
    Nivel = 0
    OrdenInf = 0
    PadreCargado = False
   rstN1.MoveNext
    'Me.Application.SysCmd acSysCmdSetStatus, "Validando " & rstN1!dsArtSuperior
Loop



'RESUMEN

Dim seg As Double
seg = DateDiff("s", ini, Now)
Dim minutos As String
Dim segundos As String
minutos = Int(seg / 60)
segundos = seg - minutos * 60
Me.Etiqueta5.Caption = "Tiempo de importacion de datos " & minutos & " minutos y " & segundos & " segundos."
Me.LblQueda.Caption = ""

Me.Application.SysCmd acSysCmdClearStatus
'Debug.Print "fin: " & DateDiff("s", ini, Now)
Dim rcs As DAO.Recordset
Set rcs = Me.Application.CurrentDb.OpenRecordset("Local")
rcs.MoveLast
rcs.Edit
rcs!TiempoImport = minutos & " minutos y " & segundos & " segundos."
rcs.Update

End Sub

Private Sub CuantoQueda(ini As Date, numreg As Integer, numtot As Integer)
'CALCULAR CUANTO TIEMPO APROXIMADAMENTE QUEDA
Dim seg As Double
seg = DateDiff("s", ini, Now)

Dim tiempoqueda As Double
tiempoqueda = (numtot - numreg) * seg / numreg


Dim minutos As String
Dim segundos As String
minutos = Int(tiempoqueda / 60)
segundos = tiempoqueda - minutos * 60
Me.LblQueda.Caption = "Tiempo aproximado para fin, " & minutos & " minutos y " & Round(segundos, 0) & " segundos."
Me.Repaint
End Sub


Private Sub Comando11_Click()
Dim Respuesta
Respuesta = MsgBox("Se va a proceder al borrado de los datos del historial y los datos de la aplicacion. ¿Desea continuar?", vbCritical + vbYesNo, "Borrar Historial")
If Respuesta = vbNo Then
    Exit Sub
End If

Dim db As Database
Set db = Me.Application.CurrentDb
Dim sqlA As String
sqlA = "DELETE Local.* FROM Local;"
db.Execute sqlA
sqlA = "DELETE ConfReal.* FROM ConfReal;"
Set db = Me.Application.CurrentDb
db.Execute sqlA
sqlA = "DELETE InfConfReal.* FROM InfConfReal;"
db.Execute sqlA
sqlA = "DELETE ComparaNOC.* FROM ComparaNOC;"
db.Execute sqlA
'MsgBox "Se han borrado todos los registros", vbInformation + vbOKOnly, "Borrar registros"
Me.Etiqueta2.Visible = True
Me.Etiqueta2.Caption = "Se ha borrado el historial y todos los registros"
Me.TimerInterval = Tiempo
Me.Etiqueta13.Caption = ""
End Sub

Private Sub Comando12_Click()
DoCmd.OpenForm "Historial"
End Sub

Private Sub Form_Load()
Dim rst As DAO.Recordset
Set rst = Me.Application.CurrentDb.OpenRecordset("Local")
If rst.RecordCount > 0 Then
    rst.MoveLast
    Me.Etiqueta13.Caption = "Archivo importado: " & rst!Archivo & vbNewLine & "Fecha de creacion: " & rst!fechacreacion & vbNewLine & "Unidad: " & rst!Unidad
End If
End Sub

Private Sub Form_Timer()
contador = contador + 1
If contador = 2 Then
    contador = 0
    Me.TimerInterval = 0
    Me.Etiqueta2.Visible = False
    Me.Etiqueta5.Caption = " "
    If CargarInicio Then
        CargarInicio = False
        Form_Panel.Secundario0.SourceObject = "ConfReal"
        Form_Panel.Marco9.Value = 1
    End If
End If
End Sub


'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************

Private Sub InfConfReal()

Me.Etiqueta5.Caption = "Validando datos..." & " >>> 0%"
Me.Repaint
Dim ini As Date
Dim contador As Integer
ini = Now
'Rellena la tabla de informe de conf real con los datos de la tabla ConfReal

Dim rstN1 As DAO.Recordset
Dim rstN2 As DAO.Recordset
Dim Norden As Integer
Dim rstN3 As DAO.Recordset
Dim rstN4 As DAO.Recordset

Dim rst As DAO.Recordset
Dim Sql1, Sql2, Sql3, Sql4 As String
'*************************
'ARTICULOS SUPERIORES
'*************************
Sql1 = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.dsArtSuperior, ConfReal.cdSistemaSuperior, ConfReal.vaNroSerieSuperior FROM ConfReal"
Set rstN1 = Me.Application.CurrentDb.OpenRecordset(Sql1)

Set rst = Me.Application.CurrentDb.OpenRecordset("InfConfReal")
Dim StrPadre As String
'Nivel 1
Me.Application.SysCmd acSysCmdClearStatus
Dim tot As Integer
tot = rstN1.RecordCount
Me.Application.SysCmd acSysCmdInitMeter, "Validando datos de configuracion real...", tot
Do While Not rstN1.EOF

        If rstN1!cdArtSuperior = "" Then
            GoTo Nuevo
        End If
        
        contador = contador + 1
        'Dim Sql9 As String
         If Not IsNull(rstN1!dsArtSuperior) Then
            Me.Etiqueta5.Caption = "Validando datos de " & rstN1!dsArtSuperior & " >>> " & Format(rstN1.AbsolutePosition / tot, "0%")
        End If
        StrPadre = rstN1!cdArtSuperior & "+" & rstN1!vaNroSerieSuperior
    
   
    'ROTAR POR LOS HIJOS DEL PRINCIPAL
    Sql3 = "SELECT DISTINCT ConfReal.cdArt, ConfReal.dsArt, ConfReal.cdSistema, ConfReal.vaNroSerie FROM ConfReal WHERE (((ConfReal.cdArtSuperior)='" & rstN1!cdArtSuperior & "') AND ((ConfReal.vaNroSerieSuperior)='" & rstN1!vaNroSerieSuperior & "') AND ((ConfReal.cdArtBajo)<>'')) ORDER BY ConfReal.cdSistema;"
       
    'Dim codpadre, codhijo As String
    Set rstN3 = Me.Application.CurrentDb.OpenRecordset(Sql3)
    
    Do While Not rstN3.EOF
        contador = contador + 1
        
        
        'ROTAR POR TODOS LOS NIETOS DEL PRINCIPAL
        Sql2 = "SELECT ConfReal.cdArt, ConfReal.dsArt, ConfReal.cdSistema, ConfReal.vaNroSerie, ConfReal.cdArtBajo, ConfReal.dsArtBajo, ConfReal.cdSistemaBajo, ConfReal.vaNroSerieBajo, ConfReal.caComp, ConfReal.cdEsNroSerie FROM ConfReal WHERE (((ConfReal.cdArt)='" & rstN3!cdArt & "') AND ((ConfReal.cdSistema)='" & rstN3!cdSistema & "') AND ((ConfReal.vaNroSerie)='" & rstN3!vaNroSerie & "')) ORDER BY ConfReal.cdArtSuperior, ConfReal.cdSistemaSuperior, ConfReal.vaNroSerieSuperior, ConfReal.cdArt, ConfReal.cdSistema, ConfReal.vaNroSerie, ConfReal.cdSistemaBajo;"
        'Debug.Print Sql2
        Set rstN2 = Me.Application.CurrentDb.OpenRecordset(Sql2)
        Dim inc2 As Integer
        
        Do While Not rstN2.EOF
        
            rst.AddNew
            
            'COMPROBAR SI EXISTE UN PADRE CON NUMERO DE ORDEN
            Sql4 = "SELECT InfConfReal.NOC, InfConfReal.CodSistema, InfConfReal.NSerie, InfConfReal.Orden FROM InfConfReal WHERE (((InfConfReal.NOC)='" & rstN2!cdArt & "') AND ((InfConfReal.CodSistema)='" & rstN2!cdSistema & "') AND ((InfConfReal.NSerie)='" & rstN2!vaNroSerie & "'));"
            Set rstN4 = Me.Application.CurrentDb.OpenRecordset(Sql4)
            If Not rstN4.BOF Then rstN4.MoveFirst
            
            If rstN4.RecordCount > 0 And Not IsNull(rstN4!Orden) Then
                rst!Orden = rstN4!Orden + inc2
                inc2 = inc2 + 5
                Else
                inc2 = 0
                Norden = Norden + 100
                rst!Orden = Norden
            End If
            
            'rst!Orden = Norden
            rst!Padre = StrPadre
            rst!NOC = rstN2!cdArtBajo
            rst!Descripcion = rstN2!dsArtBajo
            rst!NSerie = rstN2!vaNroSerieBajo
            rst!CodSistema = rstN2!cdSistemaBajo
            rst!CaComp = rstN2!CaComp
            rst!Padre1 = rstN2!cdArt & "+" & rstN2!vaNroSerie
            'Determinar el nivel
            codpadre = rstN2!cdSistema
            codhijo = rstN2!cdSistemaBajo
            Coincidencias = 0
            For x = 1 To 5
                If Left(codpadre, x) = Left(codhijo, x) Then
                    Coincidencias = Coincidencias + 1
                End If
            Next
            
            If Coincidencias = 0 Then Coincidencias = 1
            rst!Nivel = Coincidencias
            rst!Imprimir = True
            rst.Update
            
            rstN2.MoveNext
        
        Loop
        
        rstN3.MoveNext
    Loop
Nuevo:
    Norden = 0
   rstN1.MoveNext
    'Me.Application.SysCmd acSysCmdSetStatus, "Validando " & rstN1!dsArtSuperior
Loop

Dim seg As Double
seg = DateDiff("s", ini, Now)
Dim minutos As String
Dim segundos As String
minutos = Int(seg / 60)
segundos = seg - minutos * 60
Me.Etiqueta5.Caption = "Tiempo de importacion de datos " & minutos & " minutos y " & segundos & " segundos."

Me.Application.SysCmd acSysCmdClearStatus
'Debug.Print "fin: " & DateDiff("s", ini, Now)
Dim rcs As DAO.Recordset
Set rcs = Me.Application.CurrentDb.OpenRecordset("Local")
rcs.MoveLast
rcs.Edit
rcs!TiempoImport = minutos & " minutos y " & segundos & " segundos."
rcs.Update

End Sub

Sub CargarconfRealInforme()
Me.Etiqueta5.Caption = "Validando datos..." & " >>> 0%"
Me.Repaint
Dim ini As Date
Dim contador As Integer
ini = Now
'Debug.Print "inicio: " & ini
'Rellena la tabla de informe de conf real con los datos de la tabla ConfReal
Dim rstN1 As DAO.Recordset
Dim rstN2 As DAO.Recordset
Dim Norden As Integer
Dim rstN3 As DAO.Recordset
Dim rst As DAO.Recordset
Dim Sql1, Sql2, Sql3 As String
Sql1 = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.dsArtSuperior, ConfReal.cdSistemaSuperior, ConfReal.vaNroSerieSuperior FROM ConfReal"
Set rstN1 = Me.Application.CurrentDb.OpenRecordset(Sql1)
    'Debug.Print Sql1
Set rst = Me.Application.CurrentDb.OpenRecordset("InfConfReal")
Dim StrPadre As String
'Nivel 1
Me.Application.SysCmd acSysCmdClearStatus
Dim tot As Integer
tot = rstN1.RecordCount
Me.Application.SysCmd acSysCmdInitMeter, "Validando datos de configuracion real...", tot
Do While Not rstN1.EOF
        If rstN1!cdArtSuperior = "" Then
            GoTo Nuevo
        End If
        contador = contador + 1
        'Dim Sql9 As String
         If Not IsNull(rstN1!dsArtSuperior) Then
            Me.Etiqueta5.Caption = "Validando datos de " & rstN1!dsArtSuperior & " >>> " & Format(rstN1.AbsolutePosition / tot, "0%")
        End If
   rst.AddNew
    StrPadre = rstN1!cdArtSuperior & "+" & rstN1!vaNroSerieSuperior
    'Sql9 = "INSERT INTO InfConfReal ( Padre, NOC, Descripcion, NSerie, CodSistema ) VALUES ('" & StrPadre & "', '" & rstN1!cdArtSuperior & "', '" & rstN1!dsArtSuperior & "', '" & rstN1!vaNroSerieSuperior & "', '" & rstN1!cdSistemaSuperior & "');"
    'Dim db As Database
    'Set db = Me.Application.CurrentDb
    'db.Execute (Sql9)
    rst!Padre = StrPadre
    rst!NOC = rstN1!cdArtSuperior
    rst!Descripcion = rstN1!dsArtSuperior
    rst!NSerie = rstN1!vaNroSerieSuperior
    rst!CodSistema = rstN1!cdSistemaSuperior
    ''rst!cantidad = rstN1!caComp
    rst!Nivel = 1
    rst!Imprimir = True
    rst.Update
    
    Me.Application.SysCmd acSysCmdUpdateMeter, rstN1.AbsolutePosition
    
    
    'Niveles inferiores
     'Sql3 = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.vaNroSerieSuperior,ConfReal.cdSistema, ConfReal.cdArt, ConfReal.vaNroSerie, ConfReal.cdArtBajo, ConfReal.dsArtBajo, ConfReal.cdSistemaBajo, ConfReal.vaNroSerieBajo, ConfReal.cdEsNroSerie, ConfReal.caComp, ConfReal.cdEsNroSerie FROM ConfReal WHERE (((ConfReal.cdArtSuperior)='" & rstN1!cdArtSuperior & "') AND ((ConfReal.vaNroSerieSuperior)='" & rstN1!vaNroSerieSuperior & "') AND ((ConfReal.cdArtBajo)<>'')) ORDER BY ConfReal.cdSistemaBajo;"
    'Dim HH As String
    Sql3 = "SELECT DISTINCT ConfReal.cdArt, ConfReal.dsArt, ConfReal.cdSistema, ConfReal.vaNroSerie FROM ConfReal WHERE (((ConfReal.cdArtSuperior)='" & rstN1!cdArtSuperior & "') AND ((ConfReal.vaNroSerieSuperior)='" & rstN1!vaNroSerieSuperior & "') AND ((ConfReal.cdArtBajo)<>'')) ORDER BY ConfReal.cdSistemaBajo;"
    
    'Debug.Print HH
    'Dim Coincidencias, x As Integer
    
    Dim codpadre, codhijo As String
    Set rstN3 = Me.Application.CurrentDb.OpenRecordset(Sql3)
    Do While Not rstN3.EOF
        contador = contador + 1
        Norden = Norden + 100
        rst.AddNew
        rst!Orden = Norden
        rst!Padre = StrPadre
        rst!NOC = rstN3!cdArtBajo
        rst!Descripcion = rstN3!dsArtBajo
        rst!NSerie = rstN3!vaNroSerieBajo
        rst!CodSistema = rstN3!cdSistemaBajo
        rst!CaComp = rstN3!CaComp
        rst!Padre1 = rstN3!cdArt & "+" & rstN3!vaNroSerie
        If rst!Padre1 <> rst!Padre Then
            'Reordenar en el caso
            Dim StrNoc, StrSerie As String
            StrNoc = Left(rst!Padre1, 13)
            StrSerie = Right(rst!Padre1, Len(rst!Padre1) - 14)
            Dim sql9, sql10 As String
            Dim rstOrden As DAO.Recordset
            Dim RstAux As DAO.Recordset
            Dim nreg As Integer
            sql10 = "SELECT DISTINCT ConfReal.cdArtSuperior, ConfReal.vaNroSerieSuperior, ConfReal.vaNroSerie, ConfReal.cdArtBajo, ConfReal.cdSistemaBajo FROM ConfReal WHERE (((ConfReal.cdArtSuperior)='" & rstN3!cdArtSuperior & "') AND ((ConfReal.vaNroSerieSuperior)='" & rstN3!vaNroSerieSuperior & "') AND ((ConfReal.cdSistemaBajo)='" & rstN3!cdSistemaBajo & "'));"
            sql9 = "SELECT DISTINCT InfConfReal.Orden, InfConfReal.NSerie, InfConfReal.NOC FROM InfConfReal WHERE (((InfConfReal.NSerie)='" & StrSerie & "') AND ((InfConfReal.NOC)='" & StrNoc & "'));"
            'Debug.Print sql10
            Set RstAux = Me.Application.CurrentDb.OpenRecordset(sql10)
            If Not RstAux.EOF Then RstAux.MoveLast
            nreg = RstAux.RecordCount
            If nreg > 1 Then
                Set rstOrden = Me.Application.CurrentDb.OpenRecordset(sql9)
                If Not rstOrden.BOF Then rstOrden.MoveFirst
                rst!Orden = rstOrden!Orden + 50
                rstOrden.Close
            End If
            RstAux.Close
        End If
        'Determinar el nivel
        codpadre = rstN3!cdSistema
        codhijo = rstN3!cdSistemaBajo
        Coincidencias = 0
        For x = 1 To 5
            If Left(codpadre, x) = Left(codhijo, x) Then
                Coincidencias = Coincidencias + 1
            End If
        Next
        If Coincidencias = 0 Then Coincidencias = 1
        rst!Nivel = Coincidencias
        rst!Imprimir = True
        rst.Update
        
        rstN3.MoveNext
    Loop
Nuevo:
    Norden = 0
   rstN1.MoveNext
    'Me.Application.SysCmd acSysCmdSetStatus, "Validando " & rstN1!dsArtSuperior
Loop
Dim seg As Double
seg = DateDiff("s", ini, Now)
Dim minutos As String
Dim segundos As String
minutos = Int(seg / 60)
segundos = seg - minutos * 60
Me.Etiqueta5.Caption = "Tiempo de importacion de datos " & minutos & " minutos y " & segundos & " segundos."

Me.Application.SysCmd acSysCmdClearStatus
'Debug.Print "fin: " & DateDiff("s", ini, Now)
Dim rcs As DAO.Recordset
Set rcs = Me.Application.CurrentDb.OpenRecordset("Local")
rcs.MoveLast
rcs.Edit
rcs!TiempoImport = minutos & " minutos y " & segundos & " segundos."
rcs.Update
End Sub

