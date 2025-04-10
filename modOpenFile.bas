Attribute VB_Name = "modOpenFile"
Option Compare Text
Option Explicit

Private Declare Function ap_GetOpenFileName Lib "comdlg32.dll" _
                    Alias "GetOpenFileNameA" _
                        (pOpenfilename As OPENFILENAME) As Long
 
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const cdlOFNAllowMultiselect = &H200
Const cdlOFNCreatePrompt = &H2000
Const cdlOFNExplorer = &H80000
Const cdlOFNExtensionDifferent = &H400
Const cdlOFNFileMustExist = &H1000
Const cdlOFNHelpButton = &H10
Const cdlOFNHideReadOnly = &H4
Const cdlOFNLongNames = &H200000
Const cdlOFNNoChangeDir = &H8
Const CdlOFNNoDereferenceLinks = &H100000
Const cdlOFNNoLongNames = &H40000
Const CdlOFNNoReadOnlyReturn = &H8000
Const cdlOFNNoValidate = &H100
Const cdlOFNOverwritePrompt = &H2
Const cdlOFNPathMustExist = &H800
Const cdlOFNReadOnly = &H1
Const CdlOFNShareAware = &H4000

Public Function ap_OpenFile2(Optional ByVal strFileNameIn _
                                         As String = "", Optional strDialogTitle _
                                         As String = "Archivo csv")
 
    Dim lngReturn As Long
    Dim intLocNull As Integer
    Dim StrTemp As String
    Dim ofnFileInfo As OPENFILENAME
    Dim strInitialDir As String
    Dim strFileName As String
    
    '-- if a file path passed in with the name,
    '-- parse it and split it off.
    
    If InStr(strFileNameIn, "\") <> 0 Then
        
        strInitialDir = Left(strFileNameIn, InStrRev(strFileNameIn, "\"))
        strFileName = Left(Mid$(strFileNameIn, _
                                        InStrRev(strFileNameIn, "\") + 1) & _
                                        String(256, 0), 256)
        
    Else
        
        strInitialDir = Left(CurrentDb.Name, _
                                        InStrRev(CurrentDb.Name, "\") - 1)
        strFileName = Left(strFileNameIn & String(256, 0), 256)
    
    End If
       
    With ofnFileInfo
        .lStructSize = Len(ofnFileInfo)
        .lpstrFile = strFileName
        .lpstrFileTitle = String(256, 0)
        .lpstrInitialDir = strInitialDir
        .hwndOwner = Application.hWndAccessApp
        .lpstrFilter = "Archivos separados por comas (*.csv)" & Chr(0) & "*.csv" & Chr(0)
        .nFilterIndex = 1
        .nMaxFile = Len(strFileName)
        .nMaxFileTitle = ofnFileInfo.nMaxFile
        .lpstrTitle = strDialogTitle
        .flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or _
                            cdlOFNNoChangeDir
        .hInstance = 0
        .lpstrCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
        .lpfnHook = 0
    End With
    
    lngReturn = ap_GetOpenFileName(ofnFileInfo)
    
    If lngReturn = 0 Then
       StrTemp = ""
    Else
       
       '-- Trim off any null string
       StrTemp = Trim(ofnFileInfo.lpstrFile)
       intLocNull = InStr(StrTemp, Chr(0))
       
       If intLocNull Then
          StrTemp = Left(StrTemp, intLocNull - 1)
       End If
 
    End If

    ap_OpenFile2 = StrTemp

End Function

