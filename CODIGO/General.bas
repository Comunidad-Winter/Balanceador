Attribute VB_Name = "General"
Option Explicit

'Prg
Public prgRun As Boolean

'Ventana activa
Private Declare Function GetActiveWindow Lib "USER32" () As Long

'Teclado
Public Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'GetVar y WriteVar
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

'Golpes
Public Enum PartesCuerpo

    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6

End Enum

'Balance
Public Type ModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    Escudo As Double
    AtaqueArpon As Double
    DañoArpon As Double
    
    EvasionBackUp As Double
    AtaqueArmasBackUp As Double
    AtaqueProyectilesBackUp As Double
    AtaqueWrestlingBackUp As Double
    DañoArmasBackUp As Double
    DañoProyectilesBackUp As Double
    DañoWrestlingBackUp As Double
    EscudoBackUp As Double
    AtaqueArponBackUp As Double
    DañoArponBackUp As Double
    
End Type

Public ModClase(1 To 18) As ModClase

'Atacante

Public Type tAtacante
fuerza As Byte
agilidad As Byte
Nivel As Byte
SkillCombateConArmas As Byte
SkillWrestling As Byte
SkillApuñalar As Byte
MinHit As Integer
MaxHit As Integer
Raza As Byte
clase As Byte
Arma As Integer
dañoextra As Integer
End Type

Public Atacante As tAtacante

'Oponente
Public Type tOponente
agilidad As Byte
Nivel As Byte
SkillDefensaConEscudos As Byte
SkillTacticas As Byte
clase As Byte
Raza As Byte
Casco As Integer
Armadura As Integer
Escudo As Integer
DefensaExtra As Integer
End Type

Public Oponente As tOponente

'MaxObj
Public NumMaximoData As Integer
 
'Clases
Public Enum eClass
    Clerigo = 1
    Mago = 2
    Guerrero = 3
    Asesino = 4
    Ladron = 5
    Bardo = 6
    Druida = 7
    Gladiador = 8 'Cazarecompensas
    Paladin = 9
    Cazador = 10
    Pescador = 11
    Herrero = 12
    Leñador = 13
    Minero = 14
    Carpintero = 15
    Sastre = 16
    Mercenario = 17 'Drakkar
    Nigromante = 18
End Enum

Public ListaClases() As String
 
'Razas


Public Enum eRaza

    Humano = 1
    Elfo
    Drow
    gnomo
    enano
    Orco
    
End Enum


Public ListaRazas() As String
 
'LoadArmas
Public Type tDaño
ObjName As String
ObjIndex As Integer
End Type
 
Public daño() As tDaño

'LoadDaño and Def extra

Public Type tDañoExtra
ObjName As String
ObjIndex As Integer
End Type

Public dañoextra() As tDañoExtra

Public Type tDefensaExtra
ObjName As String
ObjIndex As Integer
End Type

Public defextra() As tDefensaExtra

'LoadCascos
Public Type tcasco
ObjName As String
ObjIndex As Integer
End Type
 
Public Casco() As tcasco

'LoadEscudos
Public Type tescudos
ObjName As String
ObjIndex As Integer
End Type
 
Public escudos() As tescudos

'LoadArmaduras
Public Type tarmaduras
ObjName As String
ObjIndex As Integer
End Type
 
Public armaduras() As tarmaduras

Public Type ObjData
    Name               As String 'Nombre del obj
    tipe As Byte
    Grhindex As Integer
    MaxHit As Integer
    MinHit As Integer
    MaxDef As Integer
    MinDef As Integer
    Nivel As Byte
    ObjIndex As Integer
    SubTipo As Integer
    Apuñala As Byte
    CuantoAumento As Integer
End Type

Public ObjData() As ObjData

'Path
Public DatosPath As String


'Graficos.ini
Public Type GrhData

    sX As Integer
    sY As Integer
       
    FileNum As Integer
       
    pixelWidth As Integer
    pixelHeight As Integer
       
    TileWidth As Single
    TileHeight As Single
       
    NumFrames As Integer
    Frames() As Integer
       
    speed As Single
    mini_map_color As Long

End Type

Public GrhData() As GrhData

'API needed for creating DIBSections for sampling and pixel access.
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'Reproducir sonido
Private Const SND_APPLICATION = &H80
Private Const SND_ALIAS = &H10000
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H2000
Private Const SND_PURGE = &H40
Private Const SND_RESOURCE = &H40004
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long


Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String
    '*****************************************************************
    'Gets a field from a string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    'Gets a field from a delimited string
    '*****************************************************************

    Dim i          As Long
    Dim LastPos    As Long
    Dim CurrentPos As Long
    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

End Function

Sub LoadBalance()

    On Error GoTo error

    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To 18

        With ModClase(i)
            .Evasion = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueWrestling = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .DañoArmas = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
            .DañoProyectiles = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
            .DañoWrestling = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
            .Escudo = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODESCUDO", ListaClases(i)))
            .AtaqueArpon = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODAtaqueArpon", ListaClases(i)))
            .DañoArpon = Val(GetVar(App.Path & "\Datos\Balance.dat", "MODDañoArpon", ListaClases(i)))
 
        End With

    Next i
    
    For i = 1 To 18
    
        With ModClase(i)
            .EvasionBackUp = .Evasion
            .AtaqueArmasBackUp = .AtaqueArmas
            .AtaqueProyectilesBackUp = .AtaqueProyectiles
            .AtaqueWrestlingBackUp = .AtaqueWrestling
            .DañoArmasBackUp = .DañoArmas
            .DañoProyectilesBackUp = .DañoProyectiles
            .DañoWrestlingBackUp = .DañoWrestling
            .EscudoBackUp = .Escudo
            .AtaqueArponBackUp = .AtaqueArpon
            .DañoArponBackUp = .DañoArpon
 
        End With

    Next i
    
    Exit Sub

error:
    MsgBox "error en loadbalance"
    
End Sub
Public Function RefreshListaBalanceAtacante()
    
    Dim tmpClase As Byte
    
    tmpClase = ObtenerClase(frmMain.lstClaseAtacante.List(frmMain.lstClaseAtacante.ListIndex))
    
    If tmpClase = 0 Then
        MsgBox "Debes asignar las clases para ejecutar esta acción."
        Exit Function
    End If
    
    With ModClase(tmpClase)
    
        frmMain.txtModDañoWre.Text = CStr(.DañoWrestling)
        frmMain.txtModAtaqueWre.Text = CStr(.AtaqueWrestling)
        
        frmMain.txtModDañoArmas.Text = CStr(.DañoArmas)
        frmMain.txtModAtaqueArmas.Text = CStr(.AtaqueArmas)
         
    End With

End Function
Public Function RefreshListaBalanceOponente()

    Dim tmpClase As Byte
    
    tmpClase = ObtenerClase(frmMain.lstClaseOponente.List(frmMain.lstClaseOponente.ListIndex))
    
    If tmpClase = 0 Then
        MsgBox "Debes asignar las clases para ejecutar esta acción."
        Exit Function
    End If
    
    With ModClase(tmpClase)

        frmMain.txtMODESCUDO.Text = CStr(.Escudo)
        frmMain.txtModEvasion.Text = CStr(.Evasion)
    
    End With
    
End Function

Public Function ObtenerHit(ByVal Nivel As Integer, ByVal Raza As String, ByVal clase As String) As String

    Dim i As Long
    
    Dim AumentoHIT As Long
    
    AumentoHIT = 0
    Atacante.MinHit = 0
    Atacante.MaxHit = 0
        
    For i = 1 To Nivel
        
        If i = 1 Then
            Atacante.MinHit = 1
            Atacante.MaxHit = 2
        End If
                
        If i > 1 Then
        
            Select Case UCase(clase)
        
                Case UCase("Guerrero")

                    AumentoHIT = IIf(i > 35, 2, 3)
            
                Case UCase("cazador")
                    AumentoHIT = IIf(i > 35, 2, 3)

                Case UCase("Paladin")
                    AumentoHIT = IIf(i > 35, 1, 3)

                Case UCase("ladron")
                    AumentoHIT = 1

                Case UCase("Mago")
                    AumentoHIT = 1

                Case UCase("Leñador")
                    AumentoHIT = 2

                Case UCase("Minero")
                    AumentoHIT = 2

                Case UCase("Pescador")
                    AumentoHIT = 1
                
                Case UCase("Clerigo")
                 
                    AumentoHIT = 2

                Case UCase("Druida")
                    AumentoHIT = 2

                Case UCase("Asesino")
                    AumentoHIT = IIf(i > 35, 1, 3)

                Case UCase("Bardo")
                    AumentoHIT = 2

                Case UCase("Herrero")
                    AumentoHIT = 2
                    
                Case UCase("Carpintero")
                    AumentoHIT = 2
                
                Case UCase("Gladiador")
                    AumentoHIT = IIf(i > 40, 2, 3)
                
                Case UCase("Nigromante")
                    AumentoHIT = IIf(i > 40, 1, 3)
                    
                Case UCase("Mercenario")
                    AumentoHIT = IIf(i > 30, 2, 3)
               
                Case Else
                    AumentoHIT = 2

            End Select
            
     
           Atacante.MaxHit = Atacante.MaxHit + AumentoHIT


        If i < 36 Then
                If Atacante.MaxHit > 99 Then Atacante.MaxHit = 99
            Else
                If Atacante.MaxHit > 999 Then Atacante.MaxHit = 999
            End If
            
            Atacante.MinHit = Atacante.MinHit + AumentoHIT

            If i < 36 Then
                If Atacante.MinHit > 99 Then Atacante.MinHit = 99
            Else
                If Atacante.MinHit > 999 Then Atacante.MinHit = 999
            End If
                
        End If
            
            
    Next i
    
    ObtenerHit = CStr(Atacante.MinHit & "/" & Atacante.MaxHit)
    
End Function

Function GetVar(ByVal File As String, ByVal Main As String, ByVal var As String, Optional EmptySpaces As Long = 1024) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
    GetPrivateProfileString Main, var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function
Public Function LoadListas()
    
    Dim i As Long
    Dim Contador As Long
    
    Contador = 0
    
    ReDim ListaClases(1 To 18) As String

    ListaClases(1) = "Clerigo"
    ListaClases(2) = "Mago"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Gladiador"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Cazador"
    ListaClases(11) = "Pescador"
    ListaClases(12) = "Herrero"
    ListaClases(13) = "Leñador"
    ListaClases(14) = "Minero"
    ListaClases(15) = "Carpintero"
    ListaClases(16) = "Sastre"
    ListaClases(17) = "Mercenario"
    ListaClases(18) = "Nigromante"
    
    ReDim ListaRazas(1 To 6) As String
    
    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Drow"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"
    ListaRazas(6) = "Orco"
    
    If frmMain.Visible Then
        
        'Load Atacante
        
        frmMain.lstArmasAtacante.Clear
     
        For i = 1 To NumMaximoData
            If ObjData(i).tipe = 2 Or ObjData(i).tipe = 46 Then
                Contador = Contador + 1
            End If
        Next i
        
        ReDim daño(1 To Contador)
        
        Contador = 0
        
        Call frmMain.lstArmasAtacante.AddItem(0 & " - " & "Sin arma")
        
        For i = 1 To NumMaximoData
        
            If ObjData(i).tipe = 2 Or ObjData(i).tipe = 46 Then
                Contador = Contador + 1
                daño(Contador).ObjIndex = i
                daño(Contador).ObjName = ObjData(i).Name
                Call frmMain.lstArmasAtacante.AddItem(i & " - " & ObjData(i).Name)
            End If
        Next i
    
       frmMain.lstRazaAtacante.Clear
       frmMain.lstRazaAtacante.AddItem vbNullString
       
       For i = LBound(ListaRazas) To UBound(ListaRazas)
           frmMain.lstRazaAtacante.AddItem ListaRazas(i)
       Next i
    
       frmMain.lstClaseAtacante.Clear
       frmMain.lstClaseAtacante.AddItem vbNullString
       
       For i = LBound(ListaClases) To UBound(ListaClases)
           frmMain.lstClaseAtacante.AddItem ListaClases(i)
       Next i
        
        
       frmMain.lstArmasAtacante.ListIndex = 0
        
        Contador = 0
        
        frmMain.lstDañoExtraAtacante.Clear
        
        For i = 1 To NumMaximoData
            If ObjData(i).tipe = 31 Or ObjData(i).tipe = 44 Then
                Contador = Contador + 1
            End If
        Next i
        
        ReDim dañoextra(1 To Contador)
        
        Contador = 0
        
        Call frmMain.lstDañoExtraAtacante.AddItem(0 & " - " & "Sin equipar")
        
        For i = 1 To NumMaximoData
        
            If ObjData(i).tipe = 31 Or ObjData(i).tipe = 44 Then
                Contador = Contador + 1
                dañoextra(Contador).ObjIndex = i
                dañoextra(Contador).ObjName = ObjData(i).Name
                Call frmMain.lstDañoExtraAtacante.AddItem(i & " - " & ObjData(i).Name)
            End If
        Next i

       frmMain.lstDañoExtraAtacante.ListIndex = 0
       
        'Oponente
        
        frmMain.lstClaseOponente.Clear
        frmMain.lstClaseOponente.AddItem vbNullString
        
        For i = LBound(ListaClases) To UBound(ListaClases)
            frmMain.lstClaseOponente.AddItem ListaClases(i)
        Next i
        
        frmMain.lstRazaOponente.Clear
        frmMain.lstRazaOponente.AddItem vbNullString
        
        For i = LBound(ListaRazas) To UBound(ListaRazas)
            frmMain.lstRazaOponente.AddItem ListaRazas(i)
        Next i
    
        Contador = 0
        
        frmMain.lstCascoOponente.Clear
     
        For i = 1 To NumMaximoData
            If ObjData(i).tipe = 3 And ObjData(i).SubTipo = 1 Then
                Contador = Contador + 1
            End If
        Next i
        
        ReDim Casco(1 To Contador)
        
        Contador = 0
        
        Call frmMain.lstCascoOponente.AddItem(0 & " - " & "Sin nada")
        
        For i = 1 To NumMaximoData
        
            If ObjData(i).tipe = 3 And ObjData(i).SubTipo = 1 Then
                Contador = Contador + 1
                Casco(Contador).ObjIndex = i
                Casco(Contador).ObjName = ObjData(i).Name
                Call frmMain.lstCascoOponente.AddItem(i & " - " & ObjData(i).Name)
            End If
        Next i
 
         
        Contador = 0
        
        frmMain.lstEscudoOponente.Clear
     
        For i = 1 To NumMaximoData
            If ObjData(i).tipe = 3 And ObjData(i).SubTipo = 2 Then
                Contador = Contador + 1
            End If
        Next i
        
        ReDim escudos(1 To Contador)
        
        Contador = 0
        
        Call frmMain.lstEscudoOponente.AddItem(0 & " - " & "Sin nada")
        
        For i = 1 To NumMaximoData
        
            If ObjData(i).tipe = 3 And ObjData(i).SubTipo = 2 Then
                Contador = Contador + 1
                escudos(Contador).ObjIndex = i
                escudos(Contador).ObjName = ObjData(i).Name
                Call frmMain.lstEscudoOponente.AddItem(i & " - " & ObjData(i).Name)
            End If
        Next i
 
        Contador = 0
        
        frmMain.ListaArmadurasOponente.Clear
     
        For i = 1 To NumMaximoData
            If ObjData(i).tipe = 3 And ObjData(i).SubTipo = 0 Then
                Contador = Contador + 1
            End If
        Next i
        
        ReDim armaduras(1 To Contador)
        
        Contador = 0
        
        Call frmMain.ListaArmadurasOponente.AddItem(0 & " - " & "Sin nada")
        
        For i = 1 To NumMaximoData
        
            If ObjData(i).tipe = 3 And ObjData(i).SubTipo = 0 Then
                Contador = Contador + 1
                armaduras(Contador).ObjIndex = i
                armaduras(Contador).ObjName = ObjData(i).Name
                Call frmMain.ListaArmadurasOponente.AddItem(i & " - " & ObjData(i).Name)
            End If
        Next i
        
        frmMain.lstDefensaExtraOponente.Clear
        
        Contador = 0
        
        For i = 1 To NumMaximoData
            If ObjData(i).tipe = 31 Or ObjData(i).tipe = 44 Then
                Contador = Contador + 1
            End If
        Next i
        
        ReDim defextra(1 To Contador)
        
        Contador = 0
        
        Call frmMain.lstDefensaExtraOponente.AddItem(0 & " - " & "Sin equipar")
        
        For i = 1 To NumMaximoData
        
            If ObjData(i).tipe = 31 Or ObjData(i).tipe = 44 Then
                Contador = Contador + 1
                defextra(Contador).ObjIndex = i
                defextra(Contador).ObjName = ObjData(i).Name
                Call frmMain.lstDefensaExtraOponente.AddItem(i & " - " & ObjData(i).Name)
            End If
        Next i

       frmMain.lstDefensaExtraOponente.ListIndex = 0
       frmMain.ListaArmadurasOponente.ListIndex = 0
       frmMain.lstEscudoOponente.ListIndex = 0
       frmMain.lstCascoOponente.ListIndex = 0
    
     End If
    
End Function
Sub Main()
        
    Dim i As Long
    
    DatosPath = App.Path & "\Datos\"
    
    frmMain.Show

    Call LoadObjData

    Call LoadGrhData
    
    Call LoadListas
    
    Call LoadBalance
    
    frmMain.lblcargando.Caption = "Listo para usar..."
    
    prgRun = True
    
    Do While prgRun
        
        If frmMain.Visible Then
            
            Call CheckKeys
            
        End If
        
        DoEvents
        
    Loop
    
    If prgRun = False Then
        Call UnloadAllForms
        End
    End If
    
End Sub

Sub UnloadAllForms()
    
    On Error GoTo UnloadAllForms_Err

    Dim miFrm As Form

    For Each miFrm In Forms
        Unload miFrm
        Set miFrm = Nothing
    Next

    Exit Sub

UnloadAllForms_Err:
    Resume Next
    
End Sub

Sub CheckKeys()
    
    On Error GoTo Check_Keys_Err
    
    If Not IsAppActive() Then Exit Sub
    
    If Input_Key_Get(17) Then
        Call frmMain.cmdAttack_Click
    End If
    
    Exit Sub

Check_Keys_Err:

End Sub
Private Function IsAppActive() As Boolean

    On Error GoTo IsAppActive_Err
    
    IsAppActive = (GetActiveWindow <> 0)
    
    Exit Function

IsAppActive_Err:
    IsAppActive = False
    
End Function
Public Function Input_Key_Get(ByVal key_code As Byte) As Boolean
 
 Input_Key_Get = (GetKeyState(key_code) < 0)
 
End Function
Public Function ObtenerRaza(ByVal Raza As String) As Byte

    Dim razas As String

    razas = UCase(Raza)

    Select Case UCase(razas)
        
        Case UCase("Humano")
            ObtenerRaza = 1
 
        Case UCase("Elfo")
            ObtenerRaza = 2
            
        Case UCase("Drow")
            ObtenerRaza = 3
            
        Case UCase("gnomo")
            ObtenerRaza = 4
            
        Case UCase("enano")
            ObtenerRaza = 5
            
        Case UCase("Orco")
            ObtenerRaza = 6
            
    End Select
    
End Function
Public Function ObtenerClase(ByVal clase As String) As Byte

    Dim Clases As String

    Clases = UCase(clase)

    Select Case UCase(Clases)
        
        Case UCase("CLERIGO")
            ObtenerClase = 1
        
        Case UCase("Mago")
            ObtenerClase = 2

        Case UCase("Guerrero")
            ObtenerClase = 3
            
        Case UCase("Asesino")
            ObtenerClase = 4
            
        Case UCase("ladron")
            ObtenerClase = 5
            
        Case UCase("Bardo")
            ObtenerClase = 6
            
        Case UCase("Druida")
            ObtenerClase = 7
            
        Case UCase("Gladiador")
            ObtenerClase = 8
            
        Case UCase("Paladin")
            ObtenerClase = 9
            
        Case UCase("Cazador")
            ObtenerClase = 10
            
        Case UCase("Pescador")
            ObtenerClase = 11
            
        Case UCase("Herrero")
            ObtenerClase = 12
            
        Case UCase("Leñador")
            ObtenerClase = 13
            
        Case UCase("Minero")
            ObtenerClase = 14
            
        Case UCase("Carpintero")
            ObtenerClase = 15
            
        Case UCase("Sastre")
            ObtenerClase = 16
            
        Case UCase("Mercenario")
            ObtenerClase = 17
            
        Case UCase("Nigromante")
            ObtenerClase = 18
            
    End Select
    
End Function

Sub DrawGrhtoHdc(desthDC As Long, ByVal grh_index As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean)
On Error GoTo error

Dim file_path As String
Dim src_x As Integer
Dim src_y As Integer
Dim src_width As Integer
Dim src_height As Integer
Dim hdcsrc As Long
Dim MaskDC As Long
Dim PrevObj As Long
Dim PrevObj2 As Long
Dim bRet As Boolean

If grh_index <= 0 Then Exit Sub
   'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(grh_index).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(grh_index).TileWidth * 16) + 16
        End If
    End If
    
    If v_centered Then
        If GrhData(grh_index).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(grh_index).TileHeight * 32) + 32
        End If
    End If

'If it's animated switch grh_index to first frame
If GrhData(grh_index).NumFrames <> 1 Then
grh_index = GrhData(grh_index).Frames(1)
End If

file_path = App.Path & "\Graficos\" & GrhData(grh_index).FileNum & ".bmp"
src_x = GrhData(grh_index).sX
src_y = GrhData(grh_index).sY
src_width = GrhData(grh_index).pixelWidth
src_height = GrhData(grh_index).pixelHeight
hdcsrc = CreateCompatibleDC(desthDC)

PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))

BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy

DeleteDC hdcsrc

    
Exit Sub

error:
  MsgBox grh_index
    
End Sub

Public Sub LoadObjData()

    On Error Resume Next
    
    Dim Object As Integer
    
    Dim Leer   As New clsIniReader
    Set Leer = New clsIniReader
 
    
    Call Leer.Initialize(DatosPath & "obj.dat")

    NumMaximoData = Val(Leer.GetValue("INIT", "NumOBJs"))
   
    ReDim Preserve ObjData(1 To NumMaximoData) As ObjData
                
        For Object = 1 To NumMaximoData
             
             frmMain.lblcargando = "Cargando obj.dat... " & Round(Object / NumMaximoData * 100, 2) & "%"
             
             ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")

             ObjData(Object).Grhindex = Leer.GetValue("OBJ" & Object, "GrhIndex")
             
             ObjData(Object).MaxDef = Leer.GetValue("OBJ" & Object, "MAXDEF")
             ObjData(Object).MinDef = Leer.GetValue("OBJ" & Object, "MINDEF")
             ObjData(Object).MaxHit = Leer.GetValue("OBJ" & Object, "MAXHIT")
             ObjData(Object).MinHit = Leer.GetValue("OBJ" & Object, "MINHIT")
             ObjData(Object).tipe = Leer.GetValue("OBJ" & Object, "Objtype")
     
             ObjData(Object).Nivel = Leer.GetValue("OBJ" & Object, "MinELV")
             ObjData(Object).SubTipo = Leer.GetValue("OBJ" & Object, "SubTipo")
             ObjData(Object).Apuñala = Leer.GetValue("OBJ" & Object, "Apuñala")
             ObjData(Object).ObjIndex = Object
              
             ObjData(Object).CuantoAumento = Val(Leer.GetValue("OBJ" & Object, "CuantoAumento"))
             
            DoEvents
            
        Next Object
        
    Set Leer = Nothing
    
    Exit Sub
 
ErrHandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description
    
End Sub
  
Private Function LoadGrhData() As Boolean

    On Error GoTo errorhandler
     
    Dim Grh     As Integer
    Dim Frame   As Integer
    Dim tempInt As Integer
    
    ReDim GrhData(0 To 40000) As GrhData

    Open App.Path & "\Datos\graficos.ind" For Binary Access Read As #1
       
    Seek #1, 1
       
    Get #1, , tempInt
    Get #1, , tempInt
    Get #1, , tempInt
    Get #1, , tempInt
    Get #1, , tempInt
     
    'Get first Grh Number
    Get #1, , Grh
       
    Do Until Grh <= 0
        'Get number of frames
        Get #1, , GrhData(Grh).NumFrames
           
        If GrhData(Grh).NumFrames <= 0 Then
            GoTo errorhandler

        End If
           
        ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
           
        If GrhData(Grh).NumFrames > 1 Then
           
            'Read a animation GRH set

            For Frame = 1 To GrhData(Grh).NumFrames
                Get #1, , GrhData(Grh).Frames(Frame)

                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > 40000 Then GoTo errorhandler
                
            Next Frame
           
            Get #1, , tempInt
               
            If tempInt <= 0 Then GoTo errorhandler
            GrhData(Grh).speed = CSng(tempInt)
               
            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight

            If GrhData(Grh).pixelHeight <= 0 Then GoTo errorhandler
               
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth

            If GrhData(Grh).pixelWidth <= 0 Then GoTo errorhandler
     
            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth

            If GrhData(Grh).TileWidth <= 0 Then GoTo errorhandler
     
            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight

            If GrhData(Grh).TileHeight <= 0 Then GoTo errorhandler
        Else
            'Read in normal GRH data
            Get #1, , GrhData(Grh).FileNum

            If GrhData(Grh).FileNum <= 0 Then GoTo errorhandler
     
            Get #1, , GrhData(Grh).sX

            If GrhData(Grh).sX < 0 Then GoTo errorhandler
               
            Get #1, , GrhData(Grh).sY

            If GrhData(Grh).sY < 0 Then GoTo errorhandler
     
            Get #1, , GrhData(Grh).pixelWidth

            If GrhData(Grh).pixelWidth <= 0 Then GoTo errorhandler
     
            Get #1, , GrhData(Grh).pixelHeight

            If GrhData(Grh).pixelHeight <= 0 Then GoTo errorhandler
     
            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
               
            GrhData(Grh).Frames(1) = Grh

        End If

        'Get Next Grh Number
        Get #1, , Grh
    Loop
       
Close #1


    LoadGrhData = True

Exit Function

errorhandler:

    
End Function

Sub AddtoRichTextBox(Text As String)


    With frmMain.RecTxt

    .SelFontName = "Tahoma"
    .SelFontSize = 8
    
    If (Len(.Text)) > 20000 Then .Text = vbNullString
    .SelStart = Len(frmMain.RecTxt.Text)
    .SelLength = 0
    

    .SelBold = 0
    .SelItalic = 0
            
    .SelColor = RGB(200, 185, 10)
    
    .SelText = Text & vbCrLf
 
    End With
    
End Sub

Public Function UsuarioAtacaUsuario() As Boolean

    On Error GoTo ErrHandler

    If UsuarioImpacto Then
        Call UserDañoUser
    Else
        Call AddtoRichTextBox("Has fallado el golpe")
        Call EnviarSonidoFallas
    End If

    UsuarioAtacaUsuario = True
    
    Exit Function
    
ErrHandler:

End Function

Public Function UsuarioImpacto() As Boolean

    On Error GoTo ErrHandler

    Dim ProbRechazo            As Long
    Dim Rechazo                As Boolean
    Dim ProbExito              As Long
    Dim PoderAtaque            As Long
    Dim UserPoderEvasion       As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma                   As Integer
    Dim Nudi                    As Integer
    Dim SkillTacticas          As Long
    Dim SkillDefensa           As Long

    UsuarioImpacto = False
    
    SkillTacticas = Oponente.SkillTacticas
    SkillDefensa = Oponente.SkillDefensaConEscudos

    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion
    
    If Oponente.Escudo > 0 Then
        UserPoderEvasionEscudo = PoderEvasionEscudo
        UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0
    End If
    
    If Atacante.Arma > 0 Then
        
        'Esta usando un arma ???
        If EsNudi = True Then
            Arma = 0
            Nudi = Atacante.Arma
        Else
            Arma = Atacante.Arma
            Nudi = 0
        End If
            
    End If
 
    If Arma > 0 Or Nudi > 0 Then
        
        If Arma > 0 Then
            
            'If ObjData(Arma).proyectil = 1 Then
            '    PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
            'ElseIf ObjData(Arma).proyectil = 2 Then
            '    PoderAtaque = PoderAtaqueArpon(AtacanteIndex)
            'Else
                PoderAtaque = PoderAtaqueArma 'Probabilidad de atacar
            'End If
        
        Else
        
            PoderAtaque = PoderAtaqueNudi
        End If
        
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))

    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
        

    ' el usuario esta usando un escudo ???
    If Oponente.Escudo > 0 Then
    
        'Fallo ???
        If Not UsuarioImpacto Then
            
            Dim SumaSkills As Integer
            
            SumaSkills = MaximoInt(1, SkillDefensa + SkillTacticas)

            ' Chances are rounded
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / SumaSkills))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            
            If Rechazo Then
                Call AddtoRichTextBox("El personaje rechazo el ataque con su escudo")
                Call EnviarSonidoFallasEscudo
            End If

        End If

    End If
    

    Exit Function
    
ErrHandler:

End Function
Private Sub EnviarSonidoFallas()
Call Reproducir_WAV(App.Path & "\Sonidos\2.wav", SND_FILENAME)
End Sub
Private Sub EnviarSonidoFallasEscudo()
Call Reproducir_WAV(App.Path & "\Sonidos\1.wav", SND_FILENAME)
End Sub
Private Sub EnviarSonidoGolpe()
Call Reproducir_WAV(App.Path & "\Sonidos\3.wav", SND_FILENAME)
End Sub

Private Sub asd()

End Sub
Private Sub Reproducir_WAV(Archivo As String, Flags As Long)
    Dim ret As Long
    ret = PlaySound(Archivo, ByVal 0&, Flags)
End Sub
Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a

    End If

End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b

    End If

End Function
Private Function PoderEvasionEscudo() As Long

    PoderEvasionEscudo = (Oponente.SkillDefensaConEscudos * ModClase(Oponente.clase).Escudo) * 0.5

End Function
Private Function PoderEvasion() As Long

Dim lTemp As Long

lTemp = (Oponente.SkillTacticas + Oponente.SkillTacticas / 33 * Oponente.agilidad) * ModClase(Oponente.clase).Evasion

PoderEvasion = (lTemp + (2.5 * MaximoInt(Oponente.Nivel - 12, 0)))

'Nuevo
PoderEvasion = 0.5 * PoderEvasion


End Function
Public Function EsNudi() As Boolean


    If ObjData(Atacante.Arma).tipe = 46 Then
        EsNudi = True
    Else
        EsNudi = False
    End If
    
End Function

Private Function PoderAtaqueArma() As Long

    Dim PoderAtaqueTemp As Long

    If Atacante.SkillCombateConArmas < 31 Then
        PoderAtaqueTemp = Atacante.SkillCombateConArmas * ModClase(Atacante.clase).AtaqueArmas
        
    ElseIf Atacante.SkillCombateConArmas < 61 Then
        PoderAtaqueTemp = (Atacante.SkillCombateConArmas + Atacante.agilidad) * ModClase(Atacante.clase).AtaqueArmas
    
    ElseIf Atacante.SkillCombateConArmas < 91 Then
        PoderAtaqueTemp = (Atacante.SkillCombateConArmas + 2 * Atacante.agilidad) * ModClase(Atacante.clase).AtaqueArmas
    
    Else
        PoderAtaqueTemp = (Atacante.SkillCombateConArmas + 3 * Atacante.agilidad) * ModClase(Atacante.clase).AtaqueArmas

    End If
    
    PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(Atacante.Nivel - 12, 0)))


End Function


Private Function PoderAtaqueNudi() As Long

    Dim PoderAtaqueTemp As Long

    If Atacante.SkillWrestling < 31 Then
        PoderAtaqueTemp = Atacante.SkillWrestling * ModClase(Atacante.clase).AtaqueWrestling
        
    ElseIf Atacante.SkillWrestling < 61 Then
        PoderAtaqueTemp = (Atacante.SkillWrestling + Atacante.agilidad) * ModClase(Atacante.clase).AtaqueWrestling
    
    ElseIf Atacante.SkillWrestling < 91 Then
        PoderAtaqueTemp = (Atacante.SkillWrestling + 2 * Atacante.agilidad) * ModClase(Atacante.clase).AtaqueWrestling
    
    Else
        PoderAtaqueTemp = (Atacante.SkillWrestling + 3 * Atacante.agilidad) * ModClase(Atacante.clase).AtaqueWrestling

    End If
    
    PoderAtaqueNudi = (PoderAtaqueTemp + (2.5 * MaximoInt(Atacante.Nivel - 12, 0)))


End Function

Private Function PoderAtaqueWrestling() As Long

    Dim PoderAtaqueTemp As Long

    If Atacante.SkillWrestling < 31 Then
        PoderAtaqueTemp = Atacante.SkillWrestling * ModClase(Atacante.clase).AtaqueArmas
        
    ElseIf Atacante.SkillWrestling < 61 Then
        PoderAtaqueTemp = (Atacante.SkillWrestling + Atacante.agilidad) * ModClase(Atacante.clase).AtaqueArmas
    
    ElseIf Atacante.SkillWrestling < 91 Then
        PoderAtaqueTemp = (Atacante.SkillWrestling + 2 * Atacante.agilidad) * ModClase(Atacante.clase).AtaqueArmas
    
    Else
        PoderAtaqueTemp = (Atacante.SkillWrestling + 3 * Atacante.agilidad) * ModClase(Atacante.clase).AtaqueArmas

    End If
    
    PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(Atacante.Nivel - 12, 0)))


End Function

Public Sub UserDañoUser()
    
    On Error GoTo ErrHandler

    Dim daño       As Long
    Dim lugar      As Byte
    Dim absorbido  As Long
    Dim defextra   As Integer
    
    Dim Obj As ObjData
 
    daño = CalcularDaño

    If Atacante.dañoextra > 0 Then
        Obj = ObjData(Atacante.dañoextra)
        daño = daño + RandomNumber(Obj.MinHit, Obj.MaxHit)
    End If

    If Oponente.DefensaExtra > 0 Then
        Obj = ObjData(Oponente.DefensaExtra)
        defextra = RandomNumber(Obj.MinDef, Obj.MaxDef)
    End If
    
    lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)

    If daño <= 0 Then daño = 1
    
    Select Case lugar

        Case PartesCuerpo.bCabeza

            'Si tiene casco absorbe el golpe
            If Oponente.Casco > 0 Then
                Obj = ObjData(Oponente.Casco)
                
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                absorbido = absorbido + defextra
                daño = daño - absorbido

            End If
        
        Case Else
 
            'Si tiene armadura absorbe el golpe
            If Oponente.Armadura > 0 Then
                Obj = ObjData(Oponente.Armadura)
                
                Dim Obj2 As ObjData

                If Oponente.Escudo Then
                    Obj2 = ObjData(Oponente.Escudo)
                    absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                Else
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If

                absorbido = absorbido + defextra
                daño = daño - absorbido

            End If

    End Select
    
    If daño <= 0 Then daño = 1
    
    Dim Lugartmp As String
    
    Select Case lugar
        
        Case 1
            Lugartmp = "cabeza"
            
        Case 2
            Lugartmp = "pierna izquierda"
      
        Case 3
            Lugartmp = "pierna derecha"
       
        Case 4
            Lugartmp = "brazo derecho"
       
        Case 5
            Lugartmp = "brazo izquierdo"
      
        Case 6
            Lugartmp = "torso"
            
    End Select
    
    Call AddtoRichTextBox("Golpeaste al objetivo en " & Lugartmp & " por " & daño & " puntos de vida.")
    Call EnviarSonidoGolpe
    
    If PuedeApuñalar Then
        Dim dañoApu As Long
        Dim apuñalar As Boolean
        
        Call DoApuñalar(dañoApu, apuñalar)
        
        If apuñalar Then
            daño = daño + dañoApu
            Call AddtoRichTextBox("Apuñalaste a tu objetivo por " & dañoApu & " de daño extra, en total le sacaste " & daño)
        End If

    End If

    Exit Sub
    
ErrHandler:

End Sub
 
Public Sub DoApuñalar(ByRef daño As Long, ByRef Apuñalo As Boolean)
    
    Dim Suerte As Integer
    Dim Skill  As Integer

    Skill = Atacante.SkillApuñalar
    Apuñalo = False
    
    Select Case Atacante.clase

        Case eClass.Asesino
            Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
        Case eClass.Clerigo, eClass.Paladin, eClass.Sastre
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
        Case eClass.Bardo
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)
            
    End Select
 
    If RandomNumber(0, 100) <= Suerte Then
 
       If Atacante.clase = eClass.Asesino Then
           daño = Round(daño * 1.4, 0)
       Else
           daño = Round(daño * 1.5, 0)
       End If
   
       Apuñalo = True
      
    Else
    
        Call AddtoRichTextBox("No has logrado apuñalar al objetivo.")

    End If

End Sub

Public Function PuedeApuñalar() As Boolean

    On Error GoTo errorhandler
    
    If Atacante.Arma > 0 Then
        If ObjData(Atacante.Arma).Apuñala = 1 Then
            PuedeApuñalar = Atacante.SkillApuñalar >= 5 Or Atacante.clase = eClass.Asesino

        End If

    End If

    Exit Function

errorhandler:
PuedeApuñalar = False
End Function
Public Function CalcularDaño() As Long
    
    On Error GoTo errorhandler
    
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim ModifClase As Single
    Dim DañoMaxArma As Long
    
    Dim Arma       As Integer
    Dim Nudillo As Integer
    Dim proyectil  As ObjData

    If Atacante.Arma > 0 Then
        
        If EsNudi = True Then
            Nudillo = 1
            Arma = 0
        Else
            Nudillo = 0
            Arma = 1
        End If
        
        If Arma > 0 Then
            Arma = Atacante.Arma
            Nudillo = 0
        ElseIf Nudillo > 0 Then
            Arma = Atacante.Arma
            Nudillo = 1
        End If
    
        DañoArma = RandomNumber(ObjData(Atacante.Arma).MinHit, ObjData(Atacante.Arma).MaxHit)
        DañoMaxArma = ObjData(Atacante.Arma).MaxHit
        
        'If Arma.proyectil > 0 Then
        '    If Arma.proyectil = 1 Then 'Arcos y Ballestas
        '        ModifClase = ModClase(.clase).DañoProyectiles
        '
        '        If Arma.Municion = 1 Then 'P
        '            proyectil = ObjData(.Invent.MunicionEqpObjIndex)
        '            DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
        '
        '        End If'

        '    Else ' 2 'Arrojadizas
        '        ModifClase = ModClase(.clase).AtaqueArpon
        '
        '    End If
        '
        'Else
            
            If Nudillo > 0 Then
                ModifClase = ModClase(Atacante.clase).DañoWrestling
            Else
                ModifClase = ModClase(Atacante.clase).DañoArmas
            End If
            
        'End If
        
    Else
     
        DañoArma = DañoArma + RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
        DañoMaxArma = DañoMaxArma + 3
        ModifClase = ModClase(Atacante.clase).DañoWrestling
        
    End If
    
    DañoUsuario = RandomNumber(Atacante.MinHit, Atacante.MaxHit)
    CalcularDaño = ((3 * DañoArma) + ((DañoMaxArma / 5) * MaximoInt(0, (Atacante.fuerza - 15))) + DañoUsuario) * ModifClase
 
    If ObjData(Atacante.Arma).CuantoAumento > 0 Then
        If ObjData(Atacante.Arma).SubTipo = 20 Or ObjData(Atacante.Arma).SubTipo = 21 Then
            CalcularDaño = CalcularDaño * 0.5
        End If

    End If

    Exit Function

errorhandler:
       CalcularDaño = 0
End Function
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long

    RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
writeprivateprofilestring Main, var, Value, File
End Sub
