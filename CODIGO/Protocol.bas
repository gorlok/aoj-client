Attribute VB_Name = "Protocol"
'/** author: @Oliva Juan Agustín
' ** userForos: Agushh, JAO, Thorkes
' ** date: 05/18
'\*  comments: n/a

'---------------------------------------------------------------------------------------------------------------------------
'AGUSH -Regla general- ; _
           En caso de recibir paquetes extensos, o en su defecto, que contengan cadenas de texto, deberá utilizarse _
        un BUFFER TEMPORAL. De lo contrario, existe posibilidad de que crashee el cliente (por que por ahí no se recibieron _
        aún todos los datos). La modalidad de uso de dicho buffer consiste en hacer una copia del incomingdata, _
        explicado a continuación:

'dim buffer as new clsbytequeue

'buffer.copyBuffer incomingdata

'[CUERPO DEL PROCEDIMIENTO] -> se leerían los datos recibidos del servidor. Ejemplo: buffer.readByte()

'incomingdata.copyBuffer(buffer)

'set buffer = nothing 'y finalmente deshacemos el temp buffer.
'---------------------------------------------------------------------------------------------------------------------------


'bytes que envía el cliente al servidor
Private Enum ClientSendMSG
    mlogged
    Move
    talk
    lC
    rc
    WLC
    cDir
    DirSP
    LH
    UK
    pickUp
    Drop
    Equip
    Attack
    useItem
    Salir
    commerceStart
    commerceEnd
    commerceBuy
    commerceSell
    meditate
    refreshPos
    createNewCharacter
    throwDices
    doLittleStats
    assingSkills
    doBank
    fiBank
    sellBank
    buyBank
    safeToggle 'request sv change my security attack
End Enum

'bytes que recibe el cliente del servidor
Private Enum ClientMessages
    none
    Login
    msgErr
    IndicePJ
    msg_CC
    CMP
    talk 'msj consola
    dialog 'msj pj
    MSG_CCNPC 'move npc char
    MSG_PU 'refresh char pos
    MSG_HO 'create object
    MSG_MP 'move char
    msg_bo 'delete object
    MSG_BQ 'block position
    MSG_TW 'play music format wave
    MSG_CP 'character change
    MSG_FX 'create fx
    MSG_CSI 'Change inventory slot
    MSG_SHS 'Change spell slot
    MSG_BP 'delete char
    MSG_TO1 'Messages by skills
    MSG_NPC_INV 'Received NPC Inv from commerce
    MSG_REFRESH 'refresh user stats
    MSG_FINCOM 'commerce end
    userHitNPC
    UserSwing
    tradeOk 'refresh commerce pics
    MeditateToggle
    DeleteObject
    dropDices
    littleStats
    userAtri
    UserSkills
    msgN1
    paradOk
    wBank
    oBank
    fbank
    userWork
    userRain
    finOk
    userRep
    bancoOk
    safeToggle
    updateStatus 'is blue or red ¿?
    areasChange
    navigateToggle
End Enum

Public Sub handleIncomingData()
On Error Resume Next

Dim packet As Byte

    packet = incomingData.PeekByte()
    
    Select Case (packet)
    
        Case ClientMessages.Login
             Call HandleLogged
             
        Case ClientMessages.msgErr
             Call handleMsgError
             
        Case ClientMessages.IndicePJ
             Call HandleUserCharIndexInServer
             
        Case ClientMessages.msg_CC
             Call Handlemsg_CC
            
        Case ClientMessages.CMP
             Call HandleChangeMap
             
        Case ClientMessages.talk
             Call handleMSG_TALK
             
        Case ClientMessages.dialog
             Call handleMSG_DIALOG
             
        Case ClientMessages.MSG_CCNPC
              Call handleMSG_CCNPC
            
       Case ClientMessages.MSG_PU
            Call handleMSG_PU
            
       Case ClientMessages.MSG_HO
            Call handleMSG_HO
            
       Case ClientMessages.MSG_MP
            Call HandleCharacterMove
            
       Case ClientMessages.msg_bo
            Call HandleObjectDelete
            
       Case ClientMessages.MSG_BQ
            Call HandleBlockPosition
            
       Case ClientMessages.MSG_TW
            Call HandlePlayWave
            
       Case ClientMessages.MSG_CP
            Call HandleCharacterChange
            
       Case ClientMessages.MSG_FX
            Call HandleCreateFX
            
       Case ClientMessages.MSG_CSI
            Call HandleChangeInventorySlot
            
       Case ClientMessages.MSG_SHS
            Call HandleChangeSpellSlot
            
       Case ClientMessages.MSG_BP
            Call HandleCharacterRemove
            
       Case ClientMessages.MSG_TO1
            Call handleMSG_MSJ01
            
       Case ClientMessages.MSG_NPC_INV
            Call HandleChangeNPCInventorySlot
            
       Case ClientMessages.MSG_REFRESH
            Call handleMSG_REFRESH
            
       Case ClientMessages.MSG_FINCOM
            Call HandleCommerceEnd
            
        Case ClientMessages.userHitNPC
            Call HandleUserHitNPC
            
        Case ClientMessages.UserSwing
            Call HandleUserSwing
            
        Case ClientMessages.tradeOk
            Call HandleTradeOK
            
        Case ClientMessages.MeditateToggle
            Call HandleMeditateToggle
            
        Case ClientMessages.DeleteObject
            Call handleSetNullObject
                
        Case ClientMessages.dropDices
            Call handleThrowDices
            
        Case ClientMessages.littleStats
            Call handleLittleStats
            
        Case ClientMessages.userAtri
            Call handleuserAtris
            
        Case ClientMessages.UserSkills
            Call HandleSendSkills
            
        Case ClientMessages.msgN1
            Call handleMsgN1
            
        Case ClientMessages.paradOk
            Call handleParadOk
            
        Case ClientMessages.wBank
            Call handlewBank
            
        Case ClientMessages.oBank
            Call handleoBank
            
        Case ClientMessages.fbank
            Call handlefBank
            
        Case ClientMessages.userWork
            Call handleUserWork
            
        Case ClientMessages.userRain
            Call handleUserRain
            
        Case ClientMessages.finOk
            Call handleDisconnect
            
        Case ClientMessages.userRep
            Call handleUserRep
            
        Case ClientMessages.bancoOk
            Call HandleBankOK
            
        Case ClientMessages.safeToggle
            Call handleSafeToggle
            
        Case ClientMessages.updateStatus
            Call handleUpdateStatus
            
        Case ClientMessages.areasChange
            Call handleUpdateArea
            
        Case ClientMessages.navigateToggle
            Call handleNavigateToggle
            
       Case Else
           
           'Call MsgBox("msg_id: " & packet & " no reconocido")
           Debug.Print "msg_id: " & packet & " no reconocido"
           
        Exit Sub
                
    End Select
    
        'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call handleIncomingData
    End If
    
End Sub

Private Sub HandleLogged()

    Call incomingData.ReadByte
    
    ' Variable initialization
    EngineRun = True
    Nombres = True
    
    'Set connected state
    Call SetConnected
    
    'Show tip
    If tipf = "1" And PrimeraVez Then
        Call CargarTip
        frmtip.Visible = True
        PrimeraVez = False
    End If
End Sub

Private Sub HandleUserCharIndexInServer()

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()
    UserPos.x = charlist(UserCharIndex).Pos.x
    UserPos.y = charlist(UserCharIndex).Pos.y
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)

    frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
End Sub

Public Sub Handlemsg_CC()

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler

    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    Dim CharIndex As Integer
    Dim body As Integer
    Dim head As Integer
    Dim Heading As E_Heading
    Dim x As Integer
    Dim y As Integer
    Dim weapon As Integer
    Dim shield As Integer
    Dim helmet As Integer
    Dim privs As Integer
    
    CharIndex = Buffer.ReadInteger()
    body = Buffer.ReadInteger()
    head = Buffer.ReadInteger()
    Heading = Buffer.ReadInteger()
    x = Buffer.ReadInteger()
    y = Buffer.ReadInteger()
    weapon = Buffer.ReadInteger()
    shield = Buffer.ReadInteger()
    helmet = Buffer.ReadInteger()
    
    With charlist(CharIndex)
        Call SetCharacterFx(CharIndex, Buffer.ReadInteger(), Buffer.ReadInteger())
        
        .Nombre = Buffer.ReadASCIIString()
        .Criminal = Buffer.ReadInteger()
        
        privs = Buffer.ReadInteger()
        
        If privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil
            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0
        End If
    End With
    
    Call MakeChar(CharIndex, body, head, Heading, x, y, weapon, shield, helmet)
    Call RefreshAllChars
    
    Call incomingData.CopyBuffer(Buffer)
    Set Buffer = Nothing
    
    Exit Sub
    
ErrHandler:
    Dim error As Long
    error = Err.number
    Set Buffer = Nothing
On Error GoTo 0

    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleChangeMap()
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()
    
    If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
        Call SwitchMap(UserMap)
        If bLluvia(UserMap) = 0 Then
            If bRain Then
                Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
            End If
        End If
    Else
        'no encontramos el mapa en el hd
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
        Call CloseClient
    End If
End Sub

Public Sub handleMsgError()
On Error GoTo Err

If incomingData.length < 2 Then
   Err.Raise incomingData.NotEnoughDataErrCode
Exit Sub
End If

Dim Buffer As New clsByteQueue
Buffer.CopyBuffer incomingData

Call Buffer.ReadByte
Call MsgBox(Buffer.ReadASCIIString)

incomingData.CopyBuffer Buffer

Err:

    Dim error As Long
    error = Err.number
    Set Buffer = Nothing
On Error GoTo 0

    If error <> 0 Then _
        Err.Raise error

End Sub

Public Sub handleMSG_TALK()
On Error GoTo Err

Dim Buffer As New clsByteQueue
    Buffer.CopyBuffer incomingData
    
    Buffer.ReadByte

    Dim txt As String
    Dim FONTTYPE As String
    Dim Colors() As String
                
    txt = Buffer.ReadASCIIString
    FONTTYPE = Buffer.ReadASCIIString
                
    Colors = Split(FONTTYPE, "~")

    AddtoRichTextBox frmMain.RecTxt, txt, CInt(Colors(1)), CInt(Colors(2)), CInt(Colors(3)), CByte(Colors(4))
    
    incomingData.CopyBuffer Buffer

Err:

    Dim error As Long
    error = Err.number
    Set Buffer = Nothing
On Error GoTo 0

    If error <> 0 Then _
        Err.Raise error

    
End Sub

'agush: cuando existan cadenas recibidas debemos usar un buffer temporal para no afectar negativamente _
 , por las dudas, el incomingdata
Public Sub handleMSG_DIALOG()
On Error GoTo Err

Dim Buffer As New clsByteQueue

    Buffer.CopyBuffer incomingData
    
    Buffer.ReadByte
            
    Dim txt As String
    Dim color As Long
    Dim iuser As Integer
            
    color = Buffer.ReadLong()
    txt = Buffer.ReadASCIIString()
    iuser = Buffer.ReadInteger()
            
    Dialogos.CreateDialog txt, iuser, color
    
    incomingData.CopyBuffer Buffer
    
Err:
    Dim error As Long
    error = Err.number
    Set Buffer = Nothing
On Error GoTo 0

    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub handleMSG_CCNPC() ' crea un npcchar
Dim Buffer As New clsByteQueue
Dim body As Integer, head As Integer, dir As Byte, CharIndex As Integer, x As Integer, y As Integer

    Call Buffer.CopyBuffer(incomingData)
    
    Call Buffer.ReadByte
          
    body = Buffer.ReadInteger()
    head = Buffer.ReadInteger()
    dir = Buffer.ReadInteger()
          
    CharIndex = Buffer.ReadInteger()
          
    x = Buffer.ReadInteger()
    y = Buffer.ReadInteger()
          
    Call MakeChar(CharIndex, body, head, dir, x, y, 0, 0, 0)
            
    Call RefreshAllChars
    
    Call incomingData.CopyBuffer(Buffer)
    
Err:

    Dim error As Long
    error = Err.number
    Set Buffer = Nothing
On Error GoTo 0

    If error <> 0 Then _
        Err.Raise error

End Sub

Public Sub handleMSG_PU()
Call incomingData.ReadByte

MapData(UserPos.x, UserPos.y).CharIndex = 0
UserPos.x = incomingData.ReadInteger()
UserPos.y = incomingData.ReadInteger()
            
MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
charlist(UserCharIndex).Pos = UserPos
End Sub

Private Sub HandleCharacterMove()

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim x As Integer
    Dim y As Integer
    
    CharIndex = incomingData.ReadInteger()
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    
    With charlist(CharIndex)
    
        If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .FxIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If
    End With
    
    
    If charlist(CharIndex).Nombre <> "" Then Debug.Print x & "-" & y
    Call MoveCharbyPos(CharIndex, x, y)
    
    Call RefreshAllChars
End Sub

Public Sub handleMSG_HO()

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x   As Integer
    Dim y   As Integer
    Dim Grh As Integer
    
    Grh = incomingData.ReadInteger()
    
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    
    MapData(x, y).ObjGrh.GrhIndex = Grh
    
    Call InitGrh(MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex)
End Sub

Private Sub HandleObjectDelete()

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    
    MapData(x, y).ObjGrh.GrhIndex = 0
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()

    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    
    If incomingData.ReadInteger() Then
        MapData(x, y).Blocked = 1
    Else
        MapData(x, y).Blocked = 0
    End If
End Sub

Public Sub HandlePlayWave()
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Byte
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadByte()
    srcX = CByte(incomingData.ReadInteger())
    srcY = CByte(incomingData.ReadInteger())
        
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'***************************************************
    If incomingData.length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim headIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    With charlist(CharIndex)
        tempint = incomingData.ReadInteger()
        
        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .body = BodyData(0)
            .iBody = 0
        Else
            .body = BodyData(tempint)
            .iBody = tempint
        End If
        
        
        headIndex = incomingData.ReadInteger()
        
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .head = HeadData(0)
            .iHead = 0
        Else
            .head = HeadData(headIndex)
            .iHead = headIndex
        End If
        
        .muerto = (headIndex = CASPER_HEAD)
        
        .Heading = incomingData.ReadInteger()
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Arma = WeaponAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Casco = CascoAnimData(tempint)
        
        Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
    End With
    
    Call RefreshAllChars
End Sub

Private Sub HandleCreateFX()

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetCharacterFx(CharIndex, fX, Loops)
End Sub

Public Sub handleSetNullObject()

    Call incomingData.ReadByte
    
    Dim slot As Byte
    slot = incomingData.ReadByte()
    
    Call Inventario.SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, "Nada")

End Sub


Private Sub HandleChangeSpellSlot()
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo ErrHandler
    
    Dim Buffer As New clsByteQueue
    
    Buffer.CopyBuffer incomingData
    
    Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    UserHechizos(slot) = Buffer.ReadInteger()
    
    If slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.List(slot - 1) = Buffer.ReadASCIIString()
    Else
        Call frmMain.hlst.AddItem(Buffer.ReadASCIIString())
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error

End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call EraseChar(CharIndex)
    Call RefreshAllChars
End Sub

Public Sub handleMSG_MSJ01()
Call incomingData.ReadByte

UsingSkill = CInt(incomingData.ReadLong())

            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
            End Select

End Sub

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim slot As Byte
    Dim OBJIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Integer
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim defense As Integer
    Dim value As Single
    
    On Error GoTo Err
    
    Dim Buffer As New clsByteQueue
    
    Call Buffer.CopyBuffer(incomingData)
    
    slot = Buffer.ReadByte()
    OBJIndex = Buffer.ReadInteger()
    
    If OBJIndex > 0 Then
    
    Name = Buffer.ReadASCIIString()
    Amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadInteger()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    defense = Buffer.ReadInteger()
    value = Buffer.ReadLong()
    
    Call Inventario.SetItem(slot, OBJIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, defense, value, Name)
    
    Else
    
    Call Inventario.SetItem(slot, OBJIndex, 0, 0, 0, 0, 0, 0, 0, 0, "Nada")
    
    End If
    
    Call incomingData.CopyBuffer(Buffer)

Err:

    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
    
End Sub


Private Sub HandleChangeNPCInventorySlot()

    If incomingData.length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo Err
    
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With NPCInventory(slot)
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        
        If (.Amount > 0) Then
        .Valor = Buffer.ReadLong()
        .GrhIndex = Buffer.ReadInteger()
        .OBJIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .Def = Buffer.ReadInteger()
        End If
        
    End With
    
    If frmComerciar.List1(0).ListCount >= slot Then _
        Call frmComerciar.List1(0).RemoveItem(slot - 1)
    
    Call frmComerciar.List1(0).AddItem(NPCInventory(slot).Name, slot - 1)
    
    If Not frmComerciar.Visible Then frmComerciar.Show , frmMain
    
    Call incomingData.CopyBuffer(Buffer)
    
Err:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
    
End Sub


'by agus
Public Sub handleMSG_REFRESH()
Dim id As Byte

With incomingData

     Call .ReadByte
     
     id = .ReadByte
     
     Select Case (id)
     
            Case 1 'update gold
                 UserGLD = .ReadLong()
                 frmMain.GldLbl.Caption = UserGLD
                 
            Case 2 'update user hp
                 UserMinHP = .ReadInteger()
                 UserMaxHP = .ReadInteger()
                 frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
                 
            Case 3 'update user mana
                UserMinMAN = .ReadInteger()
                UserMaxMAN = .ReadInteger()
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
                
            Case 4 'update user stamina
                UserMinSTA = .ReadInteger()
                UserMaxSTA = .ReadInteger()
                frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
                
            Case 5 'update user elv, exp and elu
                'frmMain.Label8.Caption = val(.ReadLong())
                UserLvl = val(.ReadLong)
                UserExp = .ReadLong()
                UserPasarNivel = .ReadLong()
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
                frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
                 
     
     End Select
     

End With
End Sub

Private Sub HandleCommerceEnd()

    'Remove packet ID
    Call incomingData.ReadByte
    
    'Clear item's list
    frmComerciar.List1(0).Clear
    frmComerciar.List1(1).Clear
    
    'Reset vars
    Comerciando = False
    
    'Hide form
    Unload frmComerciar
End Sub

Private Sub HandleUserHitNPC()

    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    

    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, False)
End Sub


Private Sub HandleUserSwing()
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
End Sub

Private Sub HandleTradeOK()

'we remove first byte at queue
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then
        Dim i As Long
        
        Call frmComerciar.List1(1).Clear
        
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                Call frmComerciar.List1(1).AddItem(Inventario.ItemName(i))
            Else
                Call frmComerciar.List1(1).AddItem("")
            End If
        Next i
        
        If frmComerciar.LasActionBuy Then
            frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
            frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
        Else
            frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
            frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
        End If
    End If
End Sub

Private Sub HandleMeditateToggle()
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar
End Sub

Public Sub handleThrowDices()

Call incomingData.ReadByte
Dim i As Long

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = incomingData.ReadByte
    Next i
    
            With frmCrearPersonaje
            If .Visible Then
                .lbFuerza.Caption = UserAtributos(1)
                .lbAgilidad.Caption = UserAtributos(2)
                .lbInteligencia.Caption = UserAtributos(3)
                .lbCarisma.Caption = UserAtributos(4)
                .lbConstitucion.Caption = UserAtributos(5)
            End If
        End With

End Sub

Public Sub handleLittleStats()

    If incomingData.length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    Dim Buffer As New clsByteQueue
    Buffer.CopyBuffer incomingData
    
    Call Buffer.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = Buffer.ReadLong()
        .CriminalesMatados = Buffer.ReadLong()
        .UsuariosMatados = Buffer.ReadLong()
        .NpcsMatados = Buffer.ReadLong()
        .Clase = Buffer.ReadASCIIString
        .PenaCarcel = Buffer.ReadLong()
    End With
    
    incomingData.CopyBuffer Buffer

ErrHandler:

    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
    
End Sub

Public Sub handleuserAtris()
Call incomingData.ReadByte
Dim i As Long

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = incomingData.ReadByte
    Next i
        
        LlegaronAtrib = True
End Sub

Private Sub HandleSendSkills()

    If incomingData.length < 1 + NUMSKILLS Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Err
    
    Dim Buffer As New clsByteQueue
    Buffer.CopyBuffer incomingData
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim i As Long
    
    For i = 0 To NUMSKILLS
        If (i = 0) Then
           SkillPoints = Buffer.ReadByte()
        Else
           frmEstadisticas.setSkills i, Buffer.ReadByte
        End If
    Next i
    
    frmEstadisticas.setSkFree SkillPoints
    
    LlegaronSkills = True
    
    incomingData.CopyBuffer Buffer
    
Err:

    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
    
    
End Sub

'@Thorkes: ahora manejo el npcSwing y npcHit cuerpo a cuerpo desde un mismo método.
Public Sub handleMsgN1()

'agus; remuevo de la cola al primer byte, que es la identidad del paquete
Call incomingData.ReadByte
 
Select Case (incomingData.ReadByte)

       Case CByte(0) 'npc swing
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            
       Case Else 'npc hit me!! :@
       
            Select Case incomingData.ReadLong()
                   Case bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadLong()) & "!!", 255, 0, 0, True, False, False)
                   Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadLong()) & "!!", 255, 0, 0, True, False, False)
                   Case bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadLong()) & "!!", 255, 0, 0, True, False, False)
                   Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadLong()) & "!!", 255, 0, 0, True, False, False)
                   Case bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadLong()) & "!!", 255, 0, 0, True, False, False)
                   Case bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadLong() & "!!"), 255, 0, 0, True, False, False)
             End Select

End Select


End Sub

Public Sub handleParadOk()
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado
End Sub

Public Sub handlewBank()
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmBancoObj.List1(1).Clear
    
    'Fill the inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
        Else
            frmBancoObj.List1(1).AddItem ""
        End If
    Next i
    
    Call frmBancoObj.List1(0).Clear
    
    'Fill the bank list
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        If UserBancoInventory(i).OBJIndex <> 0 Then
            frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
        Else
            frmBancoObj.List1(0).AddItem ""
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    frmBancoObj.Show , frmMain
End Sub

Public Sub handleoBank()

    If incomingData.length < 16 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Err

    Dim Buffer As New clsByteQueue
    Buffer.CopyBuffer incomingData
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    
    slot = Buffer.ReadByte()
    UserBancoInventory(slot).OBJIndex = Buffer.ReadInteger()
    
    If (UserBancoInventory(slot).OBJIndex) > 0 Then
    
        With UserBancoInventory(slot)
             .Name = Buffer.ReadASCIIString()
             .Amount = Buffer.ReadLong
             .GrhIndex = Buffer.ReadInteger()
             .OBJType = Buffer.ReadByte()
             .MaxHit = Buffer.ReadInteger()
             .MinHit = Buffer.ReadInteger()
             .Def = Buffer.ReadInteger()
             .Valor = 0
        End With
    
        If frmBancoObj.List1(0).ListCount >= slot Then _
            Call frmBancoObj.List1(0).RemoveItem(slot - 1)
            Call frmBancoObj.List1(0).AddItem(UserBancoInventory(slot).Name, slot - 1)
     Else
           With UserBancoInventory(slot)
             .Name = ""
             .Amount = 0
             .GrhIndex = 0
             .OBJType = 0
             .MaxHit = 0
             .MinHit = 0
             .Def = 0
             .Valor = 0
            Call frmBancoObj.List1(0).RemoveItem(slot - 1)
           End With
    End If
    
    incomingData.CopyBuffer Buffer
    
Err:

    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error

End Sub

Public Sub handlefBank()

    Call incomingData.ReadByte

    frmBancoObj.List1(0).Clear
    frmBancoObj.List1(1).Clear
    
    Unload frmBancoObj
    Comerciando = False

End Sub

Public Sub handleUserWork()
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadLong()

    frmMain.MousePointer = 2
    
    Select Case UsingSkill
        Case Magia
            If (macroL < 1) Then Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 255, 255, 255, 1, 0)
        Case Pesca
            If (macroL < 1) Then Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 255, 255, 255, 1, 0)
        Case Robar
            If (macroL < 1) Then Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 255, 255, 255, 1, 0)
        Case Talar
            If (macroL < 1) Then Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 255, 255, 255, 1, 0)
        Case Mineria
            If (macroL < 1) Then Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 255, 255, 255, 1, 0)
        Case FundirMetal
            If (macroL < 1) Then Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 255, 255, 255, 1, 0)
        Case Proyectiles
            If (macroL < 1) Then Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 255, 255, 255, 1, 0)
    End Select
    
    If frmMain.t_TR.Enabled = False Then frmMain.t_TR.Enabled = True
    
End Sub

Public Sub handleUserRain()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
    
    bTecho = (MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4)
    If bRain Then
        If bLluvia(UserMap) Then
            'Stop playing the rain sound
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            If bTecho Then
                Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
            Else
                Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
            End If
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    
    bRain = Not bRain
End Sub

Public Sub handleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Close connection
#If UsarWrench = 1 Then
    frmMain.Socket1.Disconnect
#Else
    If frmMain.Winsock1.State <> sckClosed Then _
        frmMain.Winsock1.Close
#End If
    
    'Hide main form
    frmMain.Visible = False
    frmMain.Label1.Visible = False
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    
    'Show connection form
    frmConnect.Visible = True
    
    'Reset global vars
    IScombate = False
    UserDescansar = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    bRain = False
    bFogata = False
    SkillPoints = 0
    
    'Delete all kind of dialogs
    Call CleanDialogs
    
    'Reset some char variables...
    For i = 1 To LastChar
        charlist(i).invisible = False
    Next i
    
    'Unload all forms except frmMain and frmConnect
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> frmMain.Name And frm.Name <> frmConnect.Name Then
            Unload frm
        End If
    Next
    
#If SeguridadAlkon Then
    Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
End Sub

Public Sub handleUserRep()

On Error GoTo Err

Dim Buffer As New clsByteQueue
Buffer.CopyBuffer incomingData

With Buffer

     Call .ReadByte

     UserReputacion.AsesinoRep = .ReadLong()
     UserReputacion.BandidoRep = .ReadLong()
     UserReputacion.BurguesRep = .ReadLong()
     UserReputacion.LadronesRep = .ReadLong()
     UserReputacion.NobleRep = .ReadLong()
     UserReputacion.PlebeRep = .ReadLong()
     
     LlegoFama = True

End With

incomingData.CopyBuffer Buffer

Err:
    Dim error As Long
    error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If error <> 0 Then _
        Err.Raise error

End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long

    If frmBancoObj.Visible Then
        
        Call frmBancoObj.List1(1).Clear
        
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                Call frmBancoObj.List1(1).AddItem(Inventario.ItemName(i))
            Else
                Call frmBancoObj.List1(1).AddItem("")
            End If
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else
            frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
        
        frmBancoObj.NoPuedeMover = False
    End If
       
End Sub

Public Sub handleSafeToggle()
Call incomingData.ReadByte

ATSecurity = incomingData.ReadByte

Select Case ATSecurity

       Case 0
           Call frmMain.DibujarSeguro
           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 0, 255, 0, True, False, False)

       Case 1
           Call frmMain.DibujarSeguro
           Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)

End Select

End Sub

Public Sub handleUpdateStatus()
Call incomingData.ReadByte

Dim CharIndex As Integer
CharIndex = incomingData.ReadInteger()

charlist(CharIndex).Criminal = incomingData.ReadByte()

End Sub

Public Sub handleUpdateArea()
Call incomingData.ReadByte

Call ModAreas.CambioDeArea(incomingData.ReadInteger, incomingData.ReadInteger)

End Sub

Public Sub handleNavigateToggle()
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando
End Sub

Public Sub writeLoginExistingChar()
 With outgoingData
 
      .WriteByte ClientSendMSG.mlogged
      .WriteASCIIString UserName
      .WriteASCIIString UserPassword

 End With
End Sub

Public Sub writeMoveTo(ByVal dir As Integer)
With outgoingData

     .WriteByte ClientSendMSG.Move
     .WriteInteger (CInt(dir))

End With
End Sub

Public Sub writeTalk(ByVal str As String)
With outgoingData
   
    .WriteByte ClientSendMSG.talk
    .WriteASCIIString (CStr(str))

End With
End Sub

Public Sub writeLeftClick(ByVal x As Integer, ByVal y As Integer)
With outgoingData

   .WriteByte ClientSendMSG.lC
   .WriteInteger CInt(x)
   .WriteInteger CInt(y)

End With
End Sub

Public Sub writeRightClick(ByVal x As Integer, ByVal y As Integer)
With outgoingData

   .WriteByte ClientSendMSG.rc
   .WriteInteger CInt(x)
   .WriteInteger CInt(y)
   
End With
End Sub

Public Sub writeWLC(ByVal x As Integer, ByVal y As Integer, ByVal useSkill As Integer)
With outgoingData
     .WriteByte ClientSendMSG.WLC
     
     .WriteInteger CInt(x)
     .WriteInteger CInt(y)
     .WriteInteger CInt(useSkill)
     
End With
End Sub

Public Sub writeChangeHeading(ByVal Heading As Integer)
With outgoingData

     .WriteByte ClientSendMSG.cDir
     .WriteInteger Heading

End With
End Sub

Public Sub writeMoveSpell(ByVal dir As Integer, ByVal h As Integer)
With outgoingData

     .WriteByte ClientSendMSG.DirSP
     .WriteInteger CInt(dir)
     .WriteInteger CInt(h)
     
End With
End Sub

Public Sub writeThrowSpell(ByVal index As Integer)
With outgoingData

     .WriteByte ClientSendMSG.LH
     .WriteInteger CInt(index)
     
End With
End Sub

Public Sub writeDOUk(ByVal val As Integer)
With outgoingData

     .WriteByte ClientSendMSG.UK
     .WriteInteger CInt(val)
     
End With
End Sub

Public Sub writePickUp()
With outgoingData

     .WriteByte ClientSendMSG.pickUp

End With
End Sub

Public Sub writeRefreshPos()
With outgoingData
     .WriteByte ClientSendMSG.refreshPos
End With
End Sub

Public Sub WriteDrop(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientSendMSG.Drop)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
        
    End With
End Sub

Public Sub WriteEquipItem(ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientSendMSG.Equip)
        Call .WriteByte(slot)
        
    End With
End Sub

Public Sub WriteAttack()
    Call outgoingData.WriteByte(ClientSendMSG.Attack)
End Sub

Public Sub WriteUseItem(ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientSendMSG.useItem)
        Call .WriteByte(slot)
    End With
End Sub

Public Sub WriteSalir()
    With outgoingData
        Call .WriteByte(ClientSendMSG.Salir)
    End With
End Sub

Public Sub writeCommerceStart()
    With outgoingData
        Call .WriteByte(ClientSendMSG.commerceStart)
    End With
End Sub

Public Sub writeCommerceEnd()
With outgoingData
     Call .WriteByte(ClientSendMSG.commerceEnd)
End With
End Sub

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientSendMSG.commerceBuy)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal Amount As Integer)

    With outgoingData
        Call .WriteByte(ClientSendMSG.commerceSell)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteMeditate()

    With outgoingData
        Call .WriteByte(ClientSendMSG.meditate)
    End With
End Sub

Public Sub writeRequestStats()
    With outgoingData
        Call .WriteByte(ClientSendMSG.doLittleStats)
    End With
End Sub


Public Sub writeCharacterCreate(ByVal Nick As String, ByVal pass As String, ByVal race As String, _
                                ByVal gender As Integer, ByVal Class As String, ByVal email As String, ByVal home As String)
    With outgoingData
    
        Call .WriteByte(ClientSendMSG.createNewCharacter)
             .WriteASCIIString Nick
             .WriteASCIIString pass
             .WriteASCIIString race
             .WriteInteger gender
             .WriteASCIIString Class
             .WriteASCIIString email
             .WriteASCIIString home
             
    End With
End Sub

Public Sub writeTirDad()
    With outgoingData
        Call .WriteByte(ClientSendMSG.throwDices)
    End With
End Sub

Public Sub writeAssingSkills(ByVal sk As Long, ByVal Amount As Byte)
    With outgoingData
        Call .WriteByte(ClientSendMSG.assingSkills)
        Call .WriteLong(sk)
        Call .WriteByte(Amount)
    End With
End Sub

Public Sub writeDoBank()
    With outgoingData
        Call .WriteByte(ClientSendMSG.doBank)
    End With
End Sub

Public Sub writeEndBank()
    With outgoingData
        Call .WriteByte(ClientSendMSG.fiBank)
    End With
End Sub

Public Sub writeSellBank(ByVal slot As Integer, ByVal cant As Long)
With outgoingData

     .WriteByte ClientSendMSG.sellBank
     .WriteInteger slot
     .WriteLong cant

End With
End Sub

Public Sub writeBuyBank(ByVal slot As Integer, ByVal cant As Long)
With outgoingData

     .WriteByte ClientSendMSG.buyBank
     .WriteInteger slot
     .WriteLong cant

End With
End Sub

Public Sub writeSafeToggle()
With outgoingData

     .WriteByte ClientSendMSG.safeToggle
     
     ATSecurity = Not ATSecurity
     
     Select Case ATSecurity

            Case 0
                 Call frmMain.DibujarSeguro
                 Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 0, 255, 0, True, False, False)

            Case 1
                 Call frmMain.DibujarSeguro
                 Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)

      End Select
      
End With
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call SendData(sndData)
    End With
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.IsWritable Then
        'Put data back in the bytequeue
        Call outgoingData.WriteASCIIStringFixed(sdData)
        
        Exit Sub
    End If
    
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If
    
    'Send data!
#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub


