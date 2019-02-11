Attribute VB_Name = "protocolCmdParse"
'Argentum Online
'
'Copyright (C) 2006 Juan Martín Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)

'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'

Option Explicit

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

Public Sub AuxWriteWhisper(ByVal UserName As String, ByVal Mensaje As String)
    If LenB(UserName) = 0 Then Exit Sub
    
    Dim i As Long
    Dim nameLength As Long
    
    If (InStrB(UserName, "+") <> 0) Then
        UserName = Replace$(UserName, "+", " ")
    End If
    
    UserName = UCase$(UserName)
    nameLength = Len(UserName)
    
    i = 1
    Do While i <= LastChar
        If UCase$(charlist(i).Nombre) = UserName Or UCase$(Left$(charlist(i).Nombre, nameLength + 2)) = UserName & " <" Then
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    
    If i <= LastChar Then
        'Call WriteWhisper(i, Mensaje)
    End If
End Sub

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modification: 26/03/2009
'Interpreta, valida y ejecuta el comando ingresado
'26/03/2009: ZaMa - Flexibilizo la cantidad de parametros de /nene,  /onlinemap y /telep
'***************************************************
    Dim TmpArgos() As String
    
    Dim Comando As String
    Dim ArgumentosAll() As String
    Dim ArgumentosRaw As String
    Dim Argumentos2() As String
    Dim Argumentos3() As String
    Dim Argumentos4() As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments As Boolean
    
    Dim tmpArr() As String
    Dim tmpInt As Integer
    
    ' TmpArgs: Un array de a lo sumo dos elementos,
    ' el primero es el comando (hasta el primer espacio)
    ' y el segundo elemento es el resto. Si no hay argumentos
    ' devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = Trim$(UCase$(TmpArgos(0)))
    
    If UBound(TmpArgos) > 0 Then
        ' El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        ' Un array separado por blancos, con tantos elementos como
        ' se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        ' 2, 3 y 4 elementos respectivamente. Eso significa
        ' que pueden tener menos, por lo que es imperativo
        ' preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0
    End If
    
    ' Sacar cartel APESTA!! (y es ilógico, estás diciendo una pausa/espacio  :rolleyes: )
    If Comando = "" Then Comando = " "
    
    If Left$(Comando, 1) = "/" Then
        ' Comando normal
        
        Select Case Comando
            Case "/SEG"
                
                
            Case "/ONLINE"
                
                
            Case "/SALIR"
               WriteSalir
                
            Case "/SALIRCLAN"
                
                
            Case "/BALANCE"
                
            Case "/QUIETO"
                
            Case "/ACOMPAÑAR"
                
            Case "/ENTRENAR"
                
            Case "/DESCANSAR"
                
            Case "/MEDITAR"
               WriteMeditate
        
            Case "/RESUCITAR"
                
            Case "/CURAR"
                              
            Case "/EST"
            
            Case "/AYUDA"
                
            Case "/COMERCIAR"
            writeCommerceStart
                
            Case "/BOVEDA"
            writeDoBank
                
            Case "/ENLISTAR"
                    
            Case "/INFORMACION"
                
            Case "/RECOMPENSA"
                
            Case "/MOTD"
                
            Case "/UPTIME"
                
            Case "/SALIRPARTY"
                
            Case "/CREARPARTY"
                
            Case "/PARTY"
        
            Case "/CMSG"
                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                   ' Call WriteGuildMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                   ' Call ShowConsoleMsg("Escriba un mensaje.")
                End If
        
            Case "/PMSG"
                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                  '  Call WritePartyMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                  '  Call ShowConsoleMsg("Escriba un mensaje.")
                End If
            
            Case "/CENTINELA"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                       ' Call WriteCentinelReport(CInt(ArgumentosRaw))
                    Else
                        'No es numerico
                      '  Call ShowConsoleMsg("El código de verificación debe ser numerico. Utilice /centinela X, siendo X el código de verificación.")
                    End If
                Else
                    'Avisar que falta el parametro
                   ' Call ShowConsoleMsg("Faltan parámetros. Utilice /centinela X, siendo X el código de verificación.")
                End If
        
            Case "/ONLINECLAN"
               ' Call WriteGuildOnline
                
            Case "/ONLINEPARTY"
               ' Call WritePartyOnline
                
            Case "/BMSG"
                If notNullArguments Then
                  '  Call WriteCouncilMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                  '  Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                
            Case "/ROL"
                If notNullArguments Then
                   ' Call WriteRoleMasterRequest(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                  '  Call ShowConsoleMsg("Escriba una pregunta.")
                End If
                
            Case "/GM"
               ' Call WriteGMRequest
                
            Case "/_BUG"
                If notNullArguments Then
                 '   Call WriteBugReport(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                  '  Call ShowConsoleMsg("Escriba una descripción del bug.")
                End If
            
            Case "/DESC"
                If UserEstado = 1 Then 'Muerto

                    Exit Sub
                End If
                
              '  Call WriteChangeDescription(ArgumentosRaw)
            
            Case "/VOTO"
                If notNullArguments Then
                    'Call WriteGuildVote(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                   ' Call ShowConsoleMsg("Faltan parámetros. Utilice /voto NICKNAME.")
                End If
               
            Case "/PENAS"
                If notNullArguments Then
                 '   Call WritePunishments(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                  '  Call ShowConsoleMsg("Faltan parámetros. Utilice /penas NICKNAME.")
                End If
                
            Case "/CONTRASEÑA"
               ' Call frmNewPassword.Show(vbModal, frmMain)
                
            Case "/RETIRAR"
                If UserEstado = 1 Then 'Muerto
                    
                    Exit Sub
                End If
                If CantidadArgumentos = 0 Then
                    ' Version sin argumentos: LeaveFaction
                  '  Call WriteLeaveFaction
                Else
                    ' Version con argumentos: BankExtractGold
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                    '    Call WriteBankExtractGold(ArgumentosRaw)
                    Else
                        'No es numerico
                     '   Call ShowConsoleMsg("Cantidad incorrecta. Utilice /retirar CANTIDAD.")
                    End If
                End If
    
            Case "/DEPOSITAR"
                If UserEstado = 1 Then 'Muerto

                    Exit Sub
                End If
                
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                     '   Call WriteBankDepositGold(ArgumentosRaw)
                    Else
                        'No es numerico
                      '  Call ShowConsoleMsg("Cantidad incorecta. Utilice /depositar CANTIDAD.")
                    End If
                Else
                    'Avisar que falta el parametro
                    'Call ShowConsoleMsg("Faltan paramtetros. Utilice /depositar CANTIDAD.")
                End If
                
            Case "/DENUNCIAR"
                If notNullArguments Then
                   ' Call WriteDenounce(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                   ' Call ShowConsoleMsg("Formule su denuncia.")
                End If
                
            Case "/FUNDARCLAN"
                If UserLvl >= 25 Then
                '    frmEligeAlineacion.Show vbModeless, frmMain
                Else
                  '  Call ShowConsoleMsg("Para fundar un clan tenés que ser nivel 25 y tener 90 skills en liderazgo.")
                End If
            
            Case "/FUNDARCLANGM"
               ' Call WriteGuildFundate(eClanType.ct_GM)
            
            Case "/ECHARPARTY"
                If notNullArguments Then
                  '  Call WritePartyKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                   ' Call ShowConsoleMsg("Faltan parámetros. Utilice /echarparty NICKNAME.")
                End If

        End Select
        
    ElseIf Left$(Comando, 1) = "\" Then
        If UserEstado = 1 Then 'Muerto

            Exit Sub
        End If
        ' Mensaje Privado
       ' Call AuxWriteWhisper(mid$(Comando, 2), ArgumentosRaw)
        
    ElseIf Left$(Comando, 1) = "-" Then
        If UserEstado = 1 Then 'Muerto

            Exit Sub
        End If
        ' Gritar
        'Call WriteYell(mid$(RawCommand, 2))
        
    Else
        ' Hablar
        Call writeTalk(RawCommand)
    End If
End Sub


Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal red As Integer = 255, Optional ByVal green As Integer = 255, Optional ByVal blue As Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/03/07
'
'***************************************************
    Call AddtoRichTextBox(frmMain.RecTxt, Message, red, green, blue, bold, italic)
End Sub


''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, ByVal TIPO As eNumber_Types) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim Minimo As Long
    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then _
        Exit Function
    
    Select Case TIPO
        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6
    End Select
    
    If val(Numero) >= Minimo And val(Numero) <= Maximo Then _
        ValidNumber = True
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal Ip As String) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(Ip, ".")
    
    If UBound(tmpArr) <> 3 Then _
        Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then _
        Exit Function
    
    validipv4str = True
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal Ip As String) As Byte()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Bytes
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr() As String
    Dim bArr(3) As Byte
    
    tmpArr = Split(Ip, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()
'***************************************************
'Author: Lucas Tavolaro Ortuz (Tavo)
'Useful for AEMAIL BUG FIX
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Strings
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr(0 To 1) As String
    Dim Pos As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    
    AEMAILSplit = tmpArr
End Function


