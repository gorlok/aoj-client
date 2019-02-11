VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7695
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdNOREAL 
         Caption         =   "/NOREAL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   86
         Top             =   6480
         Width           =   1815
      End
      Begin VB.CommandButton cmdNOCAOS 
         Caption         =   "/NOCAOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   85
         Top             =   6480
         Width           =   1815
      End
      Begin VB.CommandButton cmdKICKCONSE 
         Caption         =   "/KICKCONSE"
         CausesValidation=   0   'False
         Height          =   675
         Left            =   2520
         TabIndex        =   84
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdACEPTCONSECAOS 
         Caption         =   "/ACEPTCONSECAOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   83
         Top             =   7320
         Width           =   2295
      End
      Begin VB.CommandButton cmdACEPTCONSE 
         Caption         =   "/ACEPTCONSE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   82
         Top             =   6960
         Width           =   2295
      End
      Begin VB.ComboBox cboListaUsus 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   54
         Top             =   480
         Width           =   3675
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   3675
      End
      Begin VB.CommandButton cmdIRCERCA 
         Caption         =   "/IRCERCA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDONDE 
         Caption         =   "/DONDE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdPENAS 
         Caption         =   "/PENAS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdTELEP 
         Caption         =   "/TELEP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   49
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdSILENCIAR 
         Caption         =   "/SILENCIAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   48
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdIRA 
         Caption         =   "/IRA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   47
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCARCEL 
         Caption         =   "/CARCEL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdADVERTENCIA 
         Caption         =   "/ADVERTENCIA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   45
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdINFO 
         Caption         =   "/INFO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSTAT 
         Caption         =   "/STAT"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   43
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAL 
         Caption         =   "/BAL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   42
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdINV 
         Caption         =   "/INV"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdBOV 
         Caption         =   "/BOV"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   40
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdSKILLS 
         Caption         =   "/SKILLS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdREVIVIR 
         Caption         =   "/REVIVIR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdPERDON 
         Caption         =   "/PERDON"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   37
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdECHAR 
         Caption         =   "/ECHAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdEJECUTAR 
         Caption         =   "/EJECUTAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   35
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAN 
         Caption         =   "/BAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdUNBAN 
         Caption         =   "/UNBAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   33
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdSUM 
         Caption         =   "/SUM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdNICK2IP 
         Caption         =   "/NICK2IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdESTUPIDO 
         Caption         =   "/ESTUPIDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdNOESTUPIDO 
         Caption         =   "/NOESTUPIDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdBORRARPENA 
         Caption         =   "/BORRARPENA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   28
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTIP 
         Caption         =   "/LASTIP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCONDEN 
         Caption         =   "/CONDEN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRAJAR 
         Caption         =   "/RAJAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   25
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRAJARCLAN 
         Caption         =   "/RAJARCLAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   24
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTEMAIL 
         Caption         =   "/LASTEMAIL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   23
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   4215
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdGMSG 
         Caption         =   "/GMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdHORA 
         Caption         =   "/HORA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdRMSG 
         Caption         =   "/RMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdREALMSG 
         Caption         =   "/REALMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCAOSMSG 
         Caption         =   "/CAOSMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCIUMSG 
         Caption         =   "/CIUMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdTALKAS 
         Caption         =   "/TALKAS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdMOTDCAMBIA 
         Caption         =   "/MOTDCAMBIA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSMSG 
         Caption         =   "/SMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   5
      Left            =   120
      TabIndex        =   55
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdCC 
         Caption         =   "/CC"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   72
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdLIMPIAR 
         Caption         =   "/LIMPIAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   71
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCT 
         Caption         =   "/CT"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   70
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdDT 
         Caption         =   "/DT"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   69
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdLLUVIA 
         Caption         =   "/LLUVIA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   68
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdMASSDEST 
         Caption         =   "/MASSDEST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   67
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdPISO 
         Caption         =   "/PISO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   66
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCI 
         Caption         =   "/CI"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   65
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdDEST 
         Caption         =   "/DEST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   64
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdONLINEREAL 
         Caption         =   "/ONLINEREAL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINECAOS 
         Caption         =   "/ONLINECAOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdNENE 
         Caption         =   "/NENE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSHOW_SOS 
         Caption         =   "/SHOW SOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdTRABAJANDO 
         Caption         =   "/TRABAJANDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdOCULTANDO 
         Caption         =   "/OCULTANDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINEGM 
         Caption         =   "/ONLINEGM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   15
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CommandButton cmdONLINEMAP 
         Caption         =   "/ONLINEMAP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdBORRAR_SOS 
         Caption         =   "/BORRAR SOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   6
      Left            =   120
      TabIndex        =   56
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSHOWCMSG 
         Caption         =   "/SHOWCMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   80
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdBANCLAN 
         Caption         =   "/BANCLAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   79
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdMIEMBROSCLAN 
         Caption         =   "/MIEMBROSCLAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   78
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdBANIPRELOAD 
         Caption         =   "/BANIPRELOAD"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   77
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdBANIPLIST 
         Caption         =   "/BANIPLIST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   76
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdIP2NICK 
         Caption         =   "/IP2NICK"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   75
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdBANIP 
         Caption         =   "/BANIP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   74
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdUNBANIP 
         Caption         =   "/UNBANIP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   73
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSHOWNAME 
         Caption         =   "/SHOWNAME"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   63
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdREM 
         Caption         =   "/REM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   62
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton cmdINVISIBLE 
         Caption         =   "/INVISIBLE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   61
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSETDESC 
         Caption         =   "/SETDESC"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   60
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdNAVE 
         Caption         =   "/NAVE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   59
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCHATCOLOR 
         Caption         =   "/CHATCOLOR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   58
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdIGNORADO 
         Caption         =   "/IGNORADO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   57
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox TabStrip 
      CausesValidation=   0   'False
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   4155
      TabIndex        =   81
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
