VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.MDIForm frm_MnuPri_01 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7245
   ClientLeft      =   1035
   ClientTop       =   1635
   ClientWidth     =   12360
   Icon            =   "PltPar_frm_018.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSMAPI.MAPIMessages mps_Mensaj 
      Left            =   4380
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3810
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   6855
      Width           =   12360
      _Version        =   65536
      _ExtentX        =   21802
      _ExtentY        =   688
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnl_EntDat 
         Height          =   315
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "lm_db_db1 - prod1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_NumVer 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   30
         Width           =   2100
         _Version        =   65536
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "rev. 008-1028.1"
         ForeColor       =   32768
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Font3D          =   2
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12360
      _Version        =   65536
      _ExtentX        =   21802
      _ExtentY        =   1138
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.CommandButton cmd_EnvMail 
         Height          =   585
         Left            =   3120
         Picture         =   "PltPar_frm_018.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Calcul 
         Height          =   585
         Left            =   0
         Picture         =   "PltPar_frm_018.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   1200
         Picture         =   "PltPar_frm_018.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_CamCon 
         Height          =   585
         Left            =   600
         Picture         =   "PltPar_frm_018.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin MSMAPI.MAPISession mps_Sesion 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Menu mnuTab 
      Caption         =   "Tablas Generales"
      Begin VB.Menu mnuTab_Opcion 
         Caption         =   "Parámetros Descriptivos"
         Index           =   1
      End
      Begin VB.Menu mnuTab_Opcion 
         Caption         =   "Parámetros con Valor"
         Index           =   2
      End
   End
   Begin VB.Menu mnuCom 
      Caption         =   "Comercial"
      Begin VB.Menu mnuCom_Opcion 
         Caption         =   "&Zonas Ubicación Inmuebles"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPrd 
      Caption         =   "Producción"
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "Ejecutivos miCasita"
         Index           =   1
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "Sucursales y Agencias miCasita"
         Index           =   2
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "Productos de Créditos Hipotecarios"
         Index           =   4
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "&Mantenimiento de Empresas de Seguro"
         Index           =   6
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "&Tipos de Seguro de Desgravamen"
         Index           =   7
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "Mantenimiento Comisiones Mivivienda"
         Index           =   9
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "Mantenimiento Días Feriados"
         Index           =   10
      End
      Begin VB.Menu mnuPrd_Opcion 
         Caption         =   "Mantenimiento de Tasas"
         Index           =   11
      End
   End
   Begin VB.Menu mnuCre 
      Caption         =   "Créditos y Riesgos"
      Begin VB.Menu mnuCre_Opcion 
         Caption         =   "Giros Comerciales"
         Index           =   1
      End
      Begin VB.Menu mnuCre_Opcion 
         Caption         =   "Personas Vinculadas"
         Index           =   2
      End
   End
   Begin VB.Menu mnuOpe 
      Caption         =   "Operaciones"
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "Cuentas Bancarias"
         Index           =   1
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuOpe_Opcion 
         Caption         =   "&Tipo de Cambio"
         Index           =   3
      End
   End
   Begin VB.Menu mnuCtb 
      Caption         =   "Contabilidad"
      Begin VB.Menu mnuCtb_Opcion 
         Caption         =   "&Operaciones Contables"
         Index           =   1
      End
      Begin VB.Menu mnuCtb_Opcion 
         Caption         =   "Operaciones Financieras"
         Index           =   2
      End
      Begin VB.Menu mnuCtb_Opcion 
         Caption         =   "&Créditos"
         Index           =   3
         Begin VB.Menu mnuCtb_Credit_Clasif 
            Caption         =   "Clasficación de Créditos"
         End
         Begin VB.Menu mnuCtb_Credit_PadDeu 
            Caption         =   "Cuentas Contables Padrón Deudores"
         End
         Begin VB.Menu mnuCtb_Credit_CtaRcd 
            Caption         =   "Cuentas Contables RCD"
         End
      End
      Begin VB.Menu mnuCtb_Opcion 
         Caption         =   "&Garantías"
         Index           =   4
         Begin VB.Menu mnuCtb_Garant_TipGar 
            Caption         =   "Tipos y Atributos de Garantías"
         End
      End
   End
End
Attribute VB_Name = "frm_MnuPri_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Calcul_Click()
Dim r_lng_NumPid    As Long
   
   r_lng_NumPid = Shell("c:\WINDOWS\system32\calc.exe", vbNormalFocus)
   If r_lng_NumPid = 0 Then
      MsgBox "Error Iniciando la Aplicación", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_CamCon_Click()
   If modgen_g_str_CodUsu <> "DESARROLLO" Then
      frm_IdeUsu_02.Show 1
   End If
End Sub

Private Sub cmd_EnvMail_Click()
Dim r_str_Mensaj     As String
Dim r_str_Asunto     As String

   If MsgBox("¿Está seguro de enviar el correo ? ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   ReDim moddat_g_arr_Genera(0)

   r_str_Asunto = "TEST DE CORREO MICASITA (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   r_str_Mensaj = ""
   r_str_Mensaj = r_str_Mensaj & "NUMERO DE OPERACION : 000000001" & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "ID CLIENTE          : 1-020202020" & Chr(13)
   r_str_Mensaj = r_str_Mensaj & "NOMBRE CLIENTE      : NOMBRE CLIENTE TEST" & Chr(13)
   r_str_Mensaj = r_str_Mensaj & Chr(13)
   
   'JEFE TECNOLOGIA INFORMACION
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(810, moddat_g_arr_Genera)
   
   'DESARROLLO SISTEMAS (811)
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = Trim("JPAMPA@MICASITA.COM.PE")
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Nombre = Trim("JPAMPA")

   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, r_str_Asunto, r_str_Mensaj)
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub MDIForm_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   'pnl_Seg_NomUsu.Caption = modgen_g_str_CodUsu
   pnl_NumVer.Caption = modgen_g_str_NumRev
   pnl_EntDat.Caption = moddat_g_str_NomEsq & " - " & UCase(moddat_g_str_EntDat)
   
   'Activando por Perfiles
   'Call fs_HabSeg
   Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   Else
      Cancel = True
   End If
End Sub

Private Sub fs_HabSeg()
Dim r_int_Posici     As Integer
Dim r_str_CodMen     As String
   
   'Desactivando todas las opciones
   For r_int_Posici = 1 To mnuTab_Opcion.Count
      If mnuTab_Opcion(r_int_Posici).Caption <> "-" Then
         mnuTab_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuCom_Opcion.Count
      If mnuCom_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCom_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuPrd_Opcion.Count
      If mnuPrd_Opcion(r_int_Posici).Caption <> "-" Then
         mnuPrd_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuCre_Opcion.Count
      If mnuCre_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCre_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuOpe_Opcion.Count
      If mnuOpe_Opcion(r_int_Posici).Caption <> "-" Then
         mnuOpe_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuCtb_Opcion.Count
      If mnuCtb_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCtb_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   g_str_Parame = "SELECT PLTOPC_CODMEN, PLTOPC_CODSUB, PLTOPC_SITUAC "
   g_str_Parame = g_str_Parame & "FROM SEG_PLTOPC "
   g_str_Parame = g_str_Parame & "WHERE PLTOPC_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "AND PLTOPC_FLGMEN = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY PLTOPC_CODMEN ASC, PLTOPC_CODSUB ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      Select Case Trim(g_rst_Princi!PLTOPC_CODMEN)
         Case "MNUTAB": mnuTab_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
         Case "MNUCOM": mnuCom_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
         Case "MNUPRD": mnuPrd_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
         Case "MNUCRE": mnuCre_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
         Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
         Case "MNUCTB": mnuCtb_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
      End Select
      g_rst_Princi.MoveNext
   Loop
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
 
   'Verificando por Plantilla de Acceso
   g_str_Parame = "SELECT PLTPLA_CODMEN, PLTPLA_CODSUB "
   g_str_Parame = g_str_Parame & "FROM SEG_PLTPLA "
   g_str_Parame = g_str_Parame & "WHERE PLTPLA_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "AND PLTPLA_TIPUSU = '" & CStr(modgen_g_int_TipUsu) & "' "
   g_str_Parame = g_str_Parame & "ORDER BY PLTPLA_CODMEN ASC, PLTPLA_CODSUB ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      Select Case Trim(g_rst_Princi!PLTPLA_CODMEN)
         Case "MNUTAB": mnuTab_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
         Case "MNUCOM": mnuCom_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
         Case "MNUPRD": mnuPrd_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
         Case "MNUCRE": mnuCre_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
         Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
         Case "MNUCTB": mnuCtb_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
      End Select
      g_rst_Princi.MoveNext
   Loop
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Verificando por Personalización de Opciones
   g_str_Parame = "SELECT * FROM SEG_PLTUSU WHERE "
   g_str_Parame = g_str_Parame & "PLTUSU_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "AND PLTUSU_CODUSU = '" & CStr(modgen_g_int_TipUsu) & "' "
   g_str_Parame = g_str_Parame & "ORDER BY PLTUSU_CODMEN ASC, PLTUSU_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
            Select Case r_str_CodMen
               Case "MNUTAB": mnuTab_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCOM": mnuCom_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUPRD": mnuPrd_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCRE": mnuCre_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUOPE": mnuOpe_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCTB": mnuCtb_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
            End Select
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub mnuTab_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_ParDes_01.Show 1
      Case 2: frm_ParVal_01.Show 1
   End Select
End Sub

Private Sub mnuCom_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_PosUbi_01.Show 1
   End Select
End Sub

Private Sub mnuPrd_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_Ejecut_01.Show 1
      Case 4: frm_Produc_01.Show 1
      Case 6: frm_Seguro_01.Show 1
      Case 7: frm_Seguro_02.Show 1
      Case 9: frm_Comviv_1.Show 1
      Case 10: frm_Feriad_01.Show 1
      Case 11: frm_Produc_10.Show 1
   End Select
End Sub

Private Sub mnuCre_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_GirCom_01.Show 1
      Case 2: frm_PerVin_01.Show 1
   End Select
End Sub

Private Sub mnuOpe_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_CtaBan_1.Show 1
      Case 3: frm_TipCam_1.Show 1
   End Select
End Sub

Private Sub mnuCtb_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1: frm_OpeFin_1.Show 1
      Case 2: frm_PrdOpe_1.Show 1
   End Select
End Sub

Private Sub mnuCtb_Credit_Clasif_Click()
   frm_CalCre_1.Show 1
End Sub

Private Sub mnuCtb_Credit_PadDeu_Click()
   frm_PrdPad_1.Show 1
End Sub

Private Sub mnuCtb_Credit_CtaRcd_Click()
   frm_PrdRcd_1.Show 1
End Sub

Private Sub mnuCtb_Garant_TipGar_Click()
   frm_TipGar_01.Show 1
End Sub

