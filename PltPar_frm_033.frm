VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_MntUsu_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5970
   ClientLeft      =   3735
   ClientTop       =   2550
   ClientWidth     =   7200
   Icon            =   "PltPar_frm_033.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5955
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   10504
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel8 
         Height          =   1095
         Left            =   30
         TabIndex        =   11
         Top             =   3990
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   1931
         _StockProps     =   15
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
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   2985
         End
         Begin VB.ComboBox cmb_Plataf 
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Text            =   "cmb_Client"
            Top             =   60
            Width           =   5625
         End
         Begin VB.ComboBox cmb_TipUsu 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   5625
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Usuario:"
            Height          =   345
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Plataforma:"
            Height          =   345
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1545
         End
         Begin VB.Label Label4 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   720
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   15
         Top             =   3180
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   1349
         _StockProps     =   15
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   5700
            Picture         =   "PltPar_frm_033.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   5010
            Picture         =   "PltPar_frm_033.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   4320
            Picture         =   "PltPar_frm_033.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   6390
            Picture         =   "PltPar_frm_033.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1875
         Left            =   30
         TabIndex        =   16
         Top             =   1260
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   3307
         _StockProps     =   15
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1485
            Left            =   30
            TabIndex        =   0
            Top             =   360
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   2619
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Plataforma"
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   1590
            TabIndex        =   18
            Top             =   60
            Width           =   5205
            _Version        =   65536
            _ExtentX        =   9181
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Usuario"
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   19
         Top             =   5130
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   1349
         _StockProps     =   15
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   6390
            Picture         =   "PltPar_frm_033.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   5700
            Picture         =   "PltPar_frm_033.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   465
         Left            =   30
         TabIndex        =   20
         Top             =   750
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   820
         _StockProps     =   15
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
         Begin Threed.SSPanel pnl_NomUsu 
            Height          =   345
            Left            =   1440
            TabIndex        =   21
            Top             =   60
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "SSPanel10"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label7 
            Caption         =   "Usuario:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   0
         TabIndex        =   23
         Top             =   30
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   1191
         _StockProps     =   15
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   480
            Left            =   630
            TabIndex        =   24
            Top             =   90
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Gestión de Usuarios"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "PltPar_frm_033.frx":14B8
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_MntUsu_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Plataf()      As moddat_tpo_Genera
Dim l_str_Plataf        As String
Dim l_str_TipUsu        As String
Dim l_int_FlgCmb        As Integer

Private Sub cmb_Plataf_Change()
   l_str_Plataf = cmb_Plataf.Text
End Sub

Private Sub cmb_Plataf_Click()
   If cmb_Plataf.ListIndex > -1 Then
      cmb_TipUsu.Clear
      
      If l_int_FlgCmb Then
         Call admusu_gs_Carga_TipUsu(l_arr_Plataf(cmb_Plataf.ListIndex + 1).Genera_Codigo, cmb_TipUsu)
         Call gs_SetFocus(cmb_TipUsu)
      End If
   End If
End Sub

Private Sub cmb_Plataf_GotFocus()
   l_int_FlgCmb = True
   l_str_Plataf = cmb_Plataf.Text
End Sub

Private Sub cmb_Plataf_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Plataf, l_str_Plataf)
      l_int_FlgCmb = True
      
      cmb_TipUsu.Clear
      If cmb_Plataf.ListIndex > -1 Then
         l_str_Plataf = ""
         Call admusu_gs_Carga_TipUsu(l_arr_Plataf(cmb_Plataf.ListIndex + 1).Genera_Codigo, cmb_TipUsu)
      End If
      
      Call gs_SetFocus(cmb_TipUsu)
   End If
End Sub

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmb_TipUsu_Click()
   Call gs_SetFocus(cmb_Situac)
End Sub

Private Sub cmb_TipUsu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipUsu_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(False)
   Call gs_SetFocus(cmb_Plataf)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = "USP_BORRAR_SEG_USUTIP ("
   g_str_Parame = g_str_Parame & "'" & admusu_g_str_CodUsu & "',"
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "',"
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia_Item
   Call fs_Activa(True)
   
   Call gs_SetFocus(grd_Listad)
   
   If grd_Listad.Rows = 0 Then
      grd_Listad.Enabled = False
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      
      Call gs_SetFocus(cmd_Agrega)
   End If
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
     
   moddat_g_int_FlgGrb = 2
   
   'Consulta SQL
   g_str_Parame = "SELECT * FROM SEG_USUTIP WHERE USUTIP_CODUSU = '" & admusu_g_str_CodUsu & "' AND USUTIP_CODPLT = '" & moddat_g_str_CodGrp & "'"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   cmb_Plataf.ListIndex = gf_Busca_Arregl(l_arr_Plataf(), Trim(g_rst_Princi!USUTIP_CODPLT)) - 1
   
   l_str_TipUsu = g_rst_Princi!USUTIP_TIPUSU
   
   Call admusu_gs_Carga_TipUsu(l_arr_Plataf(cmb_Plataf.ListIndex + 1).Genera_Codigo, cmb_TipUsu)
   Call gs_BuscarCombo_Item(cmb_TipUsu, g_rst_Princi!USUTIP_TIPUSU)
   
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!USUTIP_SITUAC)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Activa(False)
   
   cmb_Plataf.Enabled = False
   Call gs_SetFocus(cmb_TipUsu)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_Plataf.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Plataforma.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Plataf)
      Exit Sub
   End If
   
   If cmb_TipUsu.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Usuario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipUsu)
      Exit Sub
   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación del Usuario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      'Consulta SQL
      g_str_Parame = "SELECT * FROM SEG_USUTIP WHERE USUTIP_CODUSU = '" & admusu_g_str_CodUsu & "' AND USUTIP_CODPLT = '" & l_arr_Plataf(cmb_Plataf.ListIndex + 1).Genera_Codigo & "'"
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "El Usuario ya tiene acceso a la Plataforma.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Plataf)
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_SEG_USUTIP ("
      g_str_Parame = g_str_Parame & "'" & admusu_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Plataf(cmb_Plataf.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_TipUsu.ItemData(cmb_TipUsu.ListIndex), "00000") & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   Call fs_Buscar
   Call cmd_Cancel_Click
   
   If moddat_g_int_FlgGrb = 1 Then
      Call cmd_Agrega_Click
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NomUsu.Caption = admusu_g_str_CodUsu & " - " & admusu_g_str_NomUsu

   Call fs_Inicia
   
   Call fs_Limpia
   Call fs_Limpia_Item
   
   Call fs_Activa(True)
   
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM SEG_USUTIP WHERE USUTIP_CODUSU = '" & admusu_g_str_CodUsu & "' ORDER BY USUTIP_CODPLT ASC"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Trim(g_rst_Princi!USUTIP_CODPLT)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Mid(moddat_gf_Consulta_ParDes("351", g_rst_Princi!USUTIP_TIPUSU), 10)
      
      grd_Listad.Col = 2
      grd_Listad.Text = admusu_gf_Consulta_NomPlt(Trim(g_rst_Princi!USUTIP_CODPLT))
      
      grd_Listad.Col = 3
      grd_Listad.Text = g_rst_Princi!USUTIP_TIPUSU
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1530
   grd_Listad.ColWidth(1) = 5190
   grd_Listad.ColWidth(2) = 0
   grd_Listad.ColWidth(3) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter

   Call admusu_gs_Carga_Plataf(cmb_Plataf, l_arr_Plataf(), 0)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
End Sub

Private Sub fs_Limpia_Item()
   cmb_Plataf.ListIndex = -1
   cmb_TipUsu.Clear
   cmb_Situac.ListIndex = -1
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   grd_Listad.Enabled = p_Habilita
   cmd_Agrega.Enabled = p_Habilita
   cmd_Editar.Enabled = p_Habilita
   cmd_Borrar.Enabled = p_Habilita
      
   cmb_Plataf.Enabled = Not p_Habilita
   cmb_TipUsu.Enabled = Not p_Habilita
   cmb_Situac.Enabled = Not p_Habilita
   
   cmd_Grabar.Enabled = Not p_Habilita
   cmd_Cancel.Enabled = Not p_Habilita
End Sub

Private Sub SSPanel13_Click()
   Call gs_SorteaGrid(grd_Listad, 0, "C")
End Sub

Private Sub SSPanel4_Click()
   Call gs_SorteaGrid(grd_Listad, 1, "C")
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

