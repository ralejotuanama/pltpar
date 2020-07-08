VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_PrdPad_1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   690
   ClientTop       =   2475
   ClientWidth     =   8175
   Icon            =   "PltPar_frm_014.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7785
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8175
      _Version        =   65536
      _ExtentX        =   14420
      _ExtentY        =   13732
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   16
         Top             =   60
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   210
            Width           =   3225
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7380
            Picture         =   "PltPar_frm_014.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   6690
            Picture         =   "PltPar_frm_014.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5970
            Picture         =   "PltPar_frm_014.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Seleccione Producto:"
            Height          =   285
            Left            =   90
            TabIndex        =   17
            Top             =   240
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   765
         Left            =   30
         TabIndex        =   18
         Top             =   3990
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   6690
            Picture         =   "PltPar_frm_014.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   7380
            Picture         =   "PltPar_frm_014.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   19
         Top             =   6960
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
            Left            =   7350
            Picture         =   "PltPar_frm_014.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_014.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2115
         Left            =   30
         TabIndex        =   20
         Top             =   4800
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   3731
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
         Begin VB.TextBox txt_CtaDes 
            Height          =   315
            Left            =   1800
            MaxLength       =   120
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1740
            Width           =   2655
         End
         Begin VB.TextBox txt_CtaDev 
            Height          =   315
            Left            =   1800
            MaxLength       =   120
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1410
            Width           =   2655
         End
         Begin VB.ComboBox cmb_SitAnt 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   420
            Width           =   6195
         End
         Begin VB.ComboBox cmb_SitAct 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   90
            Width           =   6195
         End
         Begin VB.TextBox txt_CtaCtb 
            Height          =   315
            Left            =   1800
            MaxLength       =   120
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   750
            Width           =   2685
         End
         Begin VB.Label Label7 
            Caption         =   "Cuenta Contable Des.:"
            Height          =   285
            Left            =   90
            TabIndex        =   30
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Cuenta Contable Dev.:"
            Height          =   285
            Left            =   90
            TabIndex        =   29
            Top             =   1410
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Situación Anterior:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta Contable:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Situación Actual:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda:"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   750
            Width           =   1425
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3075
         Left            =   30
         TabIndex        =   24
         Top             =   870
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   5424
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   2520
            TabIndex        =   25
            Top             =   60
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4419
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación Anterior"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   2505
            _Version        =   65536
            _ExtentX        =   4419
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación Actual"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2685
            Left            =   30
            TabIndex        =   4
            Top             =   360
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   4736
            _Version        =   393216
            Rows            =   12
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   5010
            TabIndex        =   27
            Top             =   60
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
   End
End
Attribute VB_Name = "frm_PrdPad_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_SitAct_Click()
   Call gs_SetFocus(cmb_SitAnt)
End Sub

Private Sub cmb_SitAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SitAct_Click
   End If
End Sub

Private Sub cmb_SitAnt_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_SitAnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SitAnt_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(txt_CtaCtb)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa_Editar(True)
   Call gs_SetFocus(cmb_SitAct)
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call fs_Activa(False)
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   cmb_SitAct.ListIndex = -1
   cmb_SitAnt.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   
   txt_CtaCtb.Text = ""
   txt_CtaDev.Text = ""
   txt_CtaDes.Text = ""
   
   Call fs_Activa_Editar(False)
   Call gs_SetFocus(grd_Listad)
   
   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   Dim r_str_SitAct     As String
   Dim r_str_SitAnt     As String
   Dim r_str_Moneda     As String

   grd_Listad.Col = 3
   r_str_SitAct = grd_Listad.Text
         
   grd_Listad.Col = 4
   r_str_SitAnt = grd_Listad.Text
         
   grd_Listad.Col = 5
   r_str_Moneda = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRED_PADRON_DEUDORES WHERE "
   g_str_Parame = g_str_Parame & "PRODUCTO = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "FLAG_ESTADO_ACTUAL = '" & r_str_SitAct & "' AND "
   g_str_Parame = g_str_Parame & "FLAG_ESTADO_ANT = '" & r_str_SitAnt & "' AND "
   g_str_Parame = g_str_Parame & "COD_MONEDA = '" & r_str_Moneda & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Call gs_BuscarCombo_Item(cmb_TipMon, CInt(g_rst_Princi!COD_MONEDA))
   
   Call gs_BuscarCombo_Text(cmb_SitAct, g_rst_Princi!FLAG_ESTADO_ACTUAL, 1)
   Call gs_BuscarCombo_Text(cmb_SitAnt, g_rst_Princi!FLAG_ESTADO_ANT, 1)
   
   txt_CtaCtb.Text = Trim(g_rst_Princi!CNTA_CTBL & "")
   txt_CtaDev.Text = Trim(g_rst_Princi!CNTA_CTBL_DEV & "")
   txt_CtaDes.Text = Trim(g_rst_Princi!CNTA_CTBL_DES & "")
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Activa_Editar(True)
   
   cmb_SitAct.Enabled = False
   cmb_SitAnt.Enabled = False
   cmb_TipMon.Enabled = False
   
   Call gs_SetFocus(txt_CtaCtb)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_SitAct.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SitAct)
      Exit Sub
   End If
   
   If cmb_SitAnt.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación Anterior.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SitAnt)
      Exit Sub
   End If
   
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If Len(Trim(txt_CtaCtb.Text)) = 0 Then
      MsgBox "Debe ingresar la Cuenta Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CtaCtb)
      Exit Sub
   End If

   If Not moddat_gf_CtaCtb(txt_CtaCtb) Then
      MsgBox "La Cuenta Contable no existe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CtaCtb)
      Exit Sub
   End If

   If Len(Trim(txt_CtaDev.Text)) = 0 Then
      MsgBox "Debe ingresar la Cuenta Contable (Devengados).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CtaDev)
      Exit Sub
   End If

   If Not moddat_gf_CtaCtb(txt_CtaDev) Then
      MsgBox "La Cuenta Contable (Devengados) no existe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CtaDev)
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRED_PADRON_DEUDORES WHERE "
      g_str_Parame = g_str_Parame & "PRODUCTO = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "FLAG_ESTADO_ACTUAL = '" & CStr(cmb_SitAct.ItemData(cmb_SitAct.ListIndex)) & "' AND "
      g_str_Parame = g_str_Parame & "FLAG_ESTADO_ANT = '" & CStr(cmb_SitAnt.ItemData(cmb_SitAnt.ListIndex)) & "' AND "
      g_str_Parame = g_str_Parame & "COD_MONEDA = '" & Format(cmb_TipMon.ItemData(cmb_TipMon.ListIndex), "000") & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "La Combinación de Situación para esta Moneda ya ha sido registrada.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CRED_PADRON_DEUDORES ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & Left(cmb_SitAct.Text, 1) & "', "
      g_str_Parame = g_str_Parame & "'" & Left(cmb_SitAnt.Text, 1) & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CtaCtb.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CtaDev.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_TipMon.ItemData(cmb_TipMon.ListIndex), "000") & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CtaDes.Text & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   Screen.MousePointer = 11
   
   If moddat_g_int_FlgGrb = 2 Then
      Call fs_Buscar
      Call cmd_Cancel_Click
   Else
      Call fs_Buscar
      
      Call fs_Activa(False)
      Call fs_Activa_Editar(True)
      
      cmb_SitAct.ListIndex = -1
      cmb_SitAnt.ListIndex = -1
      cmb_TipMon.ListIndex = -1
      
      txt_CtaCtb.Text = ""
      txt_CtaDev.Text = ""
      txt_CtaDes.Text = ""
      
      Call gs_SetFocus(cmb_SitAct)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Produc.ListIndex = -1
   
   cmb_SitAct.ListIndex = -1
   cmb_SitAnt.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   
   txt_CtaCtb.Text = ""
   txt_CtaDev.Text = ""
   txt_CtaDes.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Editar(False)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " - Padrón de Deudores"
   
   Call fs_Inicia
   
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 2490
   grd_Listad.ColWidth(1) = 2490
   grd_Listad.ColWidth(2) = 2670
   grd_Listad.ColWidth(3) = 0
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   Call moddat_gs_Carga_LisIte_Combo(cmb_SitAct, 1, "057")
   Call moddat_gs_Carga_LisIte_Combo(cmb_SitAnt, 1, "057")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Produc.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Activa_Editar(ByVal p_Activa As Integer)
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   cmb_SitAct.Enabled = p_Activa
   cmb_SitAnt.Enabled = p_Activa
   cmb_TipMon.Enabled = p_Activa
   txt_CtaCtb.Enabled = p_Activa
   txt_CtaDev.Enabled = p_Activa
   txt_CtaDes.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_CodPrd = Format(CInt(l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo), "000")
   moddat_g_str_NomPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Nombre
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRED_PADRON_DEUDORES WHERE "
   g_str_Parame = g_str_Parame & "PRODUCTO = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "ORDER BY FLAG_ESTADO_ACTUAL ASC , FLAG_ESTADO_ANT ASC, COD_MONEDA ASC "

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
      grd_Listad.Text = Mid(moddat_gf_Consulta_ParDes_1("057", Trim(g_rst_Princi!FLAG_ESTADO_ACTUAL), 1), 5)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Mid(moddat_gf_Consulta_ParDes_1("057", Trim(g_rst_Princi!FLAG_ESTADO_ANT), 1), 5)
      
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", Trim(g_rst_Princi!COD_MONEDA))
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!FLAG_ESTADO_ACTUAL)
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!FLAG_ESTADO_ANT)
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!COD_MONEDA)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub SSPanel3_Click()
   Call gs_SorteaGrid(grd_Listad, 2, "C")
End Sub

Private Sub SSPanel7_Click()
   Call gs_SorteaGrid(grd_Listad, 0, "C")
End Sub

Private Sub SSPanel8_Click()
   Call gs_SorteaGrid(grd_Listad, 1, "C")
End Sub

Private Sub txt_CtaCtb_GotFocus()
   Call gs_SelecTodo(txt_CtaCtb)
End Sub

Private Sub txt_CtaCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CtaDev)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_CtaDev_GotFocus()
   Call gs_SelecTodo(txt_CtaDev)
End Sub

Private Sub txt_CtaDev_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CtaDes)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_CtaDes_GotFocus()
   Call gs_SelecTodo(txt_CtaDes)
End Sub

Private Sub txt_CtaDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

