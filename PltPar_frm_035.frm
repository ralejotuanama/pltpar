VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_TipGar_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6270
   ClientLeft      =   1080
   ClientTop       =   1770
   ClientWidth     =   13260
   Icon            =   "PltPar_frm_035.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6285
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13275
      _Version        =   65536
      _ExtentX        =   23416
      _ExtentY        =   11086
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
         Height          =   1425
         Left            =   30
         TabIndex        =   12
         Top             =   3990
         Width           =   13185
         _Version        =   65536
         _ExtentX        =   23257
         _ExtentY        =   2514
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
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   5445
         End
         Begin VB.TextBox txt_CtaCtb 
            Height          =   315
            Left            =   1590
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1050
            Width           =   2205
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   720
            Width           =   11505
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   420
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta Contable:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   1080
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Código Atributo:"
            Height          =   345
            Left            =   60
            TabIndex        =   14
            Top             =   90
            Width           =   1545
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción:"
            Height          =   345
            Left            =   60
            TabIndex        =   13
            Top             =   750
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   30
         TabIndex        =   16
         Top             =   3210
         Width           =   13185
         _Version        =   65536
         _ExtentX        =   23257
         _ExtentY        =   1296
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12480
            Picture         =   "PltPar_frm_035.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   10410
            Picture         =   "PltPar_frm_035.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   11100
            Picture         =   "PltPar_frm_035.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Modificar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   11790
            Picture         =   "PltPar_frm_035.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1875
         Left            =   30
         TabIndex        =   17
         Top             =   1290
         Width           =   13185
         _Version        =   65536
         _ExtentX        =   23257
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
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   2619
            _Version        =   393216
            Rows            =   12
            Cols            =   5
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
            TabIndex        =   18
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Atributo"
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
            TabIndex        =   19
            Top             =   60
            Width           =   5205
            _Version        =   65536
            _ExtentX        =   9181
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   6780
            TabIndex        =   27
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta Contable"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   9510
            TabIndex        =   28
            Top             =   60
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   20
         Top             =   5460
         Width           =   13185
         _Version        =   65536
         _ExtentX        =   23257
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11760
            Picture         =   "PltPar_frm_035.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12450
            Picture         =   "PltPar_frm_035.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   465
         Left            =   30
         TabIndex        =   21
         Top             =   780
         Width           =   13185
         _Version        =   65536
         _ExtentX        =   23257
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
            Left            =   1590
            TabIndex        =   22
            Top             =   60
            Width           =   11535
            _Version        =   65536
            _ExtentX        =   20346
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "SSPanel10"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Caption         =   "Tipo de Garantía:"
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Top             =   90
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   24
         Top             =   60
         Width           =   13185
         _Version        =   65536
         _ExtentX        =   23257
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
            TabIndex        =   25
            Top             =   90
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Tipos de Garantía - Atributos de Garantía"
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
            Picture         =   "PltPar_frm_035.frx":14B8
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_TipGar_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Plataf()      As moddat_tpo_Genera
Dim l_str_Plataf        As String
Dim l_str_TipUsu        As String
Dim l_int_FlgCmb        As Integer

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(False)
   Call gs_SetFocus(txt_Codigo)
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_str_TipMon     As String
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   grd_Listad.Col = 4
   r_str_TipMon = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = "USP_BORRAR_GARANTIA_ATRIBUTOS ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "',"
   g_str_Parame = g_str_Parame & "'" & r_str_TipMon & "')"
   
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
   Dim r_str_TipMon     As String
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   grd_Listad.Col = 4
   r_str_TipMon = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
     
   moddat_g_int_FlgGrb = 2
   
   'Consulta SQL
   g_str_Parame = "SELECT * FROM GARANTIA_ATRIBUTOS WHERE GARANTIA_ATRIB = '" & moddat_g_str_Codigo & "' AND COD_MONEDA = '" & r_str_TipMon & "'"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   txt_Codigo.Text = Mid(g_rst_Princi!GARANTIA_ATRIB, 3, 4)
   Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!COD_MONEDA)
   
   txt_Descri.Text = Trim(g_rst_Princi!DESCRIPCION)
   txt_CtaCtb.Text = Trim(g_rst_Princi!CNTA_CTBL)
   
   Call fs_Activa(False)
   
   txt_Codigo.Enabled = False
   cmb_TipMon.Enabled = False
   
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_CodAtr     As String
   
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el Código de Atributo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If

   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
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
   
   r_str_CodAtr = moddat_g_str_CodGrp & txt_Codigo.Text
   
   If moddat_g_int_FlgGrb = 1 Then
      'Consulta SQL
      g_str_Parame = "SELECT * FROM GARANTIA_ATRIBUTOS WHERE GARANTIA_ATRIB = '" & r_str_CodAtr & "' AND COD_MONEDA = '" & Format(cmb_TipMon.ItemData(cmb_TipMon.ListIndex), "000") & "' "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "Ya está registrado este Atributo para esta Moneda.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
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
      g_str_Parame = "USP_GARANTIA_ATRIBUTOS ("
      g_str_Parame = g_str_Parame & "'" & r_str_CodAtr & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CtaCtb.Text & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_TipMon.ItemData(cmb_TipMon.ListIndex), "000") & "', "
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
   
   pnl_NomUsu.Caption = moddat_g_str_CodGrp & " - " & moddat_g_str_DesGrp

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
   
   g_str_Parame = "SELECT * FROM GARANTIA_ATRIBUTOS WHERE GARANTIA_TIPO = '" & moddat_g_str_CodGrp & "' ORDER BY GARANTIA_ATRIB ASC, COD_MONEDA ASC"
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
      grd_Listad.Text = Trim(g_rst_Princi!GARANTIA_ATRIB)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!DESCRIPCION)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Princi!CNTA_CTBL)
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!COD_MONEDA)
      
      grd_Listad.Col = 4
      grd_Listad.Text = g_rst_Princi!COD_MONEDA
      
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
   grd_Listad.ColWidth(2) = 2715
   grd_Listad.ColWidth(3) = 3285
   grd_Listad.ColWidth(4) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter

   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
End Sub

Private Sub fs_Limpia_Item()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   txt_CtaCtb.Text = ""
   cmb_TipMon.ListIndex = -1
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   grd_Listad.Enabled = p_Habilita
   cmd_Agrega.Enabled = p_Habilita
   cmd_Editar.Enabled = p_Habilita
   cmd_Borrar.Enabled = p_Habilita
      
   txt_Codigo.Enabled = Not p_Habilita
   txt_Descri.Enabled = Not p_Habilita
   txt_CtaCtb.Enabled = Not p_Habilita
   cmb_TipMon.Enabled = Not p_Habilita
   
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

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_CtaCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CtaCtb)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "(),.-_:; /@#=")
   End If
End Sub

