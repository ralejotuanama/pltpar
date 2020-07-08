VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_MntPlt_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5445
   ClientLeft      =   3180
   ClientTop       =   2430
   ClientWidth     =   7200
   Icon            =   "PltPar_frm_030.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5445
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   9604
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
         Top             =   3480
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
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1710
            MaxLength       =   250
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   390
            Width           =   5325
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1710
            MaxLength       =   8
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   60
            Width           =   1035
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   2985
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre Plataforma:"
            Height          =   345
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Código Plataforma:"
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
         Top             =   2670
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
            Picture         =   "PltPar_frm_030.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   5010
            Picture         =   "PltPar_frm_030.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   4320
            Picture         =   "PltPar_frm_030.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   6390
            Picture         =   "PltPar_frm_030.frx":092A
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
         Top             =   750
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
            Caption         =   "Código Plataforma"
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
            Caption         =   "Nombre Plataforma"
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
         Top             =   4620
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
            Picture         =   "PltPar_frm_030.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   5700
            Picture         =   "PltPar_frm_030.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   20
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   480
            Left            =   630
            TabIndex        =   21
            Top             =   90
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Maestro de Plataformas"
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
            Picture         =   "PltPar_frm_030.frx":14B8
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_MntPlt_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(False)
   Call gs_SetFocus(txt_Codigo)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   
   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = "USP_BORRAR_SEG_PLTMAE ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
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
   g_str_Parame = "SELECT * FROM SEG_PLTMAE WHERE PLTMAE_CODIGO = '" & moddat_g_str_CodGrp & "' "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   txt_Codigo.Text = Trim(g_rst_Princi!PLTMAE_CODIGO)
   txt_Descri.Text = Trim(g_rst_Princi!PLTMAE_DESCRI)
   
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!PLTMAE_SITUAC)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Activa(False)
   
   txt_Codigo.Enabled = False
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "El Código de Plataforma está vacío. Ingrese el Código de Plataforma.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   
   
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "El Nombre de la Plataforma está vacío.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      'Consulta SQL
      g_str_Parame = "SELECT * FROM SEG_PLTMAE WHERE PLTMAE_CODIGO = '" & txt_Codigo.Text & "' "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "El Código de Plataforma ya ha sido registrado.", vbExclamation, modgen_g_str_NomPlt
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
      g_str_Parame = "USP_SEG_PLTMAE ("
      g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
         
      'Datos de Auditoria
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
         If MsgBox("No se pudo completar el procedimiento USP_SEG_PLTMAE. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
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
   
   g_str_Parame = "SELECT * FROM SEG_PLTMAE ORDER BY PLTMAE_CODIGO ASC"
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
      grd_Listad.Text = Trim(g_rst_Princi!PLTMAE_CODIGO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!PLTMAE_DESCRI)
      
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
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter

   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
End Sub

Private Sub fs_Limpia_Item()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   
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
   
   txt_Codigo.Enabled = Not p_Habilita
   txt_Descri.Enabled = Not p_Habilita
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

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_")
   End If
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "(),.-_:; /@#=")
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

