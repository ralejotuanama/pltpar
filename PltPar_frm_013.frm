VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_PrdOpe_1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   2625
   ClientTop       =   2130
   ClientWidth     =   8190
   Icon            =   "PltPar_frm_013.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6795
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   11986
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
         Left            =   60
         TabIndex        =   17
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
            Top             =   180
            Width           =   3225
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7380
            Picture         =   "PltPar_frm_013.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   6690
            Picture         =   "PltPar_frm_013.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5970
            Picture         =   "PltPar_frm_013.frx":0758
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
            TabIndex        =   18
            Top             =   210
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   765
         Left            =   60
         TabIndex        =   16
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
            Left            =   5970
            Picture         =   "PltPar_frm_013.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_013.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   7350
            Picture         =   "PltPar_frm_013.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   60
         TabIndex        =   12
         Top             =   5970
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
            Picture         =   "PltPar_frm_013.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_013.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1125
         Left            =   60
         TabIndex        =   13
         Top             =   4800
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   1984
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
         Begin VB.ComboBox cmb_Operac 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   750
            Width           =   6195
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   90
            Width           =   885
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1800
            MaxLength       =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   420
            Width           =   6195
         End
         Begin VB.Label Label2 
            Caption         =   "Código Operación:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Código Acción:"
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción Acción:"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   420
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3075
         Left            =   60
         TabIndex        =   19
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
            Left            =   1500
            TabIndex        =   20
            Top             =   60
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Acción"
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
            TabIndex        =   22
            Top             =   360
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   4736
            _Version        =   393216
            Rows            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_PrdOpe_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Operac()      As moddat_tpo_Genera

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa_Editar(True)
   Call gs_SetFocus(txt_Codigo)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar el grupo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Registro
   g_str_Parame = "USP_BORRAR_CRE_OPERAC (" & "'" & moddat_g_str_CodPrd & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   Call fs_Buscar
   
   Screen.MousePointer = 0
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
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   cmb_Operac.ListIndex = -1
   
   Call fs_Activa_Editar(False)
   Call gs_SetFocus(grd_Listad)
   
   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_OPERAC WHERE "
   g_str_Parame = g_str_Parame & "OPERAC_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "OPERAC_CODACC = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   txt_Codigo.Text = Trim(g_rst_Genera!OPERAC_CODACC)
   txt_Descri.Text = Trim(g_rst_Genera!OPERAC_DESCRI)
   
   cmb_Operac.ListIndex = gf_Busca_Arregl(l_arr_Operac, g_rst_Genera!OPERAC_CODIGO) - 1
      
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa_Editar(True)
   
   txt_Codigo.Enabled = False
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Grabar_Click()
   If moddat_g_int_FlgGrb = 1 Then
      txt_Codigo.Text = Format(txt_Codigo.Text, "000000")
      
      If Len(Trim(txt_Codigo.Text)) < 6 Then
         MsgBox "El Código de Acción es de 6 dígitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción de la Acción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If

   If cmb_Operac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Operac)
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_OPERAC WHERE "
      g_str_Parame = g_str_Parame & "OPERAC_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "OPERAC_CODACC = '" & txt_Codigo.Text & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "El Código ya ha sido registrado. Por favor verifique el código e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
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
      
      g_str_Parame = "USP_CRE_OPERAC ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Operac(cmb_Operac.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
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
      
      txt_Codigo.Text = ""
      txt_Descri.Text = ""
      cmb_Operac.ListIndex = -1
      
      Call gs_SetFocus(txt_Codigo)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Produc.ListIndex = -1
   
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   cmb_Operac.ListIndex = -1
   
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
   
   Me.Caption = modgen_g_str_NomPlt & " - Operaciones Financieras"
   
   Call fs_Inicia
   
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1425
   grd_Listad.ColWidth(1) = 6130
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_Produc_1(cmb_Produc, l_arr_Produc)
   Call moddat_gs_Carga_Operac(l_arr_Operac, cmb_Operac)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Produc.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
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
   txt_Codigo.Enabled = p_Activa
   txt_Descri.Enabled = p_Activa
   cmb_Operac.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_CodPrd = Format(CInt(l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo), "000")
   moddat_g_str_NomPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Nombre
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_OPERAC WHERE "
   g_str_Parame = g_str_Parame & "OPERAC_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "ORDER BY OPERAC_CODACC ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Trim(g_rst_Genera!OPERAC_CODACC)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Genera!OPERAC_DESCRI)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub SSPanel7_Click()
   Call gs_SorteaGrid(grd_Listad, 0, "C")
End Sub

Private Sub SSPanel8_Click()
   Call gs_SorteaGrid(grd_Listad, 1, "C")
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Operac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "()-_=/&%$#@ .?¿*")
   End If
End Sub

