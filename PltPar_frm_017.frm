VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_CtaBan_1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   2445
   ClientTop       =   1380
   ClientWidth     =   8205
   Icon            =   "PltPar_frm_017.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8205
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   14473
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
         Height          =   795
         Left            =   60
         TabIndex        =   16
         Top             =   780
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   1402
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5940
            Picture         =   "PltPar_frm_017.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_017.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7350
            Picture         =   "PltPar_frm_017.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   4785
         End
         Begin VB.Label Label1 
            Caption         =   "Banco:"
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   270
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   765
         Left            =   60
         TabIndex        =   18
         Top             =   4740
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   7350
            Picture         =   "PltPar_frm_017.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_017.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   5970
            Picture         =   "PltPar_frm_017.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   60
         TabIndex        =   19
         Top             =   7380
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_017.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   7350
            Picture         =   "PltPar_frm_017.frx":17C2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1785
         Left            =   60
         TabIndex        =   20
         Top             =   5550
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   3149
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
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1410
            Width           =   3225
         End
         Begin VB.TextBox txt_CtaCtb 
            Height          =   315
            Left            =   1470
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1080
            Width           =   2505
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   750
            Width           =   3225
         End
         Begin VB.ComboBox cmb_TipCta 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   420
            Width           =   3225
         End
         Begin VB.TextBox txt_NumCta 
            Height          =   315
            Left            =   1470
            MaxLength       =   25
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   90
            Width           =   2505
         End
         Begin VB.Label Label6 
            Caption         =   "Cuenta Vigente:"
            Height          =   285
            Left            =   90
            TabIndex        =   29
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta Contable:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   1110
            Width           =   1305
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Cuenta:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   450
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Moneda:"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   780
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3075
         Left            =   60
         TabIndex        =   24
         Top             =   1620
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
            Left            =   2340
            TabIndex        =   25
            Top             =   60
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Cuenta"
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
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Cuenta:"
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
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   5490
            TabIndex        =   27
            Top             =   60
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   60
         TabIndex        =   30
         Top             =   60
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
            TabIndex        =   31
            Top             =   90
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Cuentas Bancarias"
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
            Picture         =   "PltPar_frm_017.frx":1ACC
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_CtaBan_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CodBan()     As moddat_tpo_Genera
Dim l_arr_TipCta()     As moddat_tpo_Genera

Private Sub cmb_CodBan_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
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

Private Sub cmb_TipCta_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_TipCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCta_Click
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
   Call gs_SetFocus(txt_NumCta)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodIte = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Registro
   g_str_Parame = "USP_BORRAR_MNT_CTABAN(" & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
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
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call fs_Activa(False)
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
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
   moddat_g_str_CodIte = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN WHERE "
   g_str_Parame = g_str_Parame & "CTABAN_CODBAN = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "CTABAN_NUMCTA = '" & moddat_g_str_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   txt_NumCta.Text = Trim(g_rst_Princi!CTABAN_NUMCTA)
   cmb_TipCta.ListIndex = gf_Busca_Arregl(l_arr_TipCta, g_rst_Princi!CTABAN_TIPCTA) - 1
   Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!CTABAN_TIPMON)
   txt_CtaCtb.Text = Trim(g_rst_Princi!CTABAN_CTACTB & "")
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!CTABAN_SITUAC)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Activa_Editar(True)
   
   txt_NumCta.Enabled = False
   Call gs_SetFocus(cmb_TipCta)
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_NumCta.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumCta)
      Exit Sub
   End If
   
   If cmb_TipCta.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCta)
      Exit Sub
   End If
   
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
'   If Len(Trim(txt_CtaCtb.Text)) = 0 Then
'      MsgBox "Debe ingresar la Cuenta Contable.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_CtaCtb)
'      Exit Sub
'   End If

   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar si la Cuenta Bancaria es Vigente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN WHERE "
      g_str_Parame = g_str_Parame & "CTABAN_CODBAN = '" & moddat_g_str_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "CTABAN_NUMCTA = '" & txt_NumCta.Text & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "La Cuenta ya ha sido registrada. Por favor verifique el número e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
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
      
      g_str_Parame = "USP_MNT_CTABAN ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumCta.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_TipCta(cmb_TipCta.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_CtaCtb.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1, "
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
      Call fs_Limpia
      
      Call gs_SetFocus(txt_NumCta)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   cmb_CodBan.ListIndex = -1
   
   Call fs_Limpia
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Editar(False)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_CodBan)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " - Cuentas Bancarias"
   
   Call fs_Inicia
   
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 2285
   grd_Listad.ColWidth(1) = 3155
   grd_Listad.ColWidth(2) = 2105
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   Call moddat_gs_Carga_LisIte(cmb_TipCta, l_arr_TipCta, 1, "510")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "214")
   Call moddat_gs_Carga_TipMon(cmb_TipMon, 1)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_CodBan.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub fs_Activa_Editar(ByVal p_Activa As Integer)
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   txt_NumCta.Enabled = p_Activa
   cmb_TipCta.Enabled = p_Activa
   cmb_TipMon.Enabled = p_Activa
   txt_CtaCtb.Enabled = p_Activa
   cmb_Situac.Enabled = p_Activa
   
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

Private Sub fs_Limpia()
   txt_NumCta.Text = ""
   cmb_TipCta.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   txt_CtaCtb.Text = ""
   cmb_Situac.ListIndex = -1
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_Codigo = l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo
   
   Call gs_LimpiaGrid(grd_Listad)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN WHERE "
   g_str_Parame = g_str_Parame & "CTABAN_CODBAN = '" & moddat_g_str_Codigo & "' "

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
      grd_Listad.Text = Trim(g_rst_Princi!CTABAN_NUMCTA)
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("510", Trim(g_rst_Princi!CTABAN_TIPCTA))
      
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!CTABAN_TIPMON))
      
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

Private Sub txt_CtaCtb_GotFocus()
   Call gs_SelecTodo(txt_CtaCtb)
End Sub

Private Sub txt_CtaCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumCta_GotFocus()
   Call gs_SelecTodo(txt_NumCta)
End Sub

Private Sub txt_NumCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipCta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub
