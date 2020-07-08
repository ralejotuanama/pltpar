VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ejecut_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9765
   ClientLeft      =   3585
   ClientTop       =   1845
   ClientWidth     =   10950
   Icon            =   "PltPar_frm_036.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9765
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10965
      _Version        =   65536
      _ExtentX        =   19341
      _ExtentY        =   17224
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   30
         TabIndex        =   18
         Top             =   3630
         Width           =   10875
         _Version        =   65536
         _ExtentX        =   19182
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5430
            Picture         =   "PltPar_frm_036.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Buscar "
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_Situac2 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   240
            Width           =   3285
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   8100
            Picture         =   "PltPar_frm_036.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   8790
            Picture         =   "PltPar_frm_036.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10170
            Picture         =   "PltPar_frm_036.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   9480
            Picture         =   "PltPar_frm_036.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label12 
            Caption         =   "Situación:"
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Top             =   285
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   30
         TabIndex        =   19
         Top             =   8970
         Width           =   10875
         _Version        =   65536
         _ExtentX        =   19182
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   10170
            Picture         =   "PltPar_frm_036.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   9480
            Picture         =   "PltPar_frm_036.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2835
         Left            =   30
         TabIndex        =   20
         Top             =   750
         Width           =   10875
         _Version        =   65536
         _ExtentX        =   19182
         _ExtentY        =   5001
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
            Height          =   2475
            Left            =   60
            TabIndex        =   0
            Top             =   330
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   4366
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
         Begin Threed.SSPanel pnl_Nombre 
            Height          =   285
            Left            =   1530
            TabIndex        =   21
            Top             =   60
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Ejecutivo"
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
         Begin Threed.SSPanel pnl_Situacion 
            Height          =   285
            Left            =   7650
            TabIndex        =   23
            Top             =   60
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   2475
         Left            =   30
         TabIndex        =   24
         Top             =   4410
         Width           =   10875
         _Version        =   65536
         _ExtentX        =   19182
         _ExtentY        =   4366
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
         Begin VB.ComboBox cmb_TipHor 
            Height          =   315
            Left            =   7170
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   2070
            Width           =   3585
         End
         Begin VB.ComboBox cmb_TipPer 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   2070
            Width           =   3285
         End
         Begin VB.ComboBox cmb_UsuSis 
            Height          =   315
            Left            =   7170
            TabIndex        =   12
            Text            =   "cmb_UsuSis"
            Top             =   1410
            Width           =   3585
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1410
            Width           =   3285
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   420
            Width           =   3285
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   90
            Width           =   3285
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1080
            Width           =   3285
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   7170
            MaxLength       =   12
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   1080
            Width           =   3585
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   7170
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   420
            Width           =   3585
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   750
            Width           =   3285
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1740
            Width           =   3285
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo de Horario:"
            Height          =   285
            Left            =   5580
            TabIndex        =   43
            Top             =   2100
            Width           =   1605
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo de Persona:"
            Height          =   285
            Left            =   90
            TabIndex        =   42
            Top             =   2100
            Width           =   1605
         End
         Begin VB.Label Label9 
            Caption         =   "Usuario de Sistemas:"
            Height          =   285
            Left            =   5580
            TabIndex        =   35
            Top             =   1440
            Width           =   1605
         End
         Begin VB.Label Label8 
            Caption         =   "Correo Electrónico:"
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   1440
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   450
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Código Ejecutivo:"
            Height          =   285
            Left            =   90
            TabIndex        =   30
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   5580
            TabIndex        =   29
            Top             =   450
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   780
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de DOI:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   1110
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Número de DOI:"
            Height          =   285
            Left            =   5580
            TabIndex        =   26
            Top             =   1110
            Width           =   1605
         End
         Begin VB.Label Label7 
            Caption         =   "Situación:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   1770
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   32
         Top             =   30
         Width           =   10875
         _Version        =   65536
         _ExtentX        =   19182
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
            TabIndex        =   33
            Top             =   90
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Ejecutivos miCasita"
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
            Picture         =   "PltPar_frm_036.frx":17C2
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1995
         Left            =   30
         TabIndex        =   36
         Top             =   6930
         Width           =   10875
         _Version        =   65536
         _ExtentX        =   19182
         _ExtentY        =   3519
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisEje 
            Height          =   1575
            Left            =   60
            TabIndex        =   14
            Top             =   360
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   2778
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   2130
            TabIndex        =   37
            Top             =   60
            Width           =   7155
            _Version        =   65536
            _ExtentX        =   12621
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Ejecutivo"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   9270
            TabIndex        =   39
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seleccionar"
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
Attribute VB_Name = "frm_Ejecut_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_UsuSis()   As moddat_tpo_Genera
Dim l_str_UsuSis     As String
Dim l_int_FlgCmb     As Integer

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(grd_LisEje)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmb_UsuSis_Change()
   l_str_UsuSis = cmb_UsuSis.Text
End Sub

Private Sub cmb_UsuSis_Click()
   If cmb_UsuSis.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_Situac)
      End If
   End If
End Sub

Private Sub cmb_UsuSis_GotFocus()
   Call SendMessage(cmb_UsuSis.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
   l_str_UsuSis = cmb_UsuSis.Text
End Sub

Private Sub cmb_UsuSis_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_UsuSis, l_str_UsuSis)
      l_int_FlgCmb = True
      If cmb_UsuSis.ListIndex > -1 Then
         l_str_UsuSis = ""
      End If
      
      Call gs_SetFocus(cmb_Situac)
   End If
End Sub

Private Sub cmb_UsuSis_LostFocus()
   Call SendMessage(cmb_UsuSis.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   Call fs_Activa(False)
   Call gs_SetFocus(txt_Codigo)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar la Empresa de Seguro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Borrando de Maestro
   g_str_Parame = "USP_CRE_EJECMC_BORRAR (" & "'" & moddat_g_str_Codigo & "') "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   'Borrando de Tipo de Ejecutivos
   g_str_Parame = "USP_CRE_EJETIP_BORRAR (" & "'" & moddat_g_str_Codigo & "') "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Buscar_Click()
   Call fs_Buscar
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa(True)
   Call fs_Limpia
   Call gs_SetFocus(grd_Listad)

   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
Dim r_int_Contad     As Integer
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_EJECMC WHERE "
   g_str_Parame = g_str_Parame & "EJECMC_CODEJE = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   txt_Codigo.Text = Trim(g_rst_Genera!EJECMC_CODEJE)
   txt_ApePat.Text = Trim(g_rst_Genera!EJECMC_APEPAT)
   txt_ApeMat.Text = Trim(g_rst_Genera!EJECMC_APEMAT)
   txt_Nombre.Text = Trim(g_rst_Genera!EJECMC_NOMBRE)
   Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Genera!EJECMC_TIPDOC)
   txt_NumDoc.Text = Trim(g_rst_Genera!EJECMC_NUMDOC)
   txt_DirEle.Text = Trim(g_rst_Genera!EJECMC_DIRELE & "")
   
   cmb_UsuSis.ListIndex = gf_Busca_Arregl(l_arr_UsuSis(), Trim(g_rst_Genera!EJECMC_CODUSU & "")) - 1
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Genera!EJECMC_SITUAC)
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   'Buscando Tipos de Usuario asignado
   grd_LisEje.Redraw = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_EJETIP WHERE "
   g_str_Parame = g_str_Parame & "EJETIP_CODEJE = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      For r_int_Contad = 0 To grd_LisEje.Rows - 1
         grd_LisEje.Row = r_int_Contad
         grd_LisEje.Col = 0
         If CInt(grd_LisEje.Text) = g_rst_Genera!EJETIP_TIPEJE Then
            grd_LisEje.Col = 2
            grd_LisEje.Text = "X"
            Exit For
         End If
      Next r_int_Contad
   
      g_rst_Genera.MoveNext
   Loop
   
   grd_LisEje.Redraw = True
   Call gs_UbiIniGrid(grd_LisEje)
   Call fs_Activa(False)
   
   txt_Codigo.Enabled = False
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmd_Grabar_Click()
Dim r_int_Contad     As Integer
Dim r_int_NumPue     As Integer
Dim r_str_TipEje     As String
   
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el Código del Ejecutivo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   If Len(Trim(txt_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApeMat)
      Exit Sub
   End If
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de DOI.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de DOI.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
      If Len(Trim(txt_NumDoc.Text)) > 8 Then
         MsgBox "Debe ingresar el Número de DOI correctamente.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
   End If
   If cmb_UsuSis.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Usuario de Sistemas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_UsuSis)
      Exit Sub
   End If
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
  
   r_int_NumPue = 0
   grd_LisEje.Redraw = False
   For r_int_Contad = 0 To grd_LisEje.Rows - 1
      grd_LisEje.Row = r_int_Contad
      grd_LisEje.Col = 2
      
      If grd_LisEje.Text = "X" Then
         r_int_NumPue = r_int_NumPue + 1
      End If
   Next r_int_Contad
   grd_LisEje.Redraw = True
   
   Call gs_UbiIniGrid(grd_LisEje)
   
   If r_int_NumPue = 0 Then
      MsgBox "Debe seleccionar al menos un Puesto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_LisEje)
      Exit Sub
   End If

'   If moddat_g_int_FlgGrb = 1 Then
'      'Validar que el registro no exista
'      g_str_Parame = ""
'      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_EJECMC WHERE "
'      g_str_Parame = g_str_Parame & "EJECMC_CODEJE = '" & txt_Codigo.Text & "' "
'
'      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
'          Exit Sub
'      End If
'
'      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
'         g_rst_Genera.Close
'         Set g_rst_Genera = Nothing
'         MsgBox "El Código ya ha sido registrado. Por favor verifique el código e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
'      End If
'
'      g_rst_Genera.Close
'      Set g_rst_Genera = Nothing
'   End If
      
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando en Maestro de Ejecutivos
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CRE_EJECMC ("
      g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_UsuSis(cmb_UsuSis.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_NumPue) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
      
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
   
   'Borrando de Tipo de Ejecutivos
   g_str_Parame = "USP_CRE_EJETIP_BORRAR (" & "'" & txt_Codigo.Text & "') "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   'Grabando en Tipos de Ejecutivos
   For r_int_Contad = 0 To grd_LisEje.Rows - 1
      grd_LisEje.Row = r_int_Contad
   
      grd_LisEje.Col = 0
      r_str_TipEje = grd_LisEje.Text
      
      grd_LisEje.Col = 2
      
      If grd_LisEje.Text = "X" Then
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            Screen.MousePointer = 11
            
            g_str_Parame = "USP_CRE_EJETIP ("
            g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
            g_str_Parame = g_str_Parame & r_str_TipEje & ", "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
            
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
      End If
   Next r_int_Contad
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Call cmd_Cancel_Click
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub grd_LisEje_DblClick()
   If grd_LisEje.Rows > 0 Then
      grd_LisEje.Col = 2
      If grd_LisEje.Text = "X" Then
         grd_LisEje.Text = ""
      Else
         grd_LisEje.Text = "X"
      End If
      Call gs_RefrescaGrid(grd_LisEje)
   End If
End Sub

Private Sub grd_LisEje_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_LisEje_DblClick
   End If
End Sub

Private Sub grd_LisEje_SelChange()
   If grd_LisEje.Rows > 2 Then
      grd_LisEje.RowSel = grd_LisEje.Row
   End If
End Sub

Private Sub pnl_Codigo_Click()
   If Len(Trim(pnl_Codigo.Tag)) = 0 Or pnl_Codigo.Tag = "D" Then
      pnl_Codigo.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Codigo.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Nombre_Click()
   If Len(Trim(pnl_Nombre.Tag)) = 0 Or pnl_Nombre.Tag = "D" Then
      pnl_Nombre.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Nombre.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Situacion_Click()
   If Len(Trim(pnl_Situacion.Tag)) = 0 Or pnl_Situacion.Tag = "D" Then
      pnl_Situacion.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Situacion.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_UsuSis)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & ".-@_")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 4:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Form_Load()
   Me.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt & " - Mantenimiento de Ejecutivos miCasita"
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac2, 1, "013")
   cmb_Situac2.AddItem "TODOS"
   cmb_Situac2.ItemData(cmb_Situac2.NewIndex) = 0
   cmb_Situac2.ListIndex = 0
    
   Call fs_Inicia
   Call fs_Activa(True)
   Call fs_Limpia
   Call fs_Buscar
      
   Call gs_CentraForm(Me)
   Me.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1440
   grd_Listad.ColWidth(1) = 6120
   grd_Listad.ColWidth(2) = 2910
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter

   grd_LisEje.ColWidth(0) = 2025
   grd_LisEje.ColWidth(1) = 7155
   grd_LisEje.ColWidth(2) = 1275
   grd_LisEje.ColAlignment(0) = flexAlignCenterCenter
   grd_LisEje.ColAlignment(1) = flexAlignLeftCenter
   grd_LisEje.ColAlignment(2) = flexAlignCenterCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   Call moddat_gs_Carga_UsuSis(cmb_UsuSis, l_arr_UsuSis())
   Call fs_Carga_TipEje
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = p_Activa
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   cmb_Situac2.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   txt_Codigo.Enabled = Not p_Activa
   txt_ApePat.Enabled = Not p_Activa
   txt_ApeMat.Enabled = Not p_Activa
   txt_Nombre.Enabled = Not p_Activa
   cmb_TipDoc.Enabled = Not p_Activa
   txt_NumDoc.Enabled = Not p_Activa
   txt_DirEle.Enabled = Not p_Activa
   cmb_UsuSis.Enabled = Not p_Activa
   cmb_Situac.Enabled = Not p_Activa
   cmb_TipPer.Enabled = Not p_Activa
   cmb_TipHor.Enabled = Not p_Activa
   grd_LisEje.Enabled = Not p_Activa
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
Dim r_int_Contad     As Integer
   
   txt_Codigo.Text = ""
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_DirEle.Text = ""
   cmb_UsuSis.ListIndex = -1
   cmb_Situac.ListIndex = -1
   
   grd_LisEje.Redraw = False
   For r_int_Contad = 0 To grd_LisEje.Rows - 1
      grd_LisEje.Row = r_int_Contad
      grd_LisEje.Col = 2
      grd_LisEje.Text = ""
   Next r_int_Contad
   grd_LisEje.Redraw = True
   
   Call gs_UbiIniGrid(grd_LisEje)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApePat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS)
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & ".,- _")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & ".,- _")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & ".,- _")
   End If
End Sub

Private Sub fs_Buscar()
Dim r_str_CodEje     As String
   
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_EJETIP A, CRE_EJECMC B WHERE "
   If CInt(cmb_Situac2.ItemData(cmb_Situac2.ListIndex)) <> 0 Then
      g_str_Parame = g_str_Parame & "A.EJETIP_CODEJE = B.EJECMC_CODEJE And ejecmc_situac =" & CInt(cmb_Situac2.ItemData(cmb_Situac2.ListIndex)) & " "
   Else
      g_str_Parame = g_str_Parame & "A.EJETIP_CODEJE = B.EJECMC_CODEJE "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY B.EJECMC_APEPAT ASC, B.EJECMC_APEMAT ASC, B.EJECMC_NOMBRE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   r_str_CodEje = ""
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!EJECMC_CODEJE <> r_str_CodEje Then
         r_str_CodEje = g_rst_Princi!EJECMC_CODEJE
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Trim(g_rst_Princi!EJECMC_CODEJE)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!EJECMC_APEPAT) & " " & Trim(g_rst_Princi!EJECMC_APEMAT) & " " & Trim(g_rst_Princi!EJECMC_NOMBRE)
         
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Consulta_ParDes("013", CStr(g_rst_Princi!EJECMC_SITUAC))
      End If
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

Private Sub fs_Carga_TipEje()
   Call gs_LimpiaGrid(grd_LisEje)
   
   g_str_Parame = ""
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '034' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_LisEje.Rows = grd_LisEje.Rows + 1
      grd_LisEje.Row = grd_LisEje.Rows - 1
      
      grd_LisEje.Col = 0
      grd_LisEje.Text = CLng(Trim$(g_rst_Genera!PARDES_CODITE))
      
      grd_LisEje.Col = 1
      grd_LisEje.Text = Trim$(g_rst_Genera!PARDES_DESCRI)
      
      grd_LisEje.Col = 2
      grd_LisEje.Text = ""
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   Call gs_UbiIniGrid(grd_LisEje)
End Sub

