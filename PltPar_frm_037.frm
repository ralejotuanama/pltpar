VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_PerVin_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9870
   ClientLeft      =   1725
   ClientTop       =   660
   ClientWidth     =   12720
   Icon            =   "PltPar_frm_037.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9855
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12705
      _Version        =   65536
      _ExtentX        =   22410
      _ExtentY        =   17383
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
         Height          =   765
         Left            =   30
         TabIndex        =   16
         Top             =   5070
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
            Left            =   11160
            Picture         =   "PltPar_frm_037.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   11880
            Picture         =   "PltPar_frm_037.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   10470
            Picture         =   "PltPar_frm_037.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   9780
            Picture         =   "PltPar_frm_037.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   60
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   17
         Top             =   9030
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
            Left            =   11220
            Picture         =   "PltPar_frm_037.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   11910
            Picture         =   "PltPar_frm_037.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4275
         Left            =   30
         TabIndex        =   18
         Top             =   750
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
         _ExtentY        =   7541
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
            Height          =   3915
            Left            =   60
            TabIndex        =   0
            Top             =   330
            Width           =   12525
            _ExtentX        =   22093
            _ExtentY        =   6906
            _Version        =   393216
            Rows            =   20
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   2130
            TabIndex        =   19
            Top             =   60
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Docum. de Identidad"
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   6030
            TabIndex        =   21
            Top             =   60
            Width           =   4365
            _Version        =   65536
            _ExtentX        =   7699
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Persona Vinculada"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   10380
            TabIndex        =   31
            Top             =   60
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
         Height          =   3105
         Left            =   30
         TabIndex        =   22
         Top             =   5880
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
         _ExtentY        =   5477
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
         Begin VB.CommandButton cmd_BusVin 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11760
            TabIndex        =   42
            Top             =   1920
            Width           =   405
         End
         Begin VB.ComboBox cmb_TDoVin 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1890
            Width           =   3285
         End
         Begin VB.TextBox txt_NDoVin 
            Height          =   315
            Left            =   9240
            MaxLength       =   12
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   1890
            Width           =   2415
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   75
            Left            =   60
            TabIndex        =   34
            Top             =   1740
            Width           =   12525
            _Version        =   65536
            _ExtentX        =   22093
            _ExtentY        =   132
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.23
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin VB.ComboBox cmb_RelLab 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1380
            Width           =   10545
         End
         Begin VB.ComboBox cmb_FlgAcc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1050
            Width           =   10545
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2730
            Width           =   3285
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   720
            Width           =   3285
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   9240
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   390
            Width           =   3285
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   9240
            MaxLength       =   12
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   60
            Width           =   2415
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   60
            Width           =   3285
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   390
            Width           =   3285
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   75
            Left            =   60
            TabIndex        =   39
            Top             =   2610
            Width           =   12525
            _Version        =   65536
            _ExtentX        =   22093
            _ExtentY        =   132
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.23
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnl_PerVin 
            Height          =   315
            Left            =   2010
            TabIndex        =   41
            Top             =   2220
            Width           =   10545
            _Version        =   65536
            _ExtentX        =   18600
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel3"
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
         Begin VB.Label Label12 
            Caption         =   "Número DOI Vinculado:"
            Height          =   285
            Left            =   7320
            TabIndex        =   40
            Top             =   1890
            Width           =   1845
         End
         Begin VB.Label Label11 
            Caption         =   "Persona Vinculada:"
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   2220
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de DOI Vinculado:"
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   1890
            Width           =   1665
         End
         Begin VB.Label Label8 
            Caption         =   "Relación Laboral:"
            Height          =   285
            Left            =   90
            TabIndex        =   33
            Top             =   1380
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "Accionista:"
            Height          =   285
            Left            =   90
            TabIndex        =   32
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label Label7 
            Caption         =   "Situación:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   2730
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Número de DOI:"
            Height          =   285
            Left            =   7320
            TabIndex        =   27
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de DOI:"
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   720
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   7320
            TabIndex        =   24
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   390
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   29
         Top             =   30
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
            TabIndex        =   30
            Top             =   90
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Personas Vinculadas"
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
            Picture         =   "PltPar_frm_037.frx":14B8
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_PerVin_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_FlgAcc_Click()
   If cmb_FlgAcc.ListIndex > -1 Then
      If Left(cmb_FlgAcc.Text, 1) = "2" Then
         cmb_TDoVin.Enabled = True
         txt_NDoVin.Enabled = True
         cmd_BusVin.Enabled = True
         
         Call gs_SetFocus(cmb_TDoVin)
      Else
         cmb_TDoVin.Enabled = False
         txt_NDoVin.Enabled = False
         cmd_BusVin.Enabled = False
         
         cmb_TDoVin.ListIndex = -1
         txt_NDoVin.Text = ""
         pnl_PerVin.Caption = ""
         
         Call gs_SetFocus(cmb_RelLab)
      End If
   End If
End Sub

Private Sub cmb_FlgAcc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgAcc_Click
   End If
End Sub

Private Sub cmb_RelLab_Click()
   If cmb_RelLab.ListIndex > -1 Then
      If Left(cmb_RelLab.Text, 1) = "5" Then
         cmb_TDoVin.Enabled = True
         txt_NDoVin.Enabled = True
         cmd_BusVin.Enabled = True
         
         Call gs_SetFocus(cmb_TDoVin)
      Else
         cmb_TDoVin.Enabled = False
         txt_NDoVin.Enabled = False
         cmd_BusVin.Enabled = False
         
         cmb_TDoVin.ListIndex = -1
         txt_NDoVin.Text = ""
         pnl_PerVin.Caption = ""
         
         Call gs_SetFocus(cmb_Situac)
      End If
   End If
End Sub

Private Sub cmb_RelLab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RelLab_Click
   End If
End Sub

Private Sub cmb_TDoVin_Click()
   If cmb_TDoVin.ListIndex > -1 Then
      Select Case cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex)
         Case 1:  txt_NDoVin.MaxLength = 8
         Case 2:  txt_NDoVin.MaxLength = 12
         Case 3:  txt_NDoVin.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NDoVin)
End Sub

Private Sub cmb_TDoVin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TDoVin_Click
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

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(False)
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_int_TDoVin     As Integer
   Dim r_str_NDoVin     As String

   grd_Listad.Col = 0
   
   moddat_g_int_TipDoc = CInt(Left(grd_Listad.Text, 1))
   moddat_g_str_NumDoc = Trim(Mid(grd_Listad.Text, 4))
         
   If Len(Trim(grd_Listad.Text)) > 0 Then
      grd_Listad.Col = 4
      r_int_TDoVin = CInt(grd_Listad.Text)
            
      grd_Listad.Col = 5
      r_str_NDoVin = grd_Listad.Text
   End If
   
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar la Empresa de Seguro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If r_int_TDoVin = 0 Then
      g_str_Parame = "USP_CRE_PERVIN_BORRAR ("
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "') "
   Else
      g_str_Parame = "USP_CRE_PERVIN_BORRAR_VINCUL ("
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_TDoVin) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_NDoVin & "') "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_BusVin_Click()
   pnl_PerVin.Caption = ""
   
   Screen.MousePointer = 11

   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
   g_str_Parame = g_str_Parame & "PERVIN_TDOTIT = " & CStr(cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "PERVIN_NDOTIT = '" & txt_NDoVin.Text & "' AND "
   g_str_Parame = g_str_Parame & "PERVIN_TDOVIN = 0"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Screen.MousePointer = 0
      
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      pnl_PerVin.Caption = Trim(g_rst_Genera!PERVIN_APPTIT) & " " & Trim(g_rst_Genera!PERVIN_APMTIT) & " " & Trim(g_rst_Genera!PERVIN_NOMTIT)
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Screen.MousePointer = 0
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
   Dim r_int_TDoVin     As Integer
   Dim r_str_NDoVin     As String

   grd_Listad.Col = 0
   
   moddat_g_int_TipDoc = CInt(Left(grd_Listad.Text, 1))
   moddat_g_str_NumDoc = Trim(Mid(grd_Listad.Text, 4))
         
   If Len(Trim(grd_Listad.Text)) > 0 Then
      grd_Listad.Col = 4
      r_int_TDoVin = CInt(grd_Listad.Text)
            
      grd_Listad.Col = 5
      r_str_NDoVin = grd_Listad.Text
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   
   If r_int_TDoVin = 0 Then
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
      g_str_Parame = g_str_Parame & "PERVIN_TDOTIT = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "PERVIN_NDOTIT = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "PERVIN_TDOVIN = 0"
   Else
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
      g_str_Parame = g_str_Parame & "PERVIN_TDOTIT = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "PERVIN_NDOTIT = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "PERVIN_TDOVIN = " & CStr(r_int_TDoVin) & " AND "
      g_str_Parame = g_str_Parame & "PERVIN_NDOVIN = '" & r_str_NDoVin & "' "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   Call gs_BuscarCombo_Text(cmb_FlgAcc, g_rst_Genera!PERVIN_FLGACC, 1)
   Call gs_BuscarCombo_Text(cmb_RelLab, g_rst_Genera!PERVIN_RELLAB, 1)
   
   If r_int_TDoVin = 0 Then
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Genera!PERVIN_TDOTIT)
      txt_NumDoc.Text = Trim(g_rst_Genera!PERVIN_NDOTIT)
      
      txt_ApePat.Text = Trim(g_rst_Genera!PERVIN_APPTIT)
      txt_ApeMat.Text = Trim(g_rst_Genera!PERVIN_APMTIT)
      txt_Nombre.Text = Trim(g_rst_Genera!PERVIN_NOMTIT)
   Else
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Genera!PERVIN_TDOVIN)
      txt_NumDoc.Text = Trim(g_rst_Genera!PERVIN_NDOVIN)
   
      Call gs_BuscarCombo_Item(cmb_TDoVin, g_rst_Genera!PERVIN_TDOTIT)
      txt_NDoVin.Text = Trim(g_rst_Genera!PERVIN_NDOTIT)
      
      txt_ApePat.Text = Trim(g_rst_Genera!PERVIN_APPVIN)
      txt_ApeMat.Text = Trim(g_rst_Genera!PERVIN_APMVIN)
      txt_Nombre.Text = Trim(g_rst_Genera!PERVIN_NOMVIN)
   End If
   
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Genera!PERVIN_SITUAC)
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa(False)
   
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   
   cmb_TDoVin.Enabled = False
   txt_NDoVin.Enabled = False
   cmd_BusVin.Enabled = False
   
   If r_int_TDoVin > 0 Then
      Call cmd_BusVin_Click
   End If
   
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmd_Grabar_Click()
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
   
   If cmb_FlgAcc.ListIndex = -1 Then
      MsgBox "Debe seleccionar si es Accionista.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FlgAcc)
      Exit Sub
   End If
   
   If Left(cmb_FlgAcc.Text, 1) = "2" Then
      If cmb_TDoVin.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de DOI.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TDoVin)
         Exit Sub
      End If
      
      If Len(Trim(txt_NDoVin.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de DOI.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NDoVin)
         Exit Sub
      End If
      
      If cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex) = 1 Then
         If Len(Trim(txt_NDoVin.Text)) > 8 Then
            MsgBox "Debe ingresar el Número de DOI correctamente.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NDoVin)
            Exit Sub
         End If
         
         txt_NDoVin.Text = Format(txt_NDoVin.Text, "00000000")
      End If
         
      If Len(Trim(pnl_PerVin.Caption)) = 0 Then
         MsgBox "Debe ubicar el nombre del Vinculado Titular.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_BusVin)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex) And txt_NumDoc.Text = txt_NDoVin.Text Then
         MsgBox "El vinculado no puede ser el mismo que el titular.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_BusVin)
         Exit Sub
      End If
   End If
   
   If cmb_RelLab.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Relación Laboral.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_RelLab)
      Exit Sub
   End If
   
   If Left(cmb_RelLab.Text, 1) = "5" Then
      If cmb_TDoVin.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de DOI.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TDoVin)
         Exit Sub
      End If
      
      If Len(Trim(txt_NDoVin.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de DOI.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NDoVin)
         Exit Sub
      End If
      
      If cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex) = 1 Then
         If Len(Trim(txt_NDoVin.Text)) > 8 Then
            MsgBox "Debe ingresar el Número de DOI correctamente.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NDoVin)
            Exit Sub
         End If
         
         txt_NDoVin.Text = Format(txt_NDoVin.Text, "00000000")
      End If
         
      If Len(Trim(pnl_PerVin.Caption)) = 0 Then
         MsgBox "Debe ubicar el nombre del Vinculado Titular.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_BusVin)
         Exit Sub
      End If
   
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex) And txt_NumDoc.Text = txt_NDoVin.Text Then
         MsgBox "El vinculado no puede ser el mismo que el titular.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_BusVin)
         Exit Sub
      End If
   End If
   
   If Left(cmb_FlgAcc.Text, 1) = "0" And Left(cmb_RelLab.Text, 1) = "0" Then
      MsgBox "Esta persona no puede ser registrada como no accionista y con ninguna relación laboral.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_BusVin)
      Exit Sub
   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
  
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      
      If cmb_TDoVin.ListIndex = -1 Then
         g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
         g_str_Parame = g_str_Parame & "PERVIN_TDOTIT = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
         g_str_Parame = g_str_Parame & "PERVIN_NDOTIT = '" & txt_NumDoc.Text & "' AND "
         g_str_Parame = g_str_Parame & "PERVIN_TDOVIN = 0"
      Else
         g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
         g_str_Parame = g_str_Parame & "PERVIN_TDOTIT = " & CStr(cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex)) & " AND "
         g_str_Parame = g_str_Parame & "PERVIN_NDOTIT = '" & txt_NDoVin.Text & "' AND "
         g_str_Parame = g_str_Parame & "PERVIN_TDOVIN = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
         g_str_Parame = g_str_Parame & "PERVIN_NDOVIN = '" & txt_NumDoc.Text & "' "
      End If
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "La persona ya ha sido registrada. Por favor verifique los datos e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
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
      
      g_str_Parame = "USP_CRE_PERVIN ("
      
      If cmb_TDoVin.ListIndex > -1 Then
         g_str_Parame = g_str_Parame & CStr(cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NDoVin.Text & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & Left(cmb_FlgAcc.Text, 1) & "', "
         
         If cmb_RelLab.Enabled Then
            g_str_Parame = g_str_Parame & "'" & Left(cmb_RelLab.Text, 1) & "', "
         Else
            g_str_Parame = g_str_Parame & "'0', "
         End If
         
         g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      Else
         g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
         g_str_Parame = g_str_Parame & "'" & Left(cmb_FlgAcc.Text, 1) & "', "
         g_str_Parame = g_str_Parame & "'" & Left(cmb_RelLab.Text, 1) & "', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "' ', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
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
   
   'Grabando Situación
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   If cmb_TDoVin.ListIndex = -1 Then
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         g_str_Parame = "USP_CRE_PERVIN_SITUAC ("
         
         g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
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
   Else
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         g_str_Parame = "USP_CRE_PERVIN_SITUAC_VINCUL ("
         
         g_str_Parame = g_str_Parame & CStr(cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NDoVin.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
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
   
   Screen.MousePointer = 11
   
   Call fs_Buscar
   Call cmd_Cancel_Click
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
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

Private Sub txt_NDoVin_GotFocus()
   Call gs_SelecTodo(txt_NDoVin)
End Sub

Private Sub txt_NDoVin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusVin)
   Else
      If cmb_TDoVin.ListIndex > -1 Then
         Select Case cmb_TDoVin.ItemData(cmb_TDoVin.ListIndex)
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

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_FlgAcc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & ".,- _")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApePat)
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
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   
   Call fs_Activa(True)
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 2025
   grd_Listad.ColWidth(1) = 3915
   grd_Listad.ColWidth(2) = 4365
   grd_Listad.ColWidth(3) = 1905
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoVin, 1, "230")
   
   'Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   'Call moddat_gs_Carga_TipDocIde(cmb_TDoVin, 1)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgAcc, 1, "052")
   Call moddat_gs_Carga_LisIte_Combo(cmb_RelLab, 1, "053")
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = p_Activa
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   
   cmb_TipDoc.Enabled = Not p_Activa
   txt_NumDoc.Enabled = Not p_Activa
   txt_ApePat.Enabled = Not p_Activa
   txt_ApeMat.Enabled = Not p_Activa
   txt_Nombre.Enabled = Not p_Activa
   cmb_FlgAcc.Enabled = Not p_Activa
   cmb_RelLab.Enabled = Not p_Activa
   
   cmb_TDoVin.Enabled = Not p_Activa
   txt_NDoVin.Enabled = Not p_Activa
   cmd_BusVin.Enabled = Not p_Activa
   
   cmb_Situac.Enabled = Not p_Activa
   
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   cmb_FlgAcc.ListIndex = -1
   cmb_RelLab.ListIndex = -1
   
   cmb_TDoVin.ListIndex = -1
   txt_NDoVin.Text = ""
   pnl_PerVin.Caption = ""
   
   cmb_Situac.ListIndex = -1
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PERVIN "
   g_str_Parame = g_str_Parame & "ORDER BY PERVIN_TDOTIT ASC, PERVIN_NDOTIT ASC"

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
      grd_Listad.Text = CStr(g_rst_Princi!PERVIN_TDOTIT) & " - " & Trim(g_rst_Princi!PERVIN_NDOTIT)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!PERVIN_APPTIT) & " " & Trim(g_rst_Princi!PERVIN_APMTIT) & " " & Trim(g_rst_Princi!PERVIN_NOMTIT)
      
      If g_rst_Princi!PERVIN_TDOVIN > 0 Then
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!PERVIN_TDOVIN) & " - " & Trim(g_rst_Princi!PERVIN_NDOVIN) & " / " & Trim(g_rst_Princi!PERVIN_APPVIN) & " " & Trim(g_rst_Princi!PERVIN_APMVIN) & " " & Trim(g_rst_Princi!PERVIN_NOMVIN)
      End If
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Consulta_ParDes("013", CStr(g_rst_Princi!PERVIN_SITUAC))
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!PERVIN_TDOVIN)
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!PERVIN_NDOVIN)
      
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

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub


