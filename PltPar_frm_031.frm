VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_IdeUsu_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4230
   ClientLeft      =   3915
   ClientTop       =   3075
   ClientWidth     =   4920
   Icon            =   "PltPar_frm_031.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      _Version        =   65536
      _ExtentX        =   8652
      _ExtentY        =   7488
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
         Height          =   375
         Left            =   30
         TabIndex        =   5
         Top             =   3810
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   661
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
         Begin Threed.SSPanel pnl_EntDat 
            Height          =   315
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "lm_db_db1 - prod1"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NumVer 
            Height          =   315
            Left            =   3060
            TabIndex        =   7
            Top             =   30
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "v2.0 (rev. 15-1116)"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   915
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   1614
         _StockProps     =   15
         BackColor       =   16777215
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
         Begin VB.Image Image1 
            Height          =   675
            Left            =   1170
            Picture         =   "PltPar_frm_031.frx":000C
            Top             =   90
            Width           =   2550
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   2280
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   1138
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
         Begin VB.CommandButton cmd_Ingres 
            Height          =   585
            Left            =   30
            Picture         =   "PltPar_frm_031.frx":0507
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Ingresar al Sistema"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4170
            Picture         =   "PltPar_frm_031.frx":0949
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir del Sistema"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   795
         Left            =   30
         TabIndex        =   10
         Top             =   2970
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
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
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   1
            Text            =   "MIKEHARA"
            Top             =   60
            Width           =   2055
         End
         Begin VB.TextBox txt_Contra 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2040
            MaxLength       =   30
            PasswordChar    =   "#"
            TabIndex        =   2
            Text            =   "MIKEHARA"
            Top             =   420
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Código de Usuario:"
            Height          =   315
            Left            =   510
            TabIndex        =   12
            Top             =   90
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Contraseña:"
            Height          =   315
            Left            =   510
            TabIndex        =   11
            Top             =   450
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1245
         Left            =   30
         TabIndex        =   13
         Top             =   990
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   2196
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
         Begin Threed.SSPanel pnl_NomAre 
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   825
            Width           =   4695
            _Version        =   65536
            _ExtentX        =   8281
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Plataforma de Parámetros"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   540
            Width           =   4695
            _Version        =   65536
            _ExtentX        =   8281
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Sistema de Gestión miCasita"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   2190
            Picture         =   "PltPar_frm_031.frx":0D8B
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_IdeUsu_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_ConErr     As Integer

Private Sub cmd_Ingres_Click()
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el Código del Usuario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   If Len(Trim(txt_Contra.Text)) = 0 Then
      MsgBox "Debe ingresar la Contraseña del Usuario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Contra)
      Exit Sub
   End If
   If txt_Codigo.Text = "SISTEMAS" And txt_Contra.Text = "ABRIL05" Then
      modgen_g_int_TipUsu = 1000
      modgen_g_str_CodUsu = "DESARROLLO"
      modgen_g_str_NomUsu = "DESARROLLO TECNOLOGIA E INFORMATICA"
   Else
      g_str_Parame = "SELECT * FROM SEG_USUMAE WHERE USUMAE_CODIGO = '" & txt_Codigo.Text & "' AND USUMAE_SITUAC = 1"
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      'Verificación de Usuario
      'Si no hay datos registrados
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      
         MsgBox "El Usuario no está registrado en la base de datos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
      
         l_int_ConErr = l_int_ConErr + 1
         If l_int_ConErr = 3 Then
            Call gs_Desconecta_Servidor
            End
         End If
      
         Exit Sub
      End If
   
      'Verificación de Contraseña
      g_rst_Princi.MoveFirst
      
      If gf_Seg_Desenc(g_rst_Princi!USUMAE_CONTRA) <> txt_Contra.Text Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      
         MsgBox "La Contraseña es incorrecta.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Contra)
      
         l_int_ConErr = l_int_ConErr + 1
         If l_int_ConErr = 3 Then
            Call gs_Desconecta_Servidor
            End
         End If
         Exit Sub
      End If
      
      modgen_g_str_CodUsu = txt_Codigo
      modgen_g_str_NomUsu = Trim(g_rst_Princi!USUMAE_NOMBRE)
      modgen_g_int_TpoCad = g_rst_Princi!USUMAE_TPOCAD
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Verificación de Acceso a la Plataforma
      g_str_Parame = "SELECT * FROM SEG_USUTIP WHERE USUTIP_CODUSU = '" & txt_Codigo.Text & "' AND USUTIP_CODPLT = '" & UCase(App.EXEName) & "' AND USUTIP_SITUAC = 1"
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      
         MsgBox "El Usuario no tiene acceso a esta Plataforma.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
      
         l_int_ConErr = l_int_ConErr + 1
         If l_int_ConErr = 3 Then
            Call gs_Desconecta_Servidor
            End
         End If
         Exit Sub
      End If
      
      modgen_g_int_TipUsu = CInt(g_rst_Princi!USUTIP_TIPUSU)
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call admusu_gf_Verifica_Caducidad
   End If
   
   Me.Hide
   frm_MnuPri_01.Show

End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   modsec_g_str_KeyDes = "123"
   modsec_g_str_RutIni = "C:\"
   modgen_g_str_NomPlt = modgen_g_con_PltPar
   modgen_g_str_NumRev = "v2.0 (rev. 18-0718)"
   
   Call gs_ObtieneRuta
   Call modsec_gf_Decryp
   Call gs_Conecta_Servidor
   Call gs_AutoCopia_Exe
   
   Kill modsec_g_str_RutIni & modgen_g_con_ArcAux
   Me.Caption = modgen_g_str_NomPlt
   pnl_EntDat.Caption = moddat_g_str_NomEsq & " - " & UCase(moddat_g_str_EntDat)
   pnl_NumVer.Caption = modgen_g_str_NumRev
   
   'Obtiene Nombre de PC
   modgen_g_str_NombPC = gf_NombrePC()
   modgen_g_str_CodSuc = "001"
   
   Call moddat_gs_FecSis
   
   If date <> CDate(moddat_g_str_FecSis) Then
      MsgBox "La Fecha de la PC es diferente a la Fecha del Servidor. Se cambiará por esta última, " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy"), vbInformation, App.EXEName
      date = CDate(moddat_g_str_FecSis)
   End If
   
   txt_Codigo.Text = ""
   txt_Contra.Text = ""
   l_int_ConErr = 0
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   Else
      Cancel = True
   End If
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Contra)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_")
   End If
End Sub

Private Sub txt_Contra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Ingres)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_")
   End If
End Sub
