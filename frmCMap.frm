VERSION 5.00
Begin VB.Form frmCMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Map"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   9270
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9270
      TabIndex        =   241
      Top             =   3000
      Width           =   9270
      Begin VB.Line lnKeystrokeBorder 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   9240
         X2              =   9240
         Y1              =   0
         Y2              =   250
      End
      Begin VB.Line lnKeystrokeBorder 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   6900
         X2              =   9240
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line lnKeystrokeBorder 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   6900
         X2              =   9240
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnKeystrokeBorder 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   6900
         X2              =   6900
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label lblKeystroke 
         Caption         =   "Keystroke: Spacebar"
         Height          =   195
         Left            =   6960
         TabIndex        =   243
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   2220
      End
      Begin VB.Line lnInfoBorder 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   6840
         X2              =   6840
         Y1              =   0
         Y2              =   250
      End
      Begin VB.Line lnInfoBorder 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   6840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line lnInfoBorder 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line lnInfoBorder 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   6840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblInfo 
         Caption         =   "Shows available characters in the selected font (Tahoma)"
         Height          =   195
         Left            =   60
         TabIndex        =   242
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   6765
      End
   End
   Begin VB.Timer tmrActiveWindow 
      Interval        =   100
      Left            =   8280
      Top             =   2400
   End
   Begin VB.ComboBox cmbFontName 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8050
      TabIndex        =   9
      ToolTipText     =   "Paste all characters in the characters to copy text box to active document"
      Top             =   1925
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8050
      TabIndex        =   8
      ToolTipText     =   "Copy all characters in the characters to copy text box to the clipboard"
      Top             =   1595
      Width           =   1095
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cu&t"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8050
      TabIndex        =   7
      ToolTipText     =   "Move all characters in the characters to copy text box to the clipboard"
      Top             =   1265
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   350
      Left            =   8050
      TabIndex        =   6
      ToolTipText     =   "Copy the highlighted character to the characters to copy text box"
      Top             =   935
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   350
      Left            =   8050
      TabIndex        =   5
      ToolTipText     =   "Close this window (Esc)"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Cl&ear"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8050
      TabIndex        =   4
      ToolTipText     =   "Clear all characters in the characters to copy text box"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraLargeCharBorder 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   520
      Left            =   7305
      TabIndex        =   238
      Top             =   2385
      Visible         =   0   'False
      Width           =   520
      Begin VB.Label lblLargeChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   490
         Left            =   10
         TabIndex        =   239
         Top             =   10
         UseMnemonic     =   0   'False
         Width           =   490
      End
   End
   Begin VB.Frame fraLargeCharShadow 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7440
      TabIndex        =   237
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8640
      Top             =   2400
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "&Locate"
      Enabled         =   0   'False
      Height          =   350
      Left            =   580
      TabIndex        =   11
      Top             =   2505
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtLocate 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   120
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2505
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.PictureBox picCharContainer 
      Height          =   1740
      Left            =   120
      ScaleHeight     =   1680
      ScaleWidth      =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   7735
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   -15
         TabIndex        =   240
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "!"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   225
         TabIndex        =   236
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   """"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   465
         TabIndex        =   235
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "#"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   705
         TabIndex        =   234
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   945
         TabIndex        =   233
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1185
         TabIndex        =   232
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1425
         TabIndex        =   231
         Tag             =   "&"
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "'"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   1665
         TabIndex        =   230
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "("
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   1905
         TabIndex        =   229
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ")"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   2145
         TabIndex        =   228
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   2385
         TabIndex        =   227
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   2625
         TabIndex        =   226
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ","
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   2865
         TabIndex        =   225
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   3105
         TabIndex        =   224
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   3345
         TabIndex        =   223
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "/"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   3585
         TabIndex        =   222
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   3825
         TabIndex        =   221
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   4065
         TabIndex        =   220
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   4305
         TabIndex        =   219
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   4545
         TabIndex        =   218
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   4785
         TabIndex        =   217
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   5025
         TabIndex        =   216
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   5265
         TabIndex        =   215
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   5505
         TabIndex        =   214
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   5745
         TabIndex        =   213
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   5985
         TabIndex        =   212
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   6225
         TabIndex        =   211
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ";"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   6465
         TabIndex        =   210
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   6705
         TabIndex        =   209
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   6945
         TabIndex        =   208
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   7185
         TabIndex        =   207
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   7425
         TabIndex        =   206
         Top             =   -15
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "@"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   -15
         TabIndex        =   205
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   225
         TabIndex        =   204
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   465
         TabIndex        =   203
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   705
         TabIndex        =   202
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   945
         TabIndex        =   201
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   1185
         TabIndex        =   200
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   1425
         TabIndex        =   199
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   1665
         TabIndex        =   198
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   1905
         TabIndex        =   197
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   2145
         TabIndex        =   196
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "J"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   42
         Left            =   2385
         TabIndex        =   195
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "K"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   2625
         TabIndex        =   194
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   44
         Left            =   2865
         TabIndex        =   193
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   3105
         TabIndex        =   192
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   3345
         TabIndex        =   191
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "O"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   3585
         TabIndex        =   190
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   48
         Left            =   3825
         TabIndex        =   189
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   49
         Left            =   4065
         TabIndex        =   188
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   50
         Left            =   4305
         TabIndex        =   187
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   51
         Left            =   4545
         TabIndex        =   186
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   52
         Left            =   4785
         TabIndex        =   185
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "U"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   53
         Left            =   5025
         TabIndex        =   184
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   54
         Left            =   5265
         TabIndex        =   183
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   55
         Left            =   5505
         TabIndex        =   182
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   56
         Left            =   5745
         TabIndex        =   181
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   57
         Left            =   5985
         TabIndex        =   180
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   58
         Left            =   6225
         TabIndex        =   179
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "["
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   59
         Left            =   6465
         TabIndex        =   178
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   60
         Left            =   6705
         TabIndex        =   177
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "]"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   61
         Left            =   6945
         TabIndex        =   176
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "^"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   62
         Left            =   7185
         TabIndex        =   175
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   63
         Left            =   7425
         TabIndex        =   174
         Top             =   225
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "`"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   64
         Left            =   -15
         TabIndex        =   173
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "a"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   65
         Left            =   225
         TabIndex        =   172
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "b"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   66
         Left            =   465
         TabIndex        =   171
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "c"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   67
         Left            =   705
         TabIndex        =   170
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "d"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   68
         Left            =   945
         TabIndex        =   169
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "e"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   69
         Left            =   1185
         TabIndex        =   168
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "f"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   70
         Left            =   1425
         TabIndex        =   167
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "g"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   71
         Left            =   1665
         TabIndex        =   166
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "h"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   72
         Left            =   1905
         TabIndex        =   165
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "i"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   73
         Left            =   2145
         TabIndex        =   164
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "j"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   74
         Left            =   2385
         TabIndex        =   163
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "k"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   75
         Left            =   2625
         TabIndex        =   162
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "l"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   76
         Left            =   2865
         TabIndex        =   161
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "m"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   77
         Left            =   3105
         TabIndex        =   160
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "n"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   78
         Left            =   3345
         TabIndex        =   159
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "o"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   79
         Left            =   3585
         TabIndex        =   158
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "p"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   80
         Left            =   3825
         TabIndex        =   157
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "q"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   81
         Left            =   4065
         TabIndex        =   156
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "r"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   82
         Left            =   4305
         TabIndex        =   155
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "s"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   83
         Left            =   4545
         TabIndex        =   154
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "t"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   84
         Left            =   4785
         TabIndex        =   153
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "u"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   85
         Left            =   5025
         TabIndex        =   152
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "v"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   86
         Left            =   5265
         TabIndex        =   151
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "w"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   87
         Left            =   5505
         TabIndex        =   150
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   88
         Left            =   5745
         TabIndex        =   149
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "y"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   89
         Left            =   5985
         TabIndex        =   148
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "z"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   90
         Left            =   6225
         TabIndex        =   147
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "{"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   91
         Left            =   6465
         TabIndex        =   146
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "|"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   92
         Left            =   6705
         TabIndex        =   145
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "}"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   93
         Left            =   6945
         TabIndex        =   144
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "~"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   94
         Left            =   7185
         TabIndex        =   143
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ""
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   95
         Left            =   7425
         TabIndex        =   142
         Tag             =   "Ctrl+"
         Top             =   465
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ä"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   96
         Left            =   -15
         TabIndex        =   141
         Tag             =   "Ctrl+Alt+4"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Å"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   97
         Left            =   225
         TabIndex        =   140
         Tag             =   "Alt+0129"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ç"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   98
         Left            =   465
         TabIndex        =   139
         Tag             =   "Alt+0130"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "É"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   99
         Left            =   705
         TabIndex        =   138
         Tag             =   "Alt+0131"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ñ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   100
         Left            =   945
         TabIndex        =   137
         Tag             =   "Alt+0132"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ö"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   101
         Left            =   1185
         TabIndex        =   136
         Tag             =   "Alt+0133"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ü"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   102
         Left            =   1425
         TabIndex        =   135
         Tag             =   "Alt+0134"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "á"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   103
         Left            =   1665
         TabIndex        =   134
         Tag             =   "Alt+0135"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "à"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   104
         Left            =   1905
         TabIndex        =   133
         Tag             =   "Alt+0136"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "â"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   105
         Left            =   2145
         TabIndex        =   132
         Tag             =   "Alt+0137"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ä"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   106
         Left            =   2385
         TabIndex        =   131
         Tag             =   "Alt+0138"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ã"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   107
         Left            =   2625
         TabIndex        =   130
         Tag             =   "Alt+0139"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "å"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   108
         Left            =   2865
         TabIndex        =   129
         Tag             =   "Alt+0140"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ç"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   109
         Left            =   3105
         TabIndex        =   128
         Tag             =   "Alt+0141"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "é"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   110
         Left            =   3345
         TabIndex        =   127
         Tag             =   "Alt+0142"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "è"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   111
         Left            =   3585
         TabIndex        =   126
         Tag             =   "Alt+0143"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ê"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   112
         Left            =   3825
         TabIndex        =   125
         Tag             =   "Alt+0144"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ë"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   113
         Left            =   4065
         TabIndex        =   124
         Tag             =   "Alt+0145"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "í"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   114
         Left            =   4305
         TabIndex        =   123
         Tag             =   "Alt+0146"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ì"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   115
         Left            =   4545
         TabIndex        =   122
         Tag             =   "Alt+0147"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "î"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   116
         Left            =   4785
         TabIndex        =   121
         Tag             =   "Alt+0148"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ï"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   117
         Left            =   5025
         TabIndex        =   120
         Tag             =   "Alt+0149"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ñ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   118
         Left            =   5265
         TabIndex        =   119
         Tag             =   "Alt+0150"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ó"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   119
         Left            =   5505
         TabIndex        =   118
         Tag             =   "Alt+0151"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ò"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   120
         Left            =   5745
         TabIndex        =   117
         Tag             =   "Alt+0152"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ô"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   121
         Left            =   5985
         TabIndex        =   116
         Tag             =   "Alt+0153"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ö"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   122
         Left            =   6225
         TabIndex        =   115
         Tag             =   "Alt+0154"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "õ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   123
         Left            =   6465
         TabIndex        =   114
         Tag             =   "Alt+0155"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ú"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   124
         Left            =   6705
         TabIndex        =   113
         Tag             =   "Alt+0156"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ù"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   125
         Left            =   6945
         TabIndex        =   112
         Tag             =   "Alt+0157"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "û"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   126
         Left            =   7185
         TabIndex        =   111
         Tag             =   "Alt+0158"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ü"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   127
         Left            =   7425
         TabIndex        =   110
         Tag             =   "Alt+0159"
         Top             =   705
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "†"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   128
         Left            =   -15
         TabIndex        =   109
         Tag             =   "Alt+0160"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "°"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   129
         Left            =   225
         TabIndex        =   108
         Tag             =   "Alt+0161"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¢"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   130
         Left            =   465
         TabIndex        =   107
         Tag             =   "Alt+0162"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "£"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   131
         Left            =   705
         TabIndex        =   106
         Tag             =   "£"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "§"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   132
         Left            =   945
         TabIndex        =   105
         Tag             =   "Alt+0164"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "•"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   133
         Left            =   1185
         TabIndex        =   104
         Tag             =   "Alt+0165"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¶"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   134
         Left            =   1425
         TabIndex        =   103
         Tag             =   "Ctrl+Alt+ë"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ß"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   135
         Left            =   1665
         TabIndex        =   102
         Tag             =   "Alt+0167"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "®"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   136
         Left            =   1905
         TabIndex        =   101
         Tag             =   "Alt+0168"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "©"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   137
         Left            =   2145
         TabIndex        =   100
         Tag             =   "Alt+0169"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "™"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   138
         Left            =   2385
         TabIndex        =   99
         Tag             =   "Alt+0170"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "´"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   139
         Left            =   2625
         TabIndex        =   98
         Tag             =   "Alt+0171"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¨"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   140
         Left            =   2865
         TabIndex        =   97
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≠"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   141
         Left            =   3105
         TabIndex        =   96
         Tag             =   "Alt+0173"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Æ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   142
         Left            =   3345
         TabIndex        =   95
         Tag             =   "Alt+0174"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ø"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   143
         Left            =   3585
         TabIndex        =   94
         Tag             =   "Alt+0175"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∞"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   144
         Left            =   3825
         TabIndex        =   93
         Tag             =   "Alt+0176"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "±"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   145
         Left            =   4065
         TabIndex        =   92
         Tag             =   "Alt+0177"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≤"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   146
         Left            =   4305
         TabIndex        =   91
         Tag             =   "Alt+0178"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≥"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   147
         Left            =   4545
         TabIndex        =   90
         Tag             =   "Alt+0179"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¥"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   148
         Left            =   4785
         TabIndex        =   89
         Tag             =   "Alt+0180"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "µ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   149
         Left            =   5025
         TabIndex        =   88
         Tag             =   "Alt+0181"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∂"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   150
         Left            =   5265
         TabIndex        =   87
         Tag             =   "Alt+0182"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∑"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   151
         Left            =   5505
         TabIndex        =   86
         Tag             =   "Alt+0183"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∏"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   152
         Left            =   5745
         TabIndex        =   85
         Tag             =   "Alt+0184"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "π"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   153
         Left            =   5985
         TabIndex        =   84
         Tag             =   "Alt+0185"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∫"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   154
         Left            =   6225
         TabIndex        =   83
         Tag             =   "Alt+0186"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ª"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   155
         Left            =   6465
         TabIndex        =   82
         Tag             =   "Alt+0187"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "º"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   156
         Left            =   6705
         TabIndex        =   81
         Tag             =   "Alt+0188"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ω"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   157
         Left            =   6945
         TabIndex        =   80
         Tag             =   "Alt+0189"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "æ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   158
         Left            =   7185
         TabIndex        =   79
         Tag             =   "Alt+0190"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ø"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   159
         Left            =   7425
         TabIndex        =   78
         Tag             =   "Alt+0191"
         Top             =   945
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¿"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   160
         Left            =   -15
         TabIndex        =   77
         Tag             =   "Alt+0192"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¡"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   161
         Left            =   225
         TabIndex        =   76
         Tag             =   "Shift+Ctrl+Alt+A"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¬"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   162
         Left            =   465
         TabIndex        =   75
         Tag             =   "Alt+0194"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "√"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   163
         Left            =   705
         TabIndex        =   74
         Tag             =   "Alt+0195"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ƒ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   164
         Left            =   945
         TabIndex        =   73
         Tag             =   "Alt+0196"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   165
         Left            =   1185
         TabIndex        =   72
         Tag             =   "Alt+0197"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "∆"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   166
         Left            =   1425
         TabIndex        =   71
         Tag             =   "Alt+0198"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "«"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   167
         Left            =   1665
         TabIndex        =   70
         Tag             =   "Alt+0199"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "»"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   168
         Left            =   1905
         TabIndex        =   69
         Tag             =   "Alt+0200"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "…"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   169
         Left            =   2145
         TabIndex        =   68
         Tag             =   "Shift+Ctrl+Alt+E"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   170
         Left            =   2385
         TabIndex        =   67
         Tag             =   "Alt+0202"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "À"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   171
         Left            =   2625
         TabIndex        =   66
         Tag             =   "Alt+0203"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ã"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   172
         Left            =   2865
         TabIndex        =   65
         Tag             =   "Alt+0204"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Õ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   173
         Left            =   3105
         TabIndex        =   64
         Tag             =   "Shift+Ctrl+Alt+I"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Œ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   174
         Left            =   3345
         TabIndex        =   63
         Tag             =   "Alt+0206"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "œ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   175
         Left            =   3585
         TabIndex        =   62
         Tag             =   "Alt+0207"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "–"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   176
         Left            =   3825
         TabIndex        =   61
         Tag             =   "Alt+0208"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "—"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   177
         Left            =   4065
         TabIndex        =   60
         Tag             =   "Alt+0209"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "“"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   178
         Left            =   4305
         TabIndex        =   59
         Tag             =   "Alt+0210"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "”"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   179
         Left            =   4545
         TabIndex        =   58
         Tag             =   "Shift+Ctrl+Alt+O"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‘"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   180
         Left            =   4785
         TabIndex        =   57
         Tag             =   "Alt+0212"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "’"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   181
         Left            =   5025
         TabIndex        =   56
         Tag             =   "Alt+0213"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "÷"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   182
         Left            =   5265
         TabIndex        =   55
         Tag             =   "Alt+0214"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "◊"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   183
         Left            =   5505
         TabIndex        =   54
         Tag             =   "Alt+0215"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ÿ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   184
         Left            =   5745
         TabIndex        =   53
         Tag             =   "Alt+0216"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ÿ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   185
         Left            =   5985
         TabIndex        =   52
         Tag             =   "Alt+0217"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "⁄"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   186
         Left            =   6225
         TabIndex        =   51
         Tag             =   "Shift+Ctrl+Alt+U"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "€"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   187
         Left            =   6465
         TabIndex        =   50
         Tag             =   "Alt+0219"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‹"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   188
         Left            =   6705
         TabIndex        =   49
         Tag             =   "Alt+0220"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "›"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   189
         Left            =   6945
         TabIndex        =   48
         Tag             =   "Alt+0221"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬁ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   190
         Left            =   7185
         TabIndex        =   47
         Tag             =   "Alt+0222"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬂ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   191
         Left            =   7425
         TabIndex        =   46
         Tag             =   "Alt+0223"
         Top             =   1185
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‡"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   192
         Left            =   -15
         TabIndex        =   45
         Tag             =   "Alt+0224"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "·"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   193
         Left            =   225
         TabIndex        =   44
         Tag             =   "Ctrl+Alt+A"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‚"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   194
         Left            =   465
         TabIndex        =   43
         Tag             =   "Alt+0226"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "„"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   195
         Left            =   705
         TabIndex        =   42
         Tag             =   "Alt+0227"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‰"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   196
         Left            =   945
         TabIndex        =   41
         Tag             =   "Alt+0228"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Â"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   197
         Left            =   1185
         TabIndex        =   40
         Tag             =   "Alt+0229"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ê"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   198
         Left            =   1425
         TabIndex        =   39
         Tag             =   "Alt+0230"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Á"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   199
         Left            =   1665
         TabIndex        =   38
         Tag             =   "Alt+0231"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ë"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   200
         Left            =   1905
         TabIndex        =   37
         Tag             =   "Alt+0232"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "È"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   201
         Left            =   2145
         TabIndex        =   36
         Tag             =   "Ctrl+Alt+E"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Í"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   202
         Left            =   2385
         TabIndex        =   35
         Tag             =   "Alt+0234"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Î"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   203
         Left            =   2625
         TabIndex        =   34
         Tag             =   "Alt+0235"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ï"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   204
         Left            =   2865
         TabIndex        =   33
         Tag             =   "Alt+0236"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ì"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   205
         Left            =   3105
         TabIndex        =   32
         Tag             =   "Ctrl+Alt+I"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ó"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   206
         Left            =   3345
         TabIndex        =   31
         Tag             =   "Alt+0238"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ô"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   207
         Left            =   3585
         TabIndex        =   30
         Tag             =   "Alt+0239"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ""
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   208
         Left            =   3825
         TabIndex        =   29
         Tag             =   "Alt+0240"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ò"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   209
         Left            =   4065
         TabIndex        =   28
         Tag             =   "Alt+0241"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ú"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   210
         Left            =   4305
         TabIndex        =   27
         Tag             =   "Alt+0242"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Û"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   211
         Left            =   4545
         TabIndex        =   26
         Tag             =   "Ctrl+Alt+O"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ù"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   212
         Left            =   4785
         TabIndex        =   25
         Tag             =   "Alt+0244"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ı"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   213
         Left            =   5025
         TabIndex        =   24
         Tag             =   "Alt+0245"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ˆ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   214
         Left            =   5265
         TabIndex        =   23
         Tag             =   "Alt+0246"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "˜"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   215
         Left            =   5505
         TabIndex        =   22
         Tag             =   "Alt+0247"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¯"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   216
         Left            =   5745
         TabIndex        =   21
         Tag             =   "Alt+0248"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "˘"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   217
         Left            =   5985
         TabIndex        =   20
         Tag             =   "Alt+0249"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "˙"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   218
         Left            =   6225
         TabIndex        =   19
         Tag             =   "Ctrl+Alt+U"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "˚"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   219
         Left            =   6465
         TabIndex        =   18
         Tag             =   "Alt+0251"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "¸"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   220
         Left            =   6705
         TabIndex        =   17
         Tag             =   "Alt+0252"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "˝"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   221
         Left            =   6945
         TabIndex        =   16
         Tag             =   "Alt+0253"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "˛"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   222
         Left            =   7185
         TabIndex        =   15
         Tag             =   "Alt+0254"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ˇ"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   223
         Left            =   7425
         TabIndex        =   14
         Tag             =   "Alt+0255"
         Top             =   1425
         UseMnemonic     =   0   'False
         Width           =   255
      End
   End
   Begin VB.TextBox txtCopyChars 
      Height          =   330
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdKeyboardProps 
      Caption         =   "&Keyboard Properties"
      Height          =   350
      Left            =   1560
      TabIndex        =   12
      Top             =   2505
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblCopyChars 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ch&aracters to Copy"
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   195
      Width           =   1410
   End
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Font:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   200
      Width           =   390
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPop_OnTop 
         Caption         =   "&Always on Top"
      End
      Begin VB.Menu mnuPop_Tools 
         Caption         =   "&Tools"
      End
      Begin VB.Menu mnuPop_Seperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_Keys 
         Caption         =   "&Key Codes"
      End
   End
End
Attribute VB_Name = "frmCMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-=-=-=-=-=-==-=-=-=-=-=-=-=  Info =-=-=-=-==-=-=-=-=-=-=-==-=-=-=-=-=-'
'Written by Ricky Bull 20/02/2001-23/02/2001
'Please feel free to use this form in you applications.
'Any comments/modifications can be sent to rickbull@rickmusic.co.uk
'Please leave this intro here if you use this form and have fun!
'-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-==-=-=-=-=-=-=-'
Option Explicit 'Declare all variables
'Enumerations
'Wait Constants - what the wait function time span is in
Private Enum TimeSpanConsts
    MilliSeconds = 1 'Constant for the wait function time span in MS
    Seconds = 1000 'Constant for the wait function time span in S
    Minutes = 60000 'Constant for the wait function time span in M
End Enum

'Highlight/normal constants for the small characters
Private Const NormalBackColour As Long = vbWindowBackground 'The characters normal background colour
Private Const NormalForeColour As Long = vbWindowText 'The characters text background colour
Private Const HighlightBackColour As Long = vbHighlight 'The characters normal background colour
Private Const HighlightForeColour As Long = vbHighlightText 'The characters highlighted background colour
'Highlight/normal constants for the large character (only used when locate button is pressed) _
    Default value is set to the same as the normal colour, so as not to highlight
Private Const LargeCharNormalBackColour As Long = vbWindowBackground 'The characters normal background colour
Private Const LargeCharNormalForeColour As Long = vbWindowText 'The characters text background colour
Private Const LargeCharHighlightBackColour As Long = vbWindowBackground  'The characters normal background colour
Private Const LargeCharHighlightForeColour As Long = vbWindowText  'The characters highlighted background colour

Private Const ShadowDifference As Integer = 70 'How far out the large char's shadow is
Private Const SmallCharFontSize As Integer = 8 'How big the small character's font is
Private Const BeepOnFound As Boolean = True 'Whether to beep when char is found when using locate
'Load/Exit constants
Private Const RecallText As Boolean = False   'Whether to get the last text in the chars to _
                                            copy & locate box
Private Const SaveSettings As Boolean = False 'Whether to save/load settings on exit/start

'Form height constants (for extending the form)
Private Const ShowNormal As Single = 3150 'The form's unextended height
Private Const ShowTools As Single = 500 'The form's extended added height
Private Const CharsPerRow As Integer = 32 'How many character fit on one row
Private Const CharsPerCol As Integer = 7 'How many character fit on one row

'Show cursor constants
Private Const CursorHide As Long = 0 'Constant for hiding the cursor - ShowCursor API
Private Const CursorShow As Long = 1 'Constant for showing the cursor - ShowCursor API

'Locate function constants
Private Const NoOfFlashes As Integer = 2 'how many flashes to do when locate is used (note only _
                                the LongHighLightLength const will be used if this is set to 1). Set to 1 or 2. If more than two, two will be used
Private Const ShortHighLightLength As Integer = 150 'How long the highlight should stay the first time when using the locate function
Private Const LongHighLightLength As Integer = 500 'How long the highlight should stay the second time when using the locate function
Private Const HighLightTimeSpan As Integer = MilliSeconds 'the time span used when using the locate function

'API Constants
Private Const HWND_TOPMOST As Long = -1 'Constant for making a form stay on top
Private Const HWND_NOTOPMOST As Long = -2 'Constant for making a form not on top
Private Const SWP_NOMOVE As Long = &H2 'Flags for Always On Top
Private Const SWP_NOSIZE As Long = &H1

'Form's variable
Private LastOn As Integer 'Which label has the highlight
Private LocateActive As Boolean 'Whether the locate function is being used
Private MouseIsHidden As Boolean 'Whether the mouse is hidden
Private LastActive As Long 'The last active window

'Type declarations
Private Type POINTAPI 'The type for putting the x & y co-ordinates into a variable
    X As Long 'Mouse left pos
    Y As Long 'Mouse top pos
End Type
Private Type RECT 'Type for holding Left, Top, Right & Bottom of objects (for DrawFocusRect)
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'API declarations
Private Declare Function ShowCursor Lib "user32" _
    (ByVal bShow As Long) As Long 'Used to hide or show the cursor
Private Declare Function GetCursorPos Lib "user32" _
    (lpPoint As POINTAPI) As Long 'The API Function needed for finding where the mouse is
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long 'API for making a form on top or not
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long 'API used for making buttons 3D
Private Declare Function GetActiveWindow Lib "user32" () As Long 'API used for finding the active windows hWnd
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, _
    lpRect As RECT) As Long 'API for drawing focus RECT (dotted rectangle)

Private Sub OnTop(OnTop As Boolean)

    On Error Resume Next 'Goto next line on an error

    'If the form is wanted to be on top
    If OnTop = True Then
        'Set it on top
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    'If the form isn't
    Else
        'Stop it always being on top
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    
End Sub

Private Sub cmdClear_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Clear the characters to copy text boc
    txtCopyChars.Text = ""
    
End Sub

Private Sub cmdClose_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Unload form
    Unload Me
    
End Sub

Private Sub cmdCopy_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Clear the clipboard
    Clipboard.Clear
    'Add the choosen characters to the clipboard
    Clipboard.SetText txtCopyChars.Text
    
End Sub

Private Sub cmdCut_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Clear the clipboard
    Clipboard.Clear
    'Add the choosen characters to the clipboard
    Clipboard.SetText txtCopyChars.Text
    'Clear the CopyChars text box
    txtCopyChars.Text = ""
    
End Sub

Private Sub cmdKeyboardProps_Click()
    
    On Error Resume Next 'Goto next line on an error
    Dim ReturnValue As Long 'Variable for return value from shell command
    
    'Open the keyboard control pannel
    ReturnValue = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", vbNormalFocus)
    
End Sub

Private Sub cmdLocate_Click()

    On Error Resume Next 'Goto next line on an error
    
    Call LocateChar
    
End Sub

Private Sub cmdPaste_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Add code here to set text into the form's Text box
    
End Sub

Private Sub cmdSelect_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Add the character label's caption to the characters to copy text box
    Call AddToCopyChars(lblChar(LastOn).Caption)
    
End Sub

Private Sub Form_Activate()

    On Error Resume Next 'Goto next line on an error
    
    'Set focus to the character labels
    picCharContainer.SetFocus
        
End Sub

Private Sub Form_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Take the focus off of the
    Call HideFocus
    
End Sub

Private Sub Form_Load()

    'On Error Resume Next 'Goto next line on an error
    Dim Counter As Integer 'for loops
    
    'Set the form's height to normal before changing to the _
    correct height (i.e. whether to show tools)
    Me.Height = ShowNormal
    
    'Loop for all chars
    For Counter = lblChar.LBound To lblChar.uBound
        'If the back colour is not the one it is supposed to be _
         set it to the correct one
        If lblChar(Counter).BackColor <> NormalBackColour Then _
            lblChar(Counter).BackColor = NormalBackColour
        'If the fore colour is not the one it is supposed to be _
         set it to the correct one
        If lblChar(Counter).ForeColor <> NormalForeColour Then _
            lblChar(Counter).ForeColor = NormalForeColour
    'Do the next char
    Next Counter
    
    'Initialize the last on to 0
    LastOn = 0
    LastActive = -1
    
    'Get settings from reg if is wanted
    If SaveSettings = True Then
        Call GetRegSettings
    
    'Get settings from reg if isn't wanted set needed vars
    Else
        'Set the form's height
        Call SetTools
        'Center the form
        Me.Move (Screen.Width / 2) - (Me.Width / 2), _
            (Screen.Height / 2) - (Me.Height / 2), Me.Width, ShowNormal
        'Highlight the firstcharacter
        lblChar(0).BackColor = HighlightBackColour
        lblChar(0).ForeColor = HighlightForeColour
    End If
        
    'Set the MouseIsHidden variable to not hidden
    MouseIsHidden = False
    
    'Set locate to not being used
    LocateActive = False
    
    'Add all fonts to the image combo box
    Call AddFonts
    
    'Set the font name
    cmbFontName.Text = lblChar(0).FontName
          
    'Set the colour of the large char to normal if needed
    If lblLargeChar.BackColor <> LargeCharNormalBackColour Then _
        lblLargeChar.BackColor = LargeCharNormalBackColour
    If lblLargeChar.ForeColor = LargeCharNormalForeColour Then _
        lblLargeChar.ForeColor = LargeCharNormalForeColour
    
End Sub

Private Sub AddFonts()

    On Error Resume Next 'Goto next line on an error
    Dim Counter As Integer 'For loops
    Dim FontType As Integer 'Whether the current font is True Type (1) _
                            or a Screen Font (1). This is done so that _
                            the correct image is used from the image list
    
    'Loop for all fonts on user's system
    For Counter = 0 To Screen.FontCount - 1
        'Find what type the font is
        FontType = 1
        'Add the font name + correct image
        cmbFontName.AddItem Screen.Fonts(Counter)
    'On to next font
    Next Counter
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on an error
    
    'If the right button is pressed show popupmenu
    If Button = vbRightButton Then PopupMenu mnuPop
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next 'Goto next line on an error
    
    'If the forms is big enough to show the tools options
    If Me.Height > ShowNormal Then
        'Set the tools top value
        txtLocate.Top = Me.Height - 1100
        cmdLocate.Top = Me.Height - 1100
        cmdKeyboardProps.Top = Me.Height - 1100
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next 'Goto next line on an error
    
    'If the settings are to be saved
    If SaveSettings = True Then Call SaveRegSettings
    'Clean up resources used
    Set frmCMap = Nothing
    
End Sub

Private Sub fraLargeCharShadow_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Hide it if it is clicked
    Call HideShowLargeChar(False)

End Sub

Private Sub cmbFontName_Click()

    On Error Resume Next 'Goto next line on an error
    Dim Counter As Integer 'For loops
    
    'Only change if it is different
    If lblChar(0).FontName <> cmbFontName.Text Then
        'Loop for all of the character labels
        For Counter = lblChar.LBound To lblChar.uBound
            'Set the font of the character labels to the _
            image combo box
            lblChar(Counter).FontName = cmbFontName.Text
            'Make the font the right size (if needed) as fixed-width font can change this
            If lblChar(Counter).FontSize <> SmallCharFontSize Then _
                lblChar(Counter).FontSize = SmallCharFontSize
        'Do the next label
        Next Counter
        'Set the other font names of objects on the form font
        lblLargeChar.FontName = cmbFontName.Text
        txtCopyChars.FontName = cmbFontName.Text
        txtLocate.FontName = cmbFontName.Text
        'Set the status bar
        If lblInfo.Caption <> "Shows available characters in the selected font (" _
            & cmbFontName.Text & ")" Then lblInfo.Caption = _
            "Shows available characters in the selected font (" & cmbFontName.Text & ")"
    End If
    
End Sub

Private Sub lblChar_DblClick(Index As Integer)
   
    On Error Resume Next 'Goto next line on an error
    
    'Add the character to the character to copy text box
    Call AddToCopyChars(lblChar(Index).Caption)

End Sub

Private Sub lblChar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on an error
    
    'Make the large char goto the right place if locate _
        is not active and left button is pressed
    If Button = vbLeftButton And LocateActive = False Then Call SetLargeCharPos(Index)
    
End Sub

Private Sub lblChar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    'If the left button is down check it position
    If Button = vbLeftButton Then Call CheckMousePos
    
End Sub

Private Sub lblChar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next 'Goto next line on an error
    
    'Hide the large character
    Call HideLargeChar
    
End Sub

Private Sub lblCopyChars_Click()

    On Error Resume Next 'Goto next line on an error

    'Take the focus off of the
    Call HideFocus
    
End Sub

Private Sub lblFont_Click()

    On Error Resume Next 'Goto next line on an error

    'Take the focus off of the
    Call HideFocus
    
End Sub

Private Sub mnuPop_Keys_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Give the user info on the keycodes
    MsgBox "Character Map Keycodes:" & vbNewLine & vbNewLine & _
        "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=" & vbNewLine & _
        "DIRECTION KEYS - Move Focus of Character" & vbNewLine & _
        "SHIFT + LEFT - Go to the first character on the current row" & vbNewLine & _
        "SHIFT + RIGHT - Go to the last character on the current row" & vbNewLine & _
        "SHIFT + UP - Go to the first character on the current column" & vbNewLine & _
        "SHIFT + DOWN - Go to the last character on the current column" & vbNewLine & _
        "HOME - Go to the first character on the Map" & vbNewLine & _
        "END - Go to the last character on the Map" & vbNewLine & _
        "PAGE DOWN - Move two characters down on the current row" & vbNewLine & _
        "PAGE UP - Move two characters up on the current row" & vbNewLine & _
        "ENTER - Select the current characters on the map" & vbNewLine & _
        "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=", _
        vbOKOnly, "Key Codes"

End Sub

Private Sub mnuPop_OnTop_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Invert the check
    mnuPop_OnTop.Checked = Not mnuPop_OnTop.Checked
    'Set whether form is on top according to the checked value
    Call OnTop(mnuPop_OnTop.Checked)
    
End Sub

Private Sub mnuPop_Tools_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Invert the check
    mnuPop_Tools.Checked = Not mnuPop_Tools.Checked
    Call SetTools
    
End Sub

Private Sub picCharContainer_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Take the focus off of the character labels
    Call HideFocus
        
End Sub

Private Sub picCharContainer_GotFocus()

    On Error Resume Next 'Goto next line on an error
    
    'Draw the focus rect
    Call DrawFocus
    
End Sub

Private Sub picCharContainer_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'On Error Resume Next 'Goto next line on an error
    Dim RowNo As Integer 'What row the label we are on
    Dim ColNo As Integer 'What column the label we are on
    
    'If locate isn't being used
    If LocateActive = False Then
        
        'If LEFT + NO SHIFT is pressed and the first one is not selected
        If LastOn > lblChar.LBound And Shift = 0 And KeyCode = vbKeyLeft Then
            'Move the selection to the left one
            Call ShowCharNoMouseHide(LastOn - 1)
            
        'If LEFT + SHIFT is pressed - First label of the row
        ElseIf KeyCode = vbKeyLeft And Shift = vbShiftMask Then
            'Find the last label's bottom and add 15 (to compensate _
            for all being 15 below 0 - to fit in pic box and have a nice look) _
            then divide by label's height
            RowNo = (((lblChar(LastOn).Top + lblChar(LastOn).Height) + 15) / lblChar(0).Height)
            'Use the same variable to find out how much to subtract to get to _
            the first label of the row - _
             1. Find the last label on the row (RowNo * CharsPerRow) and subtract _
                laston to find the differenct; _
             2. Substract the Different from CharactersPerRow (CharsPerRow - ((RowNo * CharsPerRow) - LastOn)) _
                - this now equals any rows differnece to get to the start; _
             3. So now all we need to do is subtract the different from the last _
                highlighted label (LastOn - (CharsPerRow - ((RowNo * CharsPerRow) - LastOn))).
            RowNo = LastOn - (CharsPerRow - ((RowNo * CharsPerRow) - LastOn))
            'Highlight the correct label
            Call ShowCharNoMouseHide(RowNo)   '(RowNo * CharsPerRow - LastOn)
        
        'If RIGHT + NO SHIFT is pressed and the last one is not selected
        ElseIf KeyCode = vbKeyRight And Shift = 0 And LastOn < lblChar.uBound Then
            'Move the selection to the right one
            Call ShowCharNoMouseHide(LastOn + 1)
        
        'If RIGHT + SHIFT is pressed - Last label of the row
        ElseIf KeyCode = vbKeyRight And Shift = vbShiftMask Then
            'Find the last label's bottom and add 15 (to compensate _
            for all being 15 below 0 - to fit in pic box and have a nice look) _
            then divide by label's height
            RowNo = (((lblChar(LastOn).Top + lblChar(LastOn).Height) + 15) / lblChar(0).Height)
            'If the highlighted character is not on the last row (this is done _
            because the last character may not be a multiple of charsperrow, i.e. _
            it may not fill the picture box/container)
            If RowNo / CharsPerCol <> 1 Then
                'Use the same variable to find out how much to add to get to _
                the first label of the row - _
                Do the same as Home + Shift then add the number of characters per row - 1 _
                (otherwise it will go over by one to get to the end
                RowNo = (LastOn - (CharsPerRow - ((RowNo * CharsPerRow) - LastOn))) + (CharsPerRow - 1)
                'Highlight the correct label
            'If it isn't
            Else
                'Make RowNo = the number of characters - the last label
                RowNo = lblChar.uBound
            End If
            Call ShowCharNoMouseHide(RowNo)   '(RowNo * CharsPerRow - LastOn)
        
        'If UP + NO SHIFT is pressed and the selection is not on the top level
        ElseIf LastOn >= 32 And Shift = 0 And KeyCode = vbKeyUp Then
            'Move it up one row
            Call ShowCharNoMouseHide(LastOn - 32)
            
        'If UP + SHIFT is pressed - Last label of the row
        ElseIf KeyCode = vbKeyUp And Shift = vbShiftMask Then
            'Find the last label's bottom and add 15 (to compensate _
            for all being 15 below 0 - to fit in pic box and have a nice look) _
            then divide by label's height
            RowNo = (((lblChar(LastOn).Top + lblChar(0).Height) + 15) / lblChar(0).Height)
            'Subtract ((the row number * the characters per row) from the current row no) _
             and add the chars per row
            RowNo = (LastOn - (RowNo * CharsPerRow)) + CharsPerRow
            Call ShowCharNoMouseHide(RowNo)
        
        'If DOWN + NO SHIFT is pressed and the selection is not on the bottom level
        ElseIf LastOn <= lblChar.uBound - 32 And Shift = 0 And KeyCode = vbKeyDown Then
            'Move it up down row
            Call ShowCharNoMouseHide(LastOn + 32)
                
        'If DOWN + SHIFT is pressed - Last label of the row
        ElseIf KeyCode = vbKeyDown And Shift = vbShiftMask Then
            'Find the last label's bottom and add 15 (to compensate _
            for all being 15 below 0 - to fit in pic box and have a nice look) _
            then divide by label's height
            RowNo = (((lblChar(LastOn).Top + lblChar(LastOn).Height) + 15) / lblChar(0).Height)
            'Subtract (the row number * the characters per row) from the current row no
            RowNo = LastOn - (RowNo * CharsPerRow)
            'ColNo = (((lblChar(LastOn).Left + lblChar(LastOn).Width) + 15) / lblChar(0).Width)
            'The column we need is the character label's count + row number
            ColNo = lblChar.Count + RowNo
            'move the large char
            Call ShowCharNoMouseHide(ColNo)
            
        'If HOME is pressed
        ElseIf KeyCode = vbKeyHome Then
            'Move it to the first character
            Call ShowCharNoMouseHide(0)
           
        'If END is pressed
        ElseIf KeyCode = vbKeyEnd Then
            'Move it to last character
            Call ShowCharNoMouseHide(lblChar.uBound)
            
        'If PAGE DOWN is pressed and the selection is not on the Penultimate level
        ElseIf LastOn <= lblChar.uBound - 64 And KeyCode = vbKeyPageDown Then
            'Move it up down row
            Call ShowCharNoMouseHide(LastOn + 64)
        
        'If PAGE UP is pressed and the selection is not on the second level
        ElseIf LastOn >= lblChar.LBound + 64 And KeyCode = vbKeyPageUp Then
            'Move it down row
            Call ShowCharNoMouseHide(LastOn - 64)
    
        'If ENTER is pressed
        ElseIf KeyCode = vbKeyReturn Then
            'Add it to the chars to copy text box
            Call AddToCopyChars(lblChar(LastOn).Caption)
    
        End If
        
        'Disable the timer so as when the mouse is moved the large _
        character does not follow it (due to the timer being enabled by the _
        call to 'SetLargeCharPos') as we only want it to follow the mouse _
        when it has actauly been clicked by the user
        'tmrMousePos.Enabled = False
        'If the mouse is hidden (from calling the 'lbl_MouseDown' routine)
        'If MouseIsHidden = True Then
            'Show it
            'Call ShowCursor(CursorShow)
            'Set the variable to not hidden
            'MouseIsHidden = False
        'End If
    End If
 
End Sub

Private Sub picCharContainer_KeyPress(KeyAscii As Integer)
    
    Dim Counter As Integer 'For loops
    
    'Loop for all chars
    For Counter = lblChar.LBound To lblChar.uBound
        'The the character on the label = the character of the keypressed
        If lblChar(Counter).Caption = Chr(KeyAscii) Then
            'Set the the large character's position
            Call ShowCharNoMouseHide(Counter)
            'Add the char to the textbox
            Call AddToCopyChars(lblChar(Counter).Caption)
            'Exit the loop as the job is donw
            Exit For
        End If
    'Onto next label if not found
    Next Counter
            
End Sub

Private Sub picCharContainer_LostFocus()

    On Error Resume Next 'Goto next line on an error
    
    'Take the focus off of the character labels
    Call HideFocus
        
End Sub

Private Sub CheckMousePos()

    On Error Resume Next 'Goto next line on an error
    Dim MousePos As POINTAPI 'Where the mouse x & y is
    Dim Counter As Integer 'For loops
    Dim OverallLeft As Single 'How far left the character is
    Dim OverallTop As Single 'How far from the top the character is
    
    'Account for the fact that the characters are in a picture box and the window _
    is not always going to be at points 0,0 on the screen. Then put it into pixels _
    rather than twips
    OverallLeft = ((picCharContainer.Left + Me.Left) / Screen.TwipsPerPixelX) + 5
    OverallTop = ((picCharContainer.Top + Me.Top) / Screen.TwipsPerPixelY) + 25
    'Put the x & y into a variable
    GetCursorPos MousePos
    
    'Loop for all characters
    For Counter = lblChar.LBound To lblChar.uBound  'CharsPerCol * CharsPerRow
        'If the x & y = the charater(counter) position (all in 1 pixels because the _
        characters actualy over lap a pixel - this will cause two characters to be displayed, _
        and keep flashing between the two)
        If MousePos.X >= ((lblChar(Counter).Left / Screen.TwipsPerPixelX) + OverallLeft) + 1 And _
            MousePos.X <= (((lblChar(Counter).Left + lblChar(Counter).Width) / Screen.TwipsPerPixelX) + OverallLeft) - 1 And _
            MousePos.Y >= ((lblChar(Counter).Top / Screen.TwipsPerPixelY) + OverallTop) + 1 And _
            MousePos.Y <= (((lblChar(Counter).Top + lblChar(Counter).Height) / Screen.TwipsPerPixelY) + OverallTop) - 1 Then
            'If the large character is not the correct one make out that it has been clicked
            If lblLargeChar.Caption <> lblChar(Counter).Caption Then Call SetLargeCharPos(Counter)
            'If the cursor is not vissible
            If MouseIsHidden = False Then
                'Show it
                Call ShowCursor(CursorHide)
                'Set variable to say it is visible
                MouseIsHidden = True
            End If
            'Stop the loop as we dont't need it now - done for effeciency
            Exit For
        Else
            'If the label being checked is the last one and the mouse is hidden
            If Counter = lblChar.uBound And MouseIsHidden = True Then
                'Show it as the mouse is not over a label
                Call ShowCursor(CursorShow)
                'Set the variable to not hidden
                MouseIsHidden = False
            End If
        End If
    'On to next character if the mouse is not over it
    Next Counter
    
End Sub

Private Sub tmrActiveWindow_Timer()

    On Error Resume Next 'Goto next line on an error
    Dim NewActive As Long
    
    NewActive = GetActiveWindow
    
    'Take the focus off of the character labels if the form is not active
    If NewActive <> Me.hwnd And LastActive <> NewActive Then
        Call HideFocus
    End If
    LastActive = NewActive
    
End Sub

Private Sub tmrWait_Timer()

    On Error Resume Next 'Goto next line on an error
    
    'This timer does nothing except turn off after _
    the interval and is only used for the 'Wait' procedure _
    - When the timer reaches 0 the wait function ends
    tmrWait.Interval = 0
    
End Sub

Private Sub txtCopyChars_Change()

    On Error Resume Next 'Goto next line on an error
    Dim Enable As Boolean 'Whether the CCP &Clear button should be enabled
    
    'If the text is not empty
    If txtCopyChars.Text <> "" Then
        'Set the variable to true
        Enable = True
    'If the text is empty
    Else
        'Set the variable to false
        Enable = False
    End If
    'Enable/Disable the cut, copy, paste & clear buttons
    cmdCut.Enabled = Enable
    cmdCopy.Enabled = Enable
    cmdPaste.Enabled = Enable
    cmdClear.Enabled = Enable
    
End Sub

Private Sub MoveLargeChar(ByVal LabelIndex As Integer)

    On Error Resume Next 'Goto next line on an error
      
    'Make the large character = the selected label's caption
    lblLargeChar.Caption = lblChar(LabelIndex).Caption
    'Move the border/frame
    fraLargeCharBorder.Move lblChar(LabelIndex).Left + 50, _
        lblChar(LabelIndex).Top + 500
    'Move the border/shadow
    fraLargeCharShadow.Move fraLargeCharBorder.Left + ShadowDifference, _
        fraLargeCharBorder.Top + ShadowDifference
    'Make the containg frame visible if needed
    If fraLargeCharBorder.Visible = False Then Call HideShowLargeChar(True)
    
End Sub

Private Sub HideShowLargeChar(ByVal ShowChar As Boolean)

    On Error Resume Next 'Goto next line on an error
    
    'If the char is to be shown mouse isn't hidden and the locate function is not being used
    If ShowChar = True And MouseIsHidden = False And LocateActive = False Then
        'Hide the cursor
        Call ShowCursor(CursorHide)
        'Set the variable to hidden
        MouseIsHidden = True
    
    'If the char isn't to be shown mouse is hidden and the locate function is being used
    ElseIf ShowChar = False And MouseIsHidden = True And LocateActive = False Then
        'Show the cursor
        Call ShowCursor(CursorShow)
        'Set the variable to not hidden
        MouseIsHidden = False
    End If
    
    'Hide or show the character according to what the user wants
    fraLargeCharBorder.Visible = ShowChar
    fraLargeCharShadow.Visible = ShowChar
        
End Sub

Private Sub lblLargeChar_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Hide it if it is clicked
    Call HideShowLargeChar(False)
    
End Sub

Private Sub txtLocate_Change()

    On Error Resume Next 'Goto next line on an error

    'If the text isn't ""
    If txtLocate.Text <> "" Then
        'Enable the command button
        cmdLocate.Enabled = True
    'If it is
    Else
        'Disable it
        cmdLocate.Enabled = False
    End If
    
End Sub


Private Sub Wait(Optional ByVal Seconds As Integer = 1, _
    Optional ByVal TimeSpan As TimeSpanConsts = Seconds)

    On Error Resume Next 'Goto next line on an error

    'Set the timer to the amount of seconds required in _
    the users choice of time span
    tmrWait.Interval = Seconds * TimeSpan
    'Start timer
    tmrWait.Enabled = True
    'Loop to keep application tied up
    Do While tmrWait.Interval > 0
        'Keep repaintinf form, etc. while in a loop so as _
        not to appeared locked-up
        DoEvents
    'End of loop
    Loop
    'Now that wait has finished turn timer off
    tmrWait.Enabled = False
    
End Sub

Private Sub GetRegSettings()

    On Error Resume Next 'Goto next line on an error
    Dim FormLeft As Single 'Form's left in registry
    Dim FormTop As Single 'Form's top in registry
    Dim MapFont  As String
    
    'Put registry options into variables
    FormLeft = GetSetting(App.EXEName, "Character Map", "Left", 0)
    FormTop = GetSetting(App.EXEName, "Character Map", "Top", 0)
    MapFont = GetSetting(App.EXEName, "Character Map", "Font", "Tahoma")
    
    
    'Get the menu states
    mnuPop_Tools.Checked = GetSetting(App.EXEName, "Character Map", "Tools", False)
    mnuPop_OnTop.Checked = GetSetting(App.EXEName, "Character Map", "On Top", False)
    
    'Display tools if needed
    If mnuPop_Tools.Checked = True Then Call SetTools
    'Set on top if needed
    If mnuPop_OnTop.Checked = True Then Call OnTop(mnuPop_OnTop.Checked)
    
    'Get character map position if wanted and form isn't off of the screen
    If FormLeft > 0 And FormLeft < Screen.Width - Me.Width And _
        FormTop > 0 And FormTop < Screen.Height - Me.Height Then
        'Move the form to the last position saved
        Me.Move FormLeft, FormTop
        
    'If character map position has not been saved or form is off the screen
    Else
        'Center the form
        Me.Move (Screen.Width / 2) - (Me.Width / 2), _
            (Screen.Height / 2) - (Me.Height / 2)
    End If
    
    
    'If the last fontname is different to the current one
    If MapFont <> cmbFontName.Text Then
        'Set the new name in the combo box
        cmbFontName.Text = MapFont
        'Make out that the combo has been clicked so all other fonts _
        will be changed
        Call cmbFontName_Click
    End If
        
    'If the copychars last text is wanted
    If RecallText = True Then
        'Set the text from registry
        Call AddToCopyChars(GetSetting(App.EXEName, "Character Map", "Copy Text", ""))
        txtLocate.Text = GetSetting(App.EXEName, "Character Map", "Locate Text", "")
        Call SetHighlight(GetSetting(App.EXEName, "Character Map", "Highlight", 0))
    Else
        'Highlight the first char if the recall text is false
        Call SetHighlight(0)
    End If
        
End Sub

Private Sub SaveRegSettings()
    
    On Error Resume Next 'Goto next line on an error
    
    'Save character map position
    SaveSetting App.EXEName, "Character Map", "Left", Me.Left
    SaveSetting App.EXEName, "Character Map", "Top", Me.Top
    'Save character map font name
    SaveSetting App.EXEName, "Character Map", "Font", lblChar(0).FontName
    'Save the copy chars text
    SaveSetting App.EXEName, "Character Map", "Copy Text", txtCopyChars.Text
    'Save the locate
    SaveSetting App.EXEName, "Character Map", "Locate Text", txtLocate.Text
    'Save the menu states
    SaveSetting App.EXEName, "Character Map", "Tools", mnuPop_Tools.Checked
    SaveSetting App.EXEName, "Character Map", "On Top", mnuPop_OnTop.Checked
    'Which character has highlight
    SaveSetting App.EXEName, "Character Map", "Highlight", LastOn
    
End Sub

Private Sub DeleteRegSettings(Optional ByVal FullDelete As Boolean = False)

    'This sub is unused by default, but if you want to remove all _
     registry settings just call it (Call DeleteRegSettings)
    On Error Resume Next 'Goto next line on an error
    
    'If all settings are to be deleted. This should only be used if _
    you haven't saved anything else in the 'Character Map' folder in the reg
    If FullDelete = True Then
        DeleteSetting App.EXEName, "Character Map"
    'If only the settings are to be deleted
    Else
        'Save character map position
        DeleteSetting App.EXEName, "Character Map", "Left"
        DeleteSetting App.EXEName, "Character Map", "Top"
        'Save character map font name
        DeleteSetting App.EXEName, "Character Map", "Font"
        'Save the copy chars text
        DeleteSetting App.EXEName, "Character Map", "Copy Text"
        'Save the locate
        DeleteSetting App.EXEName, "Character Map", "Locate Text"
        'Save the menu states
        DeleteSetting App.EXEName, "Character Map", "Tools"
        DeleteSetting App.EXEName, "Character Map", "On Top"
        'Which character has highlight
        DeleteSetting App.EXEName, "Character Map", "Highlight"
    End If
    
End Sub

Private Sub SetTools()

    On Error Resume Next 'Goto next line on an error
    
    'Show or hide the tools as nessacary
    cmdLocate.Visible = mnuPop_Tools.Checked
    txtLocate.Visible = mnuPop_Tools.Checked
    cmdKeyboardProps.Visible = mnuPop_Tools.Checked
    'If tools is checked
    If mnuPop_Tools.Checked = True Then
        'Add extra height to show tools
        Me.Height = Me.Height + ShowTools
    'If tools isn't checked
    ElseIf mnuPop_Tools.Checked = False Then
        'Subtract extra height to hide tools
        Me.Height = Me.Height - ShowTools
    End If
    
End Sub

Private Sub FlashLocated(ByVal LabelIndex As Integer)
    
    On Error Resume Next 'Goto next line on an error
    Dim Counter As Integer 'For loops
    
    'Loop for all the times that the small flash is wanted
    For Counter = 1 To NoOfFlashes - 1
        'Beep if wanted
        If BeepOnFound = True Then Beep
        'Go to the sub section of this procedure call ShowChar
        Call ShowCharNoMouseHide(LabelIndex)
        'Set the locate active to being used
        LocateActive = True
        'Change the colour of the large char to highlight
        lblLargeChar.BackColor = LargeCharHighlightBackColour
        lblLargeChar.ForeColor = LargeCharHighlightForeColour
        'Pause for a specified short time
        Call Wait(ShortHighLightLength, HighLightTimeSpan)
        'Change the colour of the large char back to normal
        lblLargeChar.BackColor = LargeCharNormalBackColour
        lblLargeChar.ForeColor = LargeCharNormalForeColour
        'Go to the sub section of this procedure call HideChar
        GoSub HideChar 'Call HideLargeChar
        'Wait again before flashin again
        Call Wait(ShortHighLightLength, HighLightTimeSpan)
        'Set the locate active to not being used so as we can use the _
        'SetLargeCharPos' routine as it checks this variable
        LocateActive = False
    Next Counter
    
    'Go to the sub section of this procedure call ShowChar
    Call ShowCharNoMouseHide(LabelIndex)
    LocateActive = True
    'Change the colour of the large char to highlight
    lblLargeChar.BackColor = LargeCharHighlightBackColour
    lblLargeChar.ForeColor = LargeCharHighlightForeColour
    'Pause for a specified long time
    Call Wait(LongHighLightLength, HighLightTimeSpan)
    'Change the colour of the large char back to normal
    lblLargeChar.BackColor = LargeCharNormalBackColour
    lblLargeChar.ForeColor = LargeCharNormalForeColour
    'Go to the sub section of this procedure call HideChar
    GoSub HideChar 'Call HideLargeChar
    'Set the locate active to not being used
    LocateActive = False
    Exit Sub

'This is used to hide the character
HideChar:
    'Show the large char
    fraLargeCharBorder.Visible = False
    fraLargeCharShadow.Visible = False
    'Return where from you came
    Return
    
End Sub

Private Sub SetLargeCharPos(ByVal LabelIndex As Integer)

    On Error Resume Next 'Goto next line on an error
    
    'Set the focus
    Call SetHighlight(LabelIndex)
    'move the character
    Call MoveLargeChar(LabelIndex)
    
End Sub

Private Sub HideLargeChar()

    On Error Resume Next 'Goto next line on an error
        
    'Make the character invisible if right button is pressed if needed
    If fraLargeCharBorder.Visible = True Then Call HideShowLargeChar(False)
    Call DrawFocus
     
End Sub

Private Sub SetHighlight(ByVal LabelIndex As Integer)

    'On Error Resume Next 'Goto next line on an error
    Dim KeystrokeCode As String 'The keycode for each symbol
    
    'If the locate function is not being used
    If LocateActive = False Then
        'If the selection is not already on this label
        If LastOn <> LabelIndex Then
            'Set the back & fore colour of the last highlighted label back to normal
            If LastOn > -1 Then
                lblChar(LastOn).BackColor = NormalBackColour
                lblChar(LastOn).ForeColor = NormalForeColour
            End If
            'Set this label's  back & fore colour to highlighted
            lblChar(LabelIndex).BackColor = HighlightBackColour
            lblChar(LabelIndex).ForeColor = HighlightForeColour
            
            'If the tag is blank
            If lblChar(LabelIndex).Tag = "" Then
                'Set thekeystroke to the caption
                KeystrokeCode = lblChar(LabelIndex).Caption
            'If it is not blank
            Else
                'Set it to the tag
                KeystrokeCode = lblChar(LabelIndex).Tag
            End If
            'Set the keystroke status bar
            If lblKeystroke.Caption <> "Keystroke: " & KeystrokeCode Then _
                lblKeystroke.Caption = "Keystroke: " & KeystrokeCode
            'Set this label as the last one with the highlight
            LastOn = LabelIndex
        End If
        'Set focus to the picture box/container so that keystrokes _
        can chnage which label has the highlight
        picCharContainer.SetFocus
    End If
    
End Sub

Private Sub HideFocus()

    On Error Resume Next 'Goto next line on an error
    
    'Make the focus rectangle invisible
    picCharContainer.Cls
    'Hide the large char
    If fraLargeCharBorder.Visible = True Then Call HideShowLargeChar(False)
    
End Sub

Private Sub ShowCharNoMouseHide(ByVal LabelIndex As Integer)
    
    On Error Resume Next 'Goto next line on an error
    
    'This is used to show the character
    'Set the highlight to the label
    Call SetHighlight(LabelIndex)
    'Make the large character = the selected label's caption
    lblLargeChar.Caption = lblChar(LabelIndex).Caption
    'Move the border/frame
    fraLargeCharBorder.Move lblChar(LabelIndex).Left + 50, _
        lblChar(LabelIndex).Top + 500
    'Move the border/shadow
    fraLargeCharShadow.Move fraLargeCharBorder.Left + ShadowDifference, _
        fraLargeCharBorder.Top + ShadowDifference
    'Make the containg frame visible
    fraLargeCharBorder.Visible = True
    fraLargeCharShadow.Visible = True

End Sub

Private Sub AddToCopyChars(ByVal Char As String)

    On Error Resume Next 'Goto next line on an error
    
    'Add the char
    txtCopyChars.Text = txtCopyChars.Text + Char
    'Set the selstart
    txtCopyChars.SelStart = Len(txtCopyChars.Text)
    
End Sub

Private Sub txtLocate_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next 'Goto next line on an error
    
    'If enter is pressed locate the char
    If KeyCode = vbKeyReturn Then Call LocateChar
      
End Sub


Private Sub LocateChar()

    On Error Resume Next 'Goto next line on an error
    Dim Counter As Integer 'for loops
    
    'Loop for all chars
    For Counter = lblChar.LBound To lblChar.uBound
        If lblChar(Counter).Caption = txtLocate.Text Then
            'Flash the located char to draw attention
            Call FlashLocated(Counter)
            'Exit the loop as the char has been found
            Exit For
        'If all chars have noe been searched with no luck
        ElseIf Counter = lblChar.uBound And lblChar(Counter).Caption <> txtLocate.Text Then
            'Tell user
            MsgBox "Sorry the specified character could not be found", _
                vbOKOnly, "Not Found"
            'Set the focus back to the text box so the user cna have another go
            txtLocate.SetFocus
        End If
    'On to next label
    Next Counter
    
End Sub

Private Sub DrawFocus()
    Dim Focus As RECT
    
    DoEvents
    With Focus
        .Left = (lblChar(LastOn).Left + 20) / Screen.TwipsPerPixelX
        .Top = (lblChar(LastOn).Top + 20) / Screen.TwipsPerPixelY
        .Right = (.Left + (lblChar(LastOn).Width / Screen.TwipsPerPixelX)) - 2
        .Bottom = (.Top + (lblChar(LastOn).Height / Screen.TwipsPerPixelY)) - 2
        Call DrawFocusRect(picCharContainer.hdc, Focus)
    End With
    
End Sub
