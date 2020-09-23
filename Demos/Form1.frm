VERSION 5.00
Object = "{798C3AED-5101-11D5-9278-0050FC0DD647}#93.1#0"; "COOLBUTTON.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CoolButton Demo"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Welcome to the CoolButton SpecialFX Button Demonstration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   5970
      Begin VB.Frame Frame2 
         Caption         =   "Features of the CoolButton SpecialFX Control"
         Height          =   5775
         Left            =   75
         TabIndex        =   1
         Top             =   180
         Width           =   5835
         Begin CoolButton.CoolCommand VOTE 
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   5100
            Width           =   5520
            _ExtentX        =   9737
            _ExtentY        =   661
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Please vote if you like this submission (Click Here to bring up Dialog)"
            Caption         =   "Please vote if you like this submission (Click Here to bring up Dialog)"
            BorderStyle3D   =   1
         End
         Begin CoolButton.CoolCommand CoolCommand9 
            Height          =   345
            Left            =   240
            TabIndex        =   13
            Top             =   3975
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Top Aligned"
            Caption         =   "Top Aligned"
            BackColor1      =   12632256
            BackColor2      =   0
            TextAlign       =   3
         End
         Begin CoolButton.CoolCommand CoolCommand7 
            Height          =   345
            Left            =   255
            TabIndex        =   11
            Top             =   3525
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Left Aligned"
            Caption         =   "Left Aligned"
            BackColor1      =   12632256
            BackColor2      =   0
            TextAlign       =   1
         End
         Begin CoolButton.CoolCommand CoolCommand6 
            Height          =   315
            Left            =   675
            TabIndex        =   9
            Top             =   2745
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   """Standard Style"" Thick Borders"
            Caption         =   """Standard Style"" Thick Borders"
            BackColor1      =   16761087
            BorderStyle3D   =   1
         End
         Begin CoolButton.CoolCommand CoolCommand5 
            Height          =   330
            Left            =   675
            TabIndex        =   8
            Top             =   2370
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Thin ""New Look"" Borders"
            Caption         =   "Thin ""New Look"" Borders"
            BackColor1      =   16761087
         End
         Begin CoolButton.CoolCommand CoolCommand3 
            Height          =   330
            Left            =   660
            TabIndex        =   6
            Top             =   1545
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Support for Cool Font FX"
            Caption         =   "Support for Cool Font FX"
         End
         Begin CoolButton.CoolCommand CoolCommand4 
            Height          =   330
            Left            =   660
            TabIndex        =   5
            Top             =   1185
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Cool New Mouse FX"
            Caption         =   "Cool New Mouse FX"
            BackColor1      =   12648447
         End
         Begin CoolButton.CoolCommand CoolCommand2 
            Height          =   345
            Left            =   660
            TabIndex        =   4
            Top             =   810
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Fully User Defineable Highlight And Normal Color Gradient Settings"
            Caption         =   "Fully User Defineable Highlight And Normal Color Gradient Settings"
            BackColor1      =   8438015
         End
         Begin CoolButton.CoolCommand CoolCommand1 
            Height          =   330
            Left            =   660
            TabIndex        =   3
            Top             =   450
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Cool New Ultra Fast Gradient FX"
            Caption         =   "Cool New Ultra Fast Gradient FX"
         End
         Begin CoolButton.CoolCommand CoolCommand8 
            Height          =   360
            Left            =   1950
            TabIndex        =   12
            Top             =   3540
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   635
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Right Aligned"
            Caption         =   "Right Aligned"
            BackColor1      =   12632256
            BackColor2      =   0
            TextAlign       =   2
         End
         Begin CoolButton.CoolCommand CoolCommand10 
            Height          =   360
            Left            =   1950
            TabIndex        =   14
            Top             =   3960
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   635
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Bottom Aligned"
            Caption         =   "Bottom Aligned"
            BackColor1      =   12632256
            BackColor2      =   0
            TextAlign       =   4
         End
         Begin CoolButton.CoolCommand CoolCommand11 
            Height          =   345
            Left            =   1050
            TabIndex        =   15
            Top             =   4365
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   609
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   1
            checkcaption    =   "Centre Aligned"
            Caption         =   "Centre Aligned"
            BackColor1      =   12632256
            BackColor2      =   0
         End
         Begin VB.Label Label4 
            Caption         =   "The Above Alignment Settings can be Used for Icons Too !"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   255
            TabIndex        =   16
            Top             =   4740
            Width           =   5505
         End
         Begin VB.Label Label3 
            Caption         =   "Icon and Text AutoAlign Support"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   10
            Top             =   3255
            Width           =   5145
         End
         Begin VB.Label Label2 
            Caption         =   "New Border Styles:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   7
            Top             =   2085
            Width           =   5145
         End
         Begin VB.Label Label1 
            Caption         =   "General Features:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   2
            Top             =   195
            Width           =   5145
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
End Sub

Private Sub VOTE_Click()
VDialog.Show 1
End Sub
