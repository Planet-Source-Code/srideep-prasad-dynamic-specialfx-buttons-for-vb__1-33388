VERSION 5.00
Begin VB.Form VDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Thanks for appreciating this code...."
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   390
      Left            =   3195
      TabIndex        =   3
      Top             =   1245
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "How much to you think, this code is worth ?"
      Height          =   1065
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   4395
      Begin VB.OptionButton Good 
         Caption         =   "&Good"
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   615
         Width           =   2070
      End
      Begin VB.OptionButton Excellent 
         Caption         =   "&Excellent"
         Height          =   270
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2565
      End
   End
End
Attribute VB_Name = "VDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Const SW_SHOWNORMAL = 1
Const CodeID = 33388


Private Sub Ok_Click()
If Excellent.Value = True Then
    GotoURL ("http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=" & Trim$(Str$(CodeID)) & "&optCodeRatingValue=5")
Else
    GotoURL ("http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=" & Trim$(Str$(CodeID)) & "&optCodeRatingValue=4")
End If
MsgBox "Thanks for the vote ! Please e-mail me at srideepprasad@yahoo.com in case of any problems", vbExclamation Or vbOKOnly, "Thanks !"
Unload Me
End Sub



Sub GotoURL(URL As String)
    Dim Res As Long
    Dim TFile As String, Browser As String, Dum As String
    
    TFile = App.Path + "\test.htm"
    Open TFile For Output As #1
    Close
    Browser = String(255, " ")
    Res = FindExecutable(TFile, Dum, Browser)
    Browser = Trim$(Browser)
    
    If Len(Browser) = 0 Then
        MsgBox "Cannot find browser"
        Exit Sub
    End If
    
    Res = ShellExecute(Me.hwnd, "open", Browser, URL, Dum, SW_SHOWNORMAL)
    If Res <= 32 Then
        MsgBox "Cannot open web page"
        Exit Sub
    End If
End Sub


