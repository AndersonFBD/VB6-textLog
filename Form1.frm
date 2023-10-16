VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H80000016&
      Caption         =   "Reset Log"
      Height          =   480
      Left            =   2895
      TabIndex        =   5
      Top             =   1020
      Width           =   990
   End
   Begin VB.Frame FraLog 
      Caption         =   "Log"
      Height          =   2205
      Left            =   180
      TabIndex        =   2
      Top             =   1710
      Width           =   4365
      Begin VB.TextBox txtLogger 
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         DragMode        =   1  'Automatic
         Height          =   1590
         Left            =   315
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   375
         Width           =   3705
      End
   End
   Begin VB.TextBox txtText1 
      Height          =   375
      Left            =   885
      TabIndex        =   1
      Top             =   255
      Width           =   3030
   End
   Begin VB.CommandButton cmdAperte 
      Caption         =   "Gravar linha"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   915
      TabIndex        =   0
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   315
      Width           =   420
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu OpenOption 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveOption 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu QuitOption 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Testing_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub cmdAperte_Click()
   txtLogger.Text = txtLogger.Text & Date$ + " " + Time$ + ": " + txtText1.Text & Constants.vbCrLf
   txtText1.Text = ""
End Sub

Private Sub cmdReset_Click()
   txtLogger.Text = ""
End Sub



Private Sub QuitOption_Click()
   End
End Sub

Private Sub txtLogger_Change()
   With txtLogger
      .SelStart = Len(.Text) 'seleciona o texto contido no log
      .SelLength = 0 'ultimo caractere na seleção do log
   End With
End Sub

Private Sub txtText1_Change()
If txtText1.Text = "" Then
         cmdAperte.Enabled = False
      Else
      
          cmdAperte.Enabled = True

   End If
End Sub


