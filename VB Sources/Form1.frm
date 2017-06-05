VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   2
      Text            =   "C:\Users\spertus\Pictures\tour eiffel\toureiffel.jpg"
      Top             =   960
      Width           =   7095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   1
      Text            =   "Image sent from VB !!"
      Top             =   2640
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Toast"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5280
      TabIndex        =   0
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Notifications title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Path to a local image:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Local Error GoTo Fin
    
    Dim toastService As New UwpWrapper.Notifications
    toastService.LaunchNotificationCustom Text1.Text, Text2.Text

Fin:

End Sub


