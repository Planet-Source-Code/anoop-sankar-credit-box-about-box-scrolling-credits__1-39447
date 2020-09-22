VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utopia App - Credits"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Website"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   113
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "(c) 2002 Anoop Sankar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "www.smilehouse.cjb.net anoop@smilehouse.cjb.net"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------
'       Copyright 2002, Anoop Sankar
'You may modify and use this box in your programs. If
'you find it suitable, you could (not compulsory) add
'my name to the credits box. But if you are giving out
'the source code, please do not remove this message. If
'you modified something, put your name below..
'
'Orginal Code : Anoop Sankar
'Modified by  : No one so far
'
'
'Last Update : May 7,2002
'Visit www.smilehouse.cjb.net for more source code
'-----------------------------------------------------

Private Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Public CrdMessage As String


Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    Website "http://www.smilehouse.cjb.net"
End Sub

Private Sub Form_Load()
        
    Show
    ReadData
    MainLoop
    
End Sub

Private Sub ReadData()
    'read data.txt
    
    Dim strLine As String
    
    Open App.Path & "\data.txt" For Input As #1
            
        While Not EOF(1)            'loop till end of file
            Line Input #1, strLine  'read line by line
            
            'next step appends the line from file and a carriage
            'return to the crdmessage string
            CrdMessage = CrdMessage & strLine & vbCrLf
        Wend
        
    Close #1
   
End Sub

Private Sub MainLoop()
    
    For j = 1 To Len(CrdMessage)
        'print letter by letter
        Text1.Text = Text1.Text & Mid(CrdMessage, j, 1)
        Text1.SelStart = Len(Text1.Text)
        
        DoEvents        'do other events
        Sleep (200)     'delay
    Next j

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End

End Sub

Public Sub Website(URL)
    ShellExecute 0&, vbNullString, URL, vbNullString, _
    vbNullString, SW_SHOWNORMAL
End Sub

