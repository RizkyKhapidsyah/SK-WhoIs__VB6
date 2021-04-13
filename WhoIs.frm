VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form WhoIs_Form 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whois Using MS WinSock Control "
   ClientHeight    =   6060
   ClientLeft      =   1155
   ClientTop       =   2355
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox WhoIs_Server 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   6
      Text            =   "WhoIs Server List"
      Top             =   120
      Width           =   3915
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   6360
      Top             =   5340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox WhoIs_Response 
      Height          =   4815
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8493
      _Version        =   393217
      BackColor       =   65535
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"WhoIs.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CLEAR_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1155
   End
   Begin VB.CommandButton Send_Query_Button1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Send Domain Query"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   5640
      Width           =   4215
   End
   Begin VB.TextBox Domain_Name 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "freevbcode.com"
      Top             =   5340
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "WhoIs Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   1110
   End
   Begin VB.Label Input_Label 
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Domain Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   840
      TabIndex        =   3
      Top             =   5400
      Width           =   1125
   End
End
Attribute VB_Name = "WhoIs_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Private Sub Form_Load()
' What to do when this program starts up

' Define a list of alternative WhoIs servers
  With WhoIs_Server
      .AddItem " whois.opensrs.net "
      .AddItem " whois.networksolutions.com "
      .AddItem " whois.nic.gov "
      .AddItem " rs.internic.net "
      .AddItem " whois.ripe.net "
      .AddItem " whois.arin.net "
      .AddItem " whois.apnic.net "
      .AddItem " whois.aunic.net "
      .ListIndex = 0
  End With
      
  End Sub

' ***********************************************************
  Private Sub Send_Query_Button1_Click()
' This code is executed when the [Send Query] button is clicked on.
' It uses the Network Solutions database for USA domains.

  Dim Selected_WhoIs_Server As String

  Input_Label.Caption = ""
  WhoIs_Response = ""
  
' Initialize Winsock prior to attempting connection
  Winsock.Close
  Winsock.LocalPort = 0

' Get address of selected WhoIs server
  Selected_WhoIs_Server = Trim(WhoIs_Server.Text)

' Connect to selected WHOIS server database, port 43
  Winsock.Connect Selected_WhoIs_Server, 43

  End Sub

  Private Sub CLEAR_Button_Click()
' Clear out the WHOIS info display and domain n

  WhoIs_Response = ""
  
  End Sub


  Private Sub Winsock_Connect()
' Connect WinSock and send domain name query

  If Trim(Domain_Name) = "" Then
     WhoIs_Response = " No domain name was entered."
     Beep
     Exit Sub
  End If
  
  Winsock.SendData Trim(Domain_Name) & vbCrLf
  
  End Sub

  Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
' Handle the incoming Winsock data stream

  Dim WhoIs_Data As String
  
  On Error GoTo ERROR_HANDLER ' Set error trap
  
  Winsock.GetData WhoIs_Data
  Input_Label.Caption = Input_Label.Caption & WhoIs_Data
  WhoIs_Response = Input_Label.Caption
  Exit Sub
  
' If some kind of error occurs, then display info about it.
ERROR_HANDLER:
  WhoIs_Response = Error
  End Sub

