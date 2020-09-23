VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bebo Winsock Login"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "UserID"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtPW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Password"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2670
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
Winsock1.Close
Call Winsock1.Connect("bebo.com", 80)
End Sub

Private Sub Winsock1_Connect()
'Connected, Post Form Data
Dim ExtraData As String, Pack As String
ExtraData$ = "FriendsMemberId=&FriendsChecksumNbr=&InviteRecipientId=&InviteChecksumNbr=&AppUrl=&Page=&QueryString=&Domain=&api_key=&AppId=&next=&canvas=null&v=&EmailUsername=" & txtID & "&Password=" & txtPW & "&SignIn=Sign+In+%3E"
Pack$ = "POST /SignIn.jsp HTTP/1.1" & vbCrLf
Pack$ = Pack$ & "Host: secure.bebo.com" & vbCrLf
Pack$ = Pack$ & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.0.3705)" & vbCrLf
Pack$ = Pack$ & "Accept: text/xml,application/xml,application/xhtml+xml,text/html;q=0.9,text/plain;q=0.8,image/png,*/*;q=0.5" & vbCrLf
Pack$ = Pack$ & "Accept-Language: en-us,en;q=0.5" & vbCrLf
Pack$ = Pack$ & "Accept-Encoding: gzip,deflate" & vbCrLf
Pack$ = Pack$ & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & vbCrLf
Pack$ = Pack$ & "Keep-Alive: 300" & vbCrLf
Pack$ = Pack$ & "Connection: keep-alive" & vbCrLf
Pack$ = Pack$ & "Referer: http://www.bebo.com/SignIn.jsp" & vbCrLf
Pack$ = Pack$ & "Cookie: Gen=M; Age=103; Username=" & txtID.Text & "; sessioncreate=20080916155210; bdaysession=7315dbe12cf6b9e1367066589; BeboLangCode=us; Password=" & txtPW.Text & "autoturnedoff" & vbCrLf
Pack$ = Pack$ & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
Pack$ = Pack$ & "Content-Length: " & Len(ExtraData) & vbCrLf
Pack$ = Pack$ & "Cache-Control: no-cache" & vbCrLf & vbCrLf
Pack$ = Pack$ & ExtraData$
Call Winsock1.SendData(Pack$)
MsgBox "Sent Packet:" & vbCrLf & vbCrLf & Pack$
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock1.GetData Data
Text1.Text = Text1.Text & vbCrLf & vbCrLf & Data
MsgBox "Recieved Packet: " & vbCrLf & vbCrLf & Data
If InStr(Text1.Text, txtID.Text) Then
Me.Caption = "Logged In...."
Text1.Text = ""
Winsock1.Close
Else
Me.Caption = "Wrong Details"
Text1.Text = ""
Winsock1.Close
End If
End Sub

