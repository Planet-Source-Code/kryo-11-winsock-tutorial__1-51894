VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Winsock Tutorial"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Get HTML Source"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdGetSource 
         Caption         =   "GO"
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtGetSource 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "URL :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5880
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Proxy (optional)"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   6135
      Begin VB.CommandButton cmdProxy 
         Caption         =   "Validate"
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtProxy 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Get Cookie"
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   6135
      Begin VB.TextBox txtCookie 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   600
         Width           =   5895
      End
      Begin VB.CommandButton cmdCookie 
         Caption         =   "Check For Cookie"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Window"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   6135
      Begin VB.TextBox txtData 
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle..."
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   6000
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------
'Check for cookie in the returned HTML code
'------------------------------------------
Private Sub cmdCookie_Click()

    txtCookie = GrabCookie(txtData) 'get cookie
    
End Sub

'-----------------------------------------
'Format URL as create a connection request
'-----------------------------------------
Private Sub cmdGetSource_Click()

    Dim CurrentServer As String
    
    CurrentServer = GetServer(txtGetSource) 'get server from URL
    
    'if the URL is blank or there is not server then do not continue
    If CurrentServer = "" Then Exit Sub
    
    'Setup the winsock and connect
    With Winsock1
    
        .Close 'if a connection if being established
               'or is already established close it
        
        'if there is a proxy in the "txtProxy" text box then use it
        If SplitProxy(txtProxy, ":", Proxy) <> "Error" Then
            
            .RemoteHost = SplitProxy(txtProxy, ":", Proxy) 'The proxy server
            .RemotePort = SplitProxy(txtProxy, ":", Port) 'The proxy port
            
        Else
        
            .RemoteHost = CurrentServer 'The server (Ex: kyro-genics.com)
            .RemotePort = 80 '80 is standard port for ALL HTML requests
            
        End If
        
        .Connect 'connect with the server and port
            
    End With
    
End Sub

'--------------------------------------------------------
'Check to make sure you are using a valid proxy (Generic)
'--------------------------------------------------------
Private Sub cmdProxy_Click()

    Dim Validate As String
    
    'call the function from the module
    Validate = SplitProxy(txtProxy, ":", Proxy)
    
    If Validate = "Error" Then 'if invalid
    
        lblStatus = "Not a valid proxy." 'status message (not important)
        
    Else 'valid
    
        lblStatus = "Valid proxy." 'status message (not important)
        
    End If
    
End Sub

'--------------------------------------------
'Kill active Winsock connections upon exiting
'--------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = True
    
    Winsock1.Close
    
    End
    
End Sub

'---------------------------
'Send a request for the HTML
'---------------------------
Private Sub Winsock1_Connect()

    'status message (not important)
    lblStatus = "Connecting to " & Winsock1.RemoteHost & " on Port 80)"
    
    txtData = Empty 'clear the Data Window
    
    Winsock1.SendData SendHeader(txtGetSource) 'Send request
    
End Sub

'----------------------------------------------
'Get and Add the Data (HTML) to the Data Window
'----------------------------------------------
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    'status message (not important)
    lblStatus = "Receiving data from " & Winsock1.RemoteHost
    
    On Error Resume Next 'on error go to next line of code
    
    Dim ReturnedHTML As String
    
    Winsock1.GetData ReturnedHTML 'Data (HTML) returned by the server
    
    txtData = txtData & ReturnedHTML 'Add the data to the Data Window
    
    txtData = FixFeed(txtData) 'fix the invalid line feed characters in the HTML
    
End Sub

'-----------------------------------------------
'If an Error occurs display error in data window
'-----------------------------------------------
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    'status message (not important)
    lblStatus = "Idle..."

    Winsock1.Close 'Kill the connection
    
    txtData = "Error : " & Description & vbCrLf 'description of the error
    txtData = txtData & "Error ID : " & Scode 'Error id
    
End Sub
