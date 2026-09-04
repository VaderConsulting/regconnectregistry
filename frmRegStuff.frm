VERSION 5.00
Begin VB.Form frmRegStuff 
   Caption         =   "RegConnectRegistry Sample"
   ClientHeight    =   4110
   ClientLeft      =   6210
   ClientTop       =   3810
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5265
   Begin VB.TextBox txtDisplay 
      Height          =   285
      Left            =   60
      TabIndex        =   10
      Top             =   3120
      Width           =   4995
   End
   Begin VB.TextBox txtSubKey 
      Height          =   300
      Left            =   60
      TabIndex        =   7
      Text            =   "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
      Top             =   1905
      Width           =   5055
   End
   Begin VB.Frame frmKeys 
      Caption         =   "Select Top Level Key to Open"
      Height          =   945
      Left            =   45
      TabIndex        =   4
      Top             =   645
      Width           =   5010
      Begin VB.OptionButton optHKEY 
         Caption         =   "HKEY_USERS"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   570
         Width           =   2175
      End
      Begin VB.OptionButton optHKEY 
         Caption         =   "HKEY_LOCAL_MACHINE"
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin VB.TextBox txtRemMachName 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "c01900883"
      Top             =   210
      Width           =   2295
   End
   Begin VB.TextBox txtValueName 
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Text            =   "Siteloc"
      Top             =   2505
      Width           =   4995
   End
   Begin VB.CommandButton cmdQueryValue 
      Caption         =   "Query Value"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   3510
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Result:"
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   2865
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Value under SubKey you wish to query"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   2265
      Width           =   2730
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SubKey you wish to query:"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1665
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remote machine name:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "frmRegStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQueryValue_Click()
    
    Dim lRetVal As Long         'used to hold return value for all API calls
    
    Dim sRemMachName As String  'used by RegConnectRegistry
    Dim lTopLevelKey As Long    'used by RegConnectRegistry
    Dim lHKeyhandle As Long     'used by RegConnectRegistry & RegOpenKeyEx
    
    Dim sKeyName As String      'used by RegOpenKeyEx
    Dim lhkey As Long           'used by RegOpenKeyEx & RegQueryValueEx & RegCloseKey
    
    Dim sValueName As String    'used by RegQueryValueEx
    Dim vValue As String        'used by RegQueryValueEx
    
    sRemMachName = txtRemMachName.Text    'Gets user entered name of remote machine
    sKeyName = txtSubKey.Text             'Gets user entered subkey name
    sValueName = txtValueName.Text        'Gets user entered value name
    
    'Get user's choice of Top Level Keys by looping through array of option buttons
    'only two choices under Winnt
    Dim opt As OptionButton
    For Each opt In optHKEY
        If opt.Value = True Then
            Select Case opt.Index
                Case 0
                lTopLevelKey = HKEY_LOCAL_MACHINE
                Case 1
                lTopLevelKey = HKEY_USERS
            End Select
            Exit For
        End If
    Next opt

    'Get handle of a top level registry key on remote machine
    lRetVal = RegConnectRegistry(sRemMachName, lTopLevelKey, lHKeyhandle)
    'Debug.Print "RegConnectRegistry returned:  " & lRetVal
    'Debug.Print "lHKeyhandle = " & lHKeyhandle
    
    'Get handle of the key which contains the value you need to check
    lRetVal = RegOpenKeyEx(lHKeyhandle, sKeyName, 0, KEY_ALL_ACCESS, lhkey)
    'Debug.Print "RegOpenKeyEx returned: " & lRetVal
    'Debug.Print "lhkey = " & lhkey
    
    'Get the value
    lRetVal = QueryValueEx(lhkey, sValueName, vValue)
    'Debug.Print "QueryValueEx returned:  " & lRetVal
    'Debug.Print "vValue = " & vValue
    
    'Always a good idea to clean up
    RegCloseKey (lhkey)
    
    If lRetVal = 0 Then    'QueryValueEx succeeded
        'Display the value obtained
        txtDisplay.Text = vValue
    Else
        'Display error message
        Dim msg As String
        msg = "An Error occured, Return value = " & lRetVal
        txtDisplay.Text = msg
    End If
End Sub

