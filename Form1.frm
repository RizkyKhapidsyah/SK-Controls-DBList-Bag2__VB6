VERSION 5.00
Object = "{8C04E108-27EB-11D3-902A-00805F4936F9}#8.0#0"; "DLISTBOX.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OCX 'R US"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin DListBoxProject.DListBox DListBox1 
      Height          =   2385
      Left            =   150
      TabIndex        =   6
      Top             =   210
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   4207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DGetListBoxItemDescription"
      Height          =   330
      Left            =   90
      TabIndex        =   5
      Top             =   3840
      Width           =   3165
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DGetListBoxItemNumber"
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   4725
      Width           =   3165
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DGetListBoxItemKey"
      Height          =   330
      Left            =   75
      TabIndex        =   3
      Top             =   4275
      Width           =   3165
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DPopulateWithoutKey"
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   3240
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DPopulateWithKey"
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   2790
      Width           =   3165
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Toggle Select All"
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   5175
      Width           =   3165
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   135
      X2              =   3285
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   150
      X2              =   3300
      Y1              =   3690
      Y2              =   3690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private conADODB As ADODB.Connection

Private Sub cmdSelectAll_Click()
    DListBox1.DSelectAll = Not DListBox1.DSelectAll
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass
    Dim bstaADODBa As Boolean
    
    With DListBox1
        .DConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                             "Data Source=" & App.Path & "\" & "TEST.MDB"
        .DSource = "qryNameAll"
        .DPopulateWithKeyValue
    End With

    Screen.MousePointer = vbDefault
End Sub


Private Sub Command2_Click()
    Screen.MousePointer = vbHourglass
    With DListBox1
        .DConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                             "Data Source=" & App.Path & "\" & "TEST.MDB"
        .DSource = "qryCodeAll"
        .DPopulateWithoutKeyValue
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
    Dim vKeyData() As Variant
    Dim lKeyCount As Long
    Dim lLoop As Long
    DListBox1.DGetListBoxItemKey vKeyData, lKeyCount
    If lKeyCount > 0 Then
       For lLoop = 0 To lKeyCount - 1
           MsgBox vKeyData(lLoop)
       Next lLoop
    End If
End Sub


Private Sub Command4_Click()
    Dim vKeyData() As Variant
    Dim lKeyCount As Long
    Dim lLoop As Long
    DListBox1.DGetListBoxItemNumber vKeyData, lKeyCount
    If lKeyCount > 0 Then
       For lLoop = 0 To lKeyCount - 1
           MsgBox vKeyData(lLoop)
       Next lLoop
    End If
End Sub


Private Sub Command5_Click()
    Dim vKeyData() As Variant
    Dim lKeyCount As Long
    Dim lLoop As Long
    DListBox1.DGetListBoxItemDescription vKeyData, lKeyCount
    If lKeyCount > 0 Then
       For lLoop = 0 To lKeyCount - 1
           MsgBox vKeyData(lLoop)
       Next lLoop
    End If
End Sub

Private Sub Form_Load()
    Set conADODB = New ADODB.Connection
    With conADODB
         .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\" & "TEST.MDB"
         .Open
    End With
    Command1_Click
End Sub


