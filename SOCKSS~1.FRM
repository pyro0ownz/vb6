VERSION 5.00
Begin VB.Form Pyro0ownz 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ×º°”˜§º¢kz 5 §¢åñ Ðå† ßî†¢H`”°º×"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   Icon            =   "SOCKSS~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Picture         =   "SOCKSS~1.frx":628A
   ScaleHeight     =   3585
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00800080&
      Caption         =   "Clear"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3000
      Width           =   615
   End
   Begin VB.PictureBox go 
      Height          =   480
      Left            =   6000
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   23
      Top             =   2040
      Width           =   1200
   End
   Begin VB.PictureBox goto 
      Height          =   480
      Left            =   5640
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   21
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   960
      MousePointer    =   2  'Cross
      TabIndex        =   17
      Text            =   "1"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      MousePointer    =   2  'Cross
      TabIndex        =   16
      Text            =   "1"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   600
      MousePointer    =   2  'Cross
      TabIndex        =   15
      Text            =   "95"
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00800080&
      Caption         =   "Save"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00800080&
      Caption         =   "Stop"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00800080&
      Caption         =   "Scan"
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1425
      Left            =   3120
      MousePointer    =   2  'Cross
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox sock 
      Height          =   480
      Index           =   0
      Left            =   5280
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   24
      Top             =   1440
      Width           =   1200
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1920
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Text            =   "1080"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3240
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Text            =   "0.01"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2520
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Text            =   "100"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Text            =   "200"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FF80&
      BorderStyle     =   2  'Dash
      X1              =   2760
      X2              =   2760
      Y1              =   3120
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FF80&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   120
      Y1              =   3120
      Y2              =   3600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FF80&
      BorderStyle     =   2  'Dash
      X1              =   2640
      X2              =   120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Version Two With Proxy Counter "
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Caption         =   "Proxies Found"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label counter1 
      BackColor       =   &H80000007&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   1560
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   3000
      X2              =   3240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   1680
      X2              =   1920
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      FillColor       =   &H0080FF80&
      Height          =   495
      Left            =   2880
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Pyro0ownz"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Status"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label stat 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "  Port"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Sockets"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Interval"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Range"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "Pyro0ownz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'first socks 5 scanner by pyro'
'this is a simple source i learned from'
'it adds anything with an open port to listbox'
'even if its dead and it gets a reply it adds'
'will update to where it only adds green to the list'

Option Explicit
Dim a2, a3 As Boolean
Dim location As String
Private list As String
Dim proxy(0 To 100) As String
Dim i As Integer
Private Sub Form_Load()
On Error Resume Next
MsgBox "Yup this my first scanner And the first prog i coded myself"
MsgBox "This is the second version it now has a proxy counter with other stuff added"
MsgBox "www.myspace.com/imbackshhh Hope you enjoy Peace"
    For i = 1 To 100
        Load sock(i)
    Next i
End Sub
Private Sub Form_Unload(Cancel As Integer)
Exit Sub
End Sub
Private Sub Command1_Click()
    Command1.Enabled = False
      On Error Resume Next
        For i = 0 To Text6
         If Text6 = 0 Then Exit Sub
           sock(i).Close
             proxy(i) = Text1 & "." & Text2 & "." & Text3 & "." & Text4
              sock(i).Connect Text1 & "." & Text2 & "." & Text3 & "." & Text4, Text5
               stat = "Pyro0ownz: " & proxy(i) & ":" & Text5
                counter1 = List1.ListCount
                 If Text3 = "255" Then
                  Text3 = "0"
                   a2 = True
                   End If
                  If a2 = True Then
                 Text3.Text = Text3.Text + 1
                a2 = False
               End If
              If Text3 = "255" Then
              Text3 = "0"
             Command1.Enabled = True
            MsgBox "Scan finished"
           Call Pause(10000)
          a2 = False
         End If
          If Text4 = "255" Then
           Text4 = "0"
            a2 = True
             End If
              If Text4 = "255" Then
               Text4 = "0"
                a2 = True
                 End If
                  Text4.Text = Text4.Text + 1
                   Call Pause(Text7)
                    Next i
                   If Command2.Caption = "Stop" Then
                  Call Pause(1)
                Command1.Value = True
              End If
            End Sub
Private Sub Command2_Click()
On Error Resume Next
         If Command2.Caption = "Stop" Then
          Command2.Caption = "Reset"
    Else
         Command1.Enabled = True
        Command2.Caption = "Stop"
       End If
End Sub
Private Sub Command3_Click()
    On Error GoTo pyro
       Dim Nbr As Long
    
         With go
        .DialogTitle = "Save Proxy"
        .Filter = "Text files (*.txt)|*.txt|all files or custom(*.*)|*.*"
        .ShowSave
        Open .FileName For Output As #1
        For Nbr = 0 To List1.ListCount - 1
            Print #1, List1.list(Nbr)
        Next Nbr
        Close #1
    End With
    Exit Sub

pyro:
    Exit Sub
    End Sub
Private Sub Command4_Click()
On Error Resume Next
List1.Clear
counter1 = 0
MsgBox "Your List Has Been Cleared"
End Sub
Private Sub sock_Connect(Index As Integer)

On Error Resume Next

    List1.AddItem (proxy(Index) & ":" & Text5)
    
End Sub
Public Sub Pause(interval)
On Error Resume Next

Dim X

    X = Timer
    Do While Timer - X < Val(interval)
        DoEvents
    Loop

End Sub
Private Sub Text1_Change()
On Error Resume Next
If Text1 = "" Then
Text1 = "0"
End If
If Text1 > 255 Then
Text1 = 0
MsgBox "255 is max number"
End If
' if any number is greater then 255 it goes to zero and a message box pops up'
End Sub
Private Sub Text2_Change()
On Error Resume Next
If Text2 = "" Then
Text2 = "0"
End If
If Text2 > 255 Then
Text2 = "0"
MsgBox "255 is max number"
End If
' if any number is greater then 255 it goes to zero and a message box pops up'
End Sub
Private Sub Text3_Change()
On Error Resume Next
If Text3 = "" Then
Text3 = "0"
End If
If Text3 > 255 Then
Text3 = "0"
MsgBox "255 is max number"
End If
' if any number is greater then 255 it goes to zero and a message box pops up'
End Sub
Private Sub Text4_Change()
On Error Resume Next
If Text4 = "" Then
Text4 = "0"
End If
If Text4 > 255 Then
Text4 = "0"
MsgBox "255 is max number"
End If
' if any number is greater then 255 it goes to zero and a message box pops up'
End Sub
Private Sub Text6_Change()
On Error Resume Next
If Text6 > 100 Then
MsgBox "Can Only Load 100 Sockets"
End If
If Text6 > 101 Then
Text6 = 100
End If
End Sub
