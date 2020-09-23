VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SetLocalTime API"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   3165
   Begin VB.TextBox txtDay 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMonth 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtYear 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtMinute 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   2160
   End
   Begin VB.TextBox txtHour 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SetLocalTime API"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Day:"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Month:"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Year:"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Minute:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Hour:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim strTime As String       '-- This simulates our UDT
    Dim NewDate As Date
    Dim NewTime As Date
        
    '-- Set new Date/Time
    NewDate = DateSerial(txtYear.Text, txtMonth.Text, txtDay.Text)
    NewTime = TimeSerial(txtHour.Text, txtMinute.Text, 0)
    
    '-- Set new Time UDT
    strTime = SYSTEMTIME(Year(NewDate), Month(NewDate), Weekday(NewDate), _
        Day(NewDate), Hour(NewTime), Minute(NewTime), 0, 0)
    
    '-- Call API to Set to time
    Call SetLocalTime(strTime)
    
End Sub
Public Function SYSTEMTIME( _
    wYear As Integer, _
    wMonth As Integer, _
    wDayOfWeek As Integer, _
    wDay As Integer, _
    wHour As Integer, _
    wMinute As Integer, _
    wSecond As Integer, _
    wMilliseconds As Integer) As String
    
    
    '-- This function sets up the UDT in VBCE
    '-- with the help of
    '-- a few conversion functions
    
    SYSTEMTIME = _
        IntegerToBinaryString(wYear) & _
        IntegerToBinaryString(wMonth) & _
        IntegerToBinaryString(wDayOfWeek) & _
        IntegerToBinaryString(wDay) & _
        IntegerToBinaryString(wHour) & _
        IntegerToBinaryString(wMinute) & _
        IntegerToBinaryString(wSecond) & _
        IntegerToBinaryString(wMilliseconds)
End Function



Private Sub Form_Load()
   ' Me.Width = Screen.Width
    'Me.Height = Screen.Height
        
    txtYear.Text = Year(Now)
    txtMonth.Text = Month(Now)
    txtDay.Text = Day(Now)
    txtHour.Text = Hour(Now)
    txtMinute.Text = Minute(Now)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    App.End
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
    lblTime.Caption = Now
End Sub
