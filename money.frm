VERSION 5.00
Begin VB.Form Money 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Money"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************

Dim cOnn As ADODB.Connection
Dim cOmd As ADODB.Command
Dim rs As ADODB.Recordset

'***************

Dim A As Boolean
Dim b As String
Dim mYdSn As String
Dim strConn As String

Private Sub Form_Load()

' Check for existing dsn_connection
    
A = checkWantedAccessDSN("money") ' Dns_connection_module

        If A = True Then

            ' If exist then goto Connection_information
    
        Else

            ' If Not, Make it....!!

            Call createAccessDSN(App.pAth & "\mdb\money", "money")

        End If


'Get Connection_information

    Call inIgEt("System Dsn", "Name")
        strConn = "Data Source=" & mYdSn
        
'Set Connection
    
    Set cOnn = New ADODB.Connection
    Set cOmd = New ADODB.Command
    Set rs = New ADODB.Recordset
 
'Open connection

cOnn.ConnectionString = strConn
cOnn.Open strConn
cOmd.ActiveConnection = cOnn

rs.Open "money1", cOnn, adOpenDynamic, adLockOptimistic

Text1.Text = rs!Id
Text2.Text = rs!test

End Sub

Private Sub inIgEt(lpAppName As String, lpKeyName As String)

Dim X As Long
Dim Temp As String * 500
Dim lpDefault As String, lpFileName As String

lpFileName = App.pAth & "\ini\money.ini"
lpDefault = lpFileName
X = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, Temp, Len(Temp), lpFileName)

mYdSn = X

If X = 0 Then
   
    Beep

Else
    
   mYdSn = Trim(Temp)

End If

End Sub

Private Sub inIwrIte()

Dim lpAppName As String, lpFileName As String, lpKeyName As String, lpString As String
Dim U As Long

lpFileName = App.pAth & "\ini\bin.ini"
U = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
If U = 0 Then
Beep
MsgBox "oooooops"
End If

End Sub



