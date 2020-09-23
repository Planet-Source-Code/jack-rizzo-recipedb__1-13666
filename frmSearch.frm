VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search For Results"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TheDB As String  'data base path
Public WS1 As Workspace 'dao workspace object
Public DB1 As Database  'dao data base object
Public RS1 As Recordset 'dao recordset object

Private Sub Command1_Click()
'
' unload the form
'
Unload Me
End Sub

Private Sub Form_Load()
'
' the list box is filled from the main form
'
'set the database path

TheDB = App.Path & "\Recipes.mdb"
'
' place the form at the top right of the screen
'
Me.Left = Screen.Width - Me.Width
Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
RS1.Close  ' close the record set
DB1.Close  'close the data base
WS1.Close  'close the workspace

End Sub

Private Sub List1_Click()
'
' Set workspace, database and recordset objects
'
Set WS1 = DBEngine.Workspaces(0)
Set DB1 = WS1.OpenDatabase(TheDB, True)
Set RS1 = DB1.OpenRecordset("SELECT * FROM recipes WHERE Name = '" & List1.List(List1.ListIndex) & "';")
'
' Make sure we find the record
'
If RS1.RecordCount = 0 Then
   MsgBox "Can not Find " & List1.List(List1.ListIndex)
   RS1.Close
   DB1.Close
   Exit Sub
End If
'
' set the main form with new data
'
frmMain.Text1.Text = RS1("Name")
frmMain.Text2.Text = RS1("Author")
frmMain.Text3.Text = RS1("Ingredients")
frmMain.Text4.Text = RS1("Instructions")
RS1.Close  ' close the record set
DB1.Close  'close the data base
'
' reset main form buttons
'
frmMain.Command1.Enabled = False
frmMain.Command2.Enabled = True
frmMain.Command3.Enabled = True
frmMain.Command5.Enabled = True

End Sub
