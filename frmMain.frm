VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jack's Lil Recipe Data Base"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "&Search For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00008000&
      Caption         =   "&Print Recipe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   360
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2655
      Left            =   7080
      TabIndex        =   9
      Top             =   3840
      Width           =   3135
      Begin VB.CommandButton Command4 
         BackColor       =   &H00008000&
         Caption         =   "&Clear Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00008000&
         Caption         =   "&Delete Recipe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00008000&
         Caption         =   "&Update Recipe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00008000&
         Caption         =   "&Add Recipe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   7080
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   2895
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3600
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   2175
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Recipes in the Data Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Recipe Instructions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Recipe Ingredients:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Recipe Author:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Recipe Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WS As Workspace  'workspace for dao
Public DB As Database   'database for dao
Public RS As Recordset  'recordset for dao
Public SrchItem As String 'public item for search
Public TheDB As String   'data base path

Private Sub Command1_Click()
'
' Add a new record to the data base
'
' Check that we have all the info we need
'
If Text1.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
   MsgBox "Data input incomplete!"
   Exit Sub
End If
'
' Open the data base
'
Set DB = WS.OpenDatabase(TheDB, True)
'
' Open the record set
'
Set RS = DB.OpenRecordset("recipes", dbOpenTable)
'
' Tell DAO that this is add a new record
'
RS.AddNew
RS("Name") = Text1.Text  'set the recipe name
'
' If no Author was entered then make it unknown
' else set the author name
'
If Text2.Text = "" Then
   RS("Author") = "Unknown"
Else
   RS("Author") = Text2.Text
End If
RS("Ingredients") = Text3.Text   'set the ingridents
RS("Instructions") = Text4.Text  'set the instructions
RS.Update                        'add the record
RS.Close                         'Close the recordset
DB.Close                         'close the data base
List1.AddItem Text1.Text         'add the recipe to the list
End Sub

Private Sub Command2_Click()
'
' update an existing record in the data base
'
' Open the data base
'
Set DB = WS.OpenDatabase(TheDB, True)
'
' find the record using SQL Statment
'
Set RS = DB.OpenRecordset("SELECT * FROM recipes WHERE Name = '" & List1.List(List1.ListIndex) & "';")
'
' Make sure we have the record if not put out error
' and then close the recordset and database
'
If RS.RecordCount = 0 Then
   MsgBox "Can not Find " & List1.List(List1.ListIndex)
   RS.Close
   DB.Close
   Exit Sub
End If
'
' Tell DAO this is an edit
'
RS.Edit
'
' re set all the data information
'
RS("Name") = Text1.Text
RS("Author") = Text2.Text
RS("Ingredients") = Text3.Text
RS("Instructions") = Text4.Text
'
' Update the record then close recordset
' and database
'
RS.Update
RS.Close
DB.Close
'
' Clear the info from the display
'
Command4_Click
End Sub

Private Sub Command3_Click()
'
' Delete an existing recipe record
'
' Open the data base
'
Set DB = WS.OpenDatabase(TheDB, True)
'
' Find the record using SQL Statment
'
Set RS = DB.OpenRecordset("SELECT * FROM recipes WHERE Name = '" & List1.List(List1.ListIndex) & "';")
'
' Make sure we have a record
'
If RS.RecordCount = 0 Then
   MsgBox "Can not Find " & List1.List(List1.ListIndex)
   RS.Close
   DB.Close
   Exit Sub
End If
'
' Delete the record
' Then close the recordset and database
'
RS.Delete
RS.Close
DB.Close
'
' Remove the recipe from the list
'
List1.RemoveItem List1.ListIndex
'
' Clear the data from the display
Command4_Click
End Sub

Private Sub Command4_Click()
Dim i As Integer
'
' Clear data from the display
'
' Deselect any selected list item
'
For i = 0 To List1.ListCount - 1
   If List1.Selected(i) Then
      List1.Selected(i) = False
      Exit For
   End If
Next i
'
'clear the text boxes
'
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
'
'reset command buttons
'
Command2.Enabled = False
Command3.Enabled = False
Command1.Enabled = True
Command5.Enabled = False
'
' Move the cursor to the name text box
'
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Dim ix As Long
On Error Resume Next        'turn errors off
CD1.CancelError = True      'turn cancel on
CD1.Flags = 0               'clear flags
CD1.ShowPrinter             'show printer dialog
If Err <> 0 Then            'check for errors
    MsgBox Error(Err)        'if so display (prob cancel)
    Exit Sub                 'exit the sub
End If
If CD1.Flags And 32 = 32 Then  'check if print to file selected
    CD1.ShowSave                'yes, so show save as dialog
    If Err <> 0 Then            'check for cancel or error
       MsgBox Error(Err)        'yes display
       Exit Sub                 'exit sub
    End If
    ix = FreeFile               'get a file handle
    Open CD1.FileName For Output As #ix  'open file
    Print #ix, "     Recipe for " & Text1.Text
    Print #ix,
    Print #ix, "     Author: " & Text2.Text
    Print #ix,
    Print #ix, "     Ingredient List:"
    Print #ix,
    Print #ix, Text3.Text
    Print #ix,
    Print #ix, "     Preperation Instructions:"
    Print #ix,
    Print #ix, Text4.Text
    Close #ix                  'close file
Else
    Printer.Font = "Arial"    'set Arial font
    Printer.FontSize = "12"   'set font size for printer
    Printer.FontBold = True   'set to bold font
    Printer.Print "     Recipe for " & Text1.Text
    Printer.Print
    Printer.Print "     Author: " & Text2.Text
    Printer.Print
    Printer.Print "     Ingredient List:"
    Printer.Print
    Printer.Print Text3.Text
    Printer.Print
    Printer.Print "     Preperation Instructions:"
    Printer.Print
    Printer.Print Text4.Text
    Printer.EndDoc
End If
End Sub

Private Sub Command6_Click()
Dim i As Integer
'
' Search the name and ingreadients for a specific item
'
' Ask what to search from
'
SrchItem = InputBox("What are you Searching For?")
'
' if nothing there exit
'
If SrchItem = "" Then
   Exit Sub
End If
'
'open the database
'
Set DB = WS.OpenDatabase(TheDB, False)
'
' Find the records in the data base that match
'
Set RS = DB.OpenRecordset("SELECT * FROM recipes WHERE Name Like '*" & SrchItem & "*'" & " OR Ingredients Like '*" & SrchItem & "*';")
'
' If zero we have none
'
If RS.RecordCount = 0 Then
   MsgBox "Can not Find " & SrchItem
   RS.Close
   DB.Close
   Exit Sub
End If
'
' Move to last record to get proper record count
'
RS.MoveLast
'
' Move back to 1st record
'
RS.MoveFirst
'
' move through the record set and add the
' Recipe name to the Search Text box
'
For i = 1 To RS.RecordCount
   frmSearch.List1.AddItem RS("Name")
   RS.MoveNext
Next i
'
' Select the first entry in the list
'
frmSearch.List1.Selected(0) = True
'
' show the form as modal
' This forces the user to close the form
' before it proceeds further in the main
' form
frmSearch.Show vbModal
'
' Clear the data and reset the command buttons
'
Command4_Click
End Sub

Private Sub Form_Load()
Dim i As Integer
'
' Set the DB path
'
TheDB = App.Path & "\recipes.mdb"
'
' Center the form on the screen
'
Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2
'
' Check to see if we have a data base
' if not create it
'
If Dir(TheDB, vbNormal) = "" Then
   CreateDB TheDB
End If
'
'This sets a workspace for the database
'
Set WS = DBEngine.Workspaces(0)
'
'this opens the database
' (Second argument is True if open exclusive or False for multiuser)
'
Set DB = WS.OpenDatabase(TheDB, True)
'
'this opens a table inside the database
'
Set RS = DB.OpenRecordset("recipes", dbOpenTable)
'
' If no records in the data base exit
'
If RS.RecordCount = 0 Then
   Exit Sub
End If
'
' Move to the last record to set the record count
'
RS.MoveLast
'
' Move back to the first record
'
RS.MoveFirst
'
' Add the recipe names to the list box
'
For i = 1 To RS.RecordCount
   List1.AddItem RS("Name")
   RS.MoveNext
Next i
RS.Close     'close the record set
DB.Close     'close the database
'
'   Notice the workspace was not closed!
'   By allowing the workspace to remain open
'   subsequent data base opens occur much faster
'
' set the update, delete and print command buttons
' as unavailable.
'
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Turn error checking off
'
On Error Resume Next
RS.Close    'close the record set if open
DB.Close    'close the data base if open
WS.Close    'close the workspace
End         'end the program
End Sub

Private Sub List1_Click()
'
' this routine displays the recipe data when
' a list box item is clicked
'
' check for no items in the list box
' if so then exit the sub
'
If List1.ListIndex = -1 Then
   Exit Sub
End If
'
'  Open the data base
'
Set DB = WS.OpenDatabase(TheDB, True)
'
'  Get the recordset using an SQL Statment looking
'  for the recipe name
'
Set RS = DB.OpenRecordset("SELECT * FROM recipes WHERE Name = '" & List1.List(List1.ListIndex) & "';")
'
'  ensure we have a record returned
'
If RS.RecordCount = 0 Then
   MsgBox "Can not Find " & List1.List(List1.ListIndex)
   RS.Close
   DB.Close
   Exit Sub
End If
'
' Set the data values from those in
' the data base
'
Text1.Text = RS("Name")
Text2.Text = RS("Author")
Text3.Text = RS("Ingredients")
Text4.Text = RS("Instructions")
'
' Close the recordset and data base
'
RS.Close
DB.Close

Command1.Enabled = False   'turn add button off
Command2.Enabled = True    'turn Update Button on
Command3.Enabled = True    'turn Delete button on
Command5.Enabled = True    'turn Print button on

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'
'  This routine traps the Enter Key being pressed
'  and moves the focus to the next text box
'
'  It also checks to make sure there are no
' embeded "'" in the string that will confuse
' the DAO SQL record update.
'
If KeyAscii = 13 Then
   KeyAscii = 0
   Text2.SetFocus
End If
If Chr(KeyAscii) = "'" Then
    MsgBox "Single Quote is an Illegal Character!"
    KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'
'  This routine traps the Enter Key being pressed
'  and moves the focus to the next text box
'
If KeyAscii = 13 Then
   KeyAscii = 0
   Text3.SetFocus
End If
End Sub

Public Sub CreateDB(dbname As String)
'
' This subroutine creates the recipe data base
'
Dim DB As Database     'define the data base object
Dim WS As Workspace    'define workspace object
Dim TD As TableDef     'define table definintion object
Dim FD As Field        'define the field object
'
' Open up the workspace
'
Set WS = DBEngine.Workspaces(0)
'
' Create the Data base
'
Set DB = WS.CreateDatabase(dbname, dbLangGeneral)
'
' Create the table in the data base
'
Set TD = DB.CreateTableDef("Recipes")
'
' create the 4 fields in the table
'
Set FD = TD.CreateField("Name", dbText, 75)
FD.Required = True          'field is required
FD.AllowZeroLength = False  'do not allow zero length
TD.Fields.Append FD         'bind the field to the table
Set FD = TD.CreateField("Author", dbText, 75)
FD.Required = True
FD.DefaultValue = "Unknown"
FD.AllowZeroLength = False
TD.Fields.Append FD
Set FD = TD.CreateField("Ingredients", dbMemo)
FD.Required = True
FD.AllowZeroLength = False
TD.Fields.Append FD
Set FD = TD.CreateField("Instructions", dbMemo)
FD.Required = True
FD.AllowZeroLength = False
TD.Fields.Append FD
DB.TableDefs.Append TD 'bind the table to the database
DB.TableDefs.Refresh   'refresh the tables in the data base
DB.Close
WS.Close
End Sub
                    
                    
                    
                    

