VERSION 5.00
Begin VB.Form frmEmpMaster 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Master"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "frmEmpMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Caption         =   "Searc&h"
      Height          =   495
      Left            =   1320
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2535
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtID 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox txtAge 
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
         Width           =   600
      End
      Begin VB.TextBox txtSalary 
         Height          =   300
         Left            =   1320
         TabIndex        =   3
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   17
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   2040
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Clear All"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Update"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">|"
      Height          =   495
      Left            =   3130
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   607
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">"
      Height          =   495
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   607
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   607
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C0C0&
      Caption         =   "|<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   607
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add &New"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEmpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Contacts: tropicalwire@hotmail.com

''This program will help you to know how to
''add, delete, modify, update and view records in the
''database. This also cleary explains(helps) about the
''usage of ADO and also other useful basic functional
''codes

''This project, i have used ADO(ActiveX Data Objects
''Library 2.5)
''If you do not have then use atleast 2.1
''This you can change by selecting the Project Menu and
''then References.

''The name of the database is Employee
''The name of the database table is Master

'________________________________________________________


'This is for connecting to the database, i am
'declaring the connection as CON
Dim CON As Connection

'This is am using for moving
''in the table (move first, move next, move previous
''move last)
Dim WW As New Recordset
Private Sub cmdAddNew_Click()
    ''Here i am clearing all the fields and setting the focus
    ''to the txtID field, before that i am enabling the field
    txtID.Text = ""
    txtName.Text = ""
    txtAge.Text = ""
    txtSalary.Text = ""
    txtID.Enabled = True
    txtID.SetFocus
    cmdSave.Enabled = True
End Sub

Private Sub cmdClear_Click()
    ''Here i am clearing all the fields if you like at any time
    ''I am calling this cmdClear_Click in all the places
    ''instead of repeating the code.
    ''Here u can also use your own function.
    txtID.Text = ""
    txtName.Text = ""
    txtAge.Text = ""
    txtSalary.Text = ""
    txtID.Enabled = True
    txtID.SetFocus
End Sub
Private Sub cmdDelete_Click()
''Before Deleting, i am checking for the txtID
''and other fields are empty(Actually checking the
''txtID is enough

If txtID.Text <> "" And txtName.Text <> "" And txtAge.Text <> "" And txtSalary.Text <> "" Then
    ''See how i am using the confirmation dialog in the msgbox
    
    If MsgBox("Are you sure to delete Employee ID: " & txtID.Text & "?", vbQuestion + vbYesNoCancel, "Confirm Delete") = vbYes Then
    Dim rs As New Recordset
    rs.Open "Select * from Master where ID=" & txtID.Text & " and Name='" & txtName.Text & "' and Age=" & txtAge.Text & " and Salary=" & txtSalary.Text, CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount <> 0 Then
            rs.Delete
            rs.Close
            MsgBox "Records deleted succesfully", vbInformation
            Call cmdClear_Click
        Else
            MsgBox "Invalid Record", vbCritical, "Invalid"
        End If
    Else
    End If
Else
    MsgBox "Field/s found empty", vbCritical, "Empty Field/s"
End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
''This is done to move to your first record in
''the database table Master
On Error Resume Next
txtID.Enabled = False
If WW.RecordCount <> 0 Then
    WW.Fields.Refresh
    WW.MoveFirst
    txtID.Text = WW![ID]
    txtName.Text = WW![Name]
    txtAge.Text = WW![Age]
    txtSalary.Text = WW![Salary]
End If
End Sub

Private Sub cmdLast_Click()
''This is done to move to your last record in
''the database table Master

On Error Resume Next
txtID.Enabled = False
If WW.RecordCount <> 0 Then
    WW.Fields.Refresh
    WW.MoveLast
    txtID.Text = WW![ID]
    txtName.Text = WW![Name]
    txtAge.Text = WW![Age]
    txtSalary.Text = WW![Salary]
End If

End Sub



Private Sub cmdNext_Click()
''This is done to move to your next record in
''the database table Master

On Error Resume Next
txtID.Enabled = False
If WW.RecordCount <> 0 Then
    WW.Fields.Refresh
    WW.MoveNext
    txtID.Text = WW![ID]
    txtName.Text = WW![Name]
    txtAge.Text = WW![Age]
    txtSalary.Text = WW![Salary]
End If

End Sub

Private Sub cmdPrevious_Click()
''This is done to move to your previous record in
''the database table Master

On Error Resume Next
txtID.Enabled = False
    If WW.RecordCount <> 0 Then
    WW.Fields.Refresh
    WW.MovePrevious
    txtID.Text = WW![ID]
    txtName.Text = WW![Name]
    txtAge.Text = WW![Age]
    txtSalary.Text = WW![Salary]
End If

End Sub

Private Sub cmdSave_Click()
''Here i am checking for the fields are not empty
''before saving

If txtID.Text <> "" And txtName.Text <> "" And txtAge.Text <> "" And txtSalary.Text <> "" Then
    Dim rs As New Recordset
''Here i am checking for a similar ID in the
''database table Master.
''If it is not there i am adding a new record
''to the table, else giving error message
    rs.Open "Select * from Master where ID=" & txtID.Text, CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount = 0 Then
            rs.AddNew
            rs![ID] = txtID.Text
            rs![Name] = txtName.Text
            rs![Age] = txtAge.Text
            rs![Salary] = txtSalary.Text
            rs.Update
            rs.Close
            Call cmdClear_Click
        Else
            MsgBox "ID exits, cannot duplicate", vbCritical, "Duplication"
        End If
    Else
    MsgBox "Empty field/s found", vbCritical, "Empty field/s"
End If
End Sub

Private Sub cmdShow_Click()
On Error Resume Next
Dim X As Double
X = InputBox("Type in the Employee ID to Edit", "Edit Employee ID", Val(txtID) - 1)

Dim rs As New Recordset
rs.Open "Select * from Master where ID=" & Val(X), CON, adOpenKeyset, adLockOptimistic
If rs.RecordCount <> 0 Then
    txtID.Text = rs![ID]
    txtName.Text = rs![Name]
    txtAge.Text = rs![Age]
    txtSalary.Text = rs![Salary]
    rs.Close
Else
    MsgBox "Invalid Employee ID", vbCritical, "Invalid ID"
End If
End Sub

Private Sub cmdUpdate_Click()
''See this is simiar code to Adding a new record
''But i am not using "rs.addnew".  I am just
''useing the rs.update
If txtID.Text <> "" And txtName.Text <> "" And txtAge.Text <> "" And txtSalary.Text <> "" Then
    Dim rs As New Recordset
    rs.Open "Select * from Master where ID=" & txtID.Text, CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount <> 0 Then
        rs![ID] = txtID.Text
        rs![Name] = txtName.Text
        rs![Age] = txtAge.Text
        rs![Salary] = txtSalary.Text
        rs.Update
        rs.Close
        Call cmdClear_Click
    Else
        MsgBox "Invalid: There is no record to be updated", vbCritical, "Invalid Record"
    End If
Else
        MsgBox "Empty field/s found", vbCritical, "Field/s Empty"
End If
End Sub

Private Sub Form_Load()
''Here i am making an instance of the Connection
Set CON = New Connection

''This is nothing but a connection string for the database
''This string varies for any other database like SQL etc
CON.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\Employee.mdb"

''I am opening the record set for (move first, move next
''move last, move previous)
WW.Open "Select * from Master", CON, adOpenKeyset, adLockOptimistic
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
'This if for using the Enter Key
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
'Take a look
'This is for accepting only numbers and back space keys
ElseIf InStr(("1234567890" & vbBack & ""), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtID_GotFocus()
''This piece of code will not allow the user to
''input anything in the txtID text box
''It will check the database for being empty
''If its empty it will automatically take "1" as
''the first Employee ID
''Else it will check for the maximum number in the
''recordset and then it will add up 1 to it and show
Dim rs As New Recordset
rs.Open "Select * from Master", CON, adOpenKeyset, adLockOptimistic
If rs.RecordCount = 0 Then
    txtID.Text = 1
    rs.Close
Else
    Dim RSRS As New Recordset
    RSRS.Open "Select max(ID) as exp1 from Master", CON, adOpenKeyset, adLockOptimistic
    txtID = RSRS![exp1] + 1
    RSRS.Close
End If
SendKeys "{Tab}"
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
'This is for using the Enter key
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
Else
    'Take a look
    'This is for changing the lower case alphabets to upper case
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
'This if for using the Enter key
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
'Take a look
'This is for accepting only numbers and back space keys
ElseIf InStr(("1234567890" & vbBack & ""), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
End If
End Sub
