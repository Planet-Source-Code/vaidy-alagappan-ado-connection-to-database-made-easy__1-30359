VERSION 5.00
Begin VB.Form frmEmpMaster 
   BackColor       =   &H00008000&
   Caption         =   "Employee Master"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00008000&
      Caption         =   "&Clear All"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00008000&
      Caption         =   "&Update"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtSalary 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   600
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00008000&
      Caption         =   ">|"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   500
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00008000&
      Caption         =   ">"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   500
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00008000&
      Caption         =   "<"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   500
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00008000&
      Caption         =   "|<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   500
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00008000&
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00008000&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      BackColor       =   &H00008000&
      Caption         =   "Add &New"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00008000&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtAge 
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   600
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   2000
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   15
      Top             =   2160
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmEmpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As Connection
Dim rs As New Recordset
Dim WW As New Recordset

Private Sub cmdAdd_Click()


End Sub

Private Sub cmdAddNew_Click()
txtID.Text = ""
txtName.Text = ""
txtAge.Text = ""
txtSalary.Text = ""
txtID.Enabled = True
txtID.SetFocus
cmdSave.Enabled = True
End Sub

Private Sub cmdClear_Click()
txtID.Text = ""
txtName.Text = ""
txtAge.Text = ""
txtSalary.Text = ""
txtID.Enabled = True
txtID.SetFocus
'txtID.SetFocus
'cmdUpdate.Enabled = True
'cmdDelete.Enabled = True
'cmdAdd.Enabled = True
'cmdModify.Enabled = True
'cmdUpdate.Enabled = True
'cmdFirst.Enabled = True
'cmdNext.Enabled = True
'cmdPrevious.Enabled = True
'cmdLast.Enabled = True
End Sub

Private Sub cmdDelete_Click()
If txtID.Text <> "" And txtName.Text <> "" And txtAge.Text <> "" And txtSalary.Text <> "" Then
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

Private Sub cmdModify_Click()
'On Error Resume Next
'Dim X As Double
'X = InputBox("Type in the Employee ID to Edit", "Edit Employee ID", "1")
'Dim rs As New Recordset
'rs.Open "Select * from Master where ID=" & Val(X), CON, adOpenKeyset, adLockOptimistic
'If rs.RecordCount <> 0 Then
'cmdUpdate.Enabled = True
'txtID.Enabled = False
'txtID.Text = rs![ID]
'txtName.Text = rs![Name]
'txtAge.Text = rs![Age]
'txtSalary.Text = rs![Salary]
'rs.Close
'Else
'MsgBox "Invalid Employee ID", vbCritical, "Invalid ID"
'End If

End Sub

Private Sub cmdNext_Click()
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
If txtID.Text <> "" And txtName.Text <> "" And txtAge.Text <> "" And txtSalary.Text <> "" Then
    Dim rs As New Recordset
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

Private Sub cmdUpdate_Click()

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
'cmdUpdate.Enabled = False
'txtID.Enabled = True
End Sub

Private Sub Form_Load()
Set CON = New Connection
CON.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\Employee.mdb"
WW.Open "Select * from Master", CON, adOpenKeyset, adLockOptimistic
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{Tab}"
ElseIf InStr(("1234567890" & vbBack & ""), Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If
End Sub

Private Sub txtID_GotFocus()
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
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{Tab}"
Else
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
SendKeys "{Tab}"
ElseIf InStr(("1234567890" & vbBack & ""), Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If

End Sub
