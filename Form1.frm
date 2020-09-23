VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialogs"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   2205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Show Font"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Color"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Printer"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Save"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Open"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    On Error GoTo e_Trap
    
    FileDialog.sFilter = "Text Files (*.txt)" & Chr$(0) & "*.sky" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.Path & "\"
    sOpen = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        FileList = "Directory : " & sOpen.sLastDirectory & vbCr
        For Count = 1 To sOpen.nFilesSelected
            FileList = FileList & sOpen.sFiles(Count) & vbCr
        Next Count
        Call MsgBox(FileList, vbOKOnly + vbInformation, "Show Open Selected")
    End If
    Exit Sub
e_Trap:
    Exit Sub
    Resume

End Sub

Private Sub Command2_Click()
Dim sSave As SelectedFile
Dim Count As Integer
Dim FileList As String

    On Error GoTo e_Trap
    
    FileDialog.sFilter = "Text Files (*.txt)" & Chr$(0) & "*.sky" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sDlgTitle = "Show Save"
    FileDialog.sInitDir = App.Path & "\"
    sSave = ShowSave(Me.hWnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
        FileList = "Directory : " & sSave.sLastDirectory & vbCr
        For Count = 1 To sSave.nFilesSelected
            FileList = FileList & sSave.sFiles(Count) & vbCr
        Next Count
        Call MsgBox(FileList, vbOKOnly + vbInformation, "Show Save Selected")
    End If
    Exit Sub
e_Trap:
    Exit Sub
    Resume

End Sub

Private Sub Command3_Click()
Dim sFont As SelectedFont
    On Error GoTo e_Trap
    FontDialog.iPointSize = 12 * 10
    sFont = ShowFont(Me.hWnd, "Times New Roman")
    Exit Sub
e_Trap:
    Exit Sub
End Sub

Private Sub Command4_Click()
    On Error GoTo e_Trap
    Call ShowPrinter(Me.hWnd)
    Exit Sub
e_Trap:
    Exit Sub
End Sub

Private Sub Command5_Click()
Dim sColor As SelectedColor
    On Error GoTo e_Trap
    sColor = ShowColor(Me.hWnd)
    Exit Sub
e_Trap:
    Exit Sub
End Sub
