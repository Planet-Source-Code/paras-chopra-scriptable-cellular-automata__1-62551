VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBar 
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   75
      ScaleHeight     =   1290
      ScaleWidth      =   5340
      TabIndex        =   6
      Top             =   900
      Width           =   5340
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace &All"
         Height          =   315
         Left            =   4275
         TabIndex        =   13
         Top             =   525
         Width           =   990
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         Height          =   315
         Left            =   4275
         TabIndex        =   12
         Top             =   900
         Width           =   990
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace..."
         Height          =   315
         Left            =   4275
         TabIndex        =   11
         Top             =   150
         Width           =   990
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search Options"
         Height          =   1215
         Left            =   75
         TabIndex        =   7
         Top             =   0
         Width           =   4065
         Begin VB.CheckBox chkWholeWord 
            Caption         =   "Find Whole Word &Only"
            Height          =   240
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Width           =   1965
         End
         Begin VB.CheckBox chkMatchCase 
            Caption         =   "Match Ca&se"
            Height          =   240
            Left            =   150
            TabIndex        =   9
            Top             =   600
            Width           =   1965
         End
         Begin VB.CheckBox chkNoHighlight 
            Caption         =   "No &Highlight"
            Height          =   240
            Left            =   150
            TabIndex        =   8
            Top             =   900
            Width           =   1965
         End
      End
   End
   Begin VB.ComboBox cboReplace 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   450
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4350
      TabIndex        =   3
      Top             =   450
      Width           =   990
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Left            =   4350
      TabIndex        =   1
      Top             =   75
      Width           =   990
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   75
      Width           =   3015
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace &With:"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label lblFind 
      Caption         =   "Fin&d What:"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Private Sub cmdFind_Click()
    On Error GoTo FindError
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    If cmdFind.Caption = "&Find" Then 'If first time
        ' Get position of the searched word
        lngResult = cell.code.Find(cboFind.Text, 0, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", vbInformation, "Cell Laboratory" 'Show message
            cell.add_info "Text not found"
            cmdFind.Caption = "&Find" 'Set caption
            cell.mnu_find_next.Enabled = False 'Disable Find Next menu
        Else 'Text found
            cell.code.SetFocus 'Set focus to rtfText
            cmdReplace.Enabled = True 'Enable Replace button
            cmdReplaceAll.Enabled = True 'Enable ReplaceAll button
            cmdFind.Caption = "&Find Next" 'Set caption
            cell.mnu_find_next.Enabled = True 'Enable Find Next menu
        End If
    Else 'Find Next
        lngPos = cell.code.SelStart + cell.code.SelLength
        lngResult = cell.code.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", vbInformation, "Cell Laboratory" 'Show message
            cell.add_info "Text not found"
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
            cell.mnu_find_next.Enabled = False 'Disable Find Next menu
        Else 'Text found
            cell.code.SetFocus 'Set focus to rtfText
            cell.mnu_find_next.Enabled = True 'Enable Find Next menu
        End If
    End If
FindError:
    
End Sub

Private Sub cmdReplace_Click()
    
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    
    If cmdReplace.Caption = "&Replace..." Then 'Show replace
        cmdReplace.Top = 150 'Set cmdReplace top
        cmdReplace.Caption = "&Replace" 'Set caption
        lblReplace.Visible = True 'Show lblReplace
        cboReplace.Visible = True 'Show cboReplace
        cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        Exit Sub
    End If

    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    
    With cell
        .code.SelText = cboReplace.Text 'Replace text
        ' Find next
        lngPos = .code.SelStart + .code.SelLength
        ' Get position of the searched word
        lngResult = .code.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
           MsgBox "Text not found", vbInformation, "Cell Laboratory" 'Show message
            cell.add_info "Text not found"
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        Else 'Text found
            .code.SetFocus 'Set focus
        End If
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReplaceAll_Click()
    On Error GoTo ReplaceAllError
    Dim intCount As Integer
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    intCount = 0
    lngPos = 0
    With cell
        Do
            If .code.Find(cboFind.Text, lngPos, , intOptions) = -1 Then 'Text not fount
                If intCount > 0 Then 'Show how many replacments have been made
                   MsgBox "The specified region has been searched. " & vbCrLf & _
                    intCount & " replacements have been made.", vbInformation, "Cell Laboratory"
            cell.add_info "The specified region has been searched. " & vbCrLf & _
                    intCount & " replacements have been made."
                
                End If
                cmdFind.Caption = "&Find" 'Set caption
                cmdReplace.Enabled = False 'Disable Replace button
                cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
                Exit Do
            Else 'Text found
                lngPos = .code.SelStart + .code.SelLength
                intCount = intCount + 1 'Increase counter by 1
                .code.SelText = cboReplace.Text 'Replace text
            End If
        Loop
    End With
ReplaceAllError:
   
End Sub

Private Sub Form_Load()
    cmdReplace.Top = 525 'Set cmdReplace top
    lblReplace.Visible = False 'Hide lblReplace
    cboReplace.Visible = False 'Hide cboReplace
    cmdReplaceAll.Visible = False 'Hide cmdReplaceAll
    
    cboFind.AddItem cell.code.SelText 'Add selected text to find combobox
    cboFind.Text = cell.code.SelText 'Set text in cbo
End Sub
