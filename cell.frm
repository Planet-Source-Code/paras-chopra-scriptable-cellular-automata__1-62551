VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cell 
   Caption         =   "Cell Laboratory v1.0"
   ClientHeight    =   8430
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9420
   FillColor       =   &H00FF0000&
   Icon            =   "cell.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   628
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList main_image 
      Left            =   840
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":1E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":4624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":5476
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":7C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":A3DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":B22C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":C07E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":E830
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":10FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":13794
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":15F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":17DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cell.frx":18C1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame errorlog_frame 
      Caption         =   "Information Log"
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   6600
      Width           =   9255
      Begin VB.TextBox infolog 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "cell.frx":1B3CC
         Top             =   360
         Width           =   8895
      End
   End
   Begin MSComctlLib.ImageList main_imagelist 
      Left            =   7080
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox tab_script 
      Height          =   4215
      Left            =   5520
      ScaleHeight     =   4155
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   600
      Width           =   3375
      Begin RichTextLib.RichTextBox code 
         Height          =   3495
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6165
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"cell.frx":1B3DE
      End
      Begin MSComctlLib.Toolbar control_panel_script 
         Height          =   540
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   953
         ButtonWidth     =   873
         ButtonHeight    =   953
         Appearance      =   1
         Style           =   1
         ImageList       =   "main_image"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Run"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox tab_field 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      Begin MSComDlg.CommonDialog com 
         Left            =   1200
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox main 
         BackColor       =   &H80000009&
         Height          =   150
         Index           =   0
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   2.857
         ScaleMode       =   0  'User
         ScaleWidth      =   11
         TabIndex        =   2
         Top             =   720
         Width           =   225
      End
      Begin MSComctlLib.Toolbar control_panel_state 
         Height          =   540
         Left            =   2160
         TabIndex        =   6
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   953
         ButtonWidth     =   873
         ButtonHeight    =   953
         Appearance      =   1
         Style           =   1
         ImageList       =   "main_image"
         HotImageList    =   "main_image"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar control_panel_field 
         Height          =   540
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   953
         ButtonWidth     =   1217
         ButtonHeight    =   953
         Appearance      =   1
         Style           =   1
         ImageList       =   "main_image"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Run"
               Key             =   "run"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Clear"
               Key             =   "clear"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Random"
               Key             =   "random"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSScriptControlCtl.ScriptControl Scr 
      Left            =   1200
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   6375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11245
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Field"
            Key             =   "field"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Script"
            Key             =   "scripy"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label field_filename 
      Height          =   135
      Left            =   480
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu mnu_new 
         Caption         =   "New field"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_open 
         Caption         =   "Open field"
         Shortcut        =   ^O
      End
      Begin VB.Menu temp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_save 
         Caption         =   "Save field"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_saveas 
         Caption         =   "Save As"
         Shortcut        =   {F12}
      End
      Begin VB.Menu temp2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu undo 
         Caption         =   "Undo"
      End
      Begin VB.Menu delin 
         Caption         =   "-"
      End
      Begin VB.Menu cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu delim1 
         Caption         =   "-"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu sel_all 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnu_search 
      Caption         =   "Search"
      Begin VB.Menu mnu_find 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu_find_next 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu blahhsjdh 
         Caption         =   "-"
      End
      Begin VB.Menu replace 
         Caption         =   "Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_blah 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_go_line 
         Caption         =   "Go to line"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnu_Actions 
      Caption         =   "Actions"
      Begin VB.Menu mnu_Run 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_Clear 
         Caption         =   "Clear"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnu_Random 
         Caption         =   "Random"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu demos 
         Caption         =   "Demos"
         Begin VB.Menu game_of_life 
            Caption         =   "Game of life"
         End
         Begin VB.Menu mnu_march_left 
            Caption         =   "Marching Left"
         End
      End
      Begin VB.Menu blah 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_general_help 
         Caption         =   "Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu blahblah 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_homepage 
         Caption         =   "Program Homepage"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Rows As Integer
Public Columns As Integer
Dim neb(1 To 8) As Integer
Public sleep_value As Long
'Dim field(1 To Rows, 1 To Columns) As Integer
Dim field() As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim stopprog As Boolean
Dim internalrun As Boolean

Dim script_changed As Boolean
Dim script_selected As Boolean

Dim saved_already_field As Boolean
Dim saved_already_script As Boolean

Dim minus As Integer ' This is the "number" of lines to minus from the whole user script (including the program added one)
'this provides help to locate the error
'minus =14 in version 1.0


Dim tab_height_diff As Integer
Dim tab_width_diff As Integer
Dim info_h_diff As Integer
Dim info_w_diff As Integer

Private Sub about_Click()
ShellExecute 0, vbNullString, "http://www.paraschopra.com", vbNullString, vbNullString, 1

End Sub

Public Sub update_state(hg As Integer, wd As Integer)
 Dim count As Integer
        If IsNumeric(hg) = True And IsNumeric(wd) = True Then
            For i = 1 To Rows
                For j = 1 To Columns
                    count = count + 1
                    If count <> Rows * Columns Then
                        Unload main(count)
                    End If
                Next j
            Next i
            ReDim field(1 To CInt(hg), 1 To CInt(wd)) As Integer
            Rows = hg
            Columns = wd
            count = 0
            For i = 1 To Rows
                DoEvents
                For j = 1 To Columns - 1
                    DoEvents
                    count = count + 1
                    Load main(count)
                    main(count).Top = main(count - 1).Top
                    main(count).Left = main(count - 1).Left + main(count - 1).Width - 1
                    main(count).Visible = True
                Next j
                If i <> Rows Then
                    count = count + 1
                    Load main(count)
                    main(count).Top = main(count - j).Top + main(count - j).Height - 1
                    main(count).Visible = True
                End If
            Next i
            
            clear
        End If
End Sub

Sub clear()
stopprog = True
        DoEvents
        add_info "Paused"
        control_panel_field.Buttons(1).Caption = "Run"
        DoEvents
        aux
        add_info "Start again!"
        'control_panel_field.Buttons(1).Caption = "Pause"
End Sub

Private Sub clear_Click()
 
End Sub


    
    
    
    
Private Sub basic_Click()
ShellExecute 0, vbNullString, App.Path & "\readme.html#controls", vbNullString, vbNullString, 1

End Sub

Private Sub code_Change()
script_changed = True
End Sub


Private Sub code_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 9 Then
    lft = Left$(code.Text, code.SelStart)
    rht = Right$(code.Text, Len(code.Text) - Len(lft) - code.SelLength)
    txt = lft & vbTab & rht
    code.Text = txt
    code.SelStart = Len(lft & vbTab)
End If
If KeyCode = 17 Then KeyCode = 0
End Sub

Private Sub code_KeyPress(KeyAscii As Integer)
DoEvents
End Sub

Private Sub control_panel_field_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

Case 1:
If stopprog = True Then
run
Else
pause
End If
Case 2:
clear
Case 3:
random
End Select
End Sub

Private Sub control_panel_script_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

Case 1:
    mnu_new_Click
Case 2:
    mnu_open_Click
Case 3:
    mnu_save_Click
Case 4:


TabStrip(0).Tabs(1).Selected = True
Select_Field
mnu_Run_Click

End Select

End Sub

Private Sub control_panel_state_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1:
    mnu_open_Click
Case 2: mnu_save_Click
End Select
End Sub

Private Sub copy_Click()
SendMessage Me.code.hWnd, WM_COPY, 0&, 0& 'Copy
End Sub

Private Sub cut_Click()
SendMessage Me.code.hWnd, WM_CUT, 0&, 0& 'Cut
End Sub

Private Sub delete_Click()
    SendMessage Me.code.hWnd, WM_CLEAR, 0&, 0& 'Delete

End Sub

    Private Sub Form_Load()

tab_field.BorderStyle = 0
tab_script.BorderStyle = 0
tab_field.Visible = True
tab_script.Visible = False



tab_script.Width = TabStrip(0).Width - 20
tab_script.Height = TabStrip(0).Height - 50
tab_field.Width = TabStrip(0).Width - 20
tab_field.Height = TabStrip(0).Height - 50
tab_script.Left = tab_field.Left
tab_script.Top = tab_field.Top
code.Width = tab_script.ScaleWidth - 20
code.Height = tab_script.ScaleHeight - 1000
add_info ("Welcome to Cell Laboratory v" & App.Major & "." & App.Minor)


edit.Enabled = False
mnu_search.Enabled = False

script_changed = False
script_selected = False

saved_already_field = False
saved_already_script = False

internalrun = False

'Internal run is the game of life

minus = 14
        Dim count As Integer
        Rows = 15
        Columns = 15
        hg = Rows
        wd = Columns
sleep_value = 100
        ReDim field(1 To Rows, 1 To Columns) As Integer
        stopprog = True
       ' pause.Caption = "Pause"
       ' run.Caption = "External Run"
        'internalrun = True
        code.Text = "If curcell = 0 Then" & _
            vbCrLf & vbTab & "If sumneb = 3 Then" & _
            vbCrLf & vbTab & vbTab & "change = True" & _
            vbCrLf & vbTab & "End If" & _
            vbCrLf & "ElseIf curcell = 1 Then" & _
            vbCrLf & vbTab & "If sumneb <> 2 Then" & _
            vbCrLf & vbTab & vbTab & "If sumneb <> 3 Then change = True" & _
            vbCrLf & vbTab & "End If" & _
            vbCrLf & "Else" & _
            vbCrLf & vbTab & "change = False" & _
            vbCrLf & "End If"
        
        
        
        For i = 1 To Rows
            DoEvents
            For j = 1 To Columns - 1
                DoEvents
                count = count + 1
                Load main(count)
                main(count).Top = main(count - 1).Top
                main(count).Left = main(count - 1).Left + main(count - 1).Width - 1
                main(count).Visible = True
            Next j
            If i <> Rows Then
                count = count + 1
                Load main(count)
                main(count).Top = main(count - j).Top + main(count - j).Height - 1
                main(count).Visible = True
            End If
        Next i
        
    End Sub
    
Private Sub Form_Resize()
If Me.ScaleWidth > 100 And Me.ScaleHeight > 300 Then

TabStrip(0).Width = Me.ScaleWidth - 7
TabStrip(0).Height = Me.ScaleHeight - 137

tab_script.Width = TabStrip(0).Width - 20
tab_script.Height = TabStrip(0).Height - 50
tab_field.Width = TabStrip(0).Width - 20
tab_field.Height = TabStrip(0).Height - 50
tab_script.Left = tab_field.Left
tab_script.Top = tab_field.Top
code.Width = tab_script.ScaleWidth - 20
code.Height = tab_script.ScaleHeight - 1000

errorlog_frame.Width = TabStrip(0).Width
errorlog_frame.Top = TabStrip(0).Height + 15

infolog.Width = Screen.TwipsPerPixelX * errorlog_frame.Width - 250
End If
End Sub

    Private Sub Form_Unload(Cancel As Integer)
        If stopprog = False Then
            pause
            stoprog = True
        End If
   
      If script_changed = True Then
      ans = MsgBox("Are you sure you want to exit without saving the script?", vbYesNo, "Save?")
If ans = vbNo Then
        Do_Save
End If
End If
End
    End Sub
    
Private Sub Label4_Click()

ShellExecute 0, vbNullString, "paras_chopra@fastmail.fm", vbNullString, vbNullString, 1

End Sub

Private Sub Label5_Click()

End Sub

Private Sub game_of_life_Click()
ans = MsgBox("Please save any unsaved script." & vbCrLf & "A new script is being loaded." & vbCrLf & "Click OK to load.", vbOKCancel, "Save?")
If ans = vbOK Then
    stopprog = False
    pause
    TabStrip(0).Tabs(2).Selected = True
    code.LoadFile (App.Path & "\game_of_life.csr")
End If
End Sub

    Private Sub main_Click(Index As Integer)
        Dim i As Double
        Dim j As Integer
        i = Int(Index / Columns)
        j = ((Index / Columns) - i) * Columns
        field(i + 1, j + 1) = IIf(field(i + 1, j + 1) = 0, 1, 0)
        If field(i + 1, j + 1) = 1 Then
            main(Index).BackColor = vbBlack
        Else
            main(Index).BackColor = vbWhite
        End If
    End Sub
    
    Sub loopgen()
        On Error GoTo scError
        Dim X As Integer, Y As Integer
        Dim changes() As Integer
        Dim counter As Integer
        Dim change As Boolean
        stopprog = False
        ReDim changes(Rows * Columns) As Integer
        
        Do While 1
            If stopprog = True Then Exit Sub
            DoEvents
            For X = 1 To Rows
                DoEvents
                For Y = 1 To Columns
                    DoEvents
                    If internalrun = True Then
                        reqsum = sumofneb(X, Y)
                        If field(X, Y) = 0 Then
                            If reqsum = 3 Then
                                counter = counter + 1
                                changes(counter) = ((X - 1) * Columns) + (Y - 1)
                                
                            End If
                        Else
                            If reqsum <> 2 Then
                                If reqsum <> 3 Then
                                    counter = counter + 1
                                    changes(counter) = ((X - 1) * Columns + (Y - 1))
                                    
                                    'field(x, y) = 0 ' del
                                    'main(((x - 1) * Rows + (y - 1))).BackColor = vbWhite ' del
                                End If
                            End If
                            
                        End If
                    Else
                        codestr = ""
                        codestr = "Function change()" & vbCrLf & _
                            "Dim sumneb" & vbCrLf & "Dim curcell" & vbCrLf & "Dim neb(8)" & vbCrLf
                        codestr = codestr & "sumneb = " & sumofneb(X, Y) & vbCrLf
                        For o = 1 To 8
                            codestr = codestr & "neb(" & o & ") = " & neb(o) & vbCrLf
                        Next o
                        codestr = codestr & "curcell = " & IIf(field(X, Y) = 1, "1", "0") & vbCrLf
                        codestr = codestr & code.Text & vbCrLf
                        codestr = codestr & "End Function"
                        'code.Text = codestr
                        Scr.AddCode codestr
                        
                        change = Scr.run("change")
                        If change = True Then
                        
                            counter = counter + 1
                            changes(counter) = ((X - 1) * Columns + (Y - 1))
                        End If
                    End If
                Next Y
            Next X
            Dim i As Double
            Dim j As Integer
            If stopprog = True Then Exit Sub
            For lp = 1 To counter
                DoEvents
                If stopprog = True Then Exit Sub
                i = Int(changes(lp) / Columns)
                j = ((changes(lp) / Columns) - i) * Columns
                If field(i + 1, j + 1) = 0 Then
                    field(i + 1, j + 1) = 1
                    main(changes(lp)).BackColor = vbBlack
                Else
                    field(i + 1, j + 1) = 0
                    main(changes(lp)).BackColor = vbWhite
                End If
            Next lp
            counter = 0
            Sleep sleep_value
        Loop
scError:
        ' Use the Error object to inform the user of the
        ' error, and what line it occured in.
        pause
        add_info Scr.Error.Number & _
            ":" & Scr.Error.Description & _
            " in line " & Scr.Error.Line - minus
        TabStrip(0).SelectedItem = TabStrip(0).Tabs(2)
        MsgBox Scr.Error.Number & ":" & Scr.Error.Description & " in line " & Scr.Error.Line - minus, vbExclamation, "Error in script"
        frmGoTo.txtGo = Scr.Error.Line - minus
        Exit Sub
        
    End Sub
    
    Function sumofneb(ByVal i As Integer, ByVal j As Integer) As Integer
        
    neb(1) = field(connect_rows(i - 1), connect_columns(j - 1))
    neb(2) = field(connect_rows(i - 1), connect_columns(j))
    neb(3) = field(connect_rows(i - 1), connect_columns(j + 1))
    neb(4) = field(connect_rows(i), connect_columns(j - 1))
    neb(5) = field(connect_rows(i), connect_columns(j + 1))
    neb(6) = field(connect_rows(i + 1), connect_columns(j - 1))
    neb(7) = field(connect_rows(i + 1), connect_columns(j))
    neb(8) = field(connect_rows(i + 1), connect_columns(j + 1))
    
        
        
        For i = 1 To 8
            Summ = neb(i) + Summ
        Next i
        sumofneb = Summ
    End Function
    
    Private Sub main_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbLeftButton Then
            Dim i As Double
            Dim j As Integer
            i = Int(Index / Columns)
            j = ((Index / Columns) - i) * Columns
            field(i + 1, j + 1) = IIf(field(i + 1, j + 1) = 0, 1, 0)
            If field(i + 1, j + 1) = 1 Then
                main(Index).BackColor = vbBlack
            Else
                main(Index).BackColor = vbWhite
            End If
        End If
    End Sub
    
Sub pause()
 stopprog = IIf(stopprog = True, False, True)
        If stopprog = False Then
            add_info "Running..."
            control_panel_field.Buttons(1).Caption = "Pause"
            control_panel_script.Buttons(4).Caption = "Pause"
            mnu_Run.Caption = "Pause"
            control_panel_field.Buttons(1).Image = 13
            control_panel_script.Buttons(4).Image = 13
            loopgen
        Else
            add_info "Paused"
            control_panel_field.Buttons(1).Caption = "Run"
            control_panel_script.Buttons(4).Caption = "Run"
            mnu_Run.Caption = "Run"
            control_panel_field.Buttons(1).Image = 9
             control_panel_script.Buttons(4).Image = 9
            
        End If
End Sub

Sub random()
    Randomize
    Dim count As Integer
    For i = 1 To Rows
        DoEvents
        For j = 1 To Columns
            DoEvents
            main(count).BackColor = vbWhite
            If Rnd <= 0.5 Then
                field(i, j) = 0
            Else
                field(i, j) = 1
                main(count).BackColor = vbBlack
            End If
            count = count + 1
        Next j: Next i
End Sub
Sub run()
add_info "Running..."
        control_panel_field.Buttons(1).Caption = "Pause"
        control_panel_script.Buttons(4).Caption = "Pause"
        mnu_Run.Caption = "Pause"
        control_panel_field.Buttons(1).Image = 13
        control_panel_script.Buttons(4).Image = 13
        loopgen

End Sub

Sub start()
'add_info "Running..."
 '       pause.Caption = "Pause"
  '      loopgen
End Sub


    
    
    Sub aux()
        Dim count As Integer
        Dim p As Double
        Dim q As Integer
        
        For i = 1 To Rows
            DoEvents
            For j = 1 To Columns
                DoEvents
                main(count).BackColor = vbWhite
                p = Int(count / Columns)
                q = ((count / Columns) - p) * Columns
                field(p + 1, q + 1) = 0
                count = count + 1
            Next j
            
        Next i
        
    End Sub
    

    
Function connect_rows(i As Integer) As Integer
If i < 1 Then
connect_rows = Rows - i
ElseIf i > Rows Then
connect_rows = i - Rows
Else
connect_rows = i
End If
End Function

Function connect_columns(j As Integer) As Integer
If j < 1 Then
connect_columns = Columns - j
ElseIf j > Columns Then
connect_columns = j - Columns
Else
connect_columns = j
End If
End Function


Private Sub TabStrip1_Click(Index As Integer)

End Sub

Private Sub mnu_Clear_Click()
clear
End Sub

Private Sub mnu_find_Click()
frmFind.Show , Me
End Sub

Private Sub mnu_find_next_Click()
On Error GoTo FindNextError
    Dim lngResult As Integer
    Dim lngPos As Integer
    Dim intOptions As Integer
    ' Set search options

    lngPos = Me.code.SelStart + Me.code.SelLength
    ' Get position of the searched word
    lngResult = Me.code.Find(frmFind.cboFind.Text, lngPos, , intOptions)

    If lngResult = -1 Then 'Text not found
        MsgBox "Text not found" 'Show message]
        add_info "Text not found"
        frmFind.cmdFind.Caption = "&Find" 'Set caption
        frmFind.cmdReplace.Enabled = False 'Disable Replace button
        frmFind.cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        mnu_find_next.Enabled = False 'Disable Find Next menu
    Else
        Me.code.SetFocus 'Set focus
    End If
FindNextError:
   Exit Sub
End Sub

Private Sub mnu_general_help_Click()
ShellExecute 0, vbNullString, App.Path & "\readme.html", vbNullString, vbNullString, 1

End Sub

Private Sub mnu_go_line_Click()
frmGoTo.Show , Me
End Sub

Private Sub new_Click()

End Sub

Private Sub open_Click()

End Sub

Private Sub mnu_homepage_Click()
ShellExecute 0, vbNullString, "http://naramcheez.paraschopra.com/celllab/index.php", vbNullString, vbNullString, 1

End Sub

Private Sub mnu_march_left_Click()
ans = MsgBox("Please save any unsaved script." & vbCrLf & "A new script is being loaded." & vbCrLf & "Click OK to load.", vbOKCancel, "Save?")
If ans = vbOK Then
    stopprog = False
    pause
    TabStrip(0).Tabs(2).Selected = True
    code.LoadFile (App.Path & "\marching_left.csr")
End If
End Sub

Private Sub mnu_new_Click()
Do_New
End Sub

Private Sub mnu_open_Click()
Do_Open
End Sub

Private Sub mnu_Random_Click()
random
End Sub

Private Sub mnu_Run_Click()
If stopprog = True Then
run
Else
pause
End If
End Sub

Private Sub mnu_save_Click()
Do_Save
End Sub

Private Sub options_Click()
If stopprog = False Then
control_panel_field_ButtonClick control_panel_field.Buttons(1)
End If
options_form.Show
End Sub

Private Sub paste_Click()
SendMessage Me.code.hWnd, WM_PASTE, 0&, 0& 'Paste
End Sub

Private Sub replace_Click()
With frmFind
        .cmdReplace.Top = 150 'Set cmdReplace top
        .cmdReplace.Caption = "&Replace" 'Set caption
        .lblReplace.Visible = True 'Show lblReplace
        .cboReplace.Visible = True 'Show cboReplace
        .cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        .Show , Me
    End With
End Sub

Private Sub scripting_Click()
ShellExecute 0, vbNullString, App.Path & "\readme.html#scripting_lang", vbNullString, vbNullString, 1

End Sub

Private Sub sel_all_Click()
Me.code.SelStart = 0 'Set the start pos of the selection
Me.code.SelLength = Len(code) 'Set length of the selection
End Sub

Private Sub TabStrip_Click(Index As Integer)
Select Case TabStrip(Index).SelectedItem.Index
Case 1:
    Select_Field
Case 2:

    tab_script.Visible = True
    tab_field.Visible = False
    edit.Enabled = True
    mnu_search.Enabled = True
    script_selected = True
    mnu_open.Caption = "Open script"
    mnu_new.Caption = "New script"
    mnu_save.Caption = "Save script"
End Select

Process_Tab_Change

End Sub


Public Sub add_info(info As String)
infolog.Text = infolog.Text & vbCrLf & info
infolog.SelStart = Len(infolog)
infolog.SelLength = 1
End Sub

Private Sub TabStrip_GotFocus(Index As Integer)
If TabStrip(Index).SelectedItem.Index = 2 Then
    'Call code.SetFocus
    End If
End Sub

Private Sub undo_Click()
SendMessage Me.code.hWnd, EM_UNDO, 0, 0&
End Sub

Sub Process_Tab_Change()

End Sub

Sub view_script()
tab_script.Visible = True
    tab_field.Visible = False
    edit.Enabled = True
    script_selected = True
    
    Process_Tab_Change
End Sub

Sub Do_Save()
If script_selected = True Then
'open the open box

If code.Tag = "" Then
    Do_Save_As
Else
    code.SaveFile code.Tag
    add_info code.Tag & " saved successfully."
    Me.Caption = "Cell Laboratory- [" & code.Tag & "]"
End If



Else
    'answer = MsgBox("Are you sure you want a open a new field?", vbYesNo, "Sure?")
'Do open the open box

If field_filename = "" Then
    Do_Save_As
Else
    SaveField (field_filename)
    
End If
End If

End Sub

Sub Do_Open()
If script_changed = True And script_selected = True Then
    answer = MsgBox("Do you want to save the changes made to the script?", vbYesNoCancel, "Save?")
    If answer = vbYes Then
        script_changed = False
        Do_Save
        
    ElseIf answer = vbCancel Then
        Exit Sub
    End If
End If

'Write the procedure for opening here

If script_selected = True Then
'open the open box

com.Filter = "Script Files|*.csr|All Files|*.*"
    com.ShowOpen
    If com.Filename <> "" Then
        code.LoadFile com.Filename
        add_info com.Filename & " loaded successfully."
        Me.Caption = "Cell Laboratory- [" & com.Filename & "]"
    End If
    
Else
    'answer = MsgBox("Are you sure you want a open a new field?", vbYesNo, "Sure?")
'Do open the open box


    Dim strmain As String
    Dim strbuffer As String
    Dim strtext As String
    Dim temp() As String
    Dim hg As Integer, wd As Integer

    com.Filter = "Field Files|*.fld|All Files|*.*"
    com.ShowOpen
    If com.Filename <> "" Then
        strmain = com.Filename

        Open strmain For Input As #1
        MousePointer = vbHourglass

        Line Input #1, strbuffer
        If strbuffer <> "CLF" Then
            MsgBox "Not a valid field file!"
            Close #1
            Exit Sub
        End If
        Line Input #1, strbuffer
        Select Case strbuffer
            Case "1.0":
                Line Input #1, strbuffer
                hg = CInt(strbuffer)
                Line Input #1, strbuffer
                wd = CInt(strbuffer)
                update_state hg, wd
                Line Input #1, strbuffer
                sleep_value = CInt(strbuffer)
            For i = 1 To hg
                Line Input #1, strbuffer
                temp = Split(strbuffer, " ")
                For j = LBound(temp) To UBound(temp)
                    If temp(j) = "1" Then
                        main_Click ((i - 1) * wd + (j))
                    End If
                Next j
            Next i
            Case Default: MsgBox "Not a valid field file!"
                    Exit Sub
        End Select
'strtext = strtext & strbuffer & vbCrLf
    MousePointer = vbDefault

    
    add_info strmain & " loaded successfully."
    Close #1
    End If

End If

End Sub

Sub Do_New()
If script_changed = True And script_selected = True Then
    answer = MsgBox("Do you want to save the changes made to the script?", vbYesNoCancel, "Save?")
    If answer = vbYes Then
        script_changed = False
        Call Do_Save
        
    ElseIf answer = vbCancel Then
        Exit Sub
    End If
End If
If script_selected = True Then
code.Text = ""
code.Tag = ""
Else
answer = MsgBox("Are you sure you want a new field?", vbYesNo, "Sure?")
If answer = vbYes Then
sleep_value = 100
update_state 15, 15
clear
field_filename = ""
End If
End If
End Sub

Sub Do_Save_As()
If script_selected = True Then
'open the open box

com.Filter = "Script Files|*.csr|All Files|*.*"
    com.ShowSave
    If com.Filename <> "" Then
        code.SaveFile com.Filename
    
   add_info com.Filename & " saved successfully."
    
    code.Tag = com.Filename
    Me.Caption = "Cell Laboratory- [" & code.Tag & "]"
    End If


Else
    'answer = MsgBox("Are you sure you want a open a new field?", vbYesNo, "Sure?")
'Do open the open box


    Dim strmain As String
    Dim strbuffer As String
    Dim strtext As String
    Dim temp() As String
    Dim hg As Integer, wd As Integer

    com.Filter = "Field Files|*.fld|All Files|*.*"
    com.ShowSave
    If com.Filename <> "" Then
        strmain = com.Filename

        SaveField (strmain)
    
    Me.Caption = "Cell Laboratory- [" & strmain & "]"
    
    field_filename = strmain
    
    End If

End If

End Sub


Sub SaveField(Filename As String)
       Dim temp_str As String
       
        Open Filename For Output As #1
        MousePointer = vbHourglass
        
        Print #1, "CLF"
        Print #1, "1.0"
        Print #1, Rows
        Print #1, Columns
        Print #1, sleep_value
        For i = main.LBound To main.UBound
            
            'if main(0)
       If main(i).BackColor = vbBlack Then
        temp_str = temp_str & "1 "
        Else
        temp_str = temp_str & "0 "
        End If
        
        If (i + 1) Mod Columns = 0 Then
                
                Print #1, temp_str
                temp_str = ""
            End If
        
       Next i
'strtext = strtext & strbuffer & vbCrLf
    MousePointer = vbDefault

    
    
    Close #1
    add_info Filename & " saved successfully."
    
End Sub

Sub Select_Field()
tab_field.Visible = True
    tab_script.Visible = False
    edit.Enabled = False
    mnu_search.Enabled = False
    script_selected = False
    mnu_open.Caption = "Open field"
    mnu_new.Caption = "New field"
    mnu_save.Caption = "Save field"
End Sub
