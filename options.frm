VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form options_form 
   Caption         =   "Options"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   LinkTopic       =   "Form2"
   ScaleHeight     =   2490
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox tab_1 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   360
      Width           =   4575
      Begin VB.CommandButton apply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox wd 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox hg 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox sleep_value 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "100"
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Width:"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Height:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sleep:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
   End
   Begin MSComctlLib.TabStrip Tab 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "general"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "options_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub apply_Click()
If IsNumeric(hg) = True And IsNumeric(wd) = True Then
    cell.update_state hg, wd
    Me.Hide
Else
MsgBox "Please enter a numeric value in height/width field"
End If
If IsNumeric(sleep_value) Then
    cell.sleep_value = sleep_value
    Me.Hide
Else
MsgBox "Please enter a numeric value in sleep field"
End If
End Sub

Private Sub Form_Load()
sleep_value = cell.sleep_value
hg = cell.Rows
wd = cell.Columns
tab_1.BorderStyle = none
End Sub

Private Sub sleep_value_Click()
sleep_value.SelStart = 0
        sleep_value.SelLength = Len(sleep_value.Text)
End Sub

Private Sub TabStrip1_Click()

End Sub
