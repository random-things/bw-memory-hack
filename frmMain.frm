VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Someone's Basic BW 1.13 Memory Hacks"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameHelp 
      Caption         =   "Help / About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   41
      Top             =   360
      Width           =   9015
      Begin VB.TextBox txtAbout 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Text            =   "frmMain.frx":0442
         Top             =   3360
         Width           =   8775
      End
      Begin VB.TextBox txtHelp 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "frmMain.frx":06AF
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame frameMisc 
      Caption         =   "Miscellaneous Patches"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   36
      Top             =   360
      Width           =   9015
      Begin VB.Frame frameCOPerms 
         Caption         =   "Courtesy of Permaphrost"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   8775
         Begin VB.CommandButton cmdPlayWithSelf 
            Caption         =   "Let Me Play With Myself"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   3975
         End
         Begin VB.CommandButton cmdNoForceLeave 
            Caption         =   "Let Me Stay Even After I Die"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   3975
         End
         Begin VB.CommandButton cmdMapDownloadStatus 
            Caption         =   "Always Show Map Download Status"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   3975
         End
      End
   End
   Begin VB.Frame frameAllyMap 
      Caption         =   "Ally Map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   9015
      Begin VB.TextBox txtAllyReport 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4125
         Left            =   4680
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   4245
      End
      Begin VB.CommandButton cmdUpdateAllyMap 
         Caption         =   "Update Ally Map"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   4320
         Width           =   4245
      End
      Begin MSComctlLib.ListView listViewAllyMap 
         Height          =   4095
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   0
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "2"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "3"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "4"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "5"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "6"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "7"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "8"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "9"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "10"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "11"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "12"
            Object.Width           =   706
         EndProperty
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   2805
         Width           =   210
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   33
         Top             =   2610
         Width           =   210
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   2415
         Width           =   210
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   225
         TabIndex        =   31
         Top             =   2220
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   30
         Top             =   1995
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   29
         Top             =   1785
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   28
         Top             =   1575
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   27
         Top             =   1365
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   26
         Top             =   1155
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   25
         Top             =   930
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   24
         Top             =   735
         Width           =   105
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   23
         Top             =   510
         Width           =   105
      End
   End
   Begin VB.Frame frameGameMessage 
      Caption         =   "In-Game Message Spoofer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   11
      Top             =   5040
      Width           =   4815
      Begin VB.TextBox txtCurrentText 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtSpoofText 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton cmdSpoofText 
         Caption         =   "Do the spoof!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.Timer timerGameTextSpoof 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2280
         Top             =   840
      End
      Begin VB.TextBox txtChar 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdInsertText 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   975
         Width           =   735
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "Automatic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblCurrText 
         AutoSize        =   -1  'True
         Caption         =   "Current Text:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblSpoofText 
         AutoSize        =   -1  'True
         Caption         =   "Spoof Text:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.Frame frameNameSpoof 
      Caption         =   "Name Spoofer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   4095
      Begin VB.Timer timerNameSpoof 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1560
         Top             =   840
      End
      Begin VB.TextBox txtCurrentName 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtSpoofName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdSpoof 
         Caption         =   "Do the spoof!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblCurrName 
         AutoSize        =   -1  'True
         Caption         =   "Current Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label lblSpoofName 
         AutoSize        =   -1  'True
         Caption         =   "Spoof Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1125
      End
   End
   Begin VB.Frame frameStats 
      Caption         =   "Stats Viewer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   9015
      Begin VB.CommandButton cmdUpdateStats 
         Caption         =   "Update Player Stats"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4200
         Width           =   2175
      End
      Begin MSComctlLib.ListView listViewStats 
         Height          =   3975
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   0
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   617
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "B.net Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Player Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Race"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Minerals"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vespene"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Player Status"
            Object.Width           =   2822
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tabStats 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11456
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "In-Game Statistics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ally Map"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Miscellaneous Patches"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6495
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer timerGameState 
      Interval        =   100
      Left            =   120
      Top             =   6960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' For name spoofing
Private Const IN_GAME_NAME As Long = &H1903EE70
' For chat spoofing
Private Const IN_GAME_CHAT_BUFFER_POINTER = &H1504BD7C
' For grabbing player names
Private Const BASE_NAME_TEAM_BUFFER = &H64C4D0
' For grabbing in-game names and races
Private Const BASE_NAME_BUFFER = &H65AE20
' For grabbing minerals
Private Const BASE_MINERALS = &H508600
' For grabbing vespene
Private Const BASE_GAS = &H508630
' 0 = chat, 1 = in game
Private Const IN_GAME As Long = &H645E34
' 0 = chat, 1 = lobby
Private Const IN_LOBBY As Long = &H647BE0
' 0 = chat/lobby, 1 = countdown/game
Private Const IN_COUNTDOWN As Long = &H65CE80
' For grabbing the ally map
Private Const BASE_ALLY_MAP As Long = &H516B44

' For patching to always show download status
Private Const DOWNLOAD_STATUS_LOCATION As Long = &H457220

' For patching to allow staying in game after game is over
Private Const STICK_AROUND_PATCH_LOCATION As Long = &H4CA096

' For patching to allow starting games without an opponent
Private Const LET_ME_PLAY_WITH_MYSELF As Long = &H454813

Private Type InGamePlayerStruct
    playerStatus As Byte
    playerRace As Byte
    unknownByte As Byte
    playerName As String * 25
    playerNumber As Long
    unknownDword As Long
End Type

Private Enum GameStates
    Chat = 0
    GameLobby = 1
    Countdown = 2
    Game = 3
End Enum

Private gameState As Integer
Private playerNumber As Integer

Private gamePlayers(12) As String
Private playerIsAllyVictory(12) As Boolean

Private DOWNLOAD_STATUS_PATCH_STRING As String
Private FORCED_LEAVE_GAME_PATCH_STRING As String
Private START_WITHOUT_OPPONENT_PATCH_STRING As String

Private Mem As CMemoryPatcher

Private Sub SetGameState()
    If Mem.ReadByteFromMemory("Brood War", IN_GAME) = 1 Then
        gameState = GameStates.Game
    ElseIf Mem.ReadByteFromMemory("Brood War", IN_LOBBY) = 1 Then
        gameState = GameStates.GameLobby
    ElseIf (Mem.ReadByteFromMemory("Brood War", IN_COUNTDOWN) = 1) And (Mem.ReadByteFromMemory("Brood War", IN_GAME) = 0) Then
        gameState = GameStates.Countdown
    ElseIf (Mem.ReadByteFromMemory("Brood War", IN_LOBBY) = 0) And (Mem.ReadByteFromMemory("Brood War", IN_GAME) = 0) Then
        gameState = GameStates.Chat
    Else
        MsgBox "Unknown Game State!"
    End If
End Sub

Private Sub cmdInsertText_Click()
    If gameState = GameStates.Game And Len(txtChar.Text) > 0 Then
        txtSpoofText.Text = txtSpoofText.Text + Chr(CInt(txtChar.Text))
    End If
End Sub

Private Sub cmdMapDownloadStatus_Click()
    Call PatchMapDownloadStatus
End Sub

Private Sub cmdNoForceLeave_Click()
    Call PatchForcedLeaveGame
End Sub

Private Sub cmdPlayWithSelf_Click()
    Call PatchForceOpponent
End Sub

Private Sub cmdSpoof_Click()
    If Len(txtSpoofName.Text) > 15 Then
        MsgBox "Spoofed name must be less than 16 characters!"
        Exit Sub
    End If
    
    If gameState = GameStates.Chat Then
        Call Mem.SetPatch(txtSpoofName.Text & Chr(0), Len(txtSpoofName.Text) + 1)
        Call Mem.SetPatchAddress(IN_GAME_NAME)
        Call Mem.ApplyPatch("Brood War")
    End If
End Sub

Private Sub cmdSpoofText_Click()
    Call SetSpoofText(txtSpoofText.Text)
End Sub

Private Sub cmdUpdateAllyMap_Click()
    Call UpdateAllyMap
End Sub

Private Sub cmdUpdateStats_Click()
    Call UpdateGamePlayers
    listViewStats.ListItems.Clear
    
    Dim i As Integer
    Dim playerStruct As InGamePlayerStruct
    Dim bnetName As String, playerName As String, playerRace As Long, playerMinerals As Long, playerGas As Long, playerStatus As Integer
    For i = 0 To 7
        bnetName = Mem.ReadBytesFromMemoryUntilChar("Brood War", BASE_NAME_TEAM_BUFFER + (72 * i), Chr(0))
        If Len(bnetName) > 0 Then
            playerStruct.playerName = Mem.ReadBytesFromMemoryUntilChar("Brood War", BASE_NAME_BUFFER + ((36 * i) + 3), Chr(0))
            playerStruct.playerRace = Mem.ReadByteFromMemory("Brood War", BASE_NAME_BUFFER + ((36 * i) + 1)) + 1
            playerStruct.playerStatus = Mem.ReadByteFromMemory("Brood War", BASE_NAME_BUFFER + (36 * i))
            playerStruct.playerNumber = Mem.ReadDwordFromMemory("Brood War", BASE_NAME_BUFFER + ((36 * i) + 28))
            
            ' Because BW cuts off that structure, so this is a hacky way to get what we want.
            If playerStruct.playerNumber = 0 Then
                For k = 0 To 11
                    If RTrim(playerStruct.playerName) = gamePlayers(k) Then playerStruct.playerNumber = k + 1
                Next k
            End If
            
            playerMinerals = Mem.ReadDwordFromMemory("Brood War", BASE_MINERALS + (4 * i))
            playerGas = Mem.ReadDwordFromMemory("Brood War", BASE_GAS + (4 * i))
        
            Dim tempString As String
            Select Case playerStruct.playerStatus
                Case 1:
                    tempString = "Live computer"
                Case 2:
                    tempString = "Live human"
                Case &HA:
                    tempString = "Dead human"
                Case &HB:
                    tempString = "Dead computer"
                Case Else:
                    tempString = "Unknown (" & Hex(playerStatus) & ")"
            End Select
        
            Dim X As ListItem
            Set X = listViewStats.ListItems.Add(, , playerStruct.playerNumber)
            X.ListSubItems.Add , , bnetName
            X.ListSubItems.Add , , RTrim(playerStruct.playerName)
            X.ListSubItems.Add , , Choose(playerStruct.playerRace, "Zerg", "Terran", "Protoss")
            X.ListSubItems.Add , , playerMinerals
            X.ListSubItems.Add , , playerGas
            X.ListSubItems.Add , , tempString
        End If
    Next i
End Sub

Private Sub Form_Load()
    Set Mem = New CMemoryPatcher
    
    'Call Mem.SetPatchAddress(&H1903EE70)
    'strPatch = "God" & Chr(0)
    'Call Mem.SetPatch(strPatch, Len(strPatch))
    'MsgBox Mem.ReadBytesFromMemory("Brood War", &H1903EE70, 15)
    'MsgBox Mem.ReadBytesFromMemoryUntilChar("Brood War", &H1903EE70, Chr(0))
    'MsgBox Mem.ApplyPatch("Brood War")
    
    frameStats.ZOrder 0
End Sub

Private Sub timerGetInfo_Timer()
    txtGameName.Text = Mem.ReadBytesFromMemoryUntilChar("Brood War", IN_GAME_NAME, Chr(0))
    txtInGame.Text = Mem.ReadBytesFromMemoryUntilChar("Brood War", Mem.ReadDwordFromMemory("Brood War", IN_GAME_CHAT_BUFFER_POINTER) + &H94, Chr(0))
End Sub

Private Sub tabStats_Click()
    If tabStats.SelectedItem.Caption = "In-Game Statistics" Then
        frameStats.ZOrder 0
    ElseIf tabStats.SelectedItem.Caption = "Ally Map" Then
        frameAllyMap.ZOrder 0
    ElseIf tabStats.SelectedItem.Caption = "Miscellaneous Patches" Then
        frameMisc.ZOrder 0
    ElseIf tabStats.SelectedItem.Caption = "Help" Then
        frameHelp.ZOrder 0
    End If
End Sub

Private Sub timerGameState_Timer()
    Call SetGameState

    If gameState = GameStates.Chat Then
        cmdUpdateStats.Enabled = False
        timerNameSpoof.Enabled = True
        cmdSpoof.Enabled = True
        timerGameTextSpoof.Enabled = False
        cmdSpoofText.Enabled = False
    ElseIf gameState = GameStates.Game Then
        cmdUpdateStats.Enabled = True
        timerNameSpoof.Enabled = False
        cmdSpoof.Enabled = False
        timerGameTextSpoof.Enabled = True
        cmdSpoofText.Enabled = True
    ElseIf gameState = GameStates.GameLobby Then
        cmdUpdateStats.Enabled = False
        timerNameSpoof.Enabled = False
        cmdSpoof.Enabled = False
        timerGameTextSpoof.Enabled = False
        cmdSpoofText.Enabled = False
    ElseIf gameState = GameStates.Countdown Then
        cmdUpdateStats.Enabled = False
        timerNameSpoof.Enabled = False
        cmdSpoof.Enabled = False
        timerGameTextSpoof.Enabled = False
        cmdSpoofText.Enabled = False
    End If

    Dim statusString As String
    statusString = Choose(gameState + 1, "Not in a game.", "In the game chat lobby.", "Counting down...", "In a game.")
    
    statusBar.SimpleText = statusString
End Sub

Private Sub SetSpoofText(ByVal strSpoofText As String)
    If gameState = GameStates.Game Then
        Call Mem.SetPatch(strSpoofText, Len(strSpoofText))
        Call Mem.SetPatchAddress(Mem.ReadDwordFromMemory("Brood War", IN_GAME_CHAT_BUFFER_POINTER) + &H94)
        Call Mem.ApplyPatch("Brood War")
    End If
End Sub

Private Sub AddAllyText(ByVal strString As String)
    txtAllyReport.Text = txtAllyReport.Text & strString & vbCrLf
    txtAllyReport.SelStart = Len(txtAllyReport.Text)
End Sub

Private Sub UpdateGamePlayers()
    Dim i As Integer
    
    For i = 0 To 11
        gamePlayers(i) = Mem.ReadBytesFromMemoryUntilChar("Brood War", BASE_NAME_BUFFER + ((36 * i) + 3), Chr(0))
        playerIsAllyVictory(i) = False
    Next i
End Sub

Private Sub PatchMapDownloadStatus()
    DOWNLOAD_STATUS_PATCH_STRING = Chr(0) & Chr(&HC6) & Chr(&H41) & Chr(&H18) & Chr(&H18)
    Call Mem.SetPatch(DOWNLOAD_STATUS_PATCH_STRING, 5)
    Call Mem.SetPatchAddress(DOWNLOAD_STATUS_LOCATION)
    Call Mem.ApplyPatch("Brood War")
End Sub

Private Sub PatchForcedLeaveGame()
    FORCED_LEAVE_GAME_PATCH_STRING = Chr(&HC3)
    Call Mem.SetPatch(FORCED_LEAVE_GAME_PATCH_STRING, 1)
    Call Mem.SetPatchAddress(STICK_AROUND_PATCH_LOCATION)
    Call Mem.ApplyPatch("Brood War")
End Sub

Private Sub PatchForceOpponent()
    START_WITHOUT_OPPONENT_PATCH_STRING = Chr(&HEB)
    Call Mem.SetPatch(START_WITHOUT_OPPONENT_PATCH_STRING, 1)
    Call Mem.SetPatchAddress(LET_ME_PLAY_WITH_MYSELF)
    Call Mem.ApplyPatch("Brood War")
End Sub

Private Sub UpdateAllyMap()
    If gameState = GameStates.Game Then
        Call UpdateGamePlayers
    
        listViewAllyMap.ListItems.Clear
        txtAllyReport.Text = ""
        
        Dim X As ListItem
        Dim i As Integer, j As Integer, tempVar As Integer
        
        For i = 0 To 11
            'gamePlayers(i) = Mem.ReadBytesFromMemoryUntilChar("Brood War", BASE_NAME_BUFFER + ((36 * i) + 3), Chr(0))
            'Debug.Print "-- Player " & i + 1 & "(" & gamePlayers(i) & ")"
            
            If Len(gamePlayers(i)) > 0 Then
                Set X = listViewAllyMap.ListItems.Add(, , IIf(Mem.ReadByteFromMemory("Brood War", BASE_ALLY_MAP + (12 * i)), "X", ""))
                'Debug.Print "Player 1 = " & Mem.ReadByteFromMemory("Brood War", BASE_ALLY_MAP + (12 * i))
                For j = 1 To 11
                    tempVar = Mem.ReadByteFromMemory("Brood War", BASE_ALLY_MAP + (12 * i) + j)
                    If tempVar = 2 Then playerIsAllyVictory(i) = True
                    'Debug.Print "Player " & j + 1 & " = " & tempVar
                    'Debug.Print "tempVar -> " & tempVar
                    X.ListSubItems.Add , , IIf(tempVar, "X", "")
                    
                    'If (tempVar > 0) And (i <> j) And (Len(gamePlayers(j)) > 0) Then
                    '    AddAllyText gamePlayers(i) & " is allied with " & gamePlayers(j) & "."
                    'End If
                Next j
            Else
                listViewAllyMap.ListItems.Add , , ""
            End If
            
            If playerIsAllyVictory(i) = True Then AddAllyText gamePlayers(i) & " set allied victory."
        Next i
        
    End If
End Sub

Private Sub timerGameTextSpoof_Timer()
    Dim bufferPtr As Long
    bufferPtr = Mem.ReadDwordFromMemory("Brood War", IN_GAME_CHAT_BUFFER_POINTER)
    bufferPtr = bufferPtr + &H94
    txtCurrentText.Text = Mem.ReadBytesFromMemoryUntilChar("Brood War", bufferPtr, Chr(0))
    
    
    If chkAuto.Value = vbChecked Then
        Dim tempString As String
        tempString = txtCurrentText.Text
        tempString = Replace(tempString, "[lb]", Chr(2) & Chr(2) & Chr(2) & Chr(2))
        tempString = Replace(tempString, "[y]", Chr(3) & Chr(3) & Chr(3))
        tempString = Replace(tempString, "[w]", Chr(4) & Chr(4) & Chr(4))
        tempString = Replace(tempString, "[gy]", Chr(5) & Chr(5) & Chr(5) & Chr(5))
        tempString = Replace(tempString, "[avon]", BuildAVReport(True))
        tempString = Replace(tempString, "[avoff]", BuildAVReport(False))
                
        Call SetSpoofText(tempString)
    End If
End Sub

Private Function BuildAVReport(Optional ByVal AVOn As Boolean = True) As String
    Dim i As Integer
    BuildAVReport = "Allied Victory " & IIf(AVOn, "On: ", "Off: ")
    For i = 0 To 11
        If playerIsAllyVictory(i) = AVOn Then
            BuildAVReport = BuildAVReport & gamePlayers(i) & ", "
        End If
    Next i
End Function

Private Sub timerNameSpoof_Timer()
    txtCurrentName.Text = Mem.ReadBytesFromMemoryUntilChar("Brood War", IN_GAME_NAME, Chr(0))
End Sub
