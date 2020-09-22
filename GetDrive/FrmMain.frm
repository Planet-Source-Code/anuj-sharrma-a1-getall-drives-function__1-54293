VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00D5DDDD&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetDriveInfo 
      BackColor       =   &H00D5DDDD&
      Caption         =   "&GetDrive Info"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3945
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5910
      Top             =   345
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D5DDDD&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5715
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3945
      Width           =   1275
   End
   Begin MSComctlLib.ListView lvwGetDriveInfo 
      Height          =   2955
      Left            =   60
      TabIndex        =   1
      Top             =   870
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   14015965
      BackColor       =   16384
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Drive Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DriveType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Total No. of bytes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Total Free Space"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Free Bytes Available"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Space Used"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Sector per Cluster"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bytes per sector"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "No. Of free cluster"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total no. of cluster"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Total no. of bytes in path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Free bytes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblDeveloper 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Anuj sharma"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   270
      Left            =   4620
      TabIndex        =   3
      Top             =   4680
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get Drive Function"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   390
      Left            =   2295
      TabIndex        =   0
      Top             =   300
      Width           =   2385
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long

Private Sub cmdClose_Click()
    Unload Me
        End
End Sub

Private Sub cmdGetDriveInfo_Click()
Dim sDrive As String, strSave As String
Dim Sectors As Long, Bytes As Long, FreeC As Long, TotalC As Long, Total As Long, Freeb As Long
Dim r As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
Dim lvwItem As ListItem

    Me.AutoRedraw = True
    strSave = String(255, Chr$(0))
    ret& = GetLogicalDriveStrings(255, strSave)
    For keer = 1 To 100
        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
            sDrive = Left$(UCase(strSave), InStr(1, strSave, Chr$(0)) - 1)
            Set lvwItem = lvwGetDriveInfo.ListItems.Add(, , sDrive)
                strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
                Select Case GetDriveType(sDrive)
                    Case 2
                        lvwItem.SubItems(1) = "Removable"
                    Case 3
                        lvwItem.SubItems(1) = "Drive Fixed"
                    Case Is = 4
                        lvwItem.SubItems(1) = "Remote"
                    Case Is = 5
                        lvwItem.SubItems(1) = "Cd-Rom"
                    Case Is = 6
                        lvwItem.SubItems(1) = "Ram disk"
                    Case Else
                        lvwItem.SubItems(1) = "Unrecognized"
                End Select
    GetDiskFreeSpace sDrive, Sectors, Bytes, FreeC, TotalC
    lvwItem.SubItems(6) = Str$(Sector)
    lvwItem.SubItems(7) = Str$(Bytes)
    lvwItem.SubItems(8) = Str$(FreeC)
    lvwItem.SubItems(9) = Str$(TotalC)
    Total = rTotalc& * rSector& * rBytes&
    lvwItem.SubItems(10) = Str$(Total)
    Freeb = rFreec& * rSector& * rBytes&
    lvwItem.SubItems(11) = Str$(Freeb)
    RootPathName = sDrive
    Call GetDiskFreeSpaceEx(RootPathName, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    lvwItem.SubItems(2) = Format$(TotalBytes * 10000, "###,###,###,##0") & " bytes"
    lvwItem.SubItems(3) = Format$(TotalFreeBytes * 10000, "###,###,###,##0") & " bytes"
    lvwItem.SubItems(4) = Format$(BytesFreeToCalller * 10000, "###,###,###,##0") & " bytes"
    lvwItem.SubItems(5) = Format$((TotalBytes - TotalFreeBytes) * 10000, "###,###,###,##0") & " bytes"
    Next keer
End Sub

Private Sub Form_Load()
    Me.Caption = App.ProductName
End Sub

Private Sub Timer1_Timer()
If lblDeveloper.Left = 4620 Then
Do While True
    lblDeveloper.Left = lblDeveloper.Left - 1
    If lblDeveloper.Left = -2499 Then
        lblDeveloper.Left = 4620
    End If
    DoEvents
Loop
End If
End Sub
