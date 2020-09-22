VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Sub Form_Load()
Dim sDrive As String
Dim strSave As String

    Me.AutoRedraw = True
    strSave = String(255, Chr$(0))
    ret& = GetLogicalDriveStrings(255, strSave)
    For keer = 1 To 100
        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
        sDrive = "Drive " & Left$(UCase(strSave), InStr(1, strSave, Chr$(0)) - 1)
        MsgBox sDrive
        Select Case sDrive
            case
        End Select
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
    Next keer
End Sub

Select Case fsoDrive.DriveType
            Case 0: sDriveType = "Unknown"
            Case 1: sDriveType = "Removable Drive"
                sDrive = sDrive & sDriveType
                If fsoDrive.IsReady Then
                   'Call ScanDrive(sDrive) ' Scan drives
                End If
            Case 2: sDriveType = "Fixed Disk"
                sDrive = sDrive & sDriveType
                If fsoDrive.IsReady Then
                    If g_bStopScanPressed Then Exit Sub
                    lvwHardDrives.ListItems.Add , , fsoDrive.DriveLetter
                    If g_bStopScanPressed Then Exit Sub
                End If
            Case 3: sDriveType = "Remote Disk"
            Case 4: sDriveType = "CDROM Drive"
            Case 5: sDriveType = "RAM Disk"
        End Select
