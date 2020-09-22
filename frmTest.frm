VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Internet File ActiveX Control Ver 1.1"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraCredits 
      Caption         =   "&Credits"
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   8295
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   32
         Text            =   "frmTest.frx":0000
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Frame fraDownloadProgress 
      Caption         =   "Download Progress"
      Height          =   1335
      Left            =   4560
      TabIndex        =   21
      Top             =   3960
      Width           =   3855
      Begin ComctlLib.ProgressBar ctlProgress 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblBytesRead 
         Caption         =   ": 0 bytes"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblBytesReadLabel 
         Caption         =   "Bytes read"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdRequestInformation 
      Caption         =   "Get file information"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame fraFileInformation 
      Caption         =   "Fileinformation"
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   4335
      Begin VB.Label lblFilesize 
         Caption         =   ": Not available"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblLastModified 
         Caption         =   ": Not available"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblFileExists 
         Caption         =   ": Not available"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label FilesizeLabel 
         Caption         =   "Filesize"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLastModifiedLabel 
         Caption         =   "Last modified"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblFileExistsLabel 
         Caption         =   "File exists"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraProxy 
      Height          =   2055
      Left            =   4560
      TabIndex        =   9
      Top             =   1800
      Width           =   3855
      Begin VB.TextBox txtProxyServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtProxyPort 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2880
         TabIndex        =   10
         Top             =   1440
         Width           =   720
      End
      Begin VB.ComboBox cboConnectType 
         Height          =   315
         ItemData        =   "frmTest.frx":00EA
         Left            =   240
         List            =   "frmTest.frx":00F7
         Style           =   2  'Dropdown-Liste
         TabIndex        =   25
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblProxyServer 
         Caption         =   "Proxy"
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   2340
      End
      Begin VB.Label lblProxyPort 
         Caption         =   "Port"
         Height          =   240
         Left            =   2880
         TabIndex        =   12
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblConnectType 
         Caption         =   "Connection type"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraDownload 
      Caption         =   "Download"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
      Begin VB.TextBox txtSiteUser 
         Height          =   330
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   2145
      End
      Begin VB.TextBox txtSitePassword 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "'"
         TabIndex        =   27
         Top             =   1440
         Width           =   1545
      End
      Begin VB.TextBox txtPort 
         Height          =   330
         Left            =   3480
         TabIndex        =   6
         Text            =   "80"
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox txtUrl 
         Height          =   330
         Left            =   240
         TabIndex        =   5
         Text            =   "http://www.crossalizer.com/download/crossxr108de.exe"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblSiteUser 
         Caption         =   "Username for site"
         Height          =   240
         Left            =   240
         TabIndex        =   30
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label lblSitePassword 
         Caption         =   "Password for site"
         Height          =   240
         Left            =   2520
         TabIndex        =   29
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label lblPort 
         Caption         =   "Port"
         Height          =   240
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblUrl 
         Caption         =   "Url"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2700
      End
   End
   Begin InternetFile_ActiveX.InternetFile ctlIFile 
      Left            =   1440
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdResumeDownload 
      Caption         =   "Resume Download"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdStopDownload 
      Caption         =   "Stop Download"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdStartDownload 
      Caption         =   "Start Download"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboConnectType_Change()
  EnableControls (cboConnectType.ListIndex = 1)
End Sub

Private Sub cboConnectType_Click()
  EnableControls (cboConnectType.ListIndex = 1)
End Sub

Private Sub cboConnectType_KeyPress(KeyAscii As Integer)
  EnableControls (cboConnectType.ListIndex = 1)
End Sub

Private Sub cmdRequestInformation_Click()
  With ctlIFile
    SetDownloadParameter
    .RequestInformation
    lblFileExists.Caption = IIf(.FileExists, ": True", ": False")
    lblFilesize.Caption = IIf(.FileSize > 0, ": " & .FileSize & " bytes", ": Not available")
    lblLastModified.Caption = IIf(.LastModified <> "", ": " & .LastModified, ": Not Available")
  End With
End Sub

Private Sub cmdResumeDownload_Click()
  SetDownloadParameter
  ctlIFile.LocalFile = "c:\test.bin"
  ctlIFile.ResumeDownload
End Sub

Private Sub cmdStartDownload_Click()
  SetDownloadParameter
  ctlIFile.LocalFile = "c:\test.bin"
  ctlIFile.StartDownload
End Sub

Private Sub cmdStopDownload_Click()
  If MsgBox("Are you sure, that you want to stop the download?", vbQuestion + vbYesNo) = vbYes Then
    ctlIFile.CancelDownload
  End If
End Sub

Private Sub ctlIFile_DownloadCancelled(lPosition As Long)
  ctlProgress.Value = 0
  lblBytesRead.Caption = ": 0 bytes"
End Sub

Private Sub ctlIFile_DownloadComplete()
  MsgBox "Download complete", vbInformation
End Sub

Private Sub ctlIFile_DownloadError(sErrorDescription As String)
  MsgBox "The following Error occured:" & vbCrLf & sErrorDescription, vbExclamation
End Sub

Private Sub ctlIFile_DownloadProgress(lBytesRead As Long)
  If ctlIFile.FileSize > -1 Then
    ctlProgress.Max = ctlIFile.FileSize
    ctlProgress.Value = lBytesRead
    lblBytesRead.Caption = ": " & Format$(lBytesRead, "#,###,###") & " bytes of " & Format$(ctlIFile.FileSize, "#,###,###") & " bytes"
  Else
    lblBytesRead.Caption = ": " & Format$(lBytesRead, "#,###,###") & " bytes"
  End If
End Sub

Private Sub EnableControls(fEnabled As Boolean)
  txtProxyServer.Enabled = fEnabled
  txtProxyPort.Enabled = fEnabled
  txtProxyServer.Locked = Not fEnabled
  txtProxyPort.Locked = Not fEnabled
End Sub

Private Sub SetDownloadParameter()
  With ctlIFile
    .Url = txtUrl
    .Port = txtPort
    .SiteUser = txtSiteUser
    .SitePassword = txtSitePassword
    .ConnectType = cboConnectType.ListIndex
    .ProxyServer = txtProxyServer
    .ProxyPort = Int(IIf(Len(txtProxyPort.Text) > 0, txtProxyPort.Text, 0))
  End With
End Sub

Private Sub Form_Load()
  cboConnectType.ListIndex = 0
End Sub
