VERSION 5.00
Begin VB.Form fWindows 
   BackColor       =   &H00FFF8F8&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Your Windows"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   ForeColor       =   &H00FFFFFF&
   Icon            =   "fKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   12
      Left            =   1125
      TabIndex        =   25
      Top             =   660
      Width           =   660
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   12
      Left            =   1950
      TabIndex        =   24
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   11
      Left            =   690
      TabIndex        =   23
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Owner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   10
      Left            =   1230
      TabIndex        =   22
      Top             =   3900
      Width           =   555
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   9
      Left            =   540
      TabIndex        =   21
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   8
      Left            =   900
      TabIndex        =   20
      Top             =   2820
      Width           =   885
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Processor Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   7
      Left            =   420
      TabIndex        =   19
      Top             =   3540
      Width           =   1365
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Update Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   6
      Left            =   810
      TabIndex        =   18
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "System Directory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   5
      Left            =   270
      TabIndex        =   17
      Top             =   4620
      Width           =   1515
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "from"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   4
      Left            =   1410
      TabIndex        =   16
      Top             =   1740
      Width           =   375
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Installed on"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   3
      Left            =   795
      TabIndex        =   15
      Top             =   1380
      Width           =   990
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Build"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   2
      Left            =   1350
      TabIndex        =   14
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Update Level"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   1
      Left            =   675
      TabIndex        =   13
      Top             =   2100
      Width           =   1110
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Product Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   12
      Top             =   3180
      Width           =   1485
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   11
      Left            =   1950
      TabIndex        =   11
      Top             =   4620
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   10
      Left            =   1950
      TabIndex        =   10
      Top             =   1740
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   9
      Left            =   1950
      TabIndex        =   9
      Top             =   4260
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   8
      Left            =   1950
      TabIndex        =   8
      Top             =   3900
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   7
      Left            =   1950
      TabIndex        =   7
      Top             =   300
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   6
      Left            =   1950
      TabIndex        =   6
      Top             =   2820
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   5
      Left            =   1950
      TabIndex        =   5
      Top             =   1020
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   4
      Left            =   1950
      TabIndex        =   4
      Top             =   3540
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   3
      Left            =   1950
      TabIndex        =   3
      Top             =   2100
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   2
      Left            =   1950
      TabIndex        =   2
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   1
      Left            =   1950
      TabIndex        =   1
      Top             =   1380
      Width           =   135
   End
   Begin VB.Label lbInfo 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   1950
      TabIndex        =   0
      Top             =   3180
      Width           =   135
   End
End
Attribute VB_Name = "fWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002
Private Const RegKey                As String = "SOFTWARE\MICROSOFT\Windows NT\CurrentVersion"
Private Const RegItem1              As String = "DigitalProductId"          'binary
Private Const RegItem2              As String = "InstallDate"               'dword
Private Const RegItem3              As String = "BuildLab"                  'sz
Private Const RegItem4              As String = "CSDVersion"                'sz
Private Const RegItem5              As String = "CurrentType"               'sz
Private Const RegItem6              As String = "CurrentBuildNumber"        'sz
Private Const RegItem7              As String = "ProductId"                 'sz
Private Const RegItem8              As String = "ProductName"               'sz
Private Const RegItem9              As String = "RegisteredOwner"           'sz
Private Const RegItem10             As String = "RegisteredOrganization"    'sz
Private Const RegItem11             As String = "SourcePath"                'sz
Private Const RegItem12             As String = "SystemRoot"                'sz
Private Const RegItem13             As String = "CurrentVersion"            'sz
Private Const REG_SZ                As Long = 1
Private Const REG_DWORD             As Long = 4
Private Const REG_BINARY            As Long = 3
Private Const ERROR_SUCCESS         As Long = 0

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const CS_DROPSHADOW         As Long = &H20000
Private Const GCL_STYLE             As Long = -26

Private RegItems(3 To 13)           As String

Private Const XlatProdId            As String = "BCDFGHJKMPQRTVWXY2346789"
Private Const NotRegistered         As String = "Unknown"

Private Sub Form_Load()

  Dim n         As Long

    RegItems(3) = RegItem3
    RegItems(4) = RegItem4
    RegItems(5) = RegItem5
    RegItems(6) = RegItem6
    RegItems(7) = RegItem7
    RegItems(8) = RegItem8
    RegItems(9) = RegItem9
    RegItems(10) = RegItem10
    RegItems(11) = RegItem11
    RegItems(12) = RegItem12
    RegItems(13) = RegItem13

    For n = 1 To 13
        lbInfo(n - 1) = WinProps(n)
    Next n
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    SetClassLong hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW

End Sub

Private Function WinProps(Which As Long) As String

  Dim hKey              As Long
  Dim ProdID(0 To 164)  As Byte
  Dim RegString         As String
  Dim i                 As Long
  Dim j                 As Long
  Dim k                 As Long

    If RegOpenKey(HKEY_LOCAL_MACHINE, RegKey, hKey) = ERROR_SUCCESS Then
        k = 255
        Select Case Which
          Case 1
            If RegQueryValueEx(hKey, RegItem1, 0&, REG_BINARY, ProdID(0), k) = ERROR_SUCCESS Then
                For i = 1 To 25
                    k = 0
                    For j = 66 To 52 Step -1
                        k = k * 256 Xor CLng(ProdID(j))
                        ProdID(j) = k \ 24
                        k = k Mod 24
                    Next j
                    WinProps = IIf(i Mod 5, "", "-") & Mid$(XlatProdId, k + 1, 1) & WinProps
                Next i
                WinProps = Mid$(WinProps, 2)
              Else 'NOT REGQUERYVALUEEX(HKEY,...
                WinProps = NotRegistered
            End If
          Case 2
            If RegQueryValueEx(hKey, RegItem2, 0&, REG_DWORD, i, k) = ERROR_SUCCESS Then
                WinProps = Format$(DateAdd("s", i, DateSerial(1970, 1, 1)), "Long Date") & " at " & _
                           Format$(DateAdd("s", i, DateSerial(1970, 1, 1)), "Long Time") & " "
              Else 'NOT REGQUERYVALUEEX(HKEY,...
                WinProps = NotRegistered
            End If
          Case Else
            RegString = String$(k, 0)
            If RegQueryValueEx(hKey, RegItems(Which), 0&, REG_SZ, ByVal RegString, k) = ERROR_SUCCESS Then
                WinProps = Left$(RegString, k + (Mid$(RegString, k, 1) = Chr$(0)))
              Else 'NOT REGQUERYVALUEEX(HKEY,...
                WinProps = NotRegistered
            End If
        End Select
        RegCloseKey hKey
      Else 'NOT REGOPENKEY(HKEY_LOCAL_MACHINE,...
        WinProps = NotRegistered
    End If

End Function

':) Ulli's VB Code Formatter V2.20.2 (2006-Feb-05 12:14)  Decl: 34  Code: 76  Total: 110 Lines
':) CommentOnly: 0 (0%)  Commented: 17 (15,5%)  Empty: 13 (11,8%)  Max Logic Depth: 6
