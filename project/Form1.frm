VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Spy++ 2.7"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8700
   StartUpPosition =   2  '屏幕中心
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   8
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   59
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Form1.frx":08CA
      Left            =   6120
      List            =   "Form1.frx":08E0
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   55
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   26
      Left            =   3682
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   54
      Tag             =   "运行"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   25
      Left            =   6120
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   53
      Tag             =   "优化内存"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   24
      Left            =   6120
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   52
      Tag             =   "恢复进程"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   23
      Left            =   6120
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   51
      Tag             =   "挂起进程"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   22
      Left            =   1282
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   49
      Tag             =   "改变文字"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Index           =   1
      ItemData        =   "Form1.frx":090C
      Left            =   360
      List            =   "Form1.frx":090E
      TabIndex        =   47
      Top             =   3360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   21
      Left            =   6112
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   46
      Tag             =   "关闭光驱"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   20
      Left            =   3682
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   45
      Tag             =   "弹出光驱"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   19
      Left            =   1282
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   44
      Tag             =   "关闭显示"
      ToolTipText     =   "关闭显示器并退出Spy++"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   18
      Left            =   6112
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   43
      Tag             =   "运行脚本"
      ToolTipText     =   "将Spy++中按钮的文字作为命令，一行一个。"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2320
      Index           =   6
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   1600
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   4
      Left            =   3420
      ScaleHeight     =   22
      ScaleMode       =   0  'User
      ScaleWidth      =   53
      TabIndex        =   37
      Tag             =   "扩展"
      Top             =   1080
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   2625
      ScaleHeight     =   22
      ScaleMode       =   0  'User
      ScaleWidth      =   53
      TabIndex        =   36
      Tag             =   "进程"
      Top             =   1080
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   6000
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   34
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   17
      Left            =   6112
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   32
      Tag             =   "结束进程"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   16
      Left            =   3682
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   31
      Tag             =   "激活"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   15
      Left            =   1282
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   30
      Tag             =   "不可用"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   14
      Left            =   6112
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   29
      Tag             =   "可用"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   13
      Left            =   3682
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   28
      Tag             =   "不置顶"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   12
      Left            =   1282
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   27
      Tag             =   "置顶"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   11
      Left            =   6112
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   26
      Tag             =   "隐藏"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   10
      Left            =   3682
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   25
      Tag             =   "显示"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   9
      Left            =   1282
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   24
      Tag             =   "关闭"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   8
      Left            =   6112
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   23
      Tag             =   "还原"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   7
      Left            =   3682
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   22
      Tag             =   "最小化"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   6
      Left            =   1282
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   21
      Tag             =   "最大化"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Index           =   0
      ItemData        =   "Form1.frx":0910
      Left            =   4320
      List            =   "Form1.frx":0912
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   7680
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   1830
      ScaleHeight     =   22
      ScaleMode       =   0  'User
      ScaleWidth      =   53
      TabIndex        =   16
      Tag             =   "窗口"
      Top             =   1080
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   1035
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   15
      Tag             =   "样式"
      Top             =   1080
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   240
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   14
      Tag             =   "常规"
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5240
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4520
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3795
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3075
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2355
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1635
      Width           =   4815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "命令行"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   20
      Left            =   6120
      TabIndex        =   58
      Top             =   4200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "设置优先级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   19
      Left            =   6120
      TabIndex        =   56
      Top             =   3360
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   16
      Left            =   2760
      TabIndex        =   50
      Top             =   2520
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "此进程加载的模块："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   15
      Left            =   480
      TabIndex        =   48
      Top             =   2520
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   14
      Left            =   2760
      TabIndex        =   41
      Top             =   2040
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   13
      Left            =   2760
      TabIndex        =   40
      Top             =   1560
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "线程标识符"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   12
      Left            =   480
      TabIndex        =   39
      Top             =   2040
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "进程标识符(PID)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   11
      Left            =   480
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 总在最前"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   10
      Left            =   6240
      TabIndex        =   35
      Top             =   495
      Width           =   945
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "选择要对该窗口执行的操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   9
      Left            =   3120
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1800
      Index           =   8
      Left            =   840
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   7
      Left            =   2040
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "样式值："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   6
      Left            =   960
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   18
      Left            =   480
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "文件路径"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   960
      TabIndex        =   12
      ToolTipText     =   "点击打开文件路径"
      Top             =   5280
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "窗口矩形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   960
      TabIndex        =   10
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "窗口类名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   960
      TabIndex        =   8
      Top             =   3840
      Width           =   840
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFC0&
      Height          =   4335
      Index           =   1
      Left            =   240
      Top             =   1395
      Width           =   8175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFC0&
      Height          =   735
      Index           =   0
      Left            =   480
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "窗口句柄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "标题文字"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "鼠标坐标"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DA9536&
      BorderWidth     =   2
      Index           =   1
      X1              =   840
      X2              =   840
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DA9536&
      BorderWidth     =   2
      Index           =   0
      X1              =   600
      X2              =   1080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00DA9536&
      BorderWidth     =   2
      Height          =   5835
      Index           =   0
      Left            =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   8685
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请拖动左侧的指针到需要查看的窗口或控件上释放"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   17
      Left            =   1560
      TabIndex        =   0
      Top             =   285
      Width           =   2415
   End
   Begin VB.Menu mnuLst 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mFilePath 
         Caption         =   "打开所在目录"
      End
      Begin VB.Menu mCopyName 
         Caption         =   "复制模块名称"
      End
   End
   Begin VB.Menu mnuSpeak 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mPlay 
         Caption         =   "播放"
      End
      Begin VB.Menu mPause 
         Caption         =   "暂停"
      End
      Begin VB.Menu mContinue 
         Caption         =   "继续"
      End
      Begin VB.Menu mEnd 
         Caption         =   "结束"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mRun 
         Caption         =   "运行"
      End
      Begin VB.Menu mRunAsAdmin 
         Caption         =   "以管理员身份运行"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const VERSTR = "Spy++ 2.7"

Dim Pic(6) As StdPicture, OldIndex As Integer, hWindow As Long, hProc As Long, LastBtn As Integer, OnTop As Boolean, PID As Long, P As POINTAPI, pfnGetCommandLine As Long

'Get command line of ANY process!
Private Function GetProcessCommandLine(ByVal hProcess As Long) As String
    Dim pfnGetCommandLine As Long, hRemoteThread As Long, RemoteCmdLine As Long
    
    pfnGetCommandLine = GetProcAddress(GetModuleHandle("Kernel32.dll"), "GetCommandLineA")
    hRemoteThread = CreateRemoteThread(hProcess, ByVal 0&, 0&, ByVal pfnGetCommandLine, ByVal 0&, 0&, ByVal 0&)
    If hRemoteThread = 0 Then
        GetProcessCommandLine = ""
        Exit Function
    End If
    WaitForSingleObject hRemoteThread, &HFFFFFFFF
    GetExitCodeThread hRemoteThread, RemoteCmdLine
    CloseHandle hRemoteThread
    GetProcessCommandLine = String(MAX_PATH, vbNullChar)
    ReadProcessMemory hProcess, ByVal RemoteCmdLine, ByVal GetProcessCommandLine, MAX_PATH, ByVal 0&
End Function

Private Sub SetAeroGlass(Win10 As Boolean)
    Dim PrevWndStyle As Long, Margin As MARGINS, Accent As ACCENT_POLICY, Data As WCAD
        
    PrevWndStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, PrevWndStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, Me.BackColor, 0, LWA_COLORKEY
    
    If Win10 Then
        With Accent
            .AccentState = ACCENT_ENABLE_BLURBEHIND
            .AccentFlags = 0
            .AnimationId = 0
            .GradientColor = 0
        End With
        
        With Data
            .Attr = WCA_ACCENT_POLICY
            .pData = VarPtr(Accent)
            .cbData = 16
        End With
        
        SetWindowCompositionAttribute Me.hWnd, Data
        

    Else
        With Margin
            .m_Buttom = -1
            .m_Left = -1
            .m_Right = -1
            .m_Top = -1
        End With
        DwmExtendFrameIntoClientArea Me.hWnd, Margin
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, Str As String, tmpStr As String, Rgn As Long

    'SetAeroGlass True
    For i = 0 To 20
        'Label1(i).ForeColor = 0
    Next
    SetAeroGlass True
    
    Set Pic(0) = LoadResPicture(104, vbResBitmap)
    Set Pic(1) = LoadResPicture(105, vbResBitmap)
    Set Pic(2) = LoadResPicture(101, vbResBitmap)
    Set Pic(3) = LoadResPicture(102, vbResBitmap)
    Set Pic(4) = LoadResPicture(103, vbResBitmap)
    Set Pic(5) = LoadResPicture(106, vbResBitmap)
    Set Pic(6) = LoadResPicture(107, vbResBitmap)
    For i = 0 To 4
        Rgn = CreateRoundRectRgn(0, 0, 54, 27, 5, 5)
        SetWindowRgn Picture1(i).hWnd, Rgn, True
    Next
    Picture1(0).PaintPicture Pic(0), 0, 0
    Picture1(1).PaintPicture Pic(1), 0, 0
    Picture1(2).PaintPicture Pic(1), 0, 0
    Picture1(3).PaintPicture Pic(1), 0, 0
    Picture1(4).PaintPicture Pic(1), 0, 0
    Picture1(5).PaintPicture Pic(5), 0, 0
    For i = 6 To 26
        Rgn = CreateRoundRectRgn(0, 0, 93, 24, 5, 5)
        SetWindowRgn Picture1(i).hWnd, Rgn, True
        Picture1(i).PaintPicture Pic(2), 0, 0
    Next

    Text -1, True

    Label1(8).Caption = LoadResString(101)
    Text1(7).Text = "查找"
    
On Error GoTo Q
    Open "Spycfg.txt" For Input As #1
    Str = ""
    Do While Not EOF(1)
        Input #1, tmpStr
        If Str = "" Then
            Str = tmpStr
        Else
            Str = Str & vbCrLf & tmpStr
        End If
    Loop
    Close #1
    Text1(6).Text = Str
Q:
    pfnGetCommandLine = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetCommandLineA")

    OnTop = True
    OldIndex = 0
    LastBtn = 0
    hWindow = 0
    hProc = 0
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_No
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1
    Me.Width = Me.Width + 10
    
End Sub
Private Sub UpdateCombo()
    Select Case GetPriorityClass(hProc)
        Case REALTIME_PRIORITY_CLASS
             Combo1.ListIndex = 0
        Case HIGH_PRIORITY_CLASS
            Combo1.ListIndex = 1
        Case ABOVE_NORMAL
            Combo1.ListIndex = 2
        Case NORMAL_PRIORITY_CLASS
            Combo1.ListIndex = 3
        Case BELOW_NORMAL
            Combo1.ListIndex = 4
        Case IDLE_PRIORITY_CLASS
            Combo1.ListIndex = 5
    End Select
End Sub
Private Sub Combo1_Click()
    Select Case Combo1.ListIndex
        Case "0"
            If MsgBox("将进程的优先级设置为实时可能会导致系统不稳定，确实要更改吗？", vbYesNo, VERSTR) = vbYes Then
                If SetPriorityClass(hProc, REALTIME_PRIORITY_CLASS) <> 1 Then MsgBox "无法更改优先级！", vbInformation, VERSTR
            End If
        Case "1"
            If SetPriorityClass(hProc, HIGH_PRIORITY_CLASS) <> 1 Then MsgBox "无法更改优先级！", vbInformation, VERSTR
        Case "2"
            If SetPriorityClass(hProc, ABOVE_NORMAL) <> 1 Then MsgBox "无法更改优先级！", vbInformation, VERSTR
        Case "3"
            If SetPriorityClass(hProc, NORMAL_PRIORITY_CLASS) <> 1 Then MsgBox "无法更改优先级！", vbInformation, VERSTR
        Case "4"
            If SetPriorityClass(hProc, BELOW_NORMAL) <> 1 Then MsgBox "无法更改优先级！", vbInformation, VERSTR
        Case "5"
            If SetPriorityClass(hProc, IDLE_PRIORITY_CLASS) <> 1 Then MsgBox "无法更改优先级！", vbInformation, VERSTR
    End Select
    UpdateCombo
End Sub

Private Sub Form_Click()
    Picture1(0).SetFocus
    Text1(7).Text = "查找"
End Sub

Private Sub Text(ByVal Index As Integer, ByVal bVisible As Boolean)
    Dim lpRect As RECT, i As Integer

    Select Case Index
        Case -1
            With lpRect
                .Top = 0
                .Left = 0
                .Right = 53
                .Bottom = 22
            End With
            For i = 0 To 4
                DrawText Picture1(i).hDC, Picture1(i).Tag, -1, lpRect, DT_Mid
            Next
            
            With lpRect
                .Right = 93
                .Bottom = 24
            End With
            For i = 6 To 26
                DrawText Picture1(i).hDC, Picture1(i).Tag, -1, lpRect, DT_Mid
            Next
            Exit Sub
        Case 0
            For i = 0 To 5
                Label1(i).Visible = bVisible
                Text1(i).Visible = bVisible
            Next
        Case 1
            Label1(6).Visible = bVisible
            Label1(7).Visible = bVisible
            Label1(8).Visible = bVisible
            List1(0).Visible = bVisible
        Case 2
            Label1(9).Visible = bVisible
            For i = 6 To 17
                Picture1(i).Visible = bVisible
            Next
        Case 3
            For i = 11 To 16
                Label1(i).Visible = bVisible
            Next
            Label1(19).Visible = bVisible
            Label1(20).Visible = bVisible
            Text1(7).Visible = bVisible
            Text1(8).Visible = bVisible
            List1(1).Visible = bVisible
            For i = 23 To 25
                Picture1(i).Visible = bVisible
            Next
            Combo1.Visible = bVisible
        Case 4
            Text1(6).Visible = bVisible
            If bVisible Then Text1(6).SetFocus
            For i = 18 To 22
                Picture1(i).Visible = bVisible
            Next
            Picture1(26).Visible = bVisible
    End Select
    
    With lpRect
        .Left = 0
        .Top = 0
        If Index < 5 Then
            .Right = 53
            .Bottom = 22
        Else
            .Right = 93
            .Bottom = 24
        End If
    End With
    DrawText Picture1(Index).hDC, Picture1(Index).Tag, -1, lpRect, DT_Mid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Text1(6).Text <> "" Then
        Open "Spycfg.txt" For Output As #1
        Print #1, Text1(6).Text
        Close #1
    End If
End Sub

Private Sub Label1_Click(Index As Integer)
    If Index = 10 Then
        If OnTop Then
            Picture1(5).PaintPicture Pic(6), 0, 0
            SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_No
        Else
            Picture1(5).PaintPicture Pic(5), 0, 0
            SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_No
        End If
        OnTop = Not OnTop
    'ElseIf Index = 5 And Text1(5).Text <> "" Then
    '    Shell "C:\Windows\Explorer.EXE /Select," & Text1(5).Text, vbNormalFocus
    End If
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 18 Then Timer1.Enabled = True
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 18 Then
        Timer1.Enabled = False
        DoEvents
        Timer1_Timer
    End If
End Sub

Private Sub List1_Click(Index As Integer)
    List1(Index).ToolTipText = List1(Index).Text
End Sub

Private Sub List1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 1 And Button = 2 And List1(1).Text <> "" Then PopupMenu mnuLst
End Sub

Private Sub mCopyName_Click()
    Clipboard.Clear
    Clipboard.SetText List1(1).Text
End Sub

Private Sub mFilePath_Click()
    Shell "C:\Windows\Explorer.EXE /Select," & List1(1).Text, vbNormalFocus
End Sub

Private Sub mRun_Click()
On Error GoTo Q
    Shell Text1(6).Text, vbNormalFocus
    Exit Sub
Q:
    MsgBox Err.Description, vbInformation, VERSTR
    Err.Clear
End Sub

Private Sub mRunAsAdmin_Click()
    Dim shinfo As SHELLEXECUTEINFO
    With shinfo
        .cbSize = Len(shinfo)
        .lpFile = Text1(6).Text
        .lpVerb = "runas"
        .nShow = SW_SHOWNORMAL
    End With
    ShellExecuteEx shinfo
End Sub

Private Sub ClickMouse(Point As POINTAPI)
    Dim LastPos As POINTAPI
    GetCursorPos LastPos
    SetCursorPos Point.x, Point.y
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
    SetCursorPos LastPos.x, LastPos.y
End Sub

Private Sub Picture1_Click(Index As Integer)
    If hWindow = 0 And Index < 18 Then Exit Sub

    Select Case Index
        Case 6
            ShowWindow hWindow, SW_MAXIMIZE
        Case 7
            ShowWindow hWindow, SW_MINIMIZE
        Case 8
            ShowWindow hWindow, SW_RESTORE
        Case 9
            SendMessage hWindow, WM_CLOSE, 0, ByVal 0&
        Case 10
            ShowWindow hWindow, SW_SHOW
        Case 11
            ShowWindow hWindow, SW_HIDE
        Case 12
            SetWindowPos hWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_No
        Case 13
            SetWindowPos hWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_No
        Case 14
            EnableWindow hWindow, 1
        Case 15
            EnableWindow hWindow, 0
        Case 16
            BringWindowToTop hWindow
        Case 17
            If TerminateProcess(hProc, 0) = 0 Then MsgBox "无法结束此进程！", vbInformation, VERSTR
        Case 18
            Dim cmd() As String, i As Integer
            cmd = Split(Text1(6).Text, vbCrLf)
            For i = 0 To UBound(cmd)
                cmd(i) = Trim(cmd(i))
                Select Case Left(cmd(i), 2)
                    Case "单击"
                        Debug.Print CLng(Right(cmd(i), Len(cmd(i)) - 4))
                        Debug.Print CLng(Right(cmd(i), Len(cmd(i)) - 6))
                        'ClickMouse
                    Case "双击"
                        'ClickMouse
                        Sleep 200
                        'ClickMouse
                    Case Else
                        If Left(cmd(i), 2) = "等待" Then
                            Sleep CLng(CStr(Right(cmd(i), Len(cmd(i)) - 2)))
                        ElseIf Left(cmd(i), 4) = "改变文字" Then
                            SendMessage hWindow, WM_SETTEXT, 0, ByVal CStr(Right(cmd(i), Len(cmd(i)) - 5))
                        Else
                            Dim j As Integer, Found As Boolean
                            
                            Found = False
                            For j = 6 To 26
                                If cmd(i) = Picture1(j).Tag Then
                                    Picture1_Click j
                                    Found = True
                                    Exit For
                                End If
                            Next
                            If cmd(i) <> "" And Not Found Then
                                MsgBox cmd(i), vbInformation, "未知命令"
                            End If
                        End If
                End Select
            Next
        Case 19
            SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal 2&
            End
        Case 20
            mciSendString "Set CDAudio Door Open", vbNullString, 0, 0
        Case 21
            mciSendString "Set CDAudio Door Closed", vbNullString, 0, 0
        Case 22
            SendMessage hWindow, WM_SETTEXT, 0, ByVal Text1(6).Text
        Case 23
            If NtSuspendProcess(hProc) <> 0 Then MsgBox "无法挂起此进程！", vbInformation, VERSTR
        Case 24
            If NtResumeProcess(hProc) <> 0 Then MsgBox "无法恢复此进程！", vbInformation, VERSTR
        Case 25
            If SetProcessWorkingSetSize(hProc, -1, -1) = 0 Then MsgBox "无法减少此进程的内存占用量！", vbInformation, VERSTR
        Case 26
            PopupMenu mnuRun
    End Select
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Index
        Case Is < 5
            If Index = OldIndex Then Exit Sub
            Picture1(Index).PaintPicture Pic(0), 0, 0
            Picture1(OldIndex).PaintPicture Pic(1), 0, 0
            Text Index, True
            Text OldIndex, False
            OldIndex = Index
        Case 5
            Label1_Click 10
        Case Else
            Picture1(Index).PaintPicture Pic(4), 0, 0
            Text Index, True
    End Select
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 5 And LastBtn = 0 Then
        Picture1(Index).PaintPicture Pic(3), 0, 0
        Text Index, True
        LastBtn = Index
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LastBtn Then
        Picture1(LastBtn).PaintPicture Pic(2), 0, 0
        Text LastBtn, True
        LastBtn = 0
    End If
End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > 5 Then
        Picture1(Index).PaintPicture Pic(2), 0, 0
        Text Index, True
    End If
End Sub

Private Sub Text1_Change(Index As Integer)
    If Index <> 7 Then Exit Sub
    If Text1(7).Text = "" Then
        List1(1).ListIndex = -1
        Exit Sub
    End If
    Dim s As String, sFind As String, i As Integer, LenFind As Integer, LenS, iLoop As Integer
    sFind = UCase(Text1(7).Text)
    LenFind = Len(sFind)
    For iLoop = 0 To List1(1).ListCount - 1
        s = UCase(List1(1).List(iLoop))
        LenS = Len(s)
        For i = 6 To LenS
            If Mid(s, LenS - i + 1, 1) = "\" Then
                If Mid(s, LenS - i + 2, LenFind) = sFind Then
                    List1(1).ListIndex = iLoop
                    Exit Sub
                End If
                Exit For
            End If
        Next
    Next
    List1(1).ListIndex = -1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Index = 7 Then Text1(7).Text = ""
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyEscape And Index = 7 Then
        Text1(7).Text = "查找"
        List1(1).ListIndex = 0
        List1(1).ListIndex = -1
        Picture1(0).SetFocus
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index = 7 Then Text1(7).Text = "查找"
End Sub

Private Sub Timer1_Timer()
    Dim lpRect As RECT, tmpRect As RECT, M As MODULEENTRY32
    Dim lTemp As Long, hDC As Long, lStyle As Long, sTemp As String
    Static OldhWnd As Long

    GetCursorPos P
    Text1(0).Text = "(" & P.x & ", " & P.y & ")"
    hWindow = WindowFromPoint(P.x, P.y)
    GetWindowRect hWindow, lpRect
    lTemp = GetWindow(hWindow, GW_CHILD)

    Do While lTemp
        GetWindowRect lTemp, tmpRect
        If P.x >= tmpRect.Left And P.x <= tmpRect.Right And P.y >= tmpRect.Top And P.y <= tmpRect.Bottom Then
            hWindow = lTemp
            lpRect = tmpRect
            Exit Do
        End If
        lTemp = GetWindow(lTemp, GW_HWNDNEXT)
    Loop

    If hWindow = OldhWnd Then Exit Sub
    
    Text1(2).Text = CStr(hWindow)
    
    With lpRect
        Text1(4).Text = "(" & .Left & ", " & .Top & ") - (" & .Right & ", " & .Bottom & ")  " & .Right - .Left & " × " & .Bottom - .Top
    End With

    hDC = GetDC(0)
    lTemp = SetROP2(hDC, vbNotXorPen)
    lStyle = CreatePen(PS_SOLID, 3, 0)
    lStyle = SelectObject(hDC, lStyle)
    If OldhWnd Then
        GetWindowRect OldhWnd, tmpRect
        Rectangle hDC, tmpRect.Left, tmpRect.Top, tmpRect.Right, tmpRect.Bottom
    End If
    Rectangle hDC, lpRect.Left, lpRect.Top, lpRect.Right, lpRect.Bottom
    SetROP2 hDC, lTemp
    lStyle = SelectObject(hDC, lStyle)
    DeleteObject lStyle
    ReleaseDC 0, hDC
    OldhWnd = hWindow
    
    lTemp = SendMessage(hWindow, WM_GETTEXTLENGTH, 0, ByVal 0&) + 1
    If lTemp < 260 Then lTemp = 260 'For class name
    sTemp = String(lTemp, vbNullChar)
    SendMessage hWindow, WM_GETTEXT, lTemp, ByVal sTemp
    Text1(1).Text = sTemp
    GetClassName hWindow, sTemp, lTemp
    Text1(3).Text = sTemp

    Label1(14).Caption = CStr(GetWindowThreadProcessId(hWindow, PID))
    Label1(13).Caption = CStr(PID)
    
    If hProc <> 0 Then
        CloseHandle hProc
    End If
    hProc = OpenProcess(PROCESS_ALL_ACCESS, 0, PID)
    
    If hProc <> 0 Then
        Text1(8).Text = GetProcessCommandLine(hProc)
        Text1(5).Text = Text1(8).Text
        UpdateCombo
    End If
    
    lStyle = GetWindowLong(hWindow, GWL_STYLE)
    Label1(7).Caption = "基本样式" & lStyle
    List1(0).Clear

    If lStyle And WS_VSCROLL Then List1(0).AddItem "WS_VSCROLL"
    If lStyle And WS_VISIBLE Then List1(0).AddItem "WS_VISIBLE"
    If lStyle And WS_TILEDWINDOW Then List1(0).AddItem "WS_TILEDWINDOW"
    If lStyle And WS_TILED Then List1(0).AddItem "WS_TITLED"
    If lStyle And WS_THICKFRAME Then List1(0).AddItem "WS_THICKFRAME"
    If lStyle And WS_TABSTOP Then List1(0).AddItem "WS_TABSTOP"
    If lStyle And WS_SYSMENU Then List1(0).AddItem "WS_SYSMENU"
    If lStyle And WS_SIZEBOX Then List1(0).AddItem "WS_SIZEBOX"
    If lStyle And WS_POPUPWINDOW Then List1(0).AddItem "WS_POPUPWINDOW"
    If lStyle And WS_POPUP Then List1(0).AddItem "WS_POPUP"
    If lStyle And WS_OVERLAPPEDWINDOW Then List1(0).AddItem "WS_OVERLAPPEDWINDOW"
    If lStyle And WS_OVERLAPPED Then List1(0).AddItem "WS_OVERLAPPED"
    If lStyle And WS_MINIMIZEBOX Then List1(0).AddItem "WS_MINIMIZEBOX"
    If lStyle And WS_MINIMIZE Then List1(0).AddItem "WS_MINIZE"
    If lStyle And WS_MAXIMIZEBOX Then List1(0).AddItem "WS_MAXIMIZEBOX"
    If lStyle And WS_MAXIMIZE Then List1(0).AddItem "WS_MAXIMIZE"
    If lStyle And WS_ICONIC Then List1(0).AddItem "WS_ICONIC"
    If lStyle And WS_HSCROLL Then List1(0).AddItem "WS_HCROLL"
    If lStyle And WS_GROUP Then List1(0).AddItem "WS_GROUP"
    If lStyle And WS_DLGFRAME Then List1(0).AddItem "WS_DLGFRAME"
    If lStyle And WS_DISABLED Then List1(0).AddItem "WS_DISABLED"
    If lStyle And WS_CLIPSIBLINGS Then List1(0).AddItem "WS_CLIPSIBLINGS"
    If lStyle And WS_CLIPCHILDREN Then List1(0).AddItem "WS_CLIPCHILEREN"
    If lStyle And WS_CHILDWINDOW Then List1(0).AddItem "WS_CHILDWINDOW"
    If lStyle And WS_CHILD Then List1(0).AddItem "WS_CHILD"
    If lStyle And WS_CAPTION Then List1(0).AddItem "WS_CAPTION"
    If lStyle And WS_BORDER Then List1(0).AddItem "WS_BORDER"

    If lStyle And ES_AUTOHSCROLL Then List1(0).AddItem "ES_AUTOHSCROLL"
    If lStyle And ES_AUTOVSCROLL Then List1(0).AddItem "ES_AUTOVSCROLL"
    If lStyle And ES_CENTER Then List1(0).AddItem "ES_CENTER"
    If lStyle And ES_LEFT Then List1(0).AddItem "ES_LEFT"
    If lStyle And ES_LOWERCASE Then List1(0).AddItem "ES_LOWERCASE"
    If lStyle And ES_MULTILINE Then List1(0).AddItem "ES_MULTILINE"
    If lStyle And ES_NOHIDESEL Then List1(0).AddItem "ES_NOHIDESEL"
    If lStyle And ES_OEMCONVERT Then List1(0).AddItem "ES_OEMCONVERT"
    If lStyle And ES_PASSWORD Then List1(0).AddItem "ES_PASSWORD"
    If lStyle And ES_READONLY Then List1(0).AddItem "ES_READONLY"
    If lStyle And ES_RIGHT Then List1(0).AddItem "ES_RIGHT"
    If lStyle And ES_UPPERCASE Then List1(0).AddItem "ES_UPPERCASE"
    If lStyle And ES_WANTRETURN Then List1(0).AddItem "ES_WANTRETURN"

    lStyle = GetWindowLong(hWindow, GWL_EXSTYLE)
    Label1(7).Caption = Label1(7).Caption & "  扩展样式" & lStyle

    If lStyle And WS_EX_TRANSPARENT Then List1(0).AddItem "WS_EX_TRANSPARENT"
    If lStyle And WS_EX_TOPMOST Then List1(0).AddItem "WS_EX_TOPMOST"
    If lStyle And WS_EX_NOPARENTNOTIFY Then List1(0).AddItem "WS_EX_NOPARENTNOTIFY"
    If lStyle And WS_EX_DLGMODALFRAME Then List1(0).AddItem "WS_EX_DLGMODALFRAME"
    If lStyle And WS_EX_ACCEPTFILES Then List1(0).AddItem "WS_EX_ACCPTFILES"
    If lStyle And WS_EX_LAYERED Then List1(0).AddItem "WS_EX_LAYERED"
    
    List1(1).Clear
    M.dwSize = Len(M)
    lStyle = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, PID)
    If Module32First(lStyle, M) Then
        Do While Module32Next(lStyle, M) <> 0
            GetModuleFileName M.hModule, sTemp, 260
            List1(1).AddItem sTemp
        Loop
    End If
    CloseHandle lStyle
    Label1(16).Caption = List1(1).ListCount & "个"

    DoEvents
End Sub
