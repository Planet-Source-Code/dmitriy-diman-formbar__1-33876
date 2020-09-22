VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{851970E1-6FD3-4CD4-8805-997B249CA7F8}#3.0#0"; "FormBarAX.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin FormBarAX.FormBar FormBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   5970
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   741
      BC              =   -2147483633
      BBC             =   -2147483633
      MW              =   2000
      MH              =   420
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5445
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0168
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1399
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1470
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   10290
      TabIndex        =   0
      Top             =   0
      Width           =   10350
      Begin VB.CommandButton Command1 
         Caption         =   "Forms"
         Height          =   420
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1770
      End
   End
   Begin VB.Menu popmenuUC 
      Caption         =   "popmenuUC"
      Visible         =   0   'False
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Dim nF As frmTest
    Dim fff As ListImage
    
    
    For i = 1 To 5
    Set nF = New frmTest
    

    Set fff = ImageList1.ListImages(i)
    
    nF.Caption = nF.Caption & "NR: " & i
    FormBar1.Add_Form nF, nF.Caption, fff.Picture
    nF.Show
    Next i
    
    
    FormBar1.setMenu Me.popmenuUC
    
End Sub

Private Sub mnuclose_Click()
    Unload Me.ActiveForm
    
End Sub
