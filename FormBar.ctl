VERSION 5.00
Begin VB.UserControl FormBar 
   Alignable       =   -1  'True
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   ScaleHeight     =   1095
   ScaleWidth      =   9960
   ToolboxBitmap   =   "FormBar.ctx":0000
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   720
      ScaleHeight     =   510
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   540
      Width           =   600
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdButton 
      Height          =   420
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   645
      Visible         =   0   'False
   End
   Begin VB.Image imgScale 
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   1395
      Stretch         =   -1  'True
      Top             =   0
      Width           =   645
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FormBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'USER CONTROL
Private cForms As New Collection
Private fImages As New Collection
Private fButtons As New Collection

Private pOpMenu As Menu
Private PressedIndex As Long


Private maxW As Long 'MAX WIDTH OF A BUTTON
Private maxH As Long 'MAX HEIGHT OF A BUTTON



Private FONTNAME As String
Private FONTSIZE As Integer
Private FONTBOLD As Boolean
Private FONTUNDER As Boolean






Public Function Add_Form(frm, Caption As String, Image As IPictureDisp) 'DON'T WRITE AS Form. Will complain when compiling
    Dim clsH As clsFormHandler, curCount As Long
    Set clsH = New clsFormHandler
    
    
    clsH.setOwner Me
    
    clsH.AssignFormToClass frm
    
    cForms.Add clsH, CStr(frm.hDC)
    fImages.Add Image, CStr(frm.hDC)
    
    curCount = cmdButton.Count
    Load cmdButton(curCount)
    fButtons.Add cmdButton(curCount), CStr(frm.hDC)
    DoEvents
    RepaintBar
    
End Function

Public Function setMenu(tM)
    Set pOpMenu = tM
End Function

'theese functions are called from the clsHandlers when some form_event occurs. They are not to be called manually
'*************************************************************************************************
Public Sub Forms_Activate_Event(fHDC As Long)
    'MsgBox fHDC
    'PRESS DOWN THE BUTTON AND RELEASE ALL OTHERS
    
    Dim tmpButt As CommandButton, i As Long
    For i = 1 To fButtons.Count
        'first release all
        Set tmpButt = fButtons(i)
        SendMessage tmpButt.hwnd, BM_SETSTATE, BUTTON_UNPRESSED, 0
    Next i
    'Now press the right one
    Set tmpButt = fButtons(CStr(fHDC))
    SendMessage tmpButt.hwnd, BM_SETSTATE, BUTTON_PRESSED, 0
    PressedIndex = tmpButt.Index
    
End Sub

Public Sub Forms_Unload_Event(fHDC As Long)
    cForms.Remove CStr(fHDC)
    fImages.Remove CStr(fHDC)
    fButtons.Remove CStr(fHDC)
    
    RepaintBar
End Sub

'*************************************************************************************************

Private Function RepaintBar(Optional Resize As Boolean = False)
    Dim i As Long, newWidth As Long, newHeight As Long, bCap As String, tCap As String, curCount As Long
    
    If Not Resize Then
        Set fButtons = New Collection
    
        'unload all buttons but Index(0)
        For i = 1 To cmdButton.Count - 1
            Unload cmdButton(i)
        Next i
    End If
    
    'if no more forms...
    If cForms.Count = 0 Then Exit Function
    
    'ME.width - 5px from left - 5px from right

    newWidth = (UserControl.Width - UserControl.ScaleX(10, vbPixels, vbTwips)) / cForms.Count
    If newWidth > maxW Then newWidth = maxW
    
    newHeight = UserControl.Height - 50 - 50 '(In Twips)
    If newHeight > maxH Then newHeight = maxH
    
    'reload all buttons and assign forms hDc as key
    Dim tmpCLSH As clsFormHandler
    If Not Resize Then
        For i = 1 To cForms.Count
            curCount = cmdButton.Count
            Load cmdButton(curCount)
        
            cmdButton(curCount).Width = newWidth
        
            Set tmpCLSH = cForms(i)
            fButtons.Add cmdButton(curCount), CStr(tmpCLSH.getHDC)
        Next i
    Else
        For i = 1 To cForms.Count
            Set cmdButton(i).Picture = Nothing
            Set cmdButton(i).DownPicture = Nothing
            Set cmdButton(i).DisabledPicture = Nothing
            cmdButton(i).Width = newWidth
            cmdButton(i).Refresh
            UserControl.Refresh
        Next i
    End If
    
    
    


    'PAINT BUTTONS
    picDraw.AutoRedraw = True
    picDraw.ScaleMode = vbTwips 'ONLY TO SET CORRECT WIDTH AND HEIGHT CASE BUTTONS ARE STILL SCALED IN TWIPS
    picDraw.Width = newWidth
    picDraw.Height = newHeight
    picDraw.ScaleMode = vbPixels 'BACK TO PIXELS
    
    
    
    Dim tmpButt, curY As Long, curLeft As Long, capWidth As Long
    With picDraw.Font
        .Name = FONTNAME: .Size = FONTSIZE: .Bold = FONTBOLD: .Underline = FONTUNDER
    End With
    'after we set the font, we can calculate Y position for drawing.
    curY = (picDraw.ScaleY(picDraw.Height, vbTwips, vbPixels) - picDraw.TextHeight("Test String")) / 2 - 1
    curLeft = 10
    For i = 1 To cForms.Count
        Set tmpButt = fButtons.Item(i)
        Set tmpCLSH = cForms(i)
        
        tmpButt.Caption = "" 'NO TEXTUAL CAPTION AT ALL. ALL WILL BE DRAWN GRAPHICALLY
        picDraw.Picture = LoadPicture  'clear picdraw picture
        DoEvents
        
        picDraw.PaintPicture fImages.Item(i), 5, curY, 16, 16
        
        'Truncate caption if too long
        capWidth = picDraw.TextWidth(tmpCLSH.getCaption)
        tCap = tmpCLSH.getCaption
        If capWidth > (picDraw.ScaleX(picDraw.Width, vbTwips, vbPixels) - 5 - 16 - 3 - 3) Then '(picdraw.Width - 5 - 16 - 3 - 3) as above but - additional 3 px to create a spece from the right side
            bCap = tmpCLSH.getCaption: tCap = ""
            While picDraw.TextWidth(tCap & "...") < (picDraw.ScaleX(picDraw.Width, vbTwips, vbPixels) - 5 - 16 - 3 - 3)
                tCap = tCap & Mid(bCap, Len(tCap) + 1, 1) 'GET ONE CHAR AT THE TIME
            Wend 'WHEN LOOP ENDS Length of tcap will be ONE CHAR TOO MUCH
            If Not tCap = "" Then tCap = Mid(tCap, 1, Len(tCap) - 1) & "..." 'ADD ...
        End If
        
        picDraw.CurrentX = 5 + 16 + 3 '5 px fro left, 1 px image width and 3 px space
        picDraw.CurrentY = curY 'the same Y value for text

        picDraw.Print tCap 'draw picture
        tmpButt.Picture = picDraw.Image 'add it to the Button
        
        tmpButt.Top = 10
        tmpButt.Left = curLeft
        curLeft = i * newWidth + (10 * i) '10 * i =  some space beetween the buttons
        tmpButt.Refresh
        If Not tmpButt.Visible Then tmpButt.Visible = True
    Next i
    
    
    SetActiveButton
End Function

Private Sub cmdButton_Click(Index As Integer)
    Dim tmpCLSH
    Set tmpCLSH = cForms(Index)
    
    
    If PressedIndex <> Index Then
        If tmpCLSH.getFormState = vbMinimized Then
            tmpCLSH.setFormState vbNormal
        End If
        tmpCLSH.setActive
    Else 'BUTTON IS ALREADY PRESSED 'FEATURE... TRY TO ADD SOME COdE TO RESTORE FORM TO ITS PREVIOUS STATE 8IF It was MAX WHEN IT GOT MIN THEN WHEN PRESSING THE BUTTON IT GETS MAX AGAIN
        If tmpCLSH.getFormState = vbMinimized Then
            tmpCLSH.setFormState vbNormal
            SendMessage cmdButton(Index).hwnd, BM_SETSTATE, BUTTON_PRESSED, 0
            PressedIndex = Index
        Else
            tmpCLSH.setFormState vbMinimized
            tmpCLSH.setActive
        End If
    End If
    
    
    
    
End Sub

Private Sub UserControl_Initialize()
    'DEBUG
    FONTNAME = "MS Sans Serif"
    FONTSIZE = 8
    FONTBOLD = False: FONTUNDER = False
    maxW = 2000
    maxH = 420
End Sub

Private Sub SetActiveButton()
    Dim tmpCLSH, tmpButt, i
    DoEvents
    For i = 1 To cForms.Count
        Set tmpCLSH = cForms(i)
        If tmpCLSH.getFState = True Then 'ACTIVE
            Set tmpButt = fButtons(CStr(tmpCLSH.getHDC))
            SendMessage tmpButt.hwnd, BM_SETSTATE, BUTTON_PRESSED, 0
            PressedIndex = tmpButt.Index
        Else
            Set tmpButt = fButtons(CStr(tmpCLSH.getHDC))
            SendMessage tmpButt.hwnd, BM_SETSTATE, BUTTON_UNPRESSED, 0
        End If
    Next i
End Sub




Private Sub UserControl_Resize()
    If UserControl.Height < maxH Then UserControl.Height = maxH 'ONLY AT DESIGN TIME

    DoEvents
    If Not UserControl.Width = 0 Then
        SendMessage cmdButton(PressedIndex).hwnd, BM_SETSTATE, BUTTON_UNPRESSED, 0
        RepaintBar True
    End If
End Sub













Public Property Let BackColor(nColor As OLE_COLOR)
    UserControl.BackColor = nColor
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let ButtonBackColor(nColor As OLE_COLOR)
    picDraw.BackColor = nColor
End Property
Public Property Get ButtonBackColor() As OLE_COLOR
    ButtonBackColor = picDraw.BackColor
End Property

'WIDTH
Public Property Let MaxButtonWidth(mW As Long)
    maxW = mW
End Property
Public Property Get MaxButtonWidth() As Long
    MaxButtonWidth = maxW
End Property
'HEIGHT
Public Property Let MaxButtonHeight(mH As Long)
    maxH = mH
End Property
Public Property Get MaxButtonHeight() As Long
    MaxButtonHeight = maxH
End Property











Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BC", &H8000000F)
    
    picDraw.BackColor = PropBag.ReadProperty("BBC", &H8000000F)
    cmdButton(0).BackColor = PropBag.ReadProperty("BBC", &H8000000F)
    
    maxW = PropBag.ReadProperty("MW", 2000)
    maxH = PropBag.ReadProperty("MH", 420)
    cmdButton(0).Width = maxW: cmdButton(0).Height = maxH
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BC", UserControl.BackColor
    PropBag.WriteProperty "BBC", picDraw.BackColor
    PropBag.WriteProperty "MW", maxW
    PropBag.WriteProperty "MH", maxH
End Sub
