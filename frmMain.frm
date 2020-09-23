VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{A6BDE5D5-8F7A-11D1-9C65-4CA605C10E27}#5.0#0"; "ActiveGUI.ocx"
Begin VB.Form frmMain 
   Caption         =   "EzColorTest"
   ClientHeight    =   4170
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin ActiveGUICtl.ActiveToolbar ToolbarContainer 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   741
      Begin AIFCmp1.asxToolbar MainToolbar 
         Height          =   375
         Left            =   120
         Top             =   30
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGap       =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   18
         PlaySounds      =   0   'False
         ShowSeparators  =   -1  'True
         ButtonKey1      =   "NewFile"
         ButtonPicture1  =   "frmMain.frx":0442
         ButtonToolTipText1=   "New File"
         ButtonKey2      =   "Open"
         ButtonPicture2  =   "frmMain.frx":0764
         ButtonToolTipText2=   "Open File"
         ButtonKey3      =   "save"
         ButtonPicture3  =   "frmMain.frx":0A26
         ButtonToolTipText3=   "Save File"
         ButtonStyle4    =   2
         ButtonEnabled5  =   0   'False
         ButtonKey5      =   "Cut"
         ButtonPicture5  =   "frmMain.frx":0CE0
         ButtonToolTipText5=   "Cut Text"
         ButtonEnabled6  =   0   'False
         ButtonKey6      =   "Copy"
         ButtonPicture6  =   "frmMain.frx":0EBA
         ButtonToolTipText6=   "Copy Text"
         ButtonKey7      =   "Paste"
         ButtonPicture7  =   "frmMain.frx":117C
         ButtonToolTipText7=   "Paste Text"
         ButtonStyle8    =   2
         ButtonKey9      =   "undo"
         ButtonPicture9  =   "frmMain.frx":146E
         ButtonToolTipText9=   "Undo Action"
         ButtonKey10     =   "redo"
         ButtonPicture10 =   "frmMain.frx":15BC
         ButtonToolTipText10=   "Redo Action"
         ButtonStyle11   =   2
         ButtonKey12     =   "Bold"
         ButtonPicture12 =   "frmMain.frx":170A
         ButtonToolTipText12=   "Bold"
         ButtonKey13     =   "Italic"
         ButtonPicture13 =   "frmMain.frx":1858
         ButtonToolTipText13=   "Italic"
         ButtonKey14     =   "underline"
         ButtonPicture14 =   "frmMain.frx":19A6
         ButtonToolTipText14=   "Underline"
         ButtonStyle15   =   2
         ButtonKey16     =   "LeftJustify"
         ButtonPicture16 =   "frmMain.frx":1B18
         ButtonToolTipText16=   "Left Justify"
         ButtonKey17     =   "HtmlCenter"
         ButtonPicture17 =   "frmMain.frx":1D4E
         ButtonToolTipText17=   "Center"
         ButtonKey18     =   "rightjust"
         ButtonPicture18 =   "frmMain.frx":1F84
         ButtonToolTipText18=   "Right Justify"
      End
   End
   Begin VB.CommandButton cmdDummy 
      Height          =   315
      Left            =   14040
      TabIndex        =   3
      Top             =   10560
      Width           =   1215
   End
   Begin VB.PictureBox Container 
      BackColor       =   &H00808080&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   440
      Width           =   5655
      Begin RichTextLib.RichTextBox RichTxtBox 
         Height          =   1215
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2143
         _Version        =   327680
         BorderStyle     =   0
         HideSelection   =   0   'False
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":21BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3900
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9790
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   6960
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_Undo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit_Redo 
         Caption         =   "Redo"
      End
      Begin VB.Menu mnuEdit_Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About EzColorTest..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    --------------------------------------------------------------------------
'    EzColorTest HTML Editor Color Coding Test
'    Copyright (C) 2000  Eric Banker
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'    --------------------------------------------------------------------------

'    Some parts of this code were originally developed by Joel Dueck (BlueIce) and
'    David Hoyt(Vhtml)

Option Explicit
    
' This is for line numbering
Private Declare Function SendMessageLong Lib "User32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1

' This variable keeps track of the currently opened file name at all times. Make sure you set it
' correctly so the user know's what is going on
Public OpenFilename As String

' Editor constants
Const WM_COPY = &H301
Const WM_CUT = &H300
Const WM_CLEAR = &H303
Const WM_PASTE = &H302

Public trapUndo As Boolean           'flag to indicate whether actions should be trapped
Private UndoStack As New Collection   'collection of undo elements
Private RedoStack As New Collection   'collection of redo elements

' Keeps track of control key state
Public CtlKey As Boolean

' ################################################################
' These are the main form handler functions
' ################################################################

' This is the main load function. It sets up the colors and the richtextbox and inserts the template
' and also colorizes it

Private Sub Form_Load()
    ' Let the user know something is happening, at least that the program is starting up :)
    Screen.MousePointer = vbHourglass
        
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
    'Set EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = False
    RichTxtBox.AutoVerbMenu = False
    RichTxtBox.HideSelection = True
    
    ' Now lets add a template to the new document
    AspTemplate
    
    ' That template still needs to be color coded so lets do that
    HtmlHighlight
    
    ' Lets set the caption of the form to say that this is an untitled document
    Me.Caption = "EzColorTest: Untitled"
    
    ' Lets let the user see the text box now that everything is finished
    RichTxtBox.Visible = True
    RichTxtBox.TabStop = True
    
    ' Everything is finished so lets set the mouse pointer back so the user knows the wait is over
    Screen.MousePointer = vbNormal
    
    trapUndo = True     'Enable Undo Trapping
    RichTxtBox_Change      'Initialize First Undo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then
        CtlKey = True
    ElseIf KeyCode = vbKeyF6 And (Shift And vbAltMask) Then
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then
        CtlKey = False
    End If
End Sub

' This function simply makes sure the text box fits in the form at all times with the gutter

Private Sub Form_Resize()
On Error Resume Next
    Container.Width = Me.Width - 139
    Container.Height = Me.Height - 1400
    RichTxtBox.Width = Container.Width - 420
    RichTxtBox.Height = Container.Height - 50
    
    ' Make sure the dummy for undo never shows itself
    cmdDummy.Left = Me.Width + 1000
    cmdDummy.Top = Me.Height + 1000
    RichTxtBox.Refresh
End Sub

' ################################################################
' This is where the form handler functions stop
' ################################################################

' ################################################################
' This is where the RichTextBox handler functions go. These
' basically do everything.
' ################################################################

Private Sub RichTxtBox_Change()
    If Not trapUndo Then Exit Sub 'because trapping is disabled

    Dim newElement As New UndoElement   'create new undo element
    Dim c%, l&

    'remove all redo items because of the change
    For c% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c%

    'set the values of the new element
    newElement.SelStart = RichTxtBox.SelStart
    newElement.TextLen = Len(RichTxtBox.Text)
    newElement.Text = RichTxtBox.Text

    'add it to the undo stack
    UndoStack.Add Item:=newElement
    
    EnableControls
End Sub

' This makes sure that when the user hits the tab key that it
' indents the text

Private Sub RichTxtBox_GotFocus()
    On Error Resume Next
    Dim Control As Control
    For Each Control In Controls
        Control.TabStop = False
    Next Control
End Sub

' This makes the copy and paste menu items disabled when they can't be used

Private Sub RichTxtBox_SelChange()
Dim Ln As Long
    Ln = RichTxtBox.SelLength
    With frmMain
        ' Determine which options are available
        .mnuEdit_Cut.Enabled = Ln
        .mnuEdit_Copy.Enabled = Ln
        .mnuEdit_Paste.Enabled = Len(Clipboard.GetText(1))
        .mnuEdit_SelectAll.Enabled = CBool(Len(RichTxtBox.Text))
        .MainToolbar.ButtonEnabled(5) = Ln
        .MainToolbar.ButtonEnabled(6) = Ln
        .MainToolbar.ButtonEnabled(7) = Len(Clipboard.GetText(1))
    End With
    GetEditStatus
End Sub

' This highlights while typing
' This function in the module reads in the current character and colors the tag correctly.

Private Sub RichTxtBox_KeyPress(KeyAscii As Integer)
On Error Resume Next
    KeyAscii = KeyPressEvent(KeyAscii)
End Sub

' This enables and disables the undo/redo options
Private Sub EnableControls()
    Me.MainToolbar.ButtonEnabled(9) = UndoStack.Count > 1
    Me.MainToolbar.ButtonEnabled(10) = RedoStack.Count > 0
    Me.mnuEdit_Undo.Enabled = UndoStack.Count > 1
    Me.mnuEdit_Redo.Enabled = RedoStack.Count > 0
    
    Me.MainToolbar.ButtonEnabled(9) = Me.MainToolbar.ButtonEnabled(9)
    Me.MainToolbar.ButtonEnabled(10) = Me.MainToolbar.ButtonEnabled(10)
    Me.mnuEdit_Undo.Enabled = Me.mnuEdit_Undo.Enabled
    Me.mnuEdit_Redo.Enabled = Me.mnuEdit_Redo.Enabled
    
    RichTxtBox_SelChange
End Sub

' This does some stuff like keeps track of the control character and knows when it is
' pressed or not. It also has some examples of keyboard shortcuts. For example hitting
' ctrl + 1 through 6 puts in the <h1></h1> tag all the way through <h6></h6>
' it also does ctrl+space which puts in &nbsp; , ctrl+enter which puts in <P>
' and shift+enter which puts in <br>
' it also finds out if the cursor is in a tag

Private Sub RichTxtBox_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TypedIn As String
    If Shift And vbCtrlMask Then
        If KeyCode > vbKey0 And KeyCode < vbKey7 Then
            Dim HeadingTag As String
            HeadingTag = "<H" & CStr(KeyCode - vbKey0) & "></H" & CStr(KeyCode - vbKey0) & ">"
            InsertTag HeadingTag, True
            PlaceCursor HeadingTag, 5
            RichTxtBox.SelColor = vbBlack
        Else
            Select Case KeyCode
            Case vbKeyV
                ' User pressed Ctrl+V  - Paste
                Dim A$, S As Long
                S = RichTxtBox.SelStart ' save this since selstart moves up after the paste
                A = Clipboard.GetText(vbCFText)
                RichTxtBox.SelText = ""
                RichTxtBox.SelText = A    ' This removes any unwanted formatting (font, &c)
                HtmlColorCode S, RichTxtBox.SelStart
                
                KeyCode = 0
            Case vbKeyReturn
                InsertTag "<P>", True
                RichTxtBox.SelColor = vbBlack
                KeyCode = 0
            Case vbKeySpace
                RichTxtBox.SelColor = vbBlack
                RichTxtBox.SelText = "&nbsp;"
                KeyCode = 0
            End Select
        End If
    ElseIf Shift And vbShiftMask Then
        If KeyCode = vbKeyReturn Then
            InsertTag "<BR>", True
            RichTxtBox.SelColor = vbBlack
            KeyCode = 0
        End If
    End If
    IsOutsideTag
End Sub

' This function finds out if the user out the mouse inside a tag and sets the right color for
' the text

Private Sub RichTxtBox_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    IsOutsideTag
    'RichTxtBox.SetFocus
End Sub

' this resets the ctrl key if it was pressed to not pressed because the key was lifted
' it also finds out if the cursor is in a tag

Private Sub RichTxtBox_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then
        CtlKey = False
    End If
    IsOutsideTag
    RichTxtBox.SetFocus
End Sub

' this shows the edit menu on a right mouse click in the richtextbox. it also finds
' out if the cursor is in a tag

Private Sub RichTxtBox_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' Show edit menu
    If Button = vbRightButton Then
        PopupMenu mnuEdit
    End If
    IsOutsideTag
    'RichTxtBox.SetFocus
End Sub

' This gets the current line number and column number and displays it

Public Sub GetEditStatus()
   Dim lLine As Long, lCol As Long
   Dim cCol As Long, lChar As Long, I As Long

   lChar = RichTxtBox.SelStart + 1

   ' Get the line number
   lLine = 1 + SendMessageLong(RichTxtBox.hWnd, EM_LINEFROMCHAR, _
           RichTxtBox.SelStart, 0&)

   ' Get the Character Position
   cCol = SendMessageLong(RichTxtBox.hWnd, EM_LINELENGTH, lChar - 1, 0&)

   I = SendMessageLong(RichTxtBox.hWnd, EM_LINEINDEX, lLine - 1, 0&)
   lCol = lChar - I

   ' Caption of Label1 is set to Cursor Position.
   ' This could also be a panel in a StatusBar.
   sbStatusBar.Panels(1).Text = "Line: " & lLine & ", Col: " & lCol

End Sub

' ################################################################
' These are the functions to handle toolbar stuff
' ################################################################

Private Sub MainToolbar_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    If ButtonIndex = 1 Then
        mnuFileNew_Click
    ElseIf ButtonIndex = 2 Then
        mnuFileOpen_Click
    ElseIf ButtonIndex = 3 Then
        mnuFileSave_Click
    ElseIf ButtonIndex = 5 Then
        mnuEdit_Cut_Click
    ElseIf ButtonIndex = 6 Then
        mnuEdit_Copy_Click
    ElseIf ButtonIndex = 7 Then
        mnuEdit_Paste_Click
    ElseIf ButtonIndex = 9 Then
        Undo
    ElseIf ButtonIndex = 10 Then
        Redo
    ElseIf ButtonIndex = 12 Then
            ' Here we are inserting a tag so we need to do some special stuff.
            ' First lets create some variables
            Dim BoldTag As String
            Dim BoldHolder As String
            
            ' Now lets create the tag. First find out if the user has any highlighted text
            BoldHolder = RichTxtBox.SelText
            
            ' This makes sure the the text stays black
            RichTxtBox.SelColor = vbBlack
            
            ' Now lets create the entire tag which is the bold tag, plus the selected text, and the ending tag
            BoldTag = "<B>" & BoldHolder & "</B>"
            
            ' Now we are going to insert the tag and color code it at the same time. To do this we
            ' simply call the insert tag function of the color code control. This works like this:
            ' EZColorCode.InsertTag (EntireTag)
            ' That's all there is to it. the control handles everything else, color coding and inserting
            InsertTag BoldTag, True
            
            ' This is a nice little function I created to place the cursor inside the tag
            PlaceCursor BoldTag, 4
    ElseIf ButtonIndex = 13 Then
            ' Here we are inserting a tag so we need to do some special stuff.
            ' First lets create some variables
            Dim ItalicTag As String
            Dim ItalicHolder As String
            
            ' Now lets create the tag. First find out if the user has any highlighted text
            ItalicHolder = RichTxtBox.SelText
            
            ' This makes sure the the text stays black
            RichTxtBox.SelColor = vbBlack
            
            ' Now lets create the entire tag which is the Italic Tag, plus the selected text, and the ending tag
            ItalicTag = "<I>" & ItalicHolder & "</I>"
            
            ' Now we are going to insert the tag and color code it at the same time. To do this we
            ' simply call the insert tag function of the color code control. This works like this:
            ' EZColorCode.InsertTag (EntireTag)
            ' That's all there is to it. the control handles everything else, color coding and inserting
            InsertTag ItalicTag, True
            
            PlaceCursor ItalicTag, 4
    ElseIf ButtonIndex = 14 Then
            ' Here we are inserting a tag so we need to do some special stuff.
            ' First lets create some variables
            Dim UnderlineTag As String
            Dim UnderlineHolder As String
            
            ' Now lets create the tag. First find out if the user has any highlighted text
            UnderlineHolder = RichTxtBox.SelText
            
            ' This makes sure the the text stays black
            RichTxtBox.SelColor = vbBlack
            
            ' Now lets create the entire tag which is the Underline Tag, plus the selected text, and the ending tag
            UnderlineTag = "<U>" & UnderlineHolder & "</U>"
            
            ' Now we are going to insert the tag and color code it at the same time. To do this we
            ' simply call the insert tag function of the color code control. This works like this:
            ' EZColorCode.InsertTag (EntireTag)
            ' That's all there is to it. the control handles everything else, color coding and inserting
            InsertTag UnderlineTag, True
            PlaceCursor UnderlineTag, 4
    ElseIf ButtonIndex = 16 Then
            ' Here we are inserting a tag so we need to do some special stuff.
            ' First lets create some variables
            Dim LeftTag As String
            Dim LeftHolder As String
            
            ' Now lets create the tag. First find out if the user has any highlighted text
            LeftHolder = RichTxtBox.SelText
            
            ' This makes sure the the text stays black
            RichTxtBox.SelColor = vbBlack
            
            ' Now lets create the entire tag which is the Left Tag, plus the selected text, and the ending tag
            LeftTag = "<DIV ALIGN=""left"">" & LeftHolder & "</DIV>"
            
            ' Now we are going to insert the tag and color code it at the same time. To do this we
            ' simply call the insert tag function of the color code control. This works like this:
            ' EZColorCode.InsertTag (EntireTag)
            ' That's all there is to it. the control handles everything else, color coding and inserting
            InsertTag LeftTag, True
            
            PlaceCursor LeftTag, 6
    ElseIf ButtonIndex = 17 Then
            ' Here we are inserting a tag so we need to do some special stuff.
            ' First lets create some variables
            Dim CenterTag As String
            Dim CenterHolder As String
            
            ' Now lets create the tag. First find out if the user has any highlighted text
            CenterHolder = RichTxtBox.SelText
            
            ' This makes sure the the text stays black
            RichTxtBox.SelColor = vbBlack
            
            ' Now lets create the entire tag which is the Center Tag, plus the selected text, and the ending tag
            CenterTag = "<CENTER>" & CenterHolder & "</CENTER>"
            
            ' Now we are going to insert the tag and color code it at the same time. To do this we
            ' simply call the insert tag function of the color code control. This works like this:
            ' EZColorCode.InsertTag (EntireTag)
            ' That's all there is to it. the control handles everything else, color coding and inserting
            InsertTag CenterTag, True
            
            PlaceCursor CenterTag, 9
    ElseIf ButtonIndex = 18 Then
            ' Here we are inserting a tag so we need to do some special stuff.
            ' First lets create some variables
            Dim RightTag As String
            Dim RightHolder As String
            
            ' Now lets create the tag. First find out if the user has any highlighted text
            RightHolder = RichTxtBox.SelText
            
            ' This makes sure the the text stays black
            RichTxtBox.SelColor = vbBlack
            
            ' Now lets create the entire tag which is the Right Tag, plus the selected text, and the ending tag
            RightTag = "<P ALIGN=""right"">" & RightHolder & "</P>"
            
            ' Now we are going to insert the tag and color code it at the same time. To do this we
            ' simply call the insert tag function of the color code control. This works like this:
            ' EZColorCode.InsertTag (EntireTag)
            ' That's all there is to it. the control handles everything else, color coding and inserting
            InsertTag RightTag, True
            
            PlaceCursor RightTag, 4
    End If
    RichTxtBox.SetFocus
End Sub

' ################################################################
' This is the end of the toolbar functions
' ################################################################

' ################################################################
' These are the menu functions for the file menu
' ################################################################

' Create a new file with a template

Private Sub mnuFileNew_Click()
    ' This function may be more complicated than what is really needed. I just like making sure everything is
    ' clean and that no problems could arise when the new document is color coded.
    
    ' First lets let the user know that something is happening
    Screen.MousePointer = vbHourglass

    ' Hide what is happening from the user
    RichTxtBox.Visible = False
    
    ' This cleans the current document and removes color coding
    ' Clean the document
    RichTxtBox.SelStart = 0
    RichTxtBox.SelLength = Len(RichTxtBox.Text)
    RichTxtBox.SelColor = vbBlack
    RichTxtBox.SelStart = 0
    
    ' Now that the control has been set back to normal we can remove the text in the text box
    RichTxtBox.Text = ""
    
    ' Now that the entire thing is clean lets add in our template
    AspTemplate
    
    ' Make sure that the template is color coded.
    HtmlHighlight
    
    ' make sure that all variables are set to show that it's a new file and is also untitled
    OpenFilename = ""
    Me.Caption = "EzColorText: Untitled"
    
    ' Now that everything is finished we can return to the user and show the new template
    RichTxtBox.Visible = True
    
    ' Everything is finished so lets set the mouse pointer back so the user knows the wait is over
    Screen.MousePointer = vbNormal
    RichTxtBox.SetFocus
End Sub

' Open a file

Private Sub mnuFileOpen_Click()
    ' This gets the file name from the common dialog
    CMDialog1.DialogTitle = "Open File"
    CMDialog1.Filter = " All Web Files (*.asp, *.asa, *.htm, *.html, *.css, *.inc) | *.asp; *.asa; *.htm; *.html; *.css; *.inc; | Asp files (*.asp) | *.asp; | Asa files (*.asa) | *.asa; | Htm files (*.htm) | *.htm | Html files (*.html) | *.html; | Style Sheets (*.css) | *.css; | Include Files (*.inc) | *.inc; | All files (*.*)|*.*|"
    CMDialog1.ShowOpen
    
    ' Lets make sure the user didn't hit cancel
    If Err <> 32755 Then
        ' Let the user know something is happening
        Screen.MousePointer = vbHourglass
        
        ' Set the textbox visible to false (so they don't see the thing color code)
        RichTxtBox.Visible = False
        
        ' before we open the new document lets clean the current document and removes color coding
        RichTxtBox.SelStart = 0
        RichTxtBox.SelLength = Len(RichTxtBox.Text)
        RichTxtBox.SelColor = vbBlack
        RichTxtBox.SelStart = 0
        
        ' lets go ahead and open the file, this will remove the current file while opening also
        RichTxtBox.LoadFile CMDialog1.FileName, rtfText
        OpenFilename = CMDialog1.FileName
        
        ' now we need to color code the new document. This function in the control does just that
        HtmlHighlight
        
        ' set the caption of the form so the user knows what they are editing
        Me.Caption = "EzColorTest: " & OpenFilename
        
        ' make the text box visible again now that everything is color coded and done
        RichTxtBox.Visible = True
        
        ' Set the mouse pointer back so the user knows the wait is over
        Screen.MousePointer = vbNormal
    End If
    RichTxtBox.SetFocus
End Sub

' Save the current file

Private Sub mnuFileSave_Click()
    ' This function just takes the current html text and saves it. nothing special
    
    ' Find out if it's an open file
    If OpenFilename = "" Then
        ' It's not so show the save dialog and get the new file name
        CMDialog1.ShowSave
        RichTxtBox.SaveFile CMDialog1.FileName, rtfText
    Else
        ' It is an opened document so lets save the new information
        RichTxtBox.SaveFile OpenFilename, rtfText
    End If
    RichTxtBox.SetFocus
End Sub

' Save the current file as...

Private Sub mnuFileSaveAs_Click()
    ' This always shows the save dialog for the user to choose what filename to give the
    ' file
    CMDialog1.ShowSave
    RichTxtBox.SaveFile CMDialog1.FileName, rtfText
    RichTxtBox.SetFocus
End Sub

' Exit the program

Private Sub mnuFileExit_Click()
    ' This just gets me out of the program
    Unload Me
End Sub

' ################################################################
' This is the end of the file menu functions
' ################################################################

' ################################################################
' This is where the edit menu functions go
' ################################################################

Private Sub mnuEdit_Redo_Click()
    Redo
End Sub

Private Sub mnuEdit_Undo_Click()
    Undo
End Sub

Private Sub mnuEdit_Click()
    mnuEdit_Paste.Enabled = Clipboard.GetFormat(vbCFText)
    RichTxtBox.SetFocus
End Sub

Private Sub mnuEdit_Copy_Click()
    EditFunction WM_COPY
    RichTxtBox.SetFocus
End Sub

Private Sub mnuEdit_Cut_Click()
    EditFunction WM_CUT
    RichTxtBox.SetFocus
End Sub

Private Sub mnuEdit_Paste_Click()
    Dim A$, S As Long
    S = RichTxtBox.SelStart ' save this since selstart moves up after the paste
    A = Clipboard.GetText(vbCFText)
    RichTxtBox.SelText = ""
    RichTxtBox.SelText = A    ' This removes any unwanted formatting (font, &c)
    HtmlColorCode S, RichTxtBox.SelStart
    RichTxtBox.SetFocus
End Sub

Private Sub mnuEdit_SelectAll_Click()
    RichTxtBox.SelStart = 0
    RichTxtBox.SelLength = Len(RichTxtBox.Text)
    RichTxtBox.SetFocus
End Sub

' These below are for the edit functions and find the current state of the control
' key. I do this so that when the user hit ctrl-c to copy and then ctrl-v to paste
' the text box does not show 2 copys of the text and not 1

Private Sub EditFunction(Action As Integer)
    If Me.CtlKey = False Then
        If Action <> WM_COPY Then RichTxtBox.SelText = ""
        Call SendMessage(RichTxtBox.hWnd, Action, 0, 0&)
    End If
End Sub

' ################################################################
' This is where the edit menu functions stop
' ################################################################

' ################################################################
' This is where the help menu functions start
' ################################################################

Private Sub mnuHelpAbout_Click()
    ' Hey I need to get credit for this thing don't I?
    frmAbout.Show 0, Me
End Sub

' ################################################################
' This is where the help menu functions stop
' ################################################################

' ################################################################
' These are misc functions for editing and cursor placement
' ################################################################

' I have created a small template here that is used when the program starts and when a new file is opened.

Sub AspTemplate()
On Error Resume Next
    RichTxtBox.Text = "<%@ LANGUAGE=""VBSCRIPT"" %>" & _
    vbCrLf & "<html>" & vbCrLf & "<!------------- Created By EzColorTest ------------->" & _
    vbCrLf & "<!----------- Copyright 2000 Eric Banker ----------->" & vbCrLf & "<head>" & _
    vbCrLf & "     <title>Untitled Document</title>" & vbCrLf & "</head>" & vbCrLf & _
    "<body bgcolor=""#FFFFFF"" text=""#000000"" link=""#804040"" vlink=""#008080"" alink=""#004080"">" & vbCrLf & "<!---------------- Insert Text Here ---------------->" & _
    vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
    RichTxtBox.SetFocus
End Sub

' Move the cursor to the new position in the textbox. Note: I'm not sure why this works but it does

Public Sub PlaceCursor(Text$, Cursor As Long)
Dim T As Long
    T = RichTxtBox.SelStart
    RichTxtBox.SelStart = (T + Len(Tag)) - Cursor
End Sub

' This is for the undo/redo functions

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function

Public Sub Undo()
Dim chg$, X&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object
    If UndoStack.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
            cmdDummy.SetFocus   'change focus of form
            X& = SendMessage(RichTxtBox.hWnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            RichTxtBox.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            RichTxtBox.SelLength = objElement.TextLen - objElement2.TextLen
            RichTxtBox.SelText = ""
            X& = SendMessage(RichTxtBox.hWnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            RichTxtBox.SelStart = objElement2.SelStart
            RichTxtBox.SelLength = 0
            RichTxtBox.SelText = chg$
            RichTxtBox.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                RichTxtBox.SelLength = Len(chg$)
            Else
                RichTxtBox.SelStart = RichTxtBox.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add Item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    EnableControls
    trapUndo = True
    RichTxtBox.SetFocus
End Sub

Public Sub Redo()
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(RichTxtBox.Text)
        If DeleteFlag Then  'delete last item
            Set objElement = RedoStack(RedoStack.Count)
            RichTxtBox.SelStart = objElement.SelStart
            RichTxtBox.SelLength = Len(RichTxtBox.Text) - objElement.TextLen
            RichTxtBox.SelText = ""
        Else 'append something
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(RichTxtBox.Text, objElement.Text, objElement.SelStart + 1)
            RichTxtBox.SelStart = objElement.SelStart - Len(chg$)
            RichTxtBox.SelLength = 0
            RichTxtBox.SelText = chg$
            RichTxtBox.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                RichTxtBox.SelLength = Len(chg$)
            Else
                RichTxtBox.SelStart = RichTxtBox.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add Item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    EnableControls
    trapUndo = True
    RichTxtBox.SetFocus
End Sub

' ################################################################
' End Misc functions
' ################################################################

'    --------------------------------------------------------------------------
'    That's all there is to it. Now you have a program that you can now use
'    to create your very own Html Editor with color coding. Isn't free software
'    grand :)
'
'    If you like the program please send me an email at ebanker@gmu.edu
'    Later
'    Eric
'    --------------------------------------------------------------------------
