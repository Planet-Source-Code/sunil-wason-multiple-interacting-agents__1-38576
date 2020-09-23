VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{9546BC25-7D7D-11D3-AFEB-9ACB87FFE207}#8.0#0"; "FT.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   480
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   480
   ScaleMode       =   0  'User
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   480
   End
   Begin FormTracer.FT FT1 
      Left            =   480
      Top             =   360
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin AgentObjectsCtl.Agent AgentX 
      Left            =   840
      Top             =   120
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Name : Multiple Agents
' Author : Sunil Wason (sunilwason@yahoo.com)
' Purpose : This application shows how we can
'make use of multiple agents interact with each
'other simultaneously. This code can have
'numerous end uses especially in the field of
'presentations and advertisements where issues
'can be focussed and delivered through animated
'characters to the public.

'Although I have designed the entire code by
'myself however, I donot take full credit to
'the same since I have taken help from various
'MS Agent codes available on the PSC and
'fianally come out with this one since, I could
'not find any code which so explicitly displayed
'the use of multiple agents.

'You are required to install Microsoft Agent 2.0
'which can be downloaded from the Microsoft site
'(http://msdn.microsoft.com/workshop/imedia/agent)
'before using this program.

'You will also require to download two characters
'viz Genie & Merlin.
'These can be downloaded from
'http://www.msagentring.org/
'On downloading these characters, copy them
'to App.Path
'These charactes have not been enclosed due to
'their size.

'If your application doesnot work go to
'Project --> References and check
'Microsoft Agent Control 2.0.

Dim X1 As Integer, Y1 As Integer

Dim Genie1 As IAgentCtlCharacterEx
Dim Genie2 As IAgentCtlCharacterEx

Const BalloonOn = 1
Dim Status
Dim GeniePathName1
Dim GeniePathName2
Dim LoadRequest
Dim LoadShow
Dim LoadMove
Dim AssistantRequest1
Dim AssistantRequest2

Private Sub AgentX_Hide(ByVal CharacterID As String, ByVal Cause As Integer)

If CharacterID = "merlin" Then
    Timer1.Enabled = False
    Unload Me
    Set Form1 = Nothing
End If

End Sub

Private Sub Form_Activate()

' use WINAPI to make form "always on top"
SetWindowPos hwnd, conHwndTopmost, 0, 0, 32, 32, _
conSwpNoActivate Or conSwpShowWindow
Me.Left = Genie1.Left + 20
Me.Top = Genie1.Top + Genie1.Height + 20
        
End Sub

Private Sub Form_DblClick()

Unload Me
End

End Sub

Private Sub Form_Load()
        Dim x As Boolean, xx As Long
        x = FT1.DoTrace(255, 255, 255)
        'Timer1.Enabled = False
        AgentX.Connected = True
        GeniePathName1 = App.Path & "\genie.acs"
        Set LoadRequest = AgentX.Characters.Load("genie", GeniePathName1)
        Set Genie1 = AgentX.Characters("genie")
        Set LoadShow = Genie1.Get("state", "Showing, Speaking")
        Genie1.MoveTo 600, 200
        Genie1.Show
        Genie1.AutoPopupMenu = False
        'Hides the Agent's baloon
        'Genie1.Balloon.Style = Genie1.Balloon.Style And (Not BalloonOn)
        GeniePathName2 = App.Path & "\merlin.acs"
        Set LoadRequest = AgentX.Characters.Load("merlin", GeniePathName2)
        Set Genie2 = AgentX.Characters("merlin")
        Set LoadShow = Genie2.Get("state", "Showing, Speaking")
        Genie2.MoveTo 200, 200
        Genie2.Show
        Genie2.AutoPopupMenu = False
        'Hides the Agent's baloon
        'Genie2.Balloon.Style = Genie2.Balloon.Style And (Not BalloonOn)
        Genie1.Play ("GestureRight")
        Set AssistantRequest1 = Genie1.Speak("Hi Merlin, \pau=200\ how are you?")
        Genie2.Wait AssistantRequest1
        Genie2.Play ("GestureLeft")
        Set AssistantRequest2 = Genie2.Speak("I am fine \pau=200\ Genie")
        Genie2.Play ("GestureLeftReturn")
        Genie1.Wait AssistantRequest2
        Genie1.Play ("Congratulate")
        Set AssistantRequest1 = Genie1.Speak("I see, \pau=400\ we seem to be in \emp\ talking terms now")
        Genie2.Wait AssistantRequest1
        Genie2.Play ("Gestureup")
        Set AssistantRequest2 = Genie2.Speak("God save the world with you around")
        Genie2.Play ("Restpose")
        Genie1.Wait AssistantRequest2
        Genie1.Play ("GetAttention")
        Set AssistantRequest1 = Genie1.Speak("There he goes again. How can I be of your service, \pau=400\ Merlin.")
        Genie1.Play ("GetAttentionReturn")
        Genie2.Play "Hear_4"
        Genie2.Wait AssistantRequest1
        Genie2.Play "Hear_1"
        Genie2.Play "read"
        Set AssistantRequest2 = Genie2.Speak("Wait a minute")
        Genie2.MoveTo 500, 200
        Set AssistantRequest2 = Genie2.Speak("You sure can be. Just wait and watch me")
        Genie1.Wait AssistantRequest2
        Genie1.Play ("Explain")
        Genie1.Play ("Restpose")
        Genie2.Play ("DoMagic1")
        Genie2.Play ("DoMagicReturn")
        Set AssistantRequest2 = Genie2.Speak("Buzz off from my sight")
        Genie1.Wait AssistantRequest2
        Set AssistantRequest1 = Genie1.Speak("Oh ho, not again")
        Genie2.Wait AssistantRequest1
        Genie1.Speak ".", "Talk.wav"
        Set AssistantRequest1 = Genie1.Speak("\emp\ But before \pau=100\ I leave! \emp\This one is for you")
        Genie2.Wait AssistantRequest1
        Genie1.Play ("LookRightBlink")
        Genie1.Play ("DoMagic1")
        Genie1.Play ("DoMagic2")
        Genie2.Wait AssistantRequest1
        Genie1.Hide
        Genie2.Play ("Pleased")
        Genie2.Wait AssistantRequest1
        Set AssistantRequest2 = Genie2.Speak("Some Genies can be a real pain. \emp\ \pau=200\ \pit=100\ Hey, what's happenning to me.")
        Genie2.Speak ".", "Talk.wav"
        Genie2.Speak ("Where am I being sent to !!")
        Genie2.MoveTo 1000, 200
        Genie2.Speak "\emp\ Man, its pretty dark here."
        Genie2.MoveTo 200, 200
        Genie2.Speak "Much better. Darkness is pretty scarry. Genies can't do any magic better than this. Bye"
        Genie2.Hide
        
        'The Speech Modifiers:
'   1.  \emp\       Emphasise the next word
'   2.  \pau=m\  e.g. '\pau=100\ Pause for m milliseconds
'   3.  \pit=p\   Pitch voice to p Hertz (1 - 400)
'   4.  \spd=s\ Set speed to s Words per minute
'                   (50-250)  -Thanks Ric for that one
'These are the only three I've ever seen and know

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'   Allows dragging of the form from any visible area.  There are a
'   Number of ways of doing this, but i like this code.
If Button = 0 Then
Y1 = y
X1 = x
End If
If Button = 1 Then
Me.Left = Me.Left - (X1 - x)
Me.Top = Me.Top - (Y1 - y)
End If
End Sub

Private Sub Timer1_Timer()

'For positioning the form
Dim A As Long
Dim B As Long
Const Speed = 75
A = (Genie1.Left * 15) + 750
B = (Genie1.Top * 15) + 1900
' determine if cursor is exactly at the center of the form
' else change form position horizontally
If Form1.Left <> A Then
    If Form1.Left > A Then
        Form1.Left = Form1.Left - Speed
    ElseIf Form1.Left < A Then
        Form1.Left = Form1.Left + Speed
    End If
End If

' determine if cursor is exactly at the center of the form
' else change form position vertically
If Form1.Top <> B Then
        If Form1.Top > B Then
        Form1.Top = Form1.Top - Speed
    ElseIf Form1.Top < B Then
        Form1.Top = Form1.Top + Speed
    End If
End If


End Sub
