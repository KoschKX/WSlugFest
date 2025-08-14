Option Explicit
Randomize
LoadVPM "01530000","WPC.VBS",3.1

' LoadVPM subroutine that initializes the VPinMAME controller and other essential elements
Sub LoadVPM(VPMver,VBSfile,VBSver)
	On Error Resume Next
		If ScriptEngineMajorVersion<5 Then MsgBox"VB Script Engine 5.0 or higher required"
		ExecuteGlobal GetTextFile(VBSfile)
		If Err Then MsgBox"Unable to open "&VBSfile&". Ensure that it is in the same folder as this table."&vbNewLine&Err.Description:Err.Clear
		Set Controller=CreateObject("VPinMAME.Controller")
		If Err Then MsgBox"Can't Load VPinMAME."&vbNewLine&Err.Description
		If VPMver>"" Then If Controller.Version<VPMver Or Err Then MsgBox"VPinMAME ver "&VPMver&" required.":Err.Clear
		If VPinMAMEDriverVer<VBSver Or Err Then MsgBox VBSFile&" ver "&VBSver&" or higher required."
	On Error Goto 0
End Sub

Const UseSolenoids=1,UseLamps=1,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="Coin3"

' Define solenoid callback functions to handle solenoid activations '
SolCallBack(1)="SolSwingBat"
SolCallBack(2)="SolMagSwitch"
SolCallback(7)="vpmSolSound ""Knocker""," 
SolCallback(17)="vpmFlasher M1,"	
SolCallback(18)="vpmFlasher M2,"	
SolCallback(19)="vpmFlasher M3,"	
SolCallback(20)="vpmFlasher Flasher20,"
SolCallBack(22)="SolPitchFast"
SolCallBack(23)="SolPitchMedium"
SolCallBack(24)="SolPitchSlow"
SolCallback(25)="vpmFlasher Flasher25,"
SolCallback(26)="vpmFlasher Flasher26,"
SolCallBack(27)="SolCard"
SolCallBack(28)="vpmFlasher M4,"

' Subroutine to handle swinging bat solenoid activation '
Sub SolSwingBat(Enabled)
	If Enabled Then
		Flipper1.RotateToEnd
		PlaySound"FlipperUp"
		Flipper1.TimerEnabled=1
	End If
End Sub

' Flipper timer subroutine to return the flipper to the starting position '
Sub Flipper1_Timer
	Flipper1.TimerEnabled=0
	Flipper1.RotateToStart
	PlaySound"FlipperDown"
End Sub

' Subroutine for the magnetic switch, setting the position of the magnet '
Sub SolMagSwitch(Enabled)
	If Enabled Then
		If Players=2 Then
			If MagSelect="NONE" Then
				mMagnet.X=500
				mMagnet.Y=1610
				mMagnet.Size=90
				Exit Sub
			End If
			If MagSelect="LEFT" Then
				mMagnet.X=459
				mMagnet.Y=1610
				mMagnet.Size=90
				mMagnet.MagnetOn=True
				Exit Sub
			End If
			If MagSelect="RIGHT" Then
				mMagnet.X=541
				mMagnet.Y=1610
				mMagnet.Size=90
				mMagnet.MagnetOn=True
				Exit Sub
			End If
		Else
			CPUC=RND
			If CPUC<=.33 Then
				mMagnet.X=459
				mMagnet.Y=1610
				mMagnet.Size=90
				mMagnet.MagnetOn=True
				Exit Sub
			End If
			If CPUC<=.66 Then
				mMagnet.X=500
				mMagnet.Y=1610
				mMagnet.Size=90
				Exit Sub
			End If
			mMagnet.X=541
			mMagnet.Y=1610
			mMagnet.Size=90
			mMagnet.MagnetOn=True
			Exit Sub
		End If
	End If
End Sub

' Subroutine for triggering a magnetic field deactivation '
Sub Trigger5_Hit
	mMagnet.MagnetOn=False
	MagSelect="NONE"
End Sub

' Subroutine for ramp solenoid activation '
Sub SolRamp(Enabled)
	If Enabled Then
		If Wall4.IsDropped=1 Then
			Wall4.IsDropped=0
			Controller.Switch(55)=0
		Else
			Wall4.IsDropped=1
			Controller.Switch(55)=1
		End If
	End If
End Sub

' Subroutine for fast pitch solenoid activation '
Sub SolPitchFast(Enabled)
	If Enabled Then
		If bsTrough.Balls Then
			Wall8.IsDropped=1
			Wall9.IsDropped=0
			bsTrough.ExitSol_On
			vpmTimer.PulseSw 54
			Speed=3
		Else
			PlaySound"SolOn"
		End If
	End If
End Sub

' Similar solenoid activation subroutines for medium pitch '
Sub SolPitchMedium(Enabled)
	If Enabled Then
		If bsTrough.Balls Then
			Wall8.IsDropped=1
			Wall9.IsDropped=0
			bsTrough.ExitSol_On
			vpmTimer.PulseSw 54
			Speed=2
		Else
			PlaySound"SolOn"
		End If
	End If
End Sub

' Similar solenoid activation subroutines for slow pitch '
Sub SolPitchSlow(Enabled)
	If Enabled Then
		If bsTrough.Balls Then
			Wall8.IsDropped=1
			Wall9.IsDropped=0
			bsTrough.ExitSol_On
			vpmTimer.PulseSw 54
			Speed=1
		Else
			PlaySound"SolOn"
		End If
	End If
End Sub

' Subroutine for triggering when a specific switch is hit (for pitch adjustment) '
Sub Trigger3_Hit
	If ActiveBall.VelY>0 Then
		Wall9.IsDropped=1
		Wall8.IsDropped=0
		If Speed=3 Then
			ActiveBall.VelY=20
			Speed=0
		End If
		If Speed=2 Then
			ActiveBall.VelY=15
			Speed=0
		End If
		If Speed=1 Then
			ActiveBall.VelY=10
			Speed=0
		End If
	End If
End Sub

' Subroutine to handle card dispensing mechanism '
Sub SolCard(Enabled)
	If Enabled Then
		vpmTimer.PulseSw 28 'Unjam dispenser '
		Controller.Switch(23)=1'Yes there is a card there '
		CardCount=CardCount+1
		Cards=Cards+1
		If CardCount>15 Then
			CardCount=CardCount-1
		Else
			Timer1.Enabled=0
			Timer1.Enabled=1
		End If
		CardReel.SetValue CardCount
		If Cards>12 Then
			Controller.Switch(27)=1'Low Cards '
		Else
			Controller.Switch(27)=0
		End If
	Else
		Controller.Switch(23)=0
	End If
End Sub

' Timer subroutine for resetting card values '
Sub Timer1_Timer
	CardReel.SetValue 0
	Timer1.Enabled=0
End Sub

' Subroutine to set the display of elements on the controller '
Sub SetDisplayToElement(Element)
 	If Controller.Version<="01500000" Then
 		' forget it, version is to old '
 		Exit Sub
 	End If
  	Dim playerRect
 	playerRect=Controller.GetClientRect(GetPlayerHwnd)
 	Dim playerWidth, playerHeight
 	playerWidth=playerRect(2)-playerRect(0)
 	playerHeight=playerRect(3)-playerRect(1)
 	Dim Game
 	Set Game=Controller.Game
 	Dim x,y
  	x=Element.x*playerWidth/1000.0-1
 	y=Element.y*playerHeight/750.0-1
 	Game.Settings.SetDisplayPosition x,y,GetPlayerHwnd
 	Set Game=nothing
End Sub

' Subroutine to handle pausing and unpausing the table '
Sub Table1_Paused:Controller.Pause=True:End Sub
Sub Table1_unPaused:Controller.Pause=False:End Sub

' Initial variables and setup for the table '
Dim Location,bsTrough,mMagnet,mRamp,MagSelect,Speed,Cards,CardCount,RampPos,Players,CPUC
Location=500:MagSelect="NONE":Speed=0:Cards=0:CardCount=0:RampPos=0:Players=0:CPUC=0

' Table initialization subroutine '
Sub Table1_Init
	Wall9.IsDropped=1'Pitch Cover'
	On Error Resume Next
		Controller.Games("sf_l1").Settings.Value("cheat")=1
		Controller.GameName="sf_l1":If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
		Controller.SplashInfoLine="Slugfest"
		Controller.HandleMechanics=0
		Controller.HandleKeyboard=0
		Controller.ShowDMDOnly=1 
		Controller.ShowFrame=0 
		Controller.ShowTitle=0
		'SetDisplayToElement TextBox1
		Controller.Run GetPlayerHwnd:If Err Then MsgBox Err.Description:Exit Sub
		Controller.DIP(0)=&H00
	On Error Goto 0
	vpmNudge.TiltSwitch=56:vpmNudge.Sensitivity=5:PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
	Controller.Switch(24)=1

	Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,0,0,0,0,0,0,0
	bsTrough.InitKick BallRelease,180,10
    bsTrough.InitExitSnd"Popper","SolOn"
    bsTrough.Balls=1

	Set mMagnet=New cvpmMagnet
	mMagnet.InitMagnet Trigger4,3
	mMagnet.GrabCenter=0
	mMagnet.CreateEvents"mMagnet"

	Set mRamp=New cvpmMech
	mRamp.Sol1=21
	mRamp.MType=vpmMechLinear+vpmMechCircle+vpmMechOneSol
	mRamp.Length=40
	mRamp.Steps=500
	mRamp.AddSw 55,0,20
	mRamp.AddSw 55,40,80
	mRamp.Callback=GetRef("UpdateRamp")
	mRamp.Start

	Controller.Switch(54)=0
	Controller.Switch(22)=1
	vpmMapLights AllLights
	
	'CHANGES'
	LightUpdateTimer.Enabled = True

	
End Sub

' Subroutine to update the ramp position '
Sub UpdateRamp(aCurrPos,aSpeed,aLastPos)
	If aCurrPos=25 Then
		Wall4.IsDropped=1'Down'
		Wall10.IsDropped=1
		RampPos=0
	End If
	If aCurrPos=87 Then
		Wall4.IsDropped=0'Up'
		Wall10.IsDropped=0
		RampPos=1
	End If
End Sub

' Trigger for ramp hit '
Sub Trigger2_Hit
	If RampPos=1 And ActiveBall.VelY<-7 Then ActiveBall.VelZ=120
End Sub

' Displaying key help for the game '
ExtraKeyHelp=KeyName(StartGameKey)&vbTab&"1 Player Game"&vbNewLine&_
			KeyName(KeyFront)&vbTab&"2 Player Game"&vbNewLine&_
			KeyName(LeftFlipperKey)&vbTab&"Fastball"&vbNewLine&_
			KeyName(KeyUpperLeft)&vbTab&"Curveball"&vbNewLine&_
			KeyName(31)&vbTab&"Change Up"&vbNewLine&_
			KeyName(32)&vbTab&"Screwball"&vbNewLine&_
			KeyName(33)&vbTab&"Throw Out Runner"&vbNewLine&_
			KeyName(39)&vbTab&"Pinch Hit"&vbNewLine&_
			KeyName(RightFlipperKey)&vbTab&"Swing Bat"&vbNewLine&_
			KeyName(KeyUpperRight)&vbTab&"Steal Base"&vbNewLine&_
			KeyName(16)&vbTab&"Refill Cards"

' Subroutine to handle key down events (when a key is pressed) '
Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode=StartGameKey Then
		Controller.Switch(37)=1
		Players=1
	End If
	If KeyCode=KeyFront Then
		Controller.Switch(38)=1
		Players=2
	End If
	If KeyCode=LeftFlipperKey Then Controller.Switch(32)=1
	If KeyCode=KeyUpperLeft Then
		Controller.Switch(34)=1
		MagSelect="RIGHT"'Curve to right'
	End If
	If KeyCode=31 Then
		Controller.Switch(33)=1
		MagSelect="NONE"
	End If
	If KeyCode=32 Then
		Controller.Switch(35)=1
		MagSelect="LEFT"'Curve to left'
	End If
	If KeyCode=33 Then Controller.Switch(36)=1
	If KeyCode=39 Then Controller.Switch(31)=1
	If KeyCode=RightFlipperKey Then Controller.Switch(26)=1
	If KeyCode=KeyUpperRight Then Controller.Switch(25)=1
	If KeyCode=16 And Controller.Switch(27) Then'Refill Cards'
		Cards=0
		CardCount=0
		Controller.Switch(27)=0
		Timer1.Enabled=1
	End If
	If KeyDownHandler(KeyCode) Then Exit Sub
End Sub

' Subroutine to handle key up events (when a key is pressed) '
Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode=StartGameKey Then Controller.Switch(37)=0
	If KeyCode=KeyFront Then Controller.Switch(38)=0
	If KeyCode=LeftFlipperKey Then Controller.Switch(32)=0
	If KeyCode=KeyUpperLeft Then Controller.Switch(34)=0
	If KeyCode=31 Then Controller.Switch(33)=0
	If KeyCode=32 Then Controller.Switch(35)=0
	If KeyCode=33 Then Controller.Switch(36)=0
	If KeyCode=39 Then Controller.Switch(31)=0
	If KeyCode=RightFlipperKey Then Controller.Switch(26)=0
	If KeyCode=KeyUpperRight Then Controller.Switch(25)=0
	If KeyUpHandler(KeyCode) Then Exit Sub
End Sub

Sub Trigger1_Hit
	Location=ActiveBall.X
	If Location<177 Then
		vpmTimer.PulseSw 41
		Exit Sub
	End If
	If Location<308 Then
		vpmTimer.PulseSw 42
		Exit Sub
	End If
	If Location<437 Then
		vpmTimer.PulseSw 43
		Exit Sub
	End If
	If Location<557 Then
		vpmTimer.PulseSw 44
		Exit Sub
	End If
	If Location<674 Then
		vpmTimer.PulseSw 45
		Exit Sub
	End If
	If Location<802 Then
		vpmTimer.PulseSw 46
		Exit Sub
	End If
	vpmTimer.PulseSw 47
End Sub

' Subroutines for when different kickers (bumper-like mechanisms) hit '
Sub Kicker1_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker2_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker3_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker4_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker5_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker6_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker7_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker8_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker9_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker10_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker11_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker12_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker13_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker14_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker15_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker16_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker17_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub
Sub Kicker18_Hit:bsTrough.AddBall Me:KickerHit Me:End Sub

Sub Kicker19_Hit:Kicker19.DestroyBall:vpmTimer.PulseSwitch(61),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker20_Hit:Kicker20.DestroyBall:vpmTimer.PulseSwitch(61),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker21_Hit:Kicker21.DestroyBall:vpmTimer.PulseSwitch(61),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker22_Hit:Kicker22.DestroyBall:vpmTimer.PulseSwitch(61),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker23_Hit:Kicker23.DestroyBall:vpmTimer.PulseSwitch(61),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker24_Hit:Kicker24.DestroyBall:vpmTimer.PulseSwitch(61),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker25_Hit:Kicker25.DestroyBall:vpmTimer.PulseSwitch(62),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker26_Hit:Kicker26.DestroyBall:vpmTimer.PulseSwitch(62),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker27_Hit:Kicker27.DestroyBall:vpmTimer.PulseSwitch(62),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker28_Hit:Kicker28.DestroyBall:vpmTimer.PulseSwitch(62),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker29_Hit:Kicker29.DestroyBall:vpmTimer.PulseSwitch(62),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker30_Hit:Kicker30.DestroyBall:vpmTimer.PulseSwitch(62),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker31_Hit:Kicker31.DestroyBall:vpmTimer.PulseSwitch(63),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker32_Hit:Kicker32.DestroyBall:vpmTimer.PulseSwitch(63),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker33_Hit:Kicker33.DestroyBall:vpmTimer.PulseSwitch(63),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker34_Hit:Kicker34.DestroyBall:vpmTimer.PulseSwitch(63),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker35_Hit:Kicker35.DestroyBall:vpmTimer.PulseSwitch(63),100,"AddBackTrough":KickerHit Me:End Sub
Sub Kicker36_Hit:Kicker36.DestroyBall:vpmTimer.PulseSwitch(63),100,"AddBackTrough":KickerHit Me:End Sub
Sub AddBackTrough(swNo):vpmTimer.PulseSwitch(51),100,"AddTrough":End Sub
Sub AddTrough(swNo):bsTrough.AddBall 0:MagSelect="NONE":End Sub
Sub Drain_Hit:Drain.DestroyBall:vpmTimer.PulseSwitch(52),100,"AddTrough":End Sub


' Subroutines for when different spinners hit '
Sub Spinner1_Spin:SpinnerHit Me:End Sub
Sub Spinner2_Spin:SpinnerHit Me:End Sub
Sub Spinner3_Spin:SpinnerHit Me:End Sub
Sub Spinner4_Spin:SpinnerHit Me:End Sub
Sub Spinner5_Spin:SpinnerHit Me:End Sub
Sub Spinner6_Spin:SpinnerHit Me:End Sub
Sub Spinner7_Spin:SpinnerHit Me:End Sub


'CHANGES'

Sub KickerHit(Kicker)
	' PlaySound "test"
End Sub

Sub SpinnerHit(Spinner)
	' PlaySound "test"
	Refresh.State = 1
	Refresh.State = 0
End Sub

Sub LightUpdateTimer_Timer()
    UpdateAllLights
End Sub

Sub StartUpdateTimer()
    ' Create a timer that pulses every 16 milliseconds (approximately 60 FPS)
    vpmTimer.PulseSwitch 1, 16, "UpdateAllLights"
End Sub

Sub UpdateAllLights()
    Dim i
    For i = 0 To AllLights.Count - 1
        If AllLights(i).State Then
            AllLights(i).Image = "SFPlayfLa"
        Else
            AllLights(i).Image = "SFPlayf"
        End If
    Next
End Sub

' Function to change the light image
Sub ChangeLightImage(light, newImage)
    light.Image = newImage  ' Set the light's image to the new image
End Sub