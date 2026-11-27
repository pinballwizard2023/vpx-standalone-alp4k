' ***********************************************************************
'               VISUAL PINBALL X EM Script
'                  Basic EM script, 1 player only 
'		          uses core.vbs for extra functions
'
'  Route 66, an original VPX table in the Electro-Mechanical style of the 1970's
'
'  Initial starting source table was Big Deal VPX7 - version by Klodo81 2022, Big Deal version 2.1
'
'  I watched an interesting video on youtube recently that helped me better understand old EM pinball:
'	check out "Technology Connections" "Wires do the math" as he explains what's going on inside the AZTEK machine.
'
'  Thanks again to the Visual Pinball community - for making the software, making tutorials, and letting people reuse objects and code.
'
'  Version 0.0 (initial release - likely full of bugs)
'  Version 0.1 (quick fix because I left the 1 second counter as 0.1 second from testing)
'  Version 0.2 (update B2S to use illumination snippits instead of player 2 reels for the odometer so that it stops that annoying clicking)
'				Mostly visual changes. A few minor scoring changes.
'  Version 0.3 (add gas gauge, update pictures, minor playfield cleanup)
'
'	For what its worth, you have my permission to use or modify anything here.  
'	I don't need any credit and don't blame me if it doesn't work.
' ***********************************************************************


Option Explicit
Randomize

' core.vbs constants
Const BallSize = 50 ' 50 is the normal size
Const BallMass = 1  ' 1 is the normal ball mass.

' load extra vbs files
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Resume Next
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
End Sub


'* * * * * * * * * * * * * * * * * * * * *

' Constants
Const TableName = "Route66" ' file name to save config
Const cGameName = "Route66" ' B2S name

Const MaxBallsOnPlayfield = 10
Const BallsPerGame = 5          
Const TiltSensitivity = 6
Const LastMile = 2400


' Global variables

Dim BallsRemaining
Dim PlayerScore
Dim MileCount
Dim MilesThisTurn
Dim OdometerAddAmount
Dim BumperHitCount
Dim TiltMeasure
Dim GasMilesLeft


' slingshot displays
Dim Lr01flex
Dim Rr01flex

' Control variables
Dim BallsOnPlayfield

' Boolean variables
Dim isGameInPlay
Dim isTilted
Dim isProjectorEnabled
Dim isResetMilesOnStart

' defaults - the Table Options menu (F12) overrides these when the table loads
isProjectorEnabled = True
OdometerAddAmount = 1
isResetMilesOnStart = False
       
Dim BackglassStatus 'for B2s attract mode
' core.vbs variables


Dim PicId, RecentPicId, PicArray

PicArray = Array("0001","0004","0010","0015","0020","0023","0025","0030","0035","0040","0045",  _
   "0050","0055","0058","0062","0066","0070","0075","0080","0085","0090","0095", _
   "0100","0105","0110","0115","0120","0125","0130","0135","0140","0145","0150","0155","0160", _
   "0165","0170","0173","0175","0180","0185","0190","0195","0197", _
   "0200","0205","0207","0210","0215","0218","0220","0223","0225","0230","0235","0240","0245","0250", _
   "0255","0257","0260","0265","0270","0275","0280","0285","0287","0290","0295", _
   "0300","0303","0305","0308","0310","0313","0315","0320","0325","0330","0335","0340","0345","0350","0355", _
   "0360","0365","0370","0375","0380","0385","0390","0395", _
   "0400","0403","0405","0410","0415","0420","0425","0430","0435","0440","0445","0450","0455", _
   "0460","0465","0470","0475","0480","0485","0490","0490", _
   "0500","0505","0510","0515","0520","0525","0530","0535","0540","0545","0550","0560","0565", _
   "0570","0575","0578","0581","0585","0590","0595", _
   "0600","0605","0610","0615","0620","0623","0625","0630","0635","0640","0645","0650","0655", _
   "0660","0665","0670","0675","0680","0685","0690","0695", _
   "0700","0705","0710","0715","0720","0725","0730","0735","0740","0745","0750","0753","0756", _
   "0760","0765","0768","0770","0775","0780","0783","0786","0790","0795", _
   "0800","0805","0810","0815","0820","0830","0840","0850","0855","0860","0865","0870","0880","0890", _
   "0900","0910","0920","0930","0940","0945","0950","0960","0970","0980","0990", _
   "1000","1005","1010","1015","1020","1023","1030","1035","1040","1045","1048","1050","1055", _
   "1060","1070","1080","1090", _
   "1100","1110","1120","1125","1130","1139","1141","1145","1150","1155","1160","1165","1170", _
   "1175","1180","1185","1190","1195", _
   "1200","1205","1210","1215","1220","1225","1230","1235","1240","1245","1250","1255","1260", _
   "1265","1270","1275","1280","1290","1295", _
   "1300","1305","1310","1315","1320","1325","1330","1335","1340","1345","1350","1355","1360", _
   "1365","1370","1375","1380","1385","1390","1395", _
   "1400","1405","1410","1415","1420","1425","1430","1435","1440","1445","1450","1455","1460", _
   "1465","1470","1475","1480","1485","1490", _
   "1500","1505","1510","1515","1520","1525","1530","1535","1540","1545","1550", _
   "1560","1565","1570","1575","1580","1585","1590","1595", _
   "1600","1605","1610","1615","1620","1625","1630","1635","1640","1645","1650","1655","1660", _
   "1665","1670","1675","1680","1685","1690","1695", _
   "1700","1705","1710","1715","1720","1725","1730","1735","1740","1745","1750","1755","1760", _
   "1765","1770","1775","1780","1785","1790","1795","1798", _
   "1800","1805","1810","1815","1820","1825","1830","1835","1840","1845","1850","1855","1860", _
   "1865","1870","1875","1878","1880","1885","1890","1895", _
   "1900","1905","1910","1915","1920","1925","1930","1935","1940","1945","1950","1955","1960", _
   "1965","1970","1975","1980","1985","1990","1995", _
   "2000","2005","2010","2015","2020","2025","2030","2035","2040","2045","2050","2055","2060", _
   "2065","2070","2075","2080","2085","2090","2095", _
   "2100","2105","2108","2110","2115","2120","2125","2130","2135","2140","2143","2145","2150", _
   "2155","2158","2162","2165","2167","2170","2175","2180","2185","2190","2195", _
   "2200","2205","2208","2210","2215","2220","2223","2225","2230","2235","2240","2245","2250", _
   "2255","2260","2265","2270","2275","2280","2285","2290","2293","2295", _
   "2300","2305","2310","2315","2317","2320","2325","2330","2335","2340","2345","2350","2355", _
   "2358","2360","2362","2365","2370","2373","2375","2380","2385","2390","2395","2398","2400")


'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

'If Keycode = AddCreditKey Then
'	if GI006.state = 0 then 
'		GiOn
'	Else	
'		GiOff
'	end if
'End If
	If isGameInPlay Then
		If keycode = PlungerKey Then
			Plunger.Pullback
			PlaySoundAt "fx_plungerpull", plunger
		End If

		If NOT isTilted Then
			If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
			If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
			If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

			If keycode = LeftFlipperKey  Then FlippersUpLeft
			If keycode = RightFlipperKey Then FlippersUpRight
		End If

	Else
		If Keycode = AddCreditKey Then
			PlaySound "fx_coin"
		End If
		'options that used to be on the 2, 3 and 4 keys are now in the Table Options menu (F12)
		If keycode = StartGameKey Then
			If isResetMilesOnStart Then MileCount = 0

			cleanupTimer.Enabled=True
			'when the cleanupTimer finishes, it will go to Game_Init to start the game
        End If
    End If

End Sub


Sub Table1_KeyUp(ByVal keycode)

    'plunger release
    If keycode = PlungerKey Then
        Plunger.Fire
        If PlungerLaneBallTrigger.UserValue="Hit" Then
            PlaySoundAt "fx_plunger", plunger
        Else
            PlaySoundAt "fx_plunger_empty", plunger
        End If
    End If

	If isGameInPlay Then
		If keycode = LeftFlipperKey Then FlippersDownLeft
		If keycode = RightFlipperKey Then FlippersDownRight
	End If

End Sub


Sub AddABall
	PlaySound "fx_bell_low"
    PlaySound "fx_bell_high"
    BallsRemaining = BallsRemaining + 1
End Sub

'******************
' Table stop/pause
'******************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub


'*******************
'  Flipper Subs
'*******************

Sub FlippersUpLeft()
	If isTilted Then Exit Sub
	If Not isGameInPlay Then Exit Sub

    PlaySoundAt "LeftflipperupH-2dB", LeftFlipper
    PlayRepeatingSoundAt "buzz", LeftFlipper
    LeftFlipper.RotateToEnd
	LeftFlipperTop.RotateToEnd
End Sub

Sub FlippersDownLeft()
    StopSound "buzz"
    PlaySoundAt "LeftflipperdownH", LeftFlipper
	LeftFlipper.RotateToStart
	LeftFlipperTop.RotateToStart
End Sub

Sub FlippersUpRight()
	If isTilted Then Exit Sub
	If Not isGameInPlay Then Exit Sub

    PlaySoundAt "RightflipperupH-2dB", RightFlipper
    PlayRepeatingSoundAt "buzz", RightFlipper
    RightFlipper.RotateToEnd
	RightFlipperTop.RotateToEnd
End Sub

Sub FlippersDownRight()
    StopSound "buzz"
    PlaySoundAt "RightflipperdownH", RightFlipper
    RightFlipper.RotateToStart
	RightFlipperTop.RotateToStart
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipperTop_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipperTop_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'************************
'       Targets
'************************

'************************
' Rollover Targets
'************************
Sub RolloverTopLeft_Hit
	PlaySoundAt "fx_Sensor", RolloverTopLeft
	PlLeftEnter.State=1
	If PlRightEnter.State>0 Then
		AddScore 10
	Else
		AddScore 1
	End If
    If GasMilesLeft > 0 Then OdometerTimer.Enabled = True
End Sub

Sub RolloverTopRight_Hit
	PlaySoundAt "fx_Sensor", RolloverTopRight
	PlRightEnter.State=1
	If PlLeftEnter.State>0 Then
		AddScore 10
	Else
		AddScore 1
	End If
    If GasMilesLeft > 0 Then OdometerTimer.Enabled = True
End Sub

Sub RolloverOutlaneLeft_Hit
	PlaySoundAtBall "fx_Sensor"
	AddScore 100
	If PlOutlaneLeft.state = 1 Then
		AddABall()
		If PlOutlaneRight.state = 0 Then
			PlOutlaneLeft.state = 0
		End If
	Else		
		PlOutlaneLeft.State = 1
	End If
End Sub

Sub RolloverOutlaneRight_Hit
	PlaySoundAtBall "fx_Sensor"
	AddScore 100
	If PlOutlaneRight.state = 1 Then
' if we have already hit the right outlane, award an extra ball
		AddABall()
		If PlOutlaneLeft.state = 0 Then
' if we have not hit the left outlane, turn of the right extra ball light
' so: if you hit both the left and right outlanes, you'll always get an extra ball when you hit again
			PlOutlaneRight.state = 0
		End If
	Else		
		PlOutlaneRight.State = 1
	End If
End Sub

Sub RolloverStarInlaneLeft_Hit
	PlaySoundAtBall "fx_Sensor"      
	If PlLeftEnter.State=1 Then
		AddScore 200
	Else
		AddScore 10 + 90 * PlStarInlaneLeft.state + (100 * PlStarInlaneRight.state)
	End If
	PlStarInlaneLeft.state = 1
End Sub

Sub RolloverStarInlaneRight_Hit
	PlaySoundAtBall "fx_Sensor"
	If PlRightEnter.State=1 Then
		AddScore 200
	Else
		AddScore 10 + 90 * PlStarInlaneRight.state + (100 * PlStarInlaneLeft.state)
	End If
	PlStarInlaneRight.state = 1
End Sub

Sub RolloverStarTopLeft_hit  
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore (10 + 90 * PlRolloverStarTopLeft.state) * (1 + PlRolloverStarTopCenter.state) * (1 + PlRolloverStarTopRight.state)
	PlRolloverStarTopLeft.state = 1
	If PlArrowLoopLeft.state = 2 Then 
		PlArrowLoopLeft.state = 1
		AddScore 500
	End If
End Sub

Sub RolloverStarTopCenter_hit
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore (10 + 90 * PlRolloverStarTopCenter.state) * (1 + PlRolloverStarTopLeft.state) * (1 + PlRolloverStarTopRight.state)
	PlRolloverStarTopCenter.state = 1
End Sub

Sub RolloverStarTopRight_hit
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore (10 + 90 * PlRolloverStarTopRight.state) * (1 + PlRolloverStarTopCenter.state) * (1 + PlRolloverStarTopLeft.state)
	PlRolloverStarTopRight.state = 1
	If PlArrowLoopRight.state = 2 Then 
		PlArrowLoopRight.state = 1
		AddScore 500
	End If
End Sub

Sub RolloverStarCenterEnter_hit
	PlaySoundAtBall "fx_Sensor"                                                       
	AddScore 10 
	PlRolloverStarCenterEnter.state = 1
	If GasMilesLeft > 0 Then OdometerTimer.Enabled = True
	If PlArrowCenterSpinner.state = 2 Then
		PlArrowCenterSpinner.state = 1
		AddScore 100
	End If
End Sub

'************************
' Standup Hit Targets
'************************

Sub TargetTopLoop_hit
	PlaySoundAtBall "fx_TargetHit"
    AddScore 10 + 90 * PlLoop1.state
	If PlLoop4.state = 1 Then PlLoop5.state = 1
	If PlLoop3.state = 1 Then PlLoop4.state = 1
	If PlLoop2.state = 1 Then PlLoop3.state = 1
	If PlLoop1.state = 1 Then PlLoop2.state = 1
	PlLoop1.state = 1
	FlashTargetTopLoop.duration 1, 100, 0
End Sub

Sub TargetSmallUpLeft_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore 1
End Sub

Sub TargetSmallUpRight_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore 1
End Sub

Sub TargetTopLeft_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore 10 + 90 * PlTargetTopLeft.state  * (1 + PlTargetTopRight.state)
	FlashKicker.duration 1, 100, 0
	If PlTargetTopLeft.state = 1 Then PlTargetTopLeft2.state = 1
	PlTargetTopLeft.state = 1
End Sub

Sub TargetTopRight_Hit
	PlaySoundAtBall "fx_TargetHit"
	AddScore (10 + 90 * PlTargetTopRight.state) * (1 + PlTargetTopLeft.state)
	FlashKicker.duration 1, 100, 0
	If PlTargetTopRight.state = 1 Then PlTargetTopRight2.state = 1
	PlTargetTopRight.state = 1
End Sub

Sub TargetMidLeft_Hit
	PlaySoundAtBall "fx_TargetHit"
    AddScore 10 + 90 * PlArrowTargetMidLeft.state
	FlashTargetMidLeft.duration 1, 100, 0
	If PlArrowTargetMidLeft.state = 2 Then
		PlArrowTargetMidLeft.state = 1
		Addscore 500
	End If
End Sub

Sub TargetMidLeft2_Hit
	PlaySoundAtBall "fx_TargetHit"
    AddScore 300
	PlTargetMidLeft2.state = 1
End Sub

Sub TargetBumperLeft_hit
	PlaySoundAt "fx_TargetHit", TargetBumperLeft
    PlaySoundAt2 "gasstationbell", TargetBumperLeft
    AddScoreNoSound 10 + 90 * PlTargetBumperLeft.state
    AddGas()
	PlTargetBumperLeft.state = 1

	If LightBumperUpLeft.State = 0 Then
		LightBumperUpLeft.State = 1
	ElseIf LightBumperUpLeft.State = 1 Then
		If LightBumperUpRight.State > 0 Then
			LightBumperUpLeft.State = 2
		Else
			LightBumperUpLeft.State = 0
		End If
	End If

	If LightBumperLowLeft.State = 0 Then
		LightBumperLowLeft.State = 1
	ElseIf LightBumperLowLeft.State = 1 Then
		If LightBumperLowRight.State > 0 Then
			LightBumperLowLeft.State = 2
		Else
			LightBumperLowLeft.State = 0
		End If
	End If

End Sub

Sub TargetBumperRight_hit
	PlaySoundAt "fx_TargetHit", TargetBumperRight
    PlaySoundAt2 "gasstationbell", TargetBumperRight
    AddScoreNoSound 10 + 90 * PlTargetBumperRight.state
    AddGas()
	PlTargetBumperRight.state = 1

	If LightBumperUpRight.State = 0 Then
		LightBumperUpRight.State = 1
	ElseIf LightBumperUpRight.State = 1 Then
		If LightBumperUpLeft.State > 0 Then
			LightBumperUpRight.State = 2
		Else
			LightBumperUpRight.State = 0
		End If
	End If

	If LightBumperLowRight.State = 0 Then
		LightBumperLowRight.State = 1
	ElseIf LightBumperLowRight.State = 1 Then
		If LightBumperLowLeft.State > 0 Then
			LightBumperLowRight.State = 2
		Else
			LightBumperLowRight.State = 0
		End If
	End If

End Sub


'************************
'       Drop Targets
'************************

Sub TargetDropLeft1_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft1
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft2_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft2
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft3_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft3
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft4_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft4
    AddScoreNoSound 10
End Sub
Sub TargetDropLeft5_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropLeft5
    AddScoreNoSound 10
End Sub

Sub TargetDropLeft1_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft2_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft3_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft4_Dropped()
	CheckLeftDropsCompleted()
End Sub
Sub TargetDropLeft5_Dropped()
	CheckLeftDropsCompleted()
End Sub

Sub TargetDropRight1_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight1
    AddScoreNoSound 10
End Sub
Sub TargetDropRight2_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight2
    AddScoreNoSound 10
End Sub
Sub TargetDropRight3_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight3
    AddScoreNoSound 10
End Sub
Sub TargetDropRight4_Hit
	PlaySoundAt "fx_TargetDrop1", TargetDropRight4
    AddScoreNoSound 10
End Sub

Sub TargetDropRight1_Dropped()
	CheckRightDropsCompleted()
End Sub
Sub TargetDropRight2_Dropped()
	CheckRightDropsCompleted()
End Sub
Sub TargetDropRight3_Dropped()
	CheckRightDropsCompleted()
End Sub
Sub TargetDropRight4_Dropped()
	CheckRightDropsCompleted()
End Sub


Sub RaiseDropTargetsLeft
	If TargetDropLeft1.IsDropped = 1 Then
		TargetDropLeft1.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft1
	End If
	If TargetDropLeft2.IsDropped = 1 Then
		TargetDropLeft2.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft2
	End If
	If TargetDropLeft3.IsDropped = 1 Then
		TargetDropLeft3.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft3
	End If
	If TargetDropLeft4.IsDropped = 1 Then
		TargetDropLeft4.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft4
	End If
	If TargetDropLeft5.IsDropped = 1 Then
		TargetDropLeft5.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropLeft5
	End If
End Sub

Sub RaiseDropTargetsRight
	If TargetDropRight1.IsDropped = 1 Then
		TargetDropRight1.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight1
	End If
	If TargetDropRight2.IsDropped = 1 Then
		TargetDropRight2.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight2
	End If
	If TargetDropRight3.IsDropped = 1 Then
		TargetDropRight3.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight3
	End If
	If TargetDropRight4.IsDropped = 1 Then
		TargetDropRight4.IsDropped = 0
		PlaySoundAt "fx_TargetUp", TargetDropRight4
	End If
End Sub


Sub CheckLeftDropsCompleted
	If TargetDropLeft1.IsDropped = 1 and TargetDropLeft2.IsDropped = 1 and TargetDropLeft3.IsDropped = 1 _
       and TargetDropLeft4.IsDropped = 1 and TargetDropLeft5.IsDropped = 1 Then
        AddScore 1000
		LightBonus()
		RaiseDropTargetsLeft
	End If

End Sub

Sub CheckRightDropsCompleted
	If TargetDropRight1.IsDropped = 1 and TargetDropRight2.IsDropped = 1 and TargetDropRight3.IsDropped = 1 and TargetDropRight4.IsDropped = 1 Then
        AddScore 1000
		LightBonus()
		RaiseDropTargetsRight
	End If
End Sub

Sub LightBonus
	If PLBonus5.state = 1 Then
		FlasherProjector.ImageA="bonus"
		AddScore 10000
        BonusPicTimer.enabled = True
	End If
	If PlBonus4.state = 1 Then PlBonus5.state = 1
	If PlBonus3.state = 1 Then PlBonus4.state = 1
	If PlBonus2.state = 1 Then PlBonus3.state = 1
	If PlBonus1.state = 1 Then PlBonus2.state = 1
	PlBonus1.state = 1
End Sub

Sub BonusPicTimer_Timer
    BonusPicTimer.enabled = False
End Sub
'************************
'       Spinners
'************************

Sub SpinnerTopCenter_Spin()
    PlaySoundAt "fx_spinner", SpinnerTopCenter
	If PlArrowCenterSpinner.state = 2 Then
		PlArrowCenterSpinner.state = 1
		AddScore 100
	Elseif PlArrowCenterSpinner.state = 1 Then
		AddScore 10
	Else
		AddScore 1
	End If
	'OdometerTimer.Enabled = True
End Sub

Sub SpinnerToShooter_Spin()
    PlaySoundAt "fx_spinner", SpinnerToShooter
End Sub


'*****************
'       Kickers
'*****************

Sub KickerTop_Hit
    PlaySoundAt "fx_kicker_enter", kickerTop
    If isTilted Then KickerTop.kick 180, 5: Exit Sub
	KickerTop.UserValue = True
	'KickerTop.TimerEnabled=True
	If PlArrowKicker.state = 2 Then
		PlArrowKicker.state = 1	
	End If
	BonusCountTimer.Enabled = True
End Sub

Sub KickerTop_Release
	if KickerTop.UserValue = True Then
		KickerTop.kick 170, 5 
		PlaySoundAt "fx_kicker", KickerTop
        'KickerTop.TimerEnabled=True
		KickerTop.UserValue = False
		PlTargetTopCenter.state = 1
    End If
End Sub

'*********
' Bumpers
'*********

Sub BumperUpLeft_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperUpLeft
    FlashBumperUpLeft.duration 1, 100, 0
    FlashBumperUpLeft2.duration 1, 100, 0
    AddScore 10 + 90 * LightBumperUpLeft.State
	BumperHitCount = BumperHitCount + 1
	'debug.print BumperHitCount
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

Sub BumperUpRight_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperUpRight
    FlashBumperUpRight.duration 1, 100, 0
    FlashBumperUpRight2.duration 1, 100, 0
    AddScore 10 + 90 * LightBumperUpRight.State
	BumperHitCount = BumperHitCount + 1
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

Sub BumperLowLeft_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperLowLeft
    FlashBumperLowLeft.duration 1, 100, 0
    FlashBumperLowLeft2.duration 1, 100, 0
    AddScore 10 + 90 * LightBumperLowLeft.State
	BumperHitCount = BumperHitCount + 1
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

Sub BumperLowRight_Hit 
    If isTilted Then Exit Sub
    PlaySoundAt "fx_Bumper", BumperLowRight
    FlashBumperLowRight.duration 1, 100, 0
    FlashBumperLowRight2.duration 1, 100, 0
    AddScore 10 + 90 * LightBumperLowRight.State
	BumperHitCount = BumperHitCount + 1
	If BumperHitCount >= 10 Then
		LightArrow()
		BumperHitCount = 0
	End If
End Sub

'*********

Sub AddGas
    OdometerTimer.Enabled=True
    GasMilesLeft = GasMilesLeft + 10 + (PlTargetBumperLeft.State * 5) + (PlTargetBumperRight.State * 5)
    If GasMilesLeft > 50 Then
        GasMilesLeft = 50
    End If
    UpdateGasLights()
End Sub

Sub UseGas
    debugTextBox.text=gasmilesleft
    If GasMilesLeft > 0 Then
        GasMilesLeft = GasMilesLeft - 1
    End If
    UpdateGasLights()
End Sub

Sub UpdateGasLights
    debugTextBox.text=gasmilesleft
    If GasMilesLeft = 0 Then
        PlGasEmpty.state = 2
        PlArrowGasLeft.state = 2
        PlArrowGasRight.state = 2
        PlGas1.state = 0
        PlGas2.state = 0
        PlGas3.state = 0
        PlGas4.state = 0
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 5 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 2
        PlGas2.state = 0
        PlGas3.state = 0
        PlGas4.state = 0
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 10 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 0
        PlGas3.state = 0
        PlGas4.state = 0
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 15 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 2
        PlGas3.state = 0
        PlGas4.state = 0
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 20 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 1
        PlGas3.state = 0
        PlGas4.state = 0
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 25 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 1
        PlGas3.state = 2
        PlGas4.state = 0
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 30 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 1
        PlGas3.state = 1
        PlGas4.state = 0
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 35 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 1
        PlGas3.state = 1
        PlGas4.state = 2
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 40 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 1
        PlGas3.state = 1
        PlGas4.state = 1
        PlGas5.state = 0
    ElseIf GasMilesLeft <= 45 Then
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 1
        PlGas3.state = 1
        PlGas4.state = 1
        PlGas5.state = 2
    Else
        PlGasEmpty.state = 0
        PlArrowGasLeft.state = 0
        PlArrowGasRight.state = 0
        PlGas1.state = 1
        PlGas2.state = 1
        PlGas3.state = 1
        PlGas4.state = 1
        PlGas5.state = 1
    End If

End Sub

Sub SetGasMinHalfTank 
    If GasMilesLeft < 30 Then
        GasMilesLeft = 30
        UpdateGasLights()
    End If
End Sub

'* * * * * * * *
'  table_init = turn on machine
'    game_init = start button - start new game  
'      
'        turn_init = start a new turn (the next ball)
'          ball_init = create a new ball in the plunger lane
'	   ball_end = destroy a ball
'        turn_end = end a turn

'    game_end = game over 
'  table_exit = turn off machine   
'**********



'******************
' Table power on
'******************
Sub Table1_Init()

	debugtextbox.visible=0

    'VPObjects_Init
    LoadEM

    isGameInPlay = False
    
    UpdateBackglassInfoDisplay()

    'turn on GI lights
    vpmtimer.addtimer 1000, "GiOn '"

    ' start the attract mode
    StartAttractMode()

    ' Remove desktop items in FS mode
    Dim x
    
    For each x in cReels
		If Table1.ShowDT then
            x.Visible = 1
		Else
            x.Visible = 0
        End If
    Next

	FlasherProjector.ImageA = "milestart"

	'wait a bit so that the B2S is loaded first
	vpmtimer.addtimer 1000, "LoadConfig() '"	

End Sub


'****************************************
' Init for a new game - start button pressed
'****************************************
Sub Game_Init()

    isGameInPlay = True
    isTilted = False

    StopAttractMode
    GiOn
	TurnOffPlayfieldLights()

	FlasherProjector.ImageA = "milestart"
    RecentPicId = ""
    PlayerScore = 0
    GasMilesLeft = 0
	ResetScores()
	ResetOdometer()
    RaiseDropTargetsLeft()
    RaiseDropTargetsRight()
    BallsRemaining = BallsPerGame

    ' first ball.  delay just a bit to let the score reset sounds finish
    vpmtimer.addtimer 1000, "Turn_Init '"
    
End Sub

'****************************************
'  start a new turn (the next ball)
'****************************************
Sub Turn_Init()
    TiltMeasure = 0

    TurnInitResetLights
    TurnInitResetVariables
    TurnInitResetPlayfield

    UpdateBackglassInfoDisplay()

    'using the timer to pause a bit before adding the ball when the turn starts, it just feels more natural
    vpmtimer.addtimer 1000, "Ball_Init() '"	
End Sub

Sub TurnInitResetLights() 'init lights for new ball/player
	'most lights are reset by the turn end bonus count process
	LightBumperUpLeft.State = 0 
	LightBumperLowLeft.state = 0 
	LightBumperUpRight.State = 0
	LightBumperLowRight.State = 0
	PlRolloverStarTopLeft.state = 0
	PlRolloverStarTopCenter.state = 0
	PlRolloverStarTopRight.state = 0
	PlLeftEnter.state = 0
	PlRolloverStarCenterEnter.state = 0
	PlRightEnter.state = 0
	PlStarInlaneLeft.state = 0
	PlStarInlaneRight.state = 0
End Sub

Sub TurnInitResetVariables() 'init variables new ball/player
    MilesThisTurn = 0
    SetGasMinHalfTank()
End Sub

Sub TurnInitResetPlayfield()
    'RaiseDropTargetsLeft()
    'RaiseDropTargetsRight()
	LightRandomBumper()
End Sub

Sub LightRandomBumper
	Randomize timer
	Dim aRnd : aRnd = Int(4 * Rnd)
	Select Case(aRnd)
		Case 0:LightBumperLowLeft.state = 1
		Case 1:LightBumperLowRight.state = 1
		Case 2:LightBumperUpLeft.state = 1
		Case 3:LightBumperUpRight.state = 1
	End Select

End Sub

'****************************************
'  Create a new ball in the plunger lane
'****************************************
Sub Ball_Init()
    'debug.print "Ball_Init"
    
    If BallsOnPlayfield>=MaxBallsOnPlayfield Then Exit Sub
    
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    
    PlaySoundAt "fx_Ballrel", BallRelease
    BallRelease.Kick 90, 4

    BallsOnPlayfield = BallsOnPlayfield + 1
End Sub

    
Sub PlungerLaneBallTrigger_Hit()                      
	PlungerLaneBallTrigger.UserValue="Hit" 'use like a switch to know if there is a ball for the plunger to hit
	OdometerTimer.Enabled=False
    SetGasMinHalfTank()
End Sub

Sub PlungerLaneBallTrigger_UnHit()
	PlungerLaneBallTrigger.UserValue="UnHit"
End Sub


'****************************************
'  A ball is no longer in play, but there may still be other balls on the playfield
'****************************************
Sub Ball_End()
    'debug.print "Ball_End"

	If(BallsOnPlayfield = 0) Then
       'no balls left on playfield, the turn is done.
		OdometerTimer.Enabled = False
		vpmtimer.addtimer 0010, "Turn_End() '"
    End If

End Sub


'****************************************
'end the turn (no balls left on playfield)
'****************************************
Sub Turn_End()

    BallsRemaining = BallsRemaining - 1

    If isTilted Then
'un-Tilt
        isTilted = False
		TurnOffPlayfieldLights()
		EnableTable()
        UpdateBackglassInfoDisplay()
        TiltMeasure = 0
	End If

	BonusCountTimer.Enabled = True 'this will count up the bonus and go to the next turn

End Sub

Sub Game_End()

    isGameInPlay = False

	FlasherProjector.ImageA = "mileend"
    PlaySound "fx_match"

    UpdateBackglassInfoDisplay()

    FlippersDownLeft
    FlippersDownRight

    'turn off gi Lights
    GiOff

    ' start the attract mode
    vpmtimer.addtimer 1000, "StartAttractMode '"

End Sub


'******************
' Table power off
'******************
Sub table1_Exit
    SaveConfig 
    If B2SOn Then Controller.Stop 'stop the B2S controller
End Sub

'**************
'     TILT
'**************

Sub CheckTilt
    TiltMeasure = TiltMeasure + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If TiltMeasure> 15 Then
		PlaySound "fx_knocker"
        isTilted = True
		DisableTable()
        ReelTilt.SetValue 1
        If B2SOn then
            Controller.B2SSetTilt 1
        End If
		If isProjectorEnabled Then 
			Dim ballinplay:ballinplay = BallsPerGame - BallsRemaining + 1
			Select Case(ballinplay)
				Case 1:FlasherProjector.ImageA = "Tilt1"
				Case 2:FlasherProjector.ImageA = "Tilt2"
				Case 3:FlasherProjector.ImageA = "Tilt3"
				Case 4:FlasherProjector.ImageA = "Tilt4"
				Case 5:FlasherProjector.ImageA = "Tilt5"
			End Select
		End If
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If TiltMeasure > 0 Then
        TiltMeasure = TiltMeasure - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable
    GiOff
    FlippersDownLeft()
    FlippersDownRight()
    BumperUpLeft.Threshold = 100
    BumperUpRight.Threshold = 100
    BumperLowLeft.Threshold = 100
    BumperLowRight.Threshold = 100
	TurnOffPlayfieldLights()
'    LeftSlingshot.Disabled = 1
'    RightSlingshot.Disabled = 1
End Sub

Sub EnableTable
    GiOn
    BumperUpLeft.Threshold = 1
    BumperUpRight.Threshold = 1
    BumperLowLeft.Threshold = 1
    BumperLowRight.Threshold = 1
End Sub


'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlayRepeatingSoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, -1, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAt2(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallNoSpeed(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, 1, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************
'   JP's VP10 Rolling Sounds v4.0
'   JP's Ball Shadows
'   JP's Ball Speed Control
'   Rothbauer's dropping sounds
'***********************************

Const maxvel = 60 'max ball velocity
ReDim rolling(MaxBallsOnPlayfield)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to (MaxBallsOnPlayfield -1)
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate() 'call this routine from any realtime timer you may have, running at an interval of 10 is good.

    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to (MaxBallsOnPlayfield - 1)
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
        cBallShadow(b).Y = 2100 'under the apron 'may differ from table to table
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) < 0 Then Exit Sub 'there no extra balls on this table

    ' draw the ball shadow
    For b = 0 to UBound(BOT)
        cBallShadow(b).X = BOT(b).X
        cBallShadow(b).Y = BOT(b).Y
        cBallShadow(b).Height = BOT(b).Z - Ballsize / 2 + 1

    'play the rolling sound for each ball
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 5
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' new experimental speed control based on jps ball speed control
        Dim aballVel
        aballvel = BallVel(BOT(b))
        If aballvel > maxvel Then
	    BOT(b).VelX = (maxvel / aballvel)^2 * BOT(b).VelX
            BOT(b).VelY = (maxvel / aballvel)^2 * BOT(b).VelY
        End If
        'If BOT(b).VelX AND BOT(b).VelY <> 0 Then
        '    speedfactorx = ABS(maxvel / BOT(b).VelX)
        '    speedfactory = ABS(maxvel / BOT(b).VelY)
        '    If speedfactorx < 1 Then
        '        BOT(b).VelX = BOT(b).VelX * speedfactorx
        '        BOT(b).VelY = BOT(b).VelY * speedfactorx
        '    End If
        '    If speedfactory < 1 Then
        '        BOT(b).VelX = BOT(b).VelX * speedfactory
        '        BOT(b).VelY = BOT(b).VelY * speedfactory
        '    End If
        'End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***************************************
'     Collection collision sounds
'***************************************

'Sub ShooterLaneGate_Hit
'	debug.print"shooterlanegatehit hit"
'End Sub
Sub cMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub cMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub cRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub cRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub cRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub cPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub cGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub cWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

' *********************************************************************
'  Drain
' *********************************************************************

Sub Drain_Hit()
    Drain.DestroyBall
    PlaySoundAt "fx_drain", Drain

    'there is no active game
    If NOT isGameInPlay Then Exit Sub

    If BallsOnPlayfield > 0 Then
        BallsOnPlayfield = BallsOnPlayfield - 1
    End If

    Ball_End

End Sub


' ****************************************
'             Score functions
' ****************************************

Sub AddScore(Points)
    If isTilted Then Exit Sub

    If Points = 0 Then
	    'nothing
	ElseIf Points < 100 Then
	    PlaySound "fx_bigbell10"
	Elseif Points < 1000 Then
	    PlaySound "fx_bigbell100"
	Else
        PlaySound "fx_bigbell1000"
    End If
    AddScoreNoSound Points
End Sub

Sub AddScoreNoSound(Points)
    If isTilted Then Exit Sub

    PlayerScore = PlayerScore + points
    UpdateScore points

End Sub


'************************************
'       Scoring
'************************************

Sub UpdateScore(playerpoints)
    ReelScore.Addvalue playerpoints
    If B2SOn then Controller.B2SSetScorePlayer1 PlayerScore
End Sub

Sub ResetScores
    ReelScore.ResetToZero
    If B2SOn then Controller.B2SSetScorePlayer1 0
End Sub

Sub ResetOdometer
    ReelOdometer.SetValue MileCount
    If B2SOn then B2sOdometer()
End Sub

'*******************
'     Bonus
'*******************

Sub BonusCountTimer_Timer
	
	If PlLeftEnter.state > 0 AND PlRightEnter.state > 0 Then
		PlLeftEnter.state = 0 
		PlRightEnter.state = 0
		If PlRolloverStarCenterEnter.state > 0 Then
			AddScore 500
			PlRolloverStarCenterEnter.state = 0
		Else
			AddScore 100
		End If
		Exit Sub
	End If

	If PlLoop5.state > 0 Then
		AddScore 500
		PlLoop5.state = 0
		Exit Sub
	End If
	If PlLoop4.state > 0 Then
		AddScore 400
		PlLoop4.state = 0
		Exit Sub
	End If
	If PlLoop3.state > 0 Then
		AddScore 300
		PlLoop3.state = 0
		Exit Sub
	End If
	If PlLoop2.state > 0 Then
		AddScore 200
		PlLoop2.state = 0
		Exit Sub
	End If
	If PlLoop1.state > 0 Then
		AddScore 100
		PlLoop1.state = 0
		Exit Sub
	End If
	If PlTargetTopLeft2.state > 0 Then
		AddScore 100
		PlTargetTopLeft2.state = 0
		Exit Sub
	End If
	If PlTargetTopLeft.state > 0 Then
		AddScore 10
		PlTargetTopLeft.state = 0
		Exit Sub
	End If
	If PlTargetTopRight2.state > 0 Then
		AddScore 100
		PlTargetTopRight2.state = 0
		Exit Sub
	End If
	If PlTargetTopRight.state > 0 Then
		AddScore 10
		PlTargetTopRight.state = 0
		Exit Sub
	End If
	If PlTargetTopCenter.state > 0 Then
		AddScore 100
		PlTargetTopCenter.state = 0
		Exit Sub
	End If
	If PlTargetBumperLeft.state > 0 Then
		AddScore 10
		PlTargetBumperLeft.state = 0
		Exit Sub
	End If
	If PlTargetBumperRight.state > 0 Then
		AddScore 10
		PlTargetBumperRight.state = 0
		Exit Sub
	End If
	If PlTargetMidLeft2.state > 0 Then
		AddScore 10
		PlTargetMidLeft2.state = 0
		Exit Sub
	End If

	If PlBonus5.state = 1 Then
		AddScore 5000
		PlBonus5.state = 2
		Exit Sub
	End If
	If PlBonus4.state = 1 Then
		AddScore 4000
		PlBonus4.state = 2
		Exit Sub
	End If
	If PlBonus3.state = 1 Then
		AddScore 3000
		PlBonus3.state = 2
		Exit Sub
	End If
	If PlBonus2.state = 1 Then
		AddScore 2000
		PlBonus2.state = 2
		Exit Sub
	End If
	If PlBonus1.state = 1 Then
		AddScore 1000
		PlBonus1.state = 2
		Exit Sub
	End If

	If PlArrowLoopLeft.state = 1 Then
		AddScore 5000
		PlArrowLoopLeft.state = 0
		Exit Sub
	End If
	If PlArrowLoopRight.state = 1 Then
		AddScore 5000
		PlArrowLoopRight.state = 0
		Exit Sub
	End If
	If PlArrowCenterSpinner.state = 1 Then 
		AddScore 5000
		PlArrowCenterSpinner.state = 0
		Exit Sub
	End If
	If PlArrowKicker.state = 1 Then 
		AddScore 5000
		PlArrowKicker.state = 0
		Exit Sub
	End If
	If PlArrowTargetMidLeft.state = 1 Then 
		AddScore 5000
		PlArrowTargetMidLeft.state = 0
		Exit Sub
	End If

	BonusCountTimer.Enabled = False

	If PlBonus5.state = 2 Then
		PlBonus5.state = 1
	End If
	If PlBonus4.state = 2 Then
		PlBonus4.state = 1
	End If
	If PlBonus3.state = 2 Then
		PlBonus3.state = 1
	End If
	If PlBonus2.state = 2 Then
		PlBonus2.state = 1
	End If
	If PlBonus1.state = 2 Then
		PlBonus1.state = 1
	End If

	If KickerTop.UserValue = True Then
		'triggered from Kicker hit
		vpmtimer.addtimer 0100, "KickerTop_Release '"
	Else
		'triggered from turn_end
        AddScore(MilesThisTurn)
		If BallsRemaining > 0 Then
			vpmtimer.addtimer 0100, "Turn_Init '"
		Else
			vpmtimer.addtimer 0100, "Game_End '"
		End If
	End If

End Sub

'***********
' Playfield lights
'***********

Sub TurnOnPlayfieldLights()
    Dim bulb
    For each bulb in cGameLights
        bulb.State = 1
    Next
End Sub

Sub TurnOffPlayfieldLights()
    Dim bulb
    For each bulb in cGameLights
        bulb.State = 0
    Next
End Sub

'***********
' GI lights
'***********

Sub GiOn 
	'debug.print "GiOn"
	PlaySound"fx_gion"
	Dim bulb
    For each bulb in cGiLights
        bulb.State = 1
        'debug.print bulb.name
    Next
End Sub

Sub GiOff 
	'debug.print "GiOff"
	PlaySound"fx_gioff"    
    Dim bulb
    For each bulb in cGiLights
        bulb.State = 0
		'debug.print bulb.name
    Next
End Sub

' ********************************
'        Attract Mode
' ********************************
' use the"Blink Pattern" of each light

Sub StartAttractMode()
	'this is an EM machine, so pretty dull attract mode....
	TurnOnPlayfieldLights()
    Dim bulb
    For each bulb in cGameLights
        bulb.State = 2
    Next
	BackglassAttractTimer.Enabled=True
End Sub

Sub BackglassAttractTimer_Timer
	'not used
     If Not B2SOn Then Exit Sub
End Sub

Sub StopAttractMode()
    BackglassAttractTimer.Enabled=False
    BackglassAttractTitleOn()
    TurnOffPlayfieldLights()
End Sub

Sub BackglassAttractTitleOn
	'not used
     If Not B2SOn Then Exit Sub
End Sub

Sub BackglassAttractTitleOff
	'not used
     If Not B2SOn Then Exit Sub
End Sub


Sub LoadConfig

    Dim FileObj
    Dim ConfigFile, TextStr
    Dim fileline(20)

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) Then Exit Sub
    If Not FileObj.FileExists(UserDirectory & TableName & ".txt") Then Exit Sub
    Set ConfigFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
    Set TextStr = ConfigFile.OpenAsTextStream(1, 0)

    Dim aLine : aLine = 1
    do until aLine > 20
        fileline(aLine) = TextStr.ReadLine
        If TextStr.AtEndOfStream Then Exit Do
		aLine = aLine + 1
    loop

    TextStr.Close
    Set ConfigFile = Nothing
    Set FileObj = Nothing

    On Error Resume Next
    MileCount = CInt(Mid(fileline(1),7)) 

    On Error Resume Next
    PlayerScore = CInt(Mid(fileline(2),7)) 
	If PlayerScore = 0 Then PlayerScore = 1234

	'UpdateBackglassInfoDisplay()
	ReelScore.setvalue PlayerScore
	ReelOdometer.setvalue MileCount
    If B2SOn then
        Controller.B2SSetScorePlayer 1, PlayerScore
		B2sOdometer()
    end if

End Sub

Sub SaveConfig
    Dim FileObj
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory)then
        Exit Sub
    End If
	Dim ConfigFile
    Set ConfigFile = FileObj.CreateTextFile(UserDirectory & TableName & ".txt", True)
	ConfigFile.WriteLine "Miles=" & MileCount
	ConfigFile.WriteLine "Score=" & PlayerScore
    ConfigFile.Close
    Set ConfigFile = Nothing
    Set FileObj = Nothing
End Sub


'****************************************
' Realtime updates
'****************************************

Sub GameTimer_Timer
	Ballcheck
    RollingUpdate
End Sub

Sub Ballcheck
	'if any ball is off of the playfield, set the location right in front of the drain so it gets destroyed
    Dim BOT, b
    BOT = GetBalls

    For b = 0 to UBound(BOT)
		If BOT(b).X < 0 or BOT(b).X > Table1.Width or BOT(b).Y < 0 or BOT(b).Y > Table1.Height or BOT(b).Z < -10 Then
			'debug.print b & " " & BOT(b).X & " " & BOT(b).Y
			BOT(b).x = Drain.x
			BOT(b).y = Drain.y - BOT(b).Radius
            BOT(b).z = 0
		End If
		
	Next
End Sub

Sub cleanupTimer_Timer
	'starting new game. if any balls left from testing, set the location right in front of the drain so it gets destroyed
	'BUT, use a timer to spread it out because if there are a bunch of them they pile up
    Dim BOT, b
    BOT = GetBalls

    'debug.print UBound(BOT)
	If Ubound(BOT) >= 0 Then
		BOT(b).x = Drain.x
		BOT(b).y = Drain.y - BOT(b).Radius
	Else
		cleanupTimer.Enabled = false
		Game_init
	End If
End Sub

' ***********************************************************************
'  *********************************************************************
'   *********     G A M E  C O D E  S T A R T S  H E R E      *********
'  *********************************************************************
' ***********************************************************************
Sub LightArrow

	'all arrows lit
	If PlArrowLoopLeft.state = 1 And PlArrowLoopRight.state = 1 And PlArrowCenterSpinner.state = 1 And PlArrowKicker.state = 1 And PlArrowTargetMidLeft.state = 1 Then Exit Sub

	If PlArrowLoopLeft.state = 2 Then PlArrowLoopLeft.state = 0
	If PlArrowLoopRight.state = 2 Then PlArrowLoopRight.state = 0
	If PlArrowCenterSpinner.state = 2 Then PlArrowCenterSpinner.state = 0
	If PlArrowKicker.state = 2 Then PlArrowKicker.state = 0
	If PlArrowTargetMidLeft.state = 2 Then PlArrowTargetMidLeft.state = 0

    Dim isDone: isDone=False
    do until isDone

		Randomize timer
		Dim aRnd : aRnd = Int((5 * Rnd) + 1)

		If aRnd = 1 and PlArrowLoopLeft.state = 0 Then 
			PlArrowLoopLeft.state = 2
			isDone = True
		End If
		If aRnd = 2 and PlArrowLoopRight.state = 0 Then 
			PlArrowLoopRight.state = 2
			isDone = True
		End If
		If aRnd = 3 and PlArrowCenterSpinner.state = 0 Then 
			PlArrowCenterSpinner.state = 2
			isDone = True
		End If
		If aRnd = 4 and PlArrowKicker.state = 0 Then 
			PlArrowKicker.state = 2
			isDone = True
		End If
		If aRnd = 5 and PlArrowTargetMidLeft.state = 0 Then 
			PlArrowTargetMidLeft.state = 2
			isDone = True
		End If
    loop

End Sub


'**********************
'       Slingshots
'**********************

Dim LStep, RStep

Sub SlingshotLeftWall_Slingshot   'left slingshot
	'debug.print "slingshot"
    If isTilted Then Exit Sub
    PlaySoundAt "fx_slingshot", Lemk
	
	FlashSlingLowLeft.state = 1
    LStep = 0
    SlingshotLeftWall.TimerEnabled = True
    AddScore 10

End Sub

Sub SlingShotLeftWall_Timer    
    Select Case LStep
        Case 0:LeftSling2.Visible = 1:Lemk.RotX = 2
        Case 1:LeftSling3.Visible = 1:LeftSling2.Visible = 0:Lemk.RotX = 14
        Case 2:LeftSling4.Visible = 1:LeftSling3.Visible = 0:Lemk.RotX = 26
        Case 3:LeftSling3.Visible = 1:LeftSling4.Visible = 0:Lemk.RotX = 14
        Case 4:LeftSling2.Visible = 1:LeftSling3.Visible = 0:Lemk.RotX = 2
        Case 5:LeftSling2.Visible = 0:Lemk.RotX = -20:FlashSlingLowLeft.state = 0:SlingShotLeftWall.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub SlingShotRightWall_Slingshot    'right slingshot
    If isTilted Then Exit Sub
    PlaySoundAt "fx_slingshot", Remk

	FlashSlingLowRight.state = 1
    RStep = 0
    SlingShotRightWall.TimerEnabled = True
    AddScore 10

End Sub

Sub SlingShotRightWall_Timer
    Select Case RStep
        Case 0:RightSling2.Visible = 1:Remk.RotX = 2
        Case 1:RightSling3.Visible = 1:RightSling2.Visible = 0:Remk.RotX = 14
        Case 2:RightSling4.Visible = 1:RightSling3.Visible = 0:Remk.RotX = 26
        Case 3:RightSling3.Visible = 1:RightSling4.Visible = 0:Remk.RotX = 14
        Case 4:RightSling2.Visible = 1:RightSling3.Visible = 0:Remk.RotX = 2
        Case 5:RightSling2.Visible = 0:Remk.RotX = -20:FlashSlingLowRight.state = 0:SlingShotRightWall.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'***********************
'       Rubbers
'***********************

Sub LeftDropTargetSlingWall_Hit
    'upper left slingshot behind drop targets
	FlashSlingUpLeft.duration 1, 100, 0
	AddScore 1
End Sub

Sub RightDropTargetSlingWall_Hit
	'upper right slingshot behind drop targets
	FlashSlingUpRight.duration 1, 100, 0
	AddScore 1
End Sub


'***************************************************************
' Backglass display
'***************************************************************

Sub UpdateBackglassInfoDisplay()

	'handles backglass and backdrop display for things other than scores
	' shoot again || game over || credits || active player || number of players || player turn (aka ball in play) || tilt

	If isGameInPlay Then
		ReelGameOver.SetValue 0
		If B2SOn Then Controller.B2SSetGameOver 0
	Else
        ReelGameOver.SetValue 1
		If B2SOn Then Controller.B2SSetGameOver 1
    End If

    Dim ballinplay
    ballinplay = BallsPerGame - BallsRemaining + 1
	If ballinplay < 0 OR NOT isGameInPlay Then
		ballinplay = 0
	End If

    ReelBallInPlay.SetValue ballinplay
    If B2SOn Then Controller.B2SSetBallInPlay ballinplay


    If isTilted Then
        ReelTilt.SetValue 1
        If B2SOn Then Controller.B2SSetTilt 1
	Else
        ReelTilt.SetValue 0
        If B2SOn Then Controller.B2SSetTilt 0
	End If
            
End Sub

Sub OdometerTimer_Timer
	'we are traveling at the incredible speed of 1 mile per second = 3600 miles per hour - about 10 times the speed of sound.
	'so...uh...better wear a seat belt or something...

	If isGameInPlay=False OR isTilted OR isProjectorEnabled = False Then 
		OdometerTimer.Enabled=False
		Exit Sub
	End If

    If BonusPicTimer.enabled = True Then
		Exit Sub
	End If

    If GasMilesLeft <= 0 Then
		OdometerTimer.Enabled=False
        FlasherProjector.ImageA = "GetGas" & Int(10 * Rnd)
		Exit Sub
	End If

    If FlasherProjector.ImageA <> RecentPicId And RecentPicId <> "" Then FlasherProjector.ImageA = RecentPicId

	If MileCount > LastMile + 4 Then 
		MileCount = 0
	ElseIf (MileCount < 0) OR (MileCount = 0 AND OdometerAddAmount < 0) Then 
		MileCount = LastMile
	Else
		MileCount = MileCount + OdometerAddAmount
	End If

	If MileCount <= LastMile and MileCount >= 0 Then
		ReelOdometer.Setvalue MileCount
	End If
    MilesThisTurn = MilesThisTurn +  1
	B2sOdometer()
	'If B2SOn then Controller.B2SSetScorePlayer2 MileCount

	'dig this: all of our pictures are in the image manager and named "mile####"
	'we have them all listed in this array.  Each time we hit this routine, we take the current mile we are on and generate an image name
	'then we look in the array.  If we find a match, we change the picture showing in the "projector" flasher object.


	'gotta add some leading zeros to make the name format "mile####"
	If milecount < 9 Then
		PicId = "mile000"& milecount
	ElseIf milecount < 99 Then
		PicId = "mile00" & milecount
	ElseIf milecount < 999 Then
		PicId = "mile0" & milecount
	ElseIf milecount < 9999 Then
		PicId = "mile" & milecount
	End If

	'debug.print PicId
	'checking the array for a match
	dim aPicname
    For each aPicname in PicArray
        If ("mile" & aPicname) = PicId then
			'debug.print z
            RecentPicId = PicId
			FlasherProjector.ImageA = PicId
		End if
    Next

    UseGas()

End Sub

Sub B2sOdometer()

	If NOT B2SOn Then Exit Sub
	If MileCount > LastMile or MileCount < 0 Then Exit Sub

'ones
	Dim x:	x = MileCount mod 10
    Select Case x
        Case 0:Controller.B2SSetData 1, 10
		Case 1:Controller.B2SSetData 1, 1
		Case 2:Controller.B2SSetData 1, 2
		Case 3:Controller.B2SSetData 1, 3
		Case 4:Controller.B2SSetData 1, 4
		Case 5:Controller.B2SSetData 1, 5
		Case 6:Controller.B2SSetData 1, 6
		Case 7:Controller.B2SSetData 1, 7
		Case 8:Controller.B2SSetData 1, 8
		Case 9:Controller.B2SSetData 1, 9
    End Select

'tens
	x = INT(MileCount / 10) mod 10
    Select Case x
        Case 0:Controller.B2SSetData 2, 10
		Case 1:Controller.B2SSetData 2, 1
		Case 2:Controller.B2SSetData 2, 2
		Case 3:Controller.B2SSetData 2, 3
		Case 4:Controller.B2SSetData 2, 4
		Case 5:Controller.B2SSetData 2, 5
		Case 6:Controller.B2SSetData 2, 6
		Case 7:Controller.B2SSetData 2, 7
		Case 8:Controller.B2SSetData 2, 8
		Case 9:Controller.B2SSetData 2, 9
    End Select

'hundreds
	x = INT(MileCount / 100) mod 10
    Select Case x
        Case 0:Controller.B2SSetData 3, 10
		Case 1:Controller.B2SSetData 3, 1
		Case 2:Controller.B2SSetData 3, 2
		Case 3:Controller.B2SSetData 3, 3
		Case 4:Controller.B2SSetData 3, 4
		Case 5:Controller.B2SSetData 3, 5
		Case 6:Controller.B2SSetData 3, 6
		Case 7:Controller.B2SSetData 3, 7
		Case 8:Controller.B2SSetData 3, 8
		Case 9:Controller.B2SSetData 3, 9
    End Select

'thousands
	x = INT(MileCount / 1000) mod 10
	
    Select Case x
        Case 0:Controller.B2SSetData 4, 10
		Case 1:Controller.B2SSetData 4, 1
		Case 2:Controller.B2SSetData 4, 2
		Case 3:Controller.B2SSetData 4, 3
		Case 4:Controller.B2SSetData 4, 4
		Case 5:Controller.B2SSetData 4, 5
		Case 6:Controller.B2SSetData 4, 6
		Case 7:Controller.B2SSetData 4, 7
		Case 8:Controller.B2SSetData 4, 8
		Case 9:Controller.B2SSetData 4, 9
    End Select
End Sub


'*******************************************
'  Table Options (F12 menu)
'  replaces the old start keys:
'    2 = reset miles, 3 = projector off, 4 = odometer counts down
'*******************************************

Sub Table1_OptionEvent(ByVal eventId)
	Dim x

	x = Table1.Option("Slideshow/Gas/Odometer", 0, 1, 1, 1, 0, Array("Off", "On"))
	If x = 1 Then isProjectorEnabled = True Else isProjectorEnabled = False

	x = Table1.Option("Odometer Direction", 0, 1, 1, 0, 0, Array("Forward", "Reverse"))
	If x = 1 Then OdometerAddAmount = -1 Else OdometerAddAmount = 1

	x = Table1.Option("Reset Miles at Game Start", 0, 1, 1, 0, 0, Array("No", "Yes"))
	If x = 1 Then isResetMilesOnStart = True Else isResetMilesOnStart = False
End Sub
