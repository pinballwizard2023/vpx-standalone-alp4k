' ****************************************************************
'                         VISUAL PINBALL X
'                          Zandys Arcade
'                         Killer Instinct
'		plain VPX script using core.vbs for supporting functions
'                         Version 1.0
' ****************************************************************
'*******DOF Config by floeschng and DarrinVPCLE********

'E101	Flipper Left
'E102	Flipper Right
'E103	Slingshot Left
'E104	Slingshot Right
'E105	Slingshot Left Effect
'E106	Slingshot Right Effect
'E107	Bumper 3
'E109	Bumper 1
'E110	Bumper 2
'E112	Target 12
'E113	Target 5/Target 7
'E114	Target 10/Target 8
'E115	Target 9
'E116	Target 13
'E117	Target 1
'E118	Started/In Play Flag
'E119	Sabrewulf Exit (Kicker)
'E120	Target 2/Target 4
'E121	Auto Plunger+Sabrewulf Exit+Extraball Effect
'E122	Knocker
'E123	Ball Release
'E124	Spinner 2
'E125	Coins in/Ready
'E126	Jackpot/SuperCombo/UltraCombo/Winbattle
'E127	Skillshot
'E128	Switch 1
'E129	Switch 6
'E130	Switch 8
'E131	Switch 7
'E132	Switch 2 Left Outlane
'E133	Switch 4 Left Inlane
'E134	Switch 3 Right Inlane
'E135	Switch 5 Right Outlane
'E136	Spinner 1
'E137	Bumper 3 Flasher
'E138	Bumper 1 Flasher
'E139	Battle Completed
'E140	Bumper 2 Flasher
'E141	Autoplunger Kicker
'E142	Ball in Plunger Lane
'E143	Multiball Flag
'E311	Left Ramp Done
'E312	Right Ramp Done
'E323	Attract Mode
'E401   Ball 1 Locked
'E402   Ball 2 Locked
'E403   Ball 3 Locked
'E404   Flame Effect
'E405   Ball Saved
'E406   Right Blood 1
'E407   Left Blood
'E408   Right Blood 2
'E409   Ball Drain 1
'E410   Ball Drain 2
'E411   Ball Drain 3
'E412   Rainbow Lights
'E413   Ball Launch
'E414   Lock Hit  
'E415   Add Credit
'E416   Multiball
'E417   Lock Is Lit
'E418   Game Over
' *******************************
Option Explicit
Randomize

Const BallSize = 50  ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1   ' normal ball mass
Const SongVolume = 0.1 ' 1 is full volume. Value is from 0 to 1

'FlexDMD in high or normal quality
'change it to True if you have an LCD screen, 256x64
'or keep it False if you have a real DMD at 128x32 in size
Const FlexDMDHighQuality = False

' Define any Constants
Const cGameName = "KillerInstinct"
Const TableName = "KIllerInstinct"
Const myVersion = "1.0"
Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 20 ' in seconds
Const MaxMultiplier = 5  ' limit to 5x in this game, both bonus multiplier and playfield multiplier
Const BallsPerGame = 3   ' usually 3 or 5
Const MaxMultiballs = 5  ' max number of balls during multiballs
Const FlippersBlood = True 
'----- General Sound Options -----
Const VolumeDial = 0.9			  'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 1.0		  'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 1.0		  'Level of ramp rolling volume. Value between 0 and 1

' Use FlexDMD if in FS mode
Dim UseFlexDMD
If Table1.ShowDT = True then
    UseFlexDMD = False
Else
    UseFlexDMD = True
End If

' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
End Sub

'*************************************************************
' VR Room Auto-Detect
'*************************************************************
Const LiveViewVRSim = 0				' 0 = Default, 1 = View table in VR mode in "Live View Editor" 

' VR uses FlexDMD to display DMD.
Dim VR_Obj, VRMode

If RenderingMode = 2 or LiveViewVRSim = 1 Then
	VRMode = True
    UseFlexDMD = True
	lrail.Visible = 0
	rrail.Visible = 0
	For Each VR_Obj in VRCabinet:VR_Obj.Visible = 1:Next
	For Each VR_Obj in VRWorld:VR_Obj.Visible = 1:Next
Else
	VRMode = False
	For Each VR_Obj in VRCabinet:VR_Obj.Visible = 0:Next
	For Each VR_Obj in VRWorld:VR_Obj.Visible = 0:Next
End If

' Define Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim PlayfieldMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot(4)
Dim SuperJackpot
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock(4)
Dim BallsInHole

' Define Game Flags
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bBallInPlungerLane
Dim bBallSaverActive
Dim bBallSaverReady
Dim bMultiBallMode
Dim bMusicOn
Dim bSkillshotReady
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bJackpot

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs
Dim cbRight   'captive ball
Dim bsJackal
Dim x

dim countr
countr = 0
FireTimer.enabled = 1
LAVATimer.enabled = 1
ARCATimer.enabled = 1

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize

    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 50 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime

        .Random 1.5
        .InitExitSnd SoundFXDOF("fx_kicker", 123, DOFPulse, DOFContactors), SoundFXDOF("fx_solenoid", 123, DOFPulse, DOFContactors)
        .CreateEvents "plungerIM"
    End With

    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 0
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        '.CreateEvents "cbRight"
        .Start
    End With
    CapKicker1.CreateSizedBallWithMass BallSize / 2, BallMass

    ' Jackal hole
    Set bsJackal = New cvpmTrough
    With bsJackal
        .size = 5
        .Initexit JackalHole, 160, 35
        '.InitExitVariance 2, 2
        .MaxBallsPerKick = 1
    End With

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Loadhs

    ' Initalise the DMD display
    DMD_Init

    ' freeplay or coins
    bFreePlay = False 'we want coins

    if bFreePlay Then DOF 125, DOFOn

    ' Init main variables and any other flags
    bAttractMode = False
    bOnTheFirstBall = False
    bBallInPlungerLane = False
    bBallSaverActive = False
    bBallSaverReady = False
    bMultiBallMode = False
    bGameInPlay = False
    bAutoPlunger = False
    bMusicOn = True
    BallsOnPlayfield = 0
    BallsInLock(1) = 0
    BallsInLock(2) = 0
    BallsInLock(3) = 0
    BallsInLock(4) = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bJackpot = False
    bInstantInfo = False
    ' set any lights for the attract mode
    GiOff
    StartAttractMode

    ' Start the RealTime timer
    RealTime.Enabled = 1

    ' Load table color
    LoadLut
End Sub

'******************
' Captive Ball Subs
'******************
Sub CapTrigger1_Hit:cbRight.TrigHit ActiveBall:End Sub
Sub CapTrigger1_UnHit:cbRight.TrigHit 0:End Sub
Sub CapWall1_Hit:cbRight.BallHit ActiveBall:PlaySoundAtBall "Ball_Collide_1":End Sub
Sub CapKicker1a_Hit:cbRight.BallReturn Me:End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = True: SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If

    If Keycode = AddCreditKey Then
        DOF 415, DOFPulse
        Credits = Credits + 1
        if bFreePlay = False Then DOF 125, DOFOn
        If(Tilted = False)Then
            DMDFlush
            DMD "_", CL("CREDITS: " & Credits), "", eNone, eNone, eNone, 500, True, "fx_coin"
            If NOT bGameInPlay Then ShowTableInfo
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
		SoundPlungerPull()
		If VRMode = True Then
			TimerVRPlunger.Enabled = True
			TimerVRPlunger2.Enabled = False
		End If
    End If

	If keycode = LeftFlipperKey Then
		If VRMode = True Then PinCab_Flipper_Button_Left.X = PinCab_Flipper_Button_Left.X + 10
		FlipperActivate LeftFlipper, LFPress
	End If

	If keycode = RightFlipperKey Then
		If VRMode = True Then PinCab_Flipper_Button_Right.X = PinCab_Flipper_Button_Right.X - 10
		FlipperActivate RightFlipper, RFPress
	End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If

    ' Table specific

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 8:SoundNudgeLeft:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 8:SoundNudgeRight:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 9:SoundNudgeCenter:CheckTilt

        If keycode = LeftFlipperKey Then
			FlipperActivate LeftFlipper, LFPress
			SolLFlipper 1
			InstantInfoTimer.Enabled = True
		End If

        If keycode = RightFlipperKey Then
			FlipperActivate RightFlipper, RFPress
			SolRFlipper 1
			InstantInfoTimer.Enabled = True
		End If

        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    TotalGamesPlayed = TotalGamesPlayed + 1
                    DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 500, True, "so_fanfare1"
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
                        DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 500, True, "so_fanfare1"
                        If Credits < 1 And bFreePlay = False Then DOF 125, DOFOff
                        Else
                            ' Not Enough Credits to start a game.
                            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 500, True, "so_nocredits"
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                If(bFreePlay = True)Then
                    If(BallsOnPlayfield = 0)Then
                        ResetForNewGame()
                    End If
                Else
                    If(Credits > 0)Then
                        If(BallsOnPlayfield = 0)Then
                            Credits = Credits - 1
                            If Credits < 1 And bFreePlay = False Then DOF 125, DOFOff
                            ResetForNewGame()
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 500, True, "so_nocredits"
                        ShowTableInfo
                    End If
                End If
            End If
    End If ' If (GameInPlay)

'test keys
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = LeftMagnaSave Then bLutActive = False: HideLUT

     If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
        If bBallInPlungerLane Then PlaySoundAt "fx_fire", plunger
		If VRMode = True Then
			TimerVRPlunger.Enabled = False
			TimerVRPlunger2.Enabled = True
			Pincab_Shooter.Y = 0
		End If
    End If

	If keycode = LeftFlipperKey Then
		If VRMode = True Then PinCab_Flipper_Button_Left.X = PinCab_Flipper_Button_Left.X - 10
		FlipperDeActivate LeftFlipper, LFPress
	End If

	If keycode = RightFlipperKey Then
		If VRMode = True Then PinCab_Flipper_Button_Right.X = PinCab_Flipper_Button_Right.X + 10
		FlipperDeActivate RightFlipper, RFPress
	End If

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
			FlipperDeActivate LeftFlipper, LFPress
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
        If keycode = RightFlipperKey Then
			FlipperDeActivate RightFlipper, RFPress
            SolRFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
            End If
        End If
    End If
End Sub

Sub InstantInfoTimer_Timer
    InstantInfoTimer.Enabled = False
    If NOT hsbModeActive Then
        bInstantInfo = True
        DMDFlush
        InstantInfo
    End If
End Sub

Sub InstantInfo
    DMD CL("INSTANT INFO"), "", "", eNone, eNone, eNone, 800, False, ""
    DMD CL("JACKPOT VALUE"), CL(Jackpot(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("SPINNER VALUE"), CL(spinnervalue(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("BUMPER VALUE"), CL( bumpervalue(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("BONUS X"), CL(BonusMultiplier(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("PLAYFIELD X"), CL(PlayfieldMultiplier(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("LOCKED BALLS"), CL(BallsInLock(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("LANE BONUS"), CL(LaneBonus), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("TARGET BONUS"), CL(TargetBonus), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("RAMP BONUS"), CL(RampBonus), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("COMBO BREAKERS"), CL(MonstersKilled(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("HIGHEST SCORE"), CL(HighScoreName(0) & " " & HighScore(0)), "", eNone, eNone, eNone, 800, False, ""
End Sub

'*************
' Pause Table
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub Table1_Exit
    Savehs
    If UseFlexDMD Then FlexDMD.Run = False
    If B2SOn = true Then Controller.Stop
End Sub

'********************
'     Flippers
'********************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
		LF.Fire
		If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then 
			RandomSoundReflipUpLeft LeftFlipper
			DOF 101, DOFOn
		Else 
			SoundFlipperUpAttackLeft LeftFlipper
			RandomSoundFlipperUpLeft LeftFlipper
			DOF 101, DOFOn
		End If		
	Else
		LeftFlipper.RotateToStart
		If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper
			DOF 101, DOFOff
		End If
		FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub


Sub SolRFlipper(Enabled)
    If Enabled Then
		RF.Fire
		If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight RightFlipper
			DOF 102, DOFOn
		Else 
			SoundFlipperUpAttackRight RightFlipper
			RandomSoundFlipperUpRight RightFlipper
			DOF 102, DOFOn
		End If
	Else
		RightFlipper.RotateToStart
		If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
			DOF 102, DOFOff
		End If	
		FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub LeftFlipper_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch Activeball, RightFlipper, RFCount, parm
	RightFlipperCollide parm
End Sub

Dim RSplat, LSplat, R2Splat

Sub RightSplat
    RSplat = 0
    Rightblood_Timer
End Sub

Sub Rightblood_Timer
    Select Case RSplat
        Case 0:Rightblood.ImageA = "blood1":Rightblood.Visible = 1:Rightblood.TimerEnabled = 1
        Case 1:Rightblood.ImageA = "blood2"
        Case 2:Rightblood.ImageA = "blood3"
        Case 3:Rightblood.ImageA = "blood4"
        Case 4:Rightblood.ImageA = "blood5"
        Case 5:Rightblood.ImageA = "blood6"
        Case 6:Rightblood.Visible = 0:Rightblood.TimerEnabled = 0
    End Select
    RSplat = RSplat + 1
End Sub

Sub LeftSplat
    LSplat = 0
    Leftblood_Timer
End Sub

Sub Leftblood_Timer
    Select Case LSplat
        Case 0:Leftblood.ImageA = "blood1a":Leftblood.Visible = 1:Leftblood.TimerEnabled = 1
        Case 1:Leftblood.ImageA = "blood2a"
        Case 2:Leftblood.ImageA = "blood3a"
        Case 3:Leftblood.ImageA = "blood4a"
        Case 4:Leftblood.ImageA = "blood5a"
        Case 5:Leftblood.ImageA = "blood6a"
        Case 6:Leftblood.Visible = 0:Leftblood.TimerEnabled = 0
    End Select
    LSplat = LSplat + 1
End Sub

Sub Right2Splat
    R2Splat = 0
    Rightblood2_Timer
End Sub


Sub Rightblood2_Timer
    Select Case R2Splat
        Case 0:Rightblood2.ImageA = "blood1":Rightblood2.Visible = 1:Rightblood2.TimerEnabled = 1
        Case 1:Rightblood2.ImageA = "blood2"
        Case 2:Rightblood2.ImageA = "blood3"
        Case 3:Rightblood2.ImageA = "blood4"
        Case 4:Rightblood2.ImageA = "blood5"
        Case 5:Rightblood2.ImageA = "blood6"
        Case 6:Rightblood2.Visible = 0:Rightblood2.TimerEnabled = 0
    End Select
    R2Splat = R2Splat + 1
End Sub

'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity)AND(Tilt < 15)Then 'show a warning
        DMD "_", CL("CAREFUL"), "_", eNone, eBlinkFast, eNone, 500, True, ""
    End if
    If Tilt > 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
        DMDFlush
        DMD "", CL("TILT"), "", eNone, eNone, eNone, 200, False, ""
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        'Bumper1.Force = 0

        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        'Bumper1.Force = 6
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0)Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'********************
' Music as wav sounds
'********************

Dim Song
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            PlaySound Song, -1, SongVolume
        End If
    End If
End Sub

Sub PlayBattleSong
    Dim tmp
    tmp = INT(RND * 1)
    Select Case tmp
        Case 0:PlaySong "mu_battle1"
        
    End Select
End Sub

Sub PlayMultiballSong
    Dim tmp
    tmp = INT(RND * 1)
    Select Case tmp
        Case 0:PlaySong "mu_war1"
        
    End Select
End Sub

Sub ChangeSong
    If(BallsOnPlayfield = 0)Then
        PlaySong "mu_end"
        Exit Sub
    End If

    If bMultiBallMode Then
        PlayMultiballSong
    Else
        Select Case Battle(CurrentPlayer, 0)
            Case 0
                PlaySong "mu_main"
            Case 1
                PlaySong "CTHUNDER"
            Case 2
                PlaySong "RIPTOR"
            Case 3
                PlaySong "CINDER"
            Case 4
                PlaySong "SABREWULF"
            Case 5
                PlaySong "SPINAL"
            Case 6
                PlaySong "FULGORE"
            Case 7
                PlaySong "TJCOMBO"
            Case 8
                PlaySong "JAGO"
            Case 9
                PlaySong "EYEDOL"
            Case 10
                PlaySong "ORCHID"
            Case 11
                PlaySong "GLACIUS"
            Case 12
                PlaySong "ULTRATECH"
            Case 13
                Playsong "FINALBATTLE"

                PlayBattleSong
                PlayMultiballSong
        End Select
    End If
End Sub

'**********************
'     GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'**********************

Dim OldGiState
OldGiState = -1   'start witht the Gi off

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 1 Then 'we have 2 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    DOF 118, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    For each bulb in aBumperLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    DOF 118, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    For each bulb in aBumperLights
        bulb.State = 0
    Next
End Sub

' GI, light & flashers sequence effects

Sub GiEffect(n)
    Dim ii
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 10, 10
    End Select
End Sub

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 10, 10
    End Select
End Sub

Sub FlashEffect(n)
    Dim ii
    Select case n
        Case 0 ' all off
            LightSeqFlasher.Play SeqAlloff
        Case 1 'all blink
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqBlinking, , 10, 10
        Case 2 'random
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqBlinking, , 5, 10
    End Select
End Sub

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v4.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 2     'number of locked balls
Const maxvel = 40 'max ball velocity

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True

    'resets the score display, and turn off attract mode
    StopAttractMode
    GiOn

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        PlayfieldMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

    ' initialise any other flags
    Tilt = 0

    ' initialise Game variables
    Game_Init()

    ' you may wish to start some music, play a sound, do whatever at this point

    vpmtimer.addtimer 1500, "FirstBall '"
End Sub

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with

Sub FirstBall
    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' create a new ball in the shooters lane
    CreateNewBall()
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
    ' make sure the correct display is upto date
    AddScore 0

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1

    ' reduce the playfield multiplier
    SetPlayfieldMultiplier 1

    ' reset any drop targets, lights, game Mode etc..

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False

    'Reset any table specific
    ResetNewBallVariables
    ResetNewBallLights()

    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

    'and the skillshot
    bSkillShotReady = True

'Change the music ?
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield > 1 Then
        DOF 143, DOFPulse
        bMultiBallMode = True
        bAutoPlunger = True
        ChangeGi 5
    End If
End Sub

' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
    'and eject the first ball
    CreateMultiballTimer_Timer
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If BallsOnPlayfield < MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                CreateMultiballTimer.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            CreateMultiballTimer.Enabled = False
        End If
    End If
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded

Sub EndOfBall()
    Dim AwardPoints, TotalBonus, ii
    AwardPoints = 0
    TotalBonus = 0
    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

    If NOT Tilted Then

'add in any bonus points (multipled by the bonus multiplier)
'AwardPoints = BonusPoints(CurrentPlayer) * BonusMultiplier(CurrentPlayer)
'AddScore AwardPoints
'debug.print "Bonus Points = " & AwardPoints
'DMD "", CL("BONUS: " & BonusPoints(CurrentPlayer) & " X" & BonusMultiplier(CurrentPlayer) ), "", eNone, eBlink, eNone, 1000, True, ""

'Count the bonus. This table uses several bonus
'Lane Bonus
        AwardPoints = LaneBonus * 1000
        TotalBonus = AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("LANE BONUS " & LaneBonus), "", eBlink, eNone, eNone, 800, False, ""

        'Number of Target hits
        AwardPoints = TargetBonus * 2000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("TARGET BONUS " & TargetBonus), "", eBlink, eNone, eNone, 800, False, ""
        
        'Combos X 150.000
        AwardPoints = ComboHits(CurrentPlayer) * 150000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL("RAMP COMBOS"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, "mu_bonus"

        'Number of Ramps completed
        AwardPoints = RampBonus * 10000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("RAMP BONUS " & RampBonus), "", eBlink, eNone, eNone, 800, False, ""

        'Number of Monsters Killed
        AwardPoints = MonstersKilled(CurrentPlayer) * 25000
        TotalBonus = TotalBonus + AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("COMBO BREAKER " & MonstersKilled(CurrentPlayer)), "", eBlink, eNone, eNone, 800, False, ""

               Select Case Balls
 
    Case 1: PlaySound"BallLost1"
            DOF 409, DOFPulse
    Case 2: PlaySound"BallLost2"
            DOF 410, DOFPulse
    Case 3: PlaySound"BallLost3"
            DOF 411, DOFPulse
End Select

        ' calculate the totalbonus
        TotalBonus = TotalBonus * BonusMultiplier(CurrentPlayer) + BonusHeldPoints(CurrentPlayer)

        ' handle the bonus held
        ' reset the bonus held value since it has been already added to the bonus
        BonusHeldPoints(CurrentPlayer) = 0

        ' the player has won the bonus held award so do something with it :)
        If bBonusHeld Then
            If Balls = BallsPerGame Then ' this is the last ball, so if bonus held has been awarded then double the bonus
                TotalBonus = TotalBonus * 2
            End If
        Else ' this is not the last ball so save the bonus for the next ball
            BonusHeldPoints(CurrentPlayer) = TotalBonus
        End If
        bBonusHeld = False

        ' Add the bonus to the score
        DMD CL(FormatScore(TotalBonus)), CL("TOTAL BONUS " & " X" & BonusMultiplier(CurrentPlayer)), "", eBlinkFast, eNone, eNone, 1500, True, ""

        AddScore TotalBonus

        ' add a bit of a delay to allow for the bonus points to be shown & added up
        vpmtimer.addtimer 6000, "EndOfBall2 '"
    Else 'if tilted then only add a short delay
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the CurrentPlayer)
'
Sub EndOfBall2()
    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0)Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        DMD CL("EXTRA BALL"), CL("SHOOT AGAIN"), "", eNone, eNone, eBlink, 1000, True, ""

        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0)Then
        DOF 418, DOFPulse
            'debug.print "No More Balls, High Score Entry"

            ' Submit the CurrentPlayers score to the High Score system
            CheckHighScore()
        ' you may wish to play some music at this point

        Else

            ' not the last ball (for that player)
            ' if multiple players are playing then move onto the next one
            EndOfBallComplete()
        End If
    End If
End Sub

' This function is called when the end of bonus display
' (or high score entry finished) AND it either end the game or
' move onto the next player (or the next ball of the same player)
'
Sub EndOfBallComplete()
    Dim NextPlayer

    'debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1)Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        EndOfGame()

    ' you may wish to put a Game Over message on the desktop/backglass

    Else
        ' set the next player
        CurrentPlayer = NextPlayer

        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()

        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame > 1 Then
            PlaySound "vo_player" &CurrentPlayer
            DMD "_", CL("PLAYER " &CurrentPlayer), "_", eNone, eNone, eNone, 800, True, ""
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
    'debug.print "End Of Game"
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    If NOT bJustStarted Then
        ChangeSong
    End If

    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball

    ' set any lights for the attract mode
    GiOff
    StartAttractMode
' you may wish to light any Game Over Light you may have
End Sub

Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Sub Drain_Hit()
   
    ' Destroy the ball
    Drain.DestroyBall
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' pretend to knock the ball into the ball storage mech
    RandomSoundDrain Drain
    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True)AND(Tilted = False)Then

        ' is the ball saver active,
        If(bBallSaverActive = True)Then

            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            ' we kick the ball with the autoplunger
            bAutoPlunger = True
            ' you may wish to put something on a display or play a sound at this point
            DMD "_", CL("BALL SAVED"), "_", eNone, eBlinkfast, eNone, 800, True, "Fighton"
            DOF 405, DOFPulse
        	DOF 142, DOFOff
        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1)Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True)then
                    ' not in multiball mode any more
                    bMultiBallMode = False
                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
                    ResetJackpotLights
                    Select Case Battle(CurrentPlayer, 0)
                        Case 13:WinBattle
                    End Select
                    ChangeGi white
                    ChangeSong
                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0)Then
                ' End Mode and timers
                ChangeSong
                ChangeGi white
                ' Show the end of ball animation
                ' and continue with the end of ball
                ' DMD something?
                StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, since there is no animation then move to the end of ball
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySoundAt "fx_sensor", swPlungerRest
    bBallInPlungerLane = True
	DOF 142, DOFOn
    ' turn on Launch light is there is one
    'LaunchLight.State = 2

    'be sure to update the Scoreboard after the animations, if any

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"

activeball.vely = -30
'        PlungerIM.AutoFire
        DOF 121, DOFPulse
		DOF 141, DOFPulse
        PlaySoundAt "fx_fire", swPlungerRest
        bAutoPlunger = False
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True)AND(BallSaverTime <> 0)And(bBallSaverActive = False)Then
        EnableBallSaver BallSaverTime
    Else
        ' show the message to shoot the ball in case the player has fallen sleep
        swPlungerRest.TimerEnabled = 1
    End If
    'Start the Selection of the skillshot if ready
    If bSkillShotReady Then
        UpdateSkillshot()
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
	DOF 142, DOFOff
	DOF 413, DOFPulse
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ResetSkillShotTimer.Enabled = 1
    End If
    If bMultiballMode Then
        If BallsOnPlayfield = 2 Then
            ChangeSong
        End If
    Else
        ChangeSong
    End If
' turn off LaunchLight
' LaunchLight.State = 0
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds

Sub swPlungerRest_Timer
    DMD "_", CL("SHOOT THE BALL"), "_", eNone, eNone, eNone, 800, True, ""
    swPlungerRest.TimerEnabled = 0
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "Ballsaver started"
    ' set our game flag
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
    BallSaverTimerExpired.Interval = 1000 * seconds
    BallSaverTimerExpired.Enabled = True
    BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
    BallSaverSpeedUpTimer.Enabled = True
    ' if you have a ball saver light you might want to turn it on at this point (or make it flash)
    LightShootAgain.BlinkInterval = 160
    LightShootAgain.State = 2
End Sub

' The ball saver timer has expired.  Turn it off AND reset the game flag
'
Sub BallSaverTimerExpired_Timer()
    'debug.print "Ballsaver ended"
    BallSaverTimerExpired.Enabled = False
    ' clear the flag
    bBallSaverActive = False
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
' In this table we use SecondRound variable to double the score points in the second round after killing Malthael
Sub AddScore(points)
    If(Tilted = False)Then
        ' add the points to the current players score variable
        Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
    End if
' you may wish to check to see if the player has gotten a replay
End Sub

' Add bonus to the bonuspoints AND update the score board
Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If(Tilted = False)Then
        ' add the bonus to the current players bonus variable
        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
    End if
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points)
    ' Jackpots only generally increment in multiball mode AND not tilted
    ' but this doesn't have to be the case
    If(Tilted = False)Then

        ' If(bMultiBallMode = True) Then
        Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
        DMD "_", CL("INCREASED JACKPOT"), "_", eNone, eNone, eNone, 800, True, ""
    ' you may wish to limit the jackpot to a upper limit, ie..
    '	If (Jackpot >= 6000) Then
    '		Jackpot = 6000
    ' 	End if
    'End if
    End if
End Sub

Sub AddSuperJackpot(points) 'not used in this table
    If(Tilted = False)Then
    End if
End Sub

Sub AddBonusMultiplier(n)
    Dim NewBonusLevel
    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <= MaxMultiplier)then
        ' then add and set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier(NewBonusLevel)
        DMD "_", CL("BONUS X " &NewBonusLevel), "_", eNone, eNone, eNone, 2000, True, "fx_bonus"
    Else
        AddScore 50000
        DMD "_", CL("50000"), "_", eNone, eNone, eNone, 800, True, ""
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    UPdateBonusXLights(Level)
End Sub

Sub UpdateBonusXLights(Level)
    ' Update the lights
    Select Case Level
        Case 1:light56.State = 0:light57.State = 0:light58.State = 0:light59.State = 0
        Case 2:light56.State = 1:light57.State = 0:light58.State = 0:light59.State = 0
        Case 3:light56.State = 0:light57.State = 1:light58.State = 0:light59.State = 0
        Case 4:light56.State = 0:light57.State = 0:light58.State = 1:light59.State = 0
        Case 5:light56.State = 0:light57.State = 0:light58.State = 0:light59.State = 1
    End Select
End Sub

Sub AddPlayfieldMultiplier(n)
    Dim NewPFLevel
    ' if not at the maximum level x
    if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier)then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
        SetPlayfieldMultiplier(NewPFLevel)
        DMD "_", CL("PLAYFIELD X " &NewPFLevel), "_", eNone, eNone, eNone, 2000, True, "fx_bonus"
    Else 'if the 5x is already lit
        AddScore 50000
        DMD "_", CL("50000"), "_", eNone, eNone, eNone, 2000, True, ""
    End if
'Start the timer to reduce the playfield x every 30 seconds
' pfxtimer.Enabled = 0
' pfxtimer.Enabled = 1
End Sub

' Set the Playfield Multiplier to the specified level AND set any lights accordingly

Sub SetPlayfieldMultiplier(Level)
    ' Set the multiplier to the specified level
    PlayfieldMultiplier(CurrentPlayer) = Level
    UpdatePFXLights(Level)
End Sub

Sub UpdatePFXLights(Level)
' in this table the multiplier is always shown in the score display sub

' Update the lights
'Select Case Level
'    Case 1:light3.State = 0:light2.State = 0:light1.State = 0:light4.State = 0
'    Case 2:light3.State = 1:light2.State = 0:light1.State = 0:light4.State = 0
'    Case 3:light3.State = 0:light2.State = 1:light1.State = 0:light4.State = 0
'    Case 4:light3.State = 0:light2.State = 0:light1.State = 1:light4.State = 0
'    Case 5:light3.State = 0:light2.State = 0:light1.State = 0:light4.State = 1
'End Select
' show the multiplier in the DMD
End Sub

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        DMD "_", CL(("EXTRA BALL WON")), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("vo_extraball", 122, DOFPulse, DOFKnocker)
        DOF 121, DOFPulse
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1 'light the shoot again lamp
        GiEffect 2
        LightEffect 2
    END If
End Sub

Sub AwardSpecial()
    DMD "_", CL(("EXTRA GAME WON")), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    Credits = Credits + 1
    If bFreePlay = False Then DOF 125, DOFOn
    LightEffect 2
    FlashEffect 2
End Sub

Sub AwardJackpot() 'award a normal jackpot, double or triple jackpot
    Dim tmp
    DMD CL(FormatScore(Jackpot(CurrentPlayer))), CL("JACKPOT"), "d_border", eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    DOF 126, DOFPulse
    tmp = INT(RND * 2)
    Select Case tmp
        Case 0:PlaySound "vo_Jackpot"
        Case 1:PlaySound "vo_Jackpot"
        Case 2:PlaySound "vo_Jackpot"
    End Select
    AddScore Jackpot(CurrentPlayer)
    LightEffect 2
    FlashEffect 2
    'sjekk for superjackpot
    EnableSuperJackpot
End Sub

Sub AwardSuperJackpot() 'this is actually 4 times a jackpot
    SuperJackpot = Jackpot(CurrentPlayer) * 4
    DMD CL(FormatScore(SuperJackpot)), CL("SUPER JACKPOT"), "d_border", eBlinkFast, eBlinkFast, eNone, 1000, True, "vo_superjackpot"
    DOF 126, DOFPulse
    AddScore SuperJackpot
    LightEffect 2
    FlashEffect 2
    'enabled jackpots again
    StartJackpots
End Sub

Sub AwardSkillshot()
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL(FormatScore(SkillshotValue(CurrentPlayer))), CL(("SKILLSHOT")), "d_border", eBlinkFast, eBlink, eNone, 1000, True, ""
    DOF 127, DOFPulse
    PlaySound "fx_skillshot"
    FlashForMs f10, 1000, 50, 0
    Playsound "target_Thunder"
    FlashForMs f22, 1000, 50, 0
    Addscore SkillShotValue(CurrentPlayer)
    ' increment the skillshot value with 250.000
    SkillShotValue(CurrentPlayer) = SkillShotValue(CurrentPlayer) + 250000
    'do some light show
    GiEffect 2
    LightEffect 2
End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "")Then HighScore(0) = CDbl(x)Else HighScore(0) = 100000 End If
    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "")Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(TableName, "HighScore2")
    If(x <> "")then HighScore(1) = CDbl(x)Else HighScore(1) = 100000 End If
    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "")then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(TableName, "HighScore3")
    If(x <> "")then HighScore(2) = CDbl(x)Else HighScore(2) = 100000 End If
    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "")then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(TableName, "HighScore4")
    If(x <> "")then HighScore(3) = CDbl(x)Else HighScore(3) = 100000 End If
    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "")then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff:End If
    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
End Sub

Sub Savehs
    SaveValue TableName, "HighScore1", HighScore(0)
    SaveValue TableName, "HighScore1Name", HighScoreName(0)
    SaveValue TableName, "HighScore2", HighScore(1)
    SaveValue TableName, "HighScore2Name", HighScoreName(1)
    SaveValue TableName, "HighScore3", HighScore(2)
    SaveValue TableName, "HighScore3Name", HighScoreName(2)
    SaveValue TableName, "HighScore4", HighScore(3)
    SaveValue TableName, "HighScore4Name", HighScoreName(3)
    SaveValue TableName, "Credits", Credits
    SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

Sub Reseths
    HighScoreName(0) = "AAA"
    HighScoreName(1) = "BBB"
    HighScoreName(2) = "CCC"
    HighScoreName(3) = "DDD"
    HighScore(0) = 100000
    HighScore(1) = 100000
    HighScore(2) = 100000
    HighScore(3) = 100000
    Savehs
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(3)
Dim hsCurrentDigit
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash

Sub CheckHighscore()
    Dim tmp
    tmp = Score(CurrentPlayer)

    If tmp > HighScore(0)Then 'add 1 credit for beating the highscore
        Credits = Credits + 1
        DOF 125, DOFOn
    End If

    If tmp > HighScore(3)Then
        PlaySound SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
        DOF 121, DOFPulse
        HighScore(3) = tmp
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    hsbModeActive = True
    'PlaySound "vo_greatscore" &RndNbr(6)
    hsLetterFlash = 0

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' < is back arrow
    hsCurrentLetter = 1
    DMDFlush()
    HighScoreDisplayNameNow()

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0)then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter > len(hsValidLetters))then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayNameNow()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<")then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3)then
                HighScoreCommitName()
            else
                HighScoreDisplayNameNow()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit > 0)then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayNameNow()
        end if
    end if
End Sub

Sub HighScoreDisplayNameNow()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreDisplayName()
    Dim i
    Dim TempTopStr
    Dim TempBotStr

    TempTopStr = "YOUR NAME:"
    dLine(0) = ExpandLine(TempTopStr)
    DMDUpdate 0

    TempBotStr = "    > "
    if(hsCurrentDigit > 0)then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit > 1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit > 2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3)then
        if(hsLetterFlash <> 0)then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit < 1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit < 2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & " <    "
    dLine(1) = ExpandLine(TempBotStr)
    DMDUpdate 1
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2)then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False

    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ")then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
    EndOfBallComplete()
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 3
        For j = 0 to 2
            If HighScore(j) < HighScore(j + 1)Then
                tmp = HighScore(j + 1)
                tmp2 = HighScoreName(j + 1)
                HighScore(j + 1) = HighScore(j)
                HighScoreName(j + 1) = HighScoreName(j)
                HighScore(j) = tmp
                HighScoreName(j) = tmp2
            End If
        Next
    Next
End Sub

' *************************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' only 5 effects: none, scroll left, scroll right, blink and blinkfast
' 3 Lines, treats all 3 lines as text.
' 1st and 2nd lines are 20 characters long
' 3rd line is just 1 character
' Example format:
' DMD "text1","text2","backpicture", eNone, eNone, eNone, 250, True, "sound"
' Short names:
' dq = display queue
' de = display effect
' *************************************************************************

Const eNone = 0        ' Instantly displayed
Const eScrollLeft = 1  ' scroll on from the right
Const eScrollRight = 2 ' scroll on from the left
Const eBlink = 3       ' Blink (blinks for 'TimeOn')
Const eBlinkFast = 4   ' Blink (blinks for 'TimeOn') at user specified intervals (fast speed)

Const dqSize = 64

Dim dqHead
Dim dqTail
Dim deSpeed
Dim deBlinkSlowRate
Dim deBlinkFastRate

Dim dCharsPerLine(2)
Dim dLine(2)
Dim deCount(2)
Dim deCountEnd(2)
Dim deBlinkCycle(2)

Dim dqText(2, 64)
Dim dqEffect(2, 64)
Dim dqTimeOn(64)
Dim dqbFlush(64)
Dim dqSound(64)

Dim FlexDMD
Dim DMDScene

Sub DMD_Init() 'default/startup values
    If UseFlexDMD Then
        Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
        If Not FlexDMD is Nothing Then
            If FlexDMDHighQuality Then
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 256
                FlexDMD.Height = 64
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + i * 12, 6, 12, 22
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + (i - 20) * 12, 34, 12, 22
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            Else
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 128
                FlexDMD.Height = 32
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + i * 6, 3, 6, 11
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + (i - 20) * 6, 17, 6, 11
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            End If
        End If
    End If

    Dim i, j
    DMDFlush()
    deSpeed = 20
    deBlinkSlowRate = 10
    deBlinkFastRate = 5
    dCharsPerLine(0) = 20 'characters lower line
    dCharsPerLine(1) = 20 'characters top line
    dCharsPerLine(2) = 1  'characters back line
    For i = 0 to 2
        dLine(i) = Space(dCharsPerLine(i))
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
        dqTimeOn(i) = 0
        dqbFlush(i) = True
        dqSound(i) = ""
    Next
    dLine(2) = " "
    For i = 0 to 2
        For j = 0 to 64
            dqText(i, j) = ""
            dqEffect(i, j) = eNone
        Next
    Next
    DMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDFlush()
    Dim i
    DMDTimer.Enabled = False
    DMDEffectTimer.Enabled = False
    dqHead = 0
    dqTail = 0
    For i = 0 to 2
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
    Next
End Sub

Sub DMDScore()
    Dim tmp, tmp1, tmp2
    if(dqHead = dqTail)Then
        tmp = RL(FormatScore(Score(Currentplayer)))
        'tmp = CL(FormatScore(Score(Currentplayer) ) )
        'tmp1 = CL("PLAYER " & CurrentPlayer & " BALL " & Balls)
        'tmp1 = FormatScore(Bonuspoints(Currentplayer) ) & " X" &BonusMultiplier(Currentplayer)

        Select Case Battle(CurrentPlayer, 0)
            Case 0:tmp1 = CL("PLAYER " & CurrentPlayer & " BALL " & Balls & " X" & PlayfieldMultiplier(CurrentPlayer))
            Case 1:tmp1 = CL("SPINNERS LEFT " & 100-SpinCount)
            Case 2:tmp1 = CL("BUMPER HITS LEFT " & 25-SuperBumperHIts)
            Case 3:tmp1 = CL("RAMP HITS LEFT " & 6-ramphits3)
            Case 4:tmp1 = CL("ORBIT HITS LEFT " & 6-orbithits)
            Case 5:tmp1 = CL("HIT THE LIGHTS")
            Case 6:tmp1 = CL("HIT THE LIGHTS")
            Case 7:tmp1 = CL("HIT THE TARGETS " & 20-TargetHits7)
            Case 8:tmp1 = CL("HIT THE TARGETS " & 6-TargetHits8)
            Case 9:tmp1 = CL("HIT THE LIT LIGHT " & 8-LightHits9)
            Case 10:tmp1 = CL("HIT THE LOOPS " & 6-loopCount)
            Case 11:tmp1 = CL("HIT THE LIT LIGHT " & 8-LightHits11)
            Case 12:tmp1 = CL("HIT RAMPS ORBITS " & 6-RampHits12)
            Case 13:tmp1 = CL("MORDEKAI BATTLE")
        End Select
        tmp2 = ""
    End If
    DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    if(dqTail < dqSize)Then
        if(Text0 = "_")Then
            dqEffect(0, dqTail) = eNone
            dqText(0, dqTail) = "_"
        Else
            dqEffect(0, dqTail) = Effect0
            dqText(0, dqTail) = ExpandLine(Text0)
        End If

        if(Text1 = "_")Then
            dqEffect(1, dqTail) = eNone
            dqText(1, dqTail) = "_"
        Else
            dqEffect(1, dqTail) = Effect1
            dqText(1, dqTail) = ExpandLine(Text1)
        End If

        if(Text2 = "_")Then
            dqEffect(2, dqTail) = eNone
            dqText(2, dqTail) = "_"
        Else
            dqEffect(2, dqTail) = Effect2
            dqText(2, dqTail) = Text2 'it is always 1 letter in this table
        End If

        dqTimeOn(dqTail) = TimeOn
        dqbFlush(dqTail) = bFlush
        dqSound(dqTail) = Sound
        dqTail = dqTail + 1
        if(dqTail = 1)Then
            DMDHead()
        End If
    End If
End Sub

Sub DMDHead()
    Dim i
    deCount(0) = 0
    deCount(1) = 0
    deCount(2) = 0
    DMDEffectTimer.Interval = deSpeed

    For i = 0 to 2
        Select Case dqEffect(i, dqHead)
            Case eNone:deCountEnd(i) = 1
            Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead))
            Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead))
            Case eBlink:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
            Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
        End Select
    Next
    if(dqSound(dqHead) <> "")Then
        PlaySound(dqSound(dqHead))
    End If
    DMDEffectTimer.Enabled = True
End Sub

Sub DMDEffectTimer_Timer()
    DMDEffectTimer.Enabled = False
    DMDProcessEffectOn()
End Sub

Sub DMDTimer_Timer()
    Dim Head
    DMDTimer.Enabled = False
    Head = dqHead
    dqHead = dqHead + 1
    if(dqHead = dqTail)Then
        if(dqbFlush(Head) = True)Then
            DMDScoreNow()
        Else
            dqHead = 0
            DMDHead()
        End If
    Else
        DMDHead()
    End If
End Sub

Sub DMDProcessEffectOn()
    Dim i
    Dim BlinkEffect
    Dim Temp

    BlinkEffect = False

    For i = 0 to 2
        if(deCount(i) <> deCountEnd(i))Then
            deCount(i) = deCount(i) + 1

            select case(dqEffect(i, dqHead))
                case eNone:
                    Temp = dqText(i, dqHead)
                case eScrollLeft:
                    Temp = Right(dLine(i), dCharsPerLine(i)- 1)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead), (dCharsPerLine(i) + 1)- deCount(i), 1)
                    Temp = Temp & Left(dLine(i), dCharsPerLine(i)- 1)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkSlowRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkFastRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
            End Select

            if(dqText(i, dqHead) <> "_")Then
                dLine(i) = Temp
                DMDUpdate i
            End If
        End If
    Next

    if(deCount(0) = deCountEnd(0))and(deCount(1) = deCountEnd(1))and(deCount(2) = deCountEnd(2))Then

        if(dqTimeOn(dqHead) = 0)Then
            DMDFlush()
        Else
            if(BlinkEffect = True)Then
                DMDTimer.Interval = 10
            Else
                DMDTimer.Interval = dqTimeOn(dqHead)
            End If

            DMDTimer.Enabled = True
        End If
    Else
        DMDEffectTimer.Enabled = True
    End If
End Sub

Function ExpandLine(TempStr) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(20)
    Else
        if Len(TempStr) > Space(20)Then
            TempStr = Left(TempStr, Space(20))
        Else
            if(Len(TempStr) < 20)Then
                TempStr = TempStr & Space(20 - Len(TempStr))
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString
    NumString = CStr(abs(Num))
    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1)) + 128) & right(NumString, Len(NumString)- i)
        end if
    Next
    FormatScore = NumString
End function

Function FL(NumString1, NumString2) 'Fill line
    Dim Temp, TempStr
    If Len(NumString1) + Len(NumString2) < 20 Then
        Temp = 20 - Len(NumString1)- Len(NumString2)
        TempStr = NumString1 & Space(Temp) & NumString2
        FL = TempStr
    End If
End Function

Function CL(NumString) 'center line
    Dim Temp, TempStr
    If Len(NumString) > 20 Then NumString = Left(NumString, 20)
    Temp = (20 - Len(NumString)) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(NumString) 'right line
    Dim Temp, TempStr
    If Len(NumString) > 20 Then NumString = Left(NumString, 20)
    Temp = 20 - Len(NumString)
    TempStr = Space(Temp) & NumString
    RL = TempStr
End Function

'**************
' Update DMD
'**************

Sub DMDUpdate(id)
    Dim digit, value
    If UseFlexDMD Then FlexDMD.LockRenderThread
    Select Case id
        Case 0 'top text line
            For digit = 0 to 19
                DMDDisplayChar mid(dLine(0), digit + 1, 1), digit
            Next
        Case 1 'bottom text line
            For digit = 20 to 39
                DMDDisplayChar mid(dLine(1), digit -19, 1), digit
            Next
        Case 2 ' back image - back animations
            If dLine(2) = "" OR dLine(2) = " " Then dLine(2) = "bkempty"
            Digits(40).ImageA = dLine(2)
            If UseFlexDMD Then DMDScene.GetImage("Back").Bitmap = FlexDMD.NewImage("", "VPX." & dLine(2) & "&dmd=2").Bitmap
    End Select
    If UseFlexDMD Then FlexDMD.UnlockRenderThread
End Sub

Sub DMDDisplayChar(achar, adigit)
    If achar = "" Then achar = " "
    achar = ASC(achar)
    Digits(adigit).ImageA = Chars(achar)
    If UseFlexDMD Then DMDScene.GetImage("Dig" & adigit).Bitmap = FlexDMD.NewImage("", "VPX." & Chars(achar) & "&dmd=2&add").Bitmap
End Sub

'****************************
' JP's new DMD using flashers
'****************************

Dim Digits, DigitsBack, Chars(255), Images(255)

DMDInit

Sub DMDInit
    Dim i
    Digits = Array(digit001, digit002, digit003, digit004, digit005, digit006, digit007, digit008, digit009, digit010, _
        digit011, digit012, digit013, digit014, digit015, digit016, digit017, digit018, digit019, digit020,            _
        digit021, digit022, digit023, digit024, digit025, digit026, digit027, digit028, digit029, digit030,            _
        digit031, digit032, digit033, digit034, digit035, digit036, digit037, digit038, digit039, digit040,            _
        digit041)
    For i = 0 to 255:Chars(i) = "d_empty":Next

    Chars(32) = "d_empty"
    Chars(33) = ""        '!
    Chars(34) = ""        '"
    Chars(35) = ""        '#
    Chars(36) = ""        '$
    Chars(37) = ""        '%
    Chars(38) = ""        '&
    Chars(39) = ""        ''
    Chars(40) = ""        '(
    Chars(41) = ""        ')
    Chars(42) = ""        '*
    Chars(43) = ""        '+
    Chars(44) = ""        '
    Chars(45) = "d_minus" '-
    Chars(46) = "d_dot"   '.
    Chars(47) = ""        '/
    Chars(48) = "d_0"     '0
    Chars(49) = "d_1"     '1
    Chars(50) = "d_2"     '2
    Chars(51) = "d_3"     '3
    Chars(52) = "d_4"     '4
    Chars(53) = "d_5"     '5
    Chars(54) = "d_6"     '6
    Chars(55) = "d_7"     '7
    Chars(56) = "d_8"     '8
    Chars(57) = "d_9"     '9
    Chars(60) = "d_less"  '<
    Chars(61) = ""        '=
    Chars(62) = "d_more"  '>
    Chars(64) = ""        '@
    Chars(65) = "d_a"     'A
    Chars(66) = "d_b"     'B
    Chars(67) = "d_c"     'C
    Chars(68) = "d_d"     'D
    Chars(69) = "d_e"     'E
    Chars(70) = "d_f"     'F
    Chars(71) = "d_g"     'G
    Chars(72) = "d_h"     'H
    Chars(73) = "d_i"     'I
    Chars(74) = "d_j"     'J
    Chars(75) = "d_k"     'K
    Chars(76) = "d_l"     'L
    Chars(77) = "d_m"     'M
    Chars(78) = "d_n"     'N
    Chars(79) = "d_o"     'O
    Chars(80) = "d_p"     'P
    Chars(81) = "d_q"     'Q
    Chars(82) = "d_r"     'R
    Chars(83) = "d_s"     'S
    Chars(84) = "d_t"     'T
    Chars(85) = "d_u"     'U
    Chars(86) = "d_v"     'V
    Chars(87) = "d_w"     'W
    Chars(88) = "d_x"     'X
    Chars(89) = "d_y"     'Y
    Chars(90) = "d_z"     'Z
    Chars(94) = "d_up"    '^
    '    Chars(95) = '_
    Chars(96) = ""
    Chars(97) = ""  'a
    Chars(98) = ""  'b
    Chars(99) = ""  'c
    Chars(100) = "" 'd
    Chars(101) = "" 'e
    Chars(102) = "" 'f
    Chars(103) = "" 'g
    Chars(104) = "" 'h
    Chars(105) = "" 'i
    Chars(106) = "" 'j
    Chars(107) = "" 'k
    Chars(108) = "" 'l
    Chars(109) = "" 'm
    Chars(110) = "" 'n
    Chars(111) = "" 'o
    Chars(112) = "" 'p
    Chars(113) = "" 'q
    Chars(114) = "" 'r
    Chars(115) = "" 's
    Chars(116) = "" 't
    Chars(117) = "" 'u
    Chars(118) = "" 'v
    Chars(119) = "" 'w
    Chars(120) = "" 'x
    Chars(121) = "" 'y
    Chars(122) = "" 'z
    Chars(123) = "" '{
    Chars(124) = "" '|
    Chars(125) = "" '}
    Chars(126) = "" '~
    'used in the FormatScore function
    Chars(176) = "d_0a" '0.
    Chars(177) = "d_1a" '1.
    Chars(178) = "d_2a" '2.
    Chars(179) = "d_3a" '3.
    Chars(180) = "d_4a" '4.
    Chars(181) = "d_5a" '5.
    Chars(182) = "d_6a" '6.
    Chars(183) = "d_7a" '7.
    Chars(184) = "d_8a" '8.
    Chars(185) = "d_9a" '9.
End Sub

'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub Realtime_Timer
p49on.blenddisablelighting = Light49.getinplayintensity * 40
p49off.blenddisablelighting = Light49.getinplayintensity + 1
p50on.blenddisablelighting = Light50.getinplayintensity * 40
p50off.blenddisablelighting = Light50.getinplayintensity + 1
p19on.blenddisablelighting = Light19.getinplayintensity * 40
p19off.blenddisablelighting = Light19.getinplayintensity + 1
p20on.blenddisablelighting = Light20.getinplayintensity * 40
p20off.blenddisablelighting = Light20.getinplayintensity + 1
p44on.blenddisablelighting = Light44.getinplayintensity * 40
p44off.blenddisablelighting = Light44.getinplayintensity + 1
p43on.blenddisablelighting = Light43.getinplayintensity * 40
p43off.blenddisablelighting = Light43.getinplayintensity + 1
p40on.blenddisablelighting = Light40.getinplayintensity * 40
p40off.blenddisablelighting = Light40.getinplayintensity + 1
p39on.blenddisablelighting = Light39.getinplayintensity * 40
p39off.blenddisablelighting = Light39.getinplayintensity + 1
p38on.blenddisablelighting = Light38.getinplayintensity * 40
p38off.blenddisablelighting = Light38.getinplayintensity + 1
p41on.blenddisablelighting = Light41.getinplayintensity * 40
p41off.blenddisablelighting = Light41.getinplayintensity + 1
p42on.blenddisablelighting = Light42.getinplayintensity * 40
p42off.blenddisablelighting = Light42.getinplayintensity + 1
p45on.blenddisablelighting = Light45.getinplayintensity * 40
p45off.blenddisablelighting = Light45.getinplayintensity + 1
p46on.blenddisablelighting = Light46.getinplayintensity * 40
p46off.blenddisablelighting = Light46.getinplayintensity + 1
p47on.blenddisablelighting = Light47.getinplayintensity * 40
p47off.blenddisablelighting = Light47.getinplayintensity + 1
p51on.blenddisablelighting = Light51.getinplayintensity * 40
p51off.blenddisablelighting = Light51.getinplayintensity + 1
p48on.blenddisablelighting = Light48.getinplayintensity * 40
p48off.blenddisablelighting = Light48.getinplayintensity + 1
p52on.blenddisablelighting = Light52.getinplayintensity * 40
p52off.blenddisablelighting = Light52.getinplayintensity + 1
p53on.blenddisablelighting = Light53.getinplayintensity * 40
p53off.blenddisablelighting = Light53.getinplayintensity + 1
p54on.blenddisablelighting = Light54.getinplayintensity * 40
p54off.blenddisablelighting = Light54.getinplayintensity + 1
p56on.blenddisablelighting = Light56.getinplayintensity * 40
p56off.blenddisablelighting = Light56.getinplayintensity + 1
p57on.blenddisablelighting = Light57.getinplayintensity * 40
p57off.blenddisablelighting = Light57.getinplayintensity + 1
p58on.blenddisablelighting = Light58.getinplayintensity * 40
p58off.blenddisablelighting = Light58.getinplayintensity + 1
p59on.blenddisablelighting = Light59.getinplayintensity * 40
p59off.blenddisablelighting = Light59.getinplayintensity + 1
p13on.blenddisablelighting = Light13.getinplayintensity * 40
p13off.blenddisablelighting = Light13.getinplayintensity + 1
p14on.blenddisablelighting = Light14.getinplayintensity * 40
p14off.blenddisablelighting = Light14.getinplayintensity + 1
p15on.blenddisablelighting = Light15.getinplayintensity * 40
p15off.blenddisablelighting = Light15.getinplayintensity + 1
p16on.blenddisablelighting = Light16.getinplayintensity * 40
p16off.blenddisablelighting = Light16.getinplayintensity + 1
p17on.blenddisablelighting = Light17.getinplayintensity * 40
p17off.blenddisablelighting = Light17.getinplayintensity + 1
p18on.blenddisablelighting = Light18.getinplayintensity * 40
p18off.blenddisablelighting = Light18.getinplayintensity + 1
p28on.blenddisablelighting = Light28.getinplayintensity * 40
p28off.blenddisablelighting = Light28.getinplayintensity + 1
p27on.blenddisablelighting = Light27.getinplayintensity * 40
p27off.blenddisablelighting = Light27.getinplayintensity + 1
p26on.blenddisablelighting = Light26.getinplayintensity * 40
p26off.blenddisablelighting = Light26.getinplayintensity + 1
p25on.blenddisablelighting = Light25.getinplayintensity * 40
p25off.blenddisablelighting = Light25.getinplayintensity + 1
p24on.blenddisablelighting = Light24.getinplayintensity * 40
p24off.blenddisablelighting = Light24.getinplayintensity + 1
p23on.blenddisablelighting = Light23.getinplayintensity * 40
p23off.blenddisablelighting = Light23.getinplayintensity + 1


    RollingUpdate
    ' add any other real time update subs, like gates or diverters
    Doorp.Roty = Doorf.CurrentAngle -90
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
End Sub

'********************************************************************************************
' Only for VPX 10.2 and higher.
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version

    If TypeName(MyLight) = "Light" Then

        If FinalState = 2 Then
            FinalState = MyLight.State 'Keep the current light state
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then

        Dim steps

        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
        If FinalState = 2 Then                      'Keep the current flasher state
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

        ' Start blink timer and create timer subroutine
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' 12 colors: red, orange, amber, yellow...
'******************************************

'colors
Const red = 1
Const orange = 2
Const amber = 3
Const yellow = 4
Const darkgreen = 5
Const green = 6
Const blue = 7
Const darkblue = 8
Const purple = 9
Const white = 10
Const teal = 11
Const ledwhite = 12

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Select Case col
        Case red
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case orange
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case amber
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 8, 0)
            n.colorfull = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 16, 0)
            n.colorfull = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(64, 0, 96)
            n.colorfull = RGB(128, 0, 192)
        Case white 'bulb
            n.color = RGB(193, 91, 0)
            n.colorfull = RGB(255, 197, 143)
        Case teal
            n.color = RGB(1, 64, 62)
            n.colorfull = RGB(2, 128, 126)
        Case ledwhite
            n.color = RGB(255, 197, 143)
            n.colorfull = RGB(255, 252, 224)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub SetFlashColor(n, col, stat) 'stat 0 = off, 1 = on, -1= no change - no blink for the flashers, use FlashForMs
    Select Case col
        Case red
            n.color = RGB(255, 0, 0)
        Case orange
            n.color = RGB(255, 64, 0)
        Case amber
            n.color = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 192)
        Case white 'bulb
            n.color = RGB(255, 197, 143)
        Case teal
            n.color = RGB(2, 128, 126)
         Case ledwhite
            n.color = RGB(255, 252, 224)
    End Select
    If stat <> -1 Then
        n.Visible = stat
    End If
End Sub

'*************************
' Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

Sub StartRainbow(n)
    DOF 412, DOFOn
    set RainbowLights = n
    RGBStep = 0
    RGBFactor = 5
    rRed = 255
    rGreen = 0
    rBlue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow()
    DOF 412, DOFOff
    Dim obj
    RainbowTimer.Enabled = 0
    RainbowTimer.Enabled = 0
End Sub

Sub RainbowTimer_Timer 'rainbow led light color changing
    Dim obj
    Select Case RGBStep
        Case 0 'Green
            rGreen = rGreen + RGBFactor
            If rGreen > 255 then
                rGreen = 255
                RGBStep = 1
            End If
        Case 1 'Red
            rRed = rRed - RGBFactor
            If rRed < 0 then
                rRed = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            rBlue = rBlue + RGBFactor
            If rBlue > 255 then
                rBlue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            rGreen = rGreen - RGBFactor
            If rGreen < 0 then
                rGreen = 0
                RGBStep = 4
            End If
        Case 4 'Red
            rRed = rRed + RGBFactor
            If rRed > 255 then
                rRed = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            rBlue = rBlue - RGBFactor
            If rBlue < 0 then
                rBlue = 0
                RGBStep = 0
            End If
    End Select
    For each obj in RainbowLights
        obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
        obj.colorfull = RGB(rRed, rGreen, rBlue)
    Next
End Sub

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    Dim ii
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1)Then
        DMD CL("LAST SCORE"), CL("PLAYER 1 " &FormatScore(Score(1))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(2)Then
        DMD CL("LAST SCORE"), CL("PLAYER 2 " &FormatScore(Score(2))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(3)Then
        DMD CL("LAST SCORE"), CL("PLAYER 3 " &FormatScore(Score(3))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(4)Then
        DMD CL("LAST SCORE"), CL("PLAYER 4 " &FormatScore(Score(4))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    DMD "", CL("GAME OVER"), "", eNone, eBlink, eNone, 2000, False, ""
    If bFreePlay Then
        DMD "", CL("FREE PLAY"), "", eNone, eBlink, eNone, 2000, False, ""
    Else
        If Credits > 0 Then
            DMD CL("CREDITS " & Credits), CL("PRESS START"), "", eNone, eBlink, eNone, 2000, False, ""
        Else
            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 2000, False, ""
        End If
    End If
    DMD "", "", "d_jppresents", eNone, eNone, eNone, 3000, False, ""
    DMD "", "", "KillerInstinct", eNone, eNone, eNone, 4000, False, ""
    DMD "", CL("ROM VERSION " &myversion), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("HIGHSCORES"), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CL("HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CL("HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(dCharsPerLine(0)), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 500, False, ""
End Sub

Sub StartAttractMode
    ChangeSong
    StartLightSeq
	DOF 323, DOFOn
    DMDFlush
    ShowTableInfo
End Sub

Sub StopAttractMode
    DMDScoreNow
	DOF 323, DOFOff
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    LightSeqAttract.UpdateInterval = 25
    LightSeqAttract.Play SeqBlinking, , 5, 150
    LightSeqAttract.Play SeqRandom, 40, , 4000
    LightSeqAttract.Play SeqAllOff
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 40, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 40, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqRightOn, 30, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqLeftOn, 30, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqStripe1VertOn, 50, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
End Sub

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels 
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
        Case 11:table1.ColorGradeImage = "LUT Warm 0"
        Case 12:table1.ColorGradeImage = "LUT Warm 1"
        Case 13:table1.ColorGradeImage = "LUT Warm 2"
        Case 14:table1.ColorGradeImage = "LUT Warm 3"
        Case 15:table1.ColorGradeImage = "LUT Warm 4"
        Case 16:table1.ColorGradeImage = "LUT Warm 5"
        Case 17:table1.ColorGradeImage = "LUT Warm 6"
        Case 18:table1.ColorGradeImage = "LUT Warm 7"
        Case 19:table1.ColorGradeImage = "LUT Warm 8"
        Case 20:table1.ColorGradeImage = "LUT Warm 9"
        Case 21:table1.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub

Dim GiIntensity
GiIntensity = 1   'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

' New LUT postit
Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea="PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
SetLUTLine ""
LUBack.imagea="PostitBL"
End Sub

'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' droptargets, animations, etc
Sub VPObjects_Init
End Sub

' tables variables and Mode init
Dim LaneBonus
Dim TargetBonus
Dim RampBonus
Dim ComboCount
Dim ComboHits(4)
Dim ComboValue(4)
Dim BumperValue(4)
Dim BumperHits
Dim SuperBumperHits
Dim SpinnerValue(4)
Dim MonstersKilled(4)
Dim SpinCount
Dim RampHits3
Dim RampHits12
Dim OrbitHits
Dim TargetHits7
Dim TargetHits8
Dim CaptiveBallHits
Dim LightHits9
Dim LightHits11
Dim loopCount
Dim BattlesWon(4)
Dim Battle(4, 15) '12 battles, 1 final battle
Dim NewBattle
Dim PowerupHits

Sub Game_Init() 'called at the start of a new game
    Dim i, j
    bExtraBallWonThisBall = False
    'Play some Music
    ChangeSong
    'Init Variables
    LaneBonus = 0 'it gets deleted when a new ball is launched
    TargetBonus = 0
    RampBonus = 0
    BumperHits = 0
    For i = 1 to 4
        SkillshotValue(i) = 500000
        Jackpot(i) = 100000
        ComboValue(i) = 500000
        MonstersKilled(i) = 0
        BallsInLock(i) = 0
        SpinnerValue(i) = 1000
        ComboHits(i) = 0
        BumperValue(i) = 210 'start at 210 and every 30 hits its value is increased by 500 points
    Next
    ResetBattles
    SpinCount = 0
    SuperBumperHits = 0
    RampHits3 = 0
    RampHits12 = 0
    OrbitHits = 0
    TargetHits7 = 0
    TargetHits8 = 0
    CaptiveBallHits = 0
    loopCount = 0
    PowerupHits = 0
    LightHits9 = 0
    LightHits11 = 0
    ComboCount = 0
    'Init Delays/Timers
    'MainMode Init()
    'Init lights
    TurnOffPlayfieldLights()
    CloseDoor
End Sub

Sub StopEndOfBallMode() 'this sub is called after the last ball is drained
    ResetSkillShotTimer_Timer
    StopBattle
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
    Dim i
    LaneBonus = 0
    TargetBonus = 0
    RampBonus = 0
    BumperHits = 0
    ' select a battle
    SelectBattle
End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
' UpdatePFXLights(PlayfieldMultiplier(CurrentPlayer)) 'ensure the multiplier is displayed right
Light13.State = 0
Light16.State = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    LightSeqSkillshot.Play SeqAllOff
    Light48.State = 2
    Light18.State = 2
    Gate2.Open = 1
    Gate3.Open = 1
    DMD CL("HIT LIT LIGHT"), CL("FOR SKILLSHOT"), "", eNone, eNone, eNone, 1500, True, ""
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    LightSeqSkillshot.StopPlay
    If Light18.State = 2 Then Light18.State = 0
    Light48.State = 0
    Gate2.Open = 0
    Gate3.Open = 0
    DMDScoreNow
End Sub

' *********************************************************************
'                        Table Object Hit Events
'
' Any target hit Sub will follow this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/Mode this trigger is a member of
' - set the "LastSwitchHit" variable in case it is needed later
' *********************************************************************

'*********************************************************
' Slingshots has been hit
' In this table the slingshots change the outlanes lights

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("Punch_Sling", 103, DOFPulse, DOFcontactors), Lemk
	RandomSoundSlingShotLeft Lemk
    DOF 105, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 210
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
    ChangeOutlanes
    ShakeLeftCat
    ShakeBatons
    FlashOrchWeapon 800, 100
    

	
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("rotate", 104, DOFPulse, DOFcontactors), Remk
	RandomSoundSlingShotRight Remk
    DOF 106, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 210
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
    ChangeOutlanes
    ShakeRightCat
    ShakeRight2Cat
	
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


Sub ChangeOutlanes
    Dim tmp
    tmp = light13.State
    light13.State = light16.State
    light16.State = tmp
End Sub

'*********
' Bumpers
'*********
' after each 30 hits the bumpers increase their score value by 500 points up to 3210
' and they increase the playfield multiplier.

Sub Bumper1_Hit
    If NOT Tilted Then
        PlaySoundAt SoundFXDOF("Bumpers_Top_1", 109, DOFPulse, DOFContactors), Bumper1
        RIPTORShake
        RIPTORARMShake
        If FlippersBlood Then Right2Splat
        If NOT bMultiBallMode Then PlaySound"TRip_"&RndNbr(6)
        FlashForMs bumpersmalllight1, 1000, 50, 0
        DOF 138, DOFPulse
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper1"
    End If
    CheckBumpers
End Sub

Sub Bumper2_Hit
    If NOT Tilted Then
        PlaySoundAt SoundFXDOF("Bumpers_Top_2", 110, DOFPulse, DOFContactors), Bumper2
        FlashForMs bumpersmalllight2, 1000, 50, 0
        DOF 140, DOFPulse
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper2"
    End If
    CheckBumpers
End Sub

Sub Bumper3_Hit
    If NOT Tilted Then
        PlaySoundAt SoundFXDOF("Bumpers_Top_3", 107, DOFPulse, DOFContactors), Bumper3
        FlashForMs bumpersmalllight3, 1000, 50, 0
        DOF 137, DOFPulse
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper3"
    End If
    CheckBumpers
End Sub

' Check the bumper hits

Sub CheckBumpers()
    ' increase the bumper hit count and increase the bumper value after each 30 hits
    BumperHits = BumperHits + 1
    If BumperHits MOD 30 = 0 Then
        If BumperValue(CurrentPlayer) < 3210 Then
            BumperValue(CurrentPlayer) = BumperValue(CurrentPlayer) + 500
        End If
        ' lit the playfield multiplier light
        light54.State = 1
    End If
End Sub

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
	DNA30 = 0 : DNA45 = (NightDay-10)/20 : DNA90 = 0 : DayNightAdjust = 0.4
Else
	DNA30 = (NightDay-10)/30 : DNA45 = (NightDay-10)/45 : DNA90 = (NightDay-10)/90 : DayNightAdjust = NightDay/25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt : For cnt = 1 to 6 : FlBumperActive(cnt) = False : Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight

FlInitBumper 1, "green"
FlInitBumper 2, "blue"
FlInitBumper 3, "red"

' ### uncomment the statement below to change the color for all bumpers ###
' Dim ind : For ind = 1 to 5 : FlInitBumper ind, "green" : next

Sub FlInitBumper(nr, col)
	FlBumperActive(nr) = True
	' store all objects in an array for use in FlFadeBumper subroutine
	FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
	Set FlBumperTop(nr) = Eval("bumpertop" & nr) : FlBumperTop(nr).material = "bumpertopmat" & nr
	Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr) : Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
	Set FlBumperDisk(nr) = Eval("bumperdisk" & nr) : Set FlBumperBase(nr) = Eval("bumperbase" & nr)
	Set FlBumperBulb(nr) = Eval("bumperbulb" & nr) : FlBumperBulb(nr).material = "bumperbulbmat" & nr
	Set FlBumperscrews(nr) = Eval("bumperscrews" & nr): FlBumperscrews(nr).material = "bumperscrew" & col
	Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
	' set the color for the two VPX lights
	select case col
		Case "red"
			FlBumperSmallLight(nr).color = RGB(255,4,0) : FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
			FlBumperBigLight(nr).color = RGB(255,32,0) : FlBumperBigLight(nr).colorfull = RGB(255,32,0)
			FlBumperHighlight(nr).color = RGB(64,255,0)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
			FlBumperSmallLight(nr).TransmissionScale = 0
		Case "blue"
			FlBumperBigLight(nr).color = RGB(32,80,255) : FlBumperBigLight(nr).colorfull = RGB(32,80,255)
			FlBumperSmallLight(nr).color = RGB(0,80,255) : FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
			FlBumperSmallLight(nr).TransmissionScale = 0 : MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
			FlBumperHighlight(nr).color = RGB(255,16,8)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "green"
			FlBumperSmallLight(nr).color = RGB(8,255,8) : FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
			FlBumperBigLight(nr).color = RGB(32,255,32) : FlBumperBigLight(nr).colorfull = RGB(32,255,32)
			FlBumperHighlight(nr).color = RGB(255,32,255) : MaterialColor "bumpertopmat" & nr, RGB(16,255,16) 
			FlBumperSmallLight(nr).TransmissionScale = 0.005
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "orange"
			FlBumperHighlight(nr).color = RGB(205,170,109)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1 
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).color = RGB(205,170,109) : FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
			FlBumperBigLight(nr).color = RGB(194,178,128) : FlBumperBigLight(nr).colorfull = RGB(205,170,109)
		Case "white"
			FlBumperBigLight(nr).color = RGB(255,230,190) : FlBumperBigLight(nr).colorfull = RGB(255,230,190)
			FlBumperHighlight(nr).color = RGB(255,180,100) : 
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
		Case "blacklight"
			FlBumperBigLight(nr).color = RGB(32,32,255) : FlBumperBigLight(nr).colorfull = RGB(32,32,255)
			FlBumperHighlight(nr).color = RGB(48,8,255) : 
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "yellow"
			FlBumperSmallLight(nr).color = RGB(255,230,4) : FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
			FlBumperBigLight(nr).color = RGB(255,240,50) : FlBumperBigLight(nr).colorfull = RGB(255,240,50)
			FlBumperHighlight(nr).color = RGB(255,255,220)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1 
			FlBumperSmallLight(nr).TransmissionScale = 0
		Case "purple"
			FlBumperBigLight(nr).color = RGB(80,32,255) : FlBumperBigLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).color = RGB(80,32,255) : FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).TransmissionScale = 0 : 
			FlBumperHighlight(nr).color = RGB(32,64,255)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
	end select
End Sub

Sub FlFadeBumper(nr, Z)
	FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
'	UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
	FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 )* DayNightAdjust	

	select case FlBumperColor(nr)

		Case "blue" :
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(38-24*Z,130 - 98*Z,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20  + 500 * Z / (0.5 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z +0.97 * Z^3)
			Flbumperbiglight(nr).intensity = 45 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 10000 * (Z^3) / (0.5 + DNA90)

		Case "green"	
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(16 + 16 * sin(Z*3.14),255,16 + 16 * sin(Z*3.14)), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
			FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 6000 * (Z^3) / (1 + DNA90)
		
		Case "red" 
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 16 - 11*Z + 16 * sin(Z*3.14),0), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30^2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z*4,8-Z*8)     
		
		Case "orange"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(205, 170 - 22*z  + 16 * sin(Z*3.14),Z*32), RGB(205,205,205), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30^2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,136, 50)

		Case "white"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
			FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 14 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255,255 - 20*Z,255-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20*Z,255-65*Z)
			MaterialColor "bumpertopmat" & nr, RGB(255,235 - z*36,220 - Z*90)

		Case "blacklight"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 1, RGB(30-27*Z^0.03,30-28*Z^0.01, 255), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z^3
			Flbumperbiglight(nr).intensity = 40 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255-240*(Z^0.1),255 - 240*(Z^0.1),255) : FlBumperSmallLight(nr).colorfull = RGB(255-200*z,255 - 200*Z,255)
			MaterialColor "bumpertopmat" & nr, RGB(255-190*Z,235 - z*180,220 + 35*Z)

		Case "yellow"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 180 + 40*z, 48* Z), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30^2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

		Case "purple" :
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(128-118*Z - 32 * sin(Z*3.14), 32-26*Z ,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 15  + 200 * Z / (0.5 + DNA30) 
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z +0.97 * Z^3)
			Flbumperbiglight(nr).intensity = 50 * Z / (1 + DNA45) 
			FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (0.5 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(128-60*Z,32,255)


	end select
End Sub

Sub BumperTimer_Timer
	dim nr
	For nr = 1 to 6
		If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
			FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
			If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1 : End If
			FlFadeBumper nr, FlBumperFadeActual(nr)
		End If
		If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
			FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
			If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0 : End If
			FlFadeBumper nr, FlBumperFadeActual(nr)
		End If
	next
End Sub

'*************************
' Top & Inlanes: Bonus X
'*************************
' lit the 2 top lane lights and the 2 inlane lights to increase the bonus multiplier

Sub sw1_Hit
    DOF 128, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light17.State = 1
    FlashForMs f8, 1000, 50, 0
    If bSkillShotReady Then
        ResetSkillShotTimer_Timer
    Else
        CheckBonusX
    End If
End Sub

Sub sw6_Hit
    DOF 129, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light18.State = 1
    FlashForMs f8, 1000, 50, 0
    If bSkillShotReady Then
        Awardskillshot
    Else
        CheckBonusX
    End If
End Sub

Sub sw4_Hit
    DOF 133, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light14.State = 1
    FlashForMs f6, 1000, 50, 0
    AddScore 5000
    CheckBonusX
' Do some sound or light effect
End Sub

Sub sw3_Hit
    DOF 134, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    Light15.State = 1
    FlashForMs f7, 1000, 50, 0
    AddScore 5000
    CheckBonusX
' Do some sound or light effect
End Sub

Sub CheckBonusX
    If Light17.State + Light18.State + Light14.State + Light15.State = 4 Then
        AddBonusMultiplier 1
        GiEffect 1
        FlashForMs Light17, 1000, 50, 0
        FlashForMs Light18, 1000, 50, 0
        FlashForMs Light14, 1000, 50, 0
        FlashForMs Light15, 1000, 50, 0
    End IF
End Sub

'************************************
' Flipper OutLanes: Virtual kickback
'************************************
' if the light is lit then activate the ballsave

Sub sw2_Hit
    DOF 132, DOFPulse
    TimerJAG.Enabled = True
    Playsound "Jaguar"
    FlashForMs flash2, 500, 50, 0
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    AddScore 50000
    ' Do some sound or light effect
    ' do some check
    If light13.State = 1 Then
        EnableBallSaver 5
    End If
End Sub

Sub sw5_Hit
    DOF 135, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    AddScore 50000
    ' Do some sound or light effect
    ' do some check
    If Light16.State = 1 Then
        EnableBallSaver 5
    End If
End Sub

'************
'  Spinners
'************

Sub spinner1_Spin
    If Tilted Then Exit Sub
    Addscore spinnervalue(CurrentPlayer)
    PlaySoundAt "Swipe_Spinner", spinner1
	SoundSpinner spinner1
    DOF 136, DOFPulse
    Select Case Battle(CurrentPlayer, 0)
        Case 1
            Addscore 3000
            SpinCount = SpinCount + 1
            CheckWinBattle
    End Select
End Sub

Sub spinner2_Spin
    If Tilted Then Exit Sub
    PlaySoundAt "Swipe_Spinner", spinner2
	SoundSpinner spinner2
    DOF 124, DOFPulse
    Addscore spinnervalue(CurrentPlayer)
    Select Case Battle(CurrentPlayer, 0)
        Case 1
            Addscore 3000
            SpinCount = SpinCount + 1
            CheckWinBattle
    End Select
End Sub

'*******************
'    RAMP COMBOS
'*******************

Sub AwardCombo
    ComboCount = ComboCount + 1
    Select Case ComboCount
        Case 1:DMD CL("HYPER COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) ) ), "", eNone, eNone, eNone, 1500, True, "vo_combo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 2:DMD CL("BRUTAL COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 2) ), "", eNone, eNone, eNone, 1500, True, "vo_2xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 3:DMD CL("TRIPLE COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 3) ), "", eNone, eNone, eNone, 1500, True, "vo_3xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 4:DMD CL("AWESOME COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 4) ), "", eNone, eNone, eNone, 1500, True, "vo_4xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 5:DMD CL("KING COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 5) ), "", eNone, eNone, eNone, 1500, True, "vo_5xcombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
        Case 6:DMD CL("SUPER COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 7) ), "", eNone, eNone, eNone, 1500, True, "vo_supercombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1:DOF 126, DOFPulse
        Case Else:DMD CL("ULTRA COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 10) ), "", eNone, eNone, eNone, 1500, True, "vo_superdupercombo":ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1:DOF 126, DOFPulse
    End Select
    AddScore ComboValue(CurrentPlayer) * ComboCount
    ComboValue(CurrentPlayer) = ComboValue(CurrentPlayer) + 100000
End Sub

Sub aComboTargets_Hit(idx) 'reset the combo count if the ball hits another target/trigger
    ComboCount = 0
End Sub

'*********************************
'      The Lock Targets
'*********************************

Sub Target13_Hit
    Playsound "target_Glacius"
    PlaySoundAt SoundFXDOF("fx_target", 116, DOFPulse, DOFTargets), Target10
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light19.State = 1
    FlashForMs f4, 1000, 50, 0
    ' do some check
    Check2BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light31.State = 2 Then
                Light31.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light31.State = 2 Then
                Light33.State = 2
                Light31.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 8:TargetHits8 = TargetHits8 + 1:Addscore 25000:CheckWinBattle
        Case 9
            If Light31.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light31.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    LastSwitchHit = "Target13"
End Sub

Sub Target1_Hit
    FULGORE2SHAKE
    FULGORESHAKE
    FlashForMs f69, 1000, 50, 0
    If NOT bMultiBallMode Then PlaySound"TFulgore_"&RndNbr(5)
    PlaySoundAt SoundFXDOF("fx_target", 117, DOFPulse, DOFTargets), Target1
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light20.State = 1
    FlashForMs f5, 1000, 50, 0
    ' do some check
    Check2BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light29.State = 2 Then
                Light29.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light29.State = 2 Then
                Light29.State = 0
                Addscore 100000
                WinBattle
            End If
        Case 8:TargetHits8 = TargetHits8 + 1:Addscore 25000:CheckWinBattle
        Case 9
            If Light29.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light29.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    LastSwitchHit = "Target1"
End Sub

Sub Check2BankTargets
    If light19.state + light20.state = 2 Then
        light19.state = 0
        light20.state = 0
        LightEffect 1
        FlashEffect 1
        Addscore 20000
        If(Light46.State = 0)AND(bMultiballMode = FALSE)Then 'lit the lock light if it is off, open the sphynx door and activate the lock switch
            Light46.State = 1
            openDoor
            
            'PlaySound "vo_lockislit"
            DMD "_", CL("LOCK IS LIT"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
            DOF 417, DOFPulse
        ElseIf light53.State = 0 Then 'lit the increase jackpot light if the lock light is lit
            light53.State = 1
        'PlaySound "vo_IncreaseJakpot"
        Else
            Addscore 30000
        End If
    End If
End Sub

'**************************
' The Lock: Main Multiball
'**************************
' the lock is a virtual lock, where the locked balls are simply counted

Sub Door_Hit
    PlaySoundAt "fx_woodhit", doorf
    OpenDoor
End Sub

Sub lock_Hit
    DOF 414, DOFPulse
    Dim delay
    delay = 500
    PlaySoundAt "fx_hole_enter", lock
    bsJackal.AddBall Me
    CloseDoor
    If(bJackpot = True)AND(light45.State = 2)Then
        light45.State = 0
        AwardJackpot
    End If
    If light46.State = 1 Then 'lock the ball
        BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
        delay = 4000
        Select Case BallsInLock(CurrentPlayer)
            Case 1:DMD "_", CL("BALL 1 LOCKED"), "_", eNone, eBlinkFast, eNone, 1000, True, "vo_ball1locked"
            DOF 401, DOFPulse
            Case 2:DMD "_", CL("BALL 2 LOCKED"), "_", eNone, eBlinkFast, eNone, 1000, True, "vo_ball2locked"
            DOF 402, DOFPulse
            Case 3:DMD "_", CL("BALL 3 LOCKED"), "_", eNone, eBlinkFast, eNone, 1000, True, "vo_ball3locked"
            DOF 403, DOFPulse
        End Select
        light46.State = 0
        If BallsInLock(CurrentPlayer) = 3 Then 'start multiball
            vpmtimer.addtimer 2000, "StartMainMultiball '"
        End If
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light37.State = 2 Then
                Light37.State = 0
                Addscore 100000
                CheckWinBattle
                Delay = 1000
            End If
        Case 6
            If Light37.State = 2 Then
                Light34.State = 2
                Light37.State = 0
                Addscore 100000
                CheckWinBattle
                Delay = 1000
            End If
        Case 9
            If Light37.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light37.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    If(Battle(CurrentPlayer, NewBattle) = 2)AND(Battle(CurrentPlayer, 0) = 0)Then 'the battle is ready, so start it
        vpmtimer.addtimer 2000, "StartBattle '"
        delay = 6000
    End If
	JackalHole.timerinterval = Delay
    JackalHole.timerenabled = True
'    vpmtimer.addtimer delay, "JackalExit '"
End Sub

Sub StartMainMultiball
    AddMultiball 3
    DMD "_", CL("MULTIBALL"), "_", eNone, eBlinkFast, eNone, 1000, True, "vo_multiball"
    DOF 416, DOFPulse
    StartJackpots
    ChangeGi 5
    'reset BallsInLock variable
    BallsInLock(CurrentPlayer) = 0
    EnableBallSaver 20
End Sub

Sub OpenDoor
	PlaySoundAt "ready", Doorf
    Doorf.RotateToEnd
    door.IsDropped = 1
End Sub

Sub CloseDoor
	PlaySoundAt "", Doorf
    Doorf.RotateToStart 
    door.IsDropped = 0
End Sub

'**********
' Jackpots
'**********
' Jackpots are enabled during the Main multiball and the wizard battles

Sub StartJackpots
    bJackpot = true
    'turn on the jackpot lights
    Select Case Battle(CurrentPlayer, 0)
        Case 9 'Anubis - jackpots on the ramps
            light44.State = 2
            light40.State = 2
        Case 10 'Osiris - jackpots on the sphynxs
            light42.State = 2
            light40.State = 2
            light45.State = 2
        Case 11 'Horus - jackpots on the ramps
            light44.State = 2
            light40.State = 2
        Case 12 'Ra - jackpots on the sphynxs
            light42.State = 2
            light40.State = 2
            light45.State = 2
        Case 13 'final battle - all jackpots on
            light42.State = 2
            light41.State = 2
            light40.State = 2
            light44.State = 2
            light45.State = 2
            Light49.State = 2
            light51.State = 2
        Case Else
            If bMultiballMode Then
                light44.State = 2
                light49.State = 2
            End If
    End Select
End Sub

Sub ResetJackpotLights 'when multiball is finished, resets jackpot and superjackpot lights
    bJackpot = False
    light42.State = 0
    light41.State = 0
    light40.State = 0
    light44.State = 0
    light45.State = 0
    Light49.State = 0
    light51.State = 0
    light52.State = 0
End Sub

Sub EnableSuperJackpot
    If bJackpot = True Then
        If light42.State + light41.State + light40.State + light44.State + light45.State + Light49.State + light51.State = 0 Then
            'PlaySound "vo_superjackpotislit"
            light48.State = 2
            light52.State = 2
        End If
    End If
End Sub

'***********************************
' Blue Targets:  The Mummy Targets
'***********************************

Sub Target2_Hit   
    PlaySoundAtBall SoundFXDOF("fx_target", 120, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target2"
    ' Do some sound or light effect
    Light23.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target4_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 120, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target4"
    ' Do some sound or light effect
    Light24.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target5_Hit
    SABREWULF1Shake
    SABREARMShake
    Playsound "Slice"
    If FlippersBlood Then RightSplat
    If NOT bMultiBallMode Then PlaySound"tSab_"&RndNbr(5)
    PlaySoundAtBall SoundFXDOF("fx_target", 113, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target5"
    ' Do some sound or light effect
    
    Light25.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target7_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 113, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target7"
    ' Do some sound or light effect
    Light26.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target10_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 114, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target10"
    ' Do some sound or light effect
    Light27.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target8_Hit
    Playsound "target_Jago"
    Playsound "Slice2"
    JAGOShake
    FlashJAGOWeapon 800, 100
    PlaySoundAtBall SoundFXDOF("fx_target", 114, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target8"
    ' Do some sound or light effect
    Light28.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Check6BankTargets
    Dim tmp
    FlashForMs f1, 1000, 50, 0
    FlashForMs f3, 1000, 50, 0
    FlashForMs f4, 1000, 50, 0
    FlashForMs f4, 1000, 50, 0
    tmp = INT(RND * 10) + 1
    ' if all 6 targets are hit then kill a monster & activate the mystery light
    If light23.state + light24.state + light25.state + light26.state + light27.state + light28.state = 6 Then
        ' kill a monster
        MonstersKilled(CurrentPlayer) = MonstersKilled(CurrentPlayer) + 1
        DMD " COMBO", " BREAKER", "monster_" &tmp, eNone, eNone, eNone, 2500, True, ""
        Playsound "combobreaker"
        LightEffect 1
        FlashEffect 1
        ' Lit the Mystery light if it is off
        If Light38.State = 1 Then
            AddScore 50000
        Else
            Light38.State = 1
            AddScore 25000
        End If
        ' reset the lights
        light23.state = 0
        light24.state = 0
        light25.state = 0
        light26.state = 0
        light27.state = 0
        light28.state = 0
    End If
End Sub

' Playfiel Multiplier timer: reduces the multiplier after 30 seconds

Sub pfxtimer_Timer
    If PlayfieldMultiplier(CurrentPlayer) > 1 Then
        PlayfieldMultiplier(CurrentPlayer) = PlayfieldMultiplier(CurrentPlayer)-1
        SetPlayfieldMultiplier PlayfieldMultiplier(CurrentPlayer)
    Else
        pfxtimer.Enabled = 0
    End If
End Sub



'*****************
'  Captive Target
'*****************

Sub Target9_Hit
    THUNDERShake
    FlashForMs f22, 1000, 50, 0
    PlaySoundAtBall SoundFXDOF("fx_target", 115, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    If bSkillShotReady Then
        Awardskillshot
        Exit Sub
    End If
    AddScore 5000 'all targets score 5000
    ' Do some sound or light effect
    ' do some check
    If(bJackpot = True)AND(light52.State = 2)Then
        AwardSuperJackpot
        light52.State = 0
        light48.State = 0
        StartJackpots
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 0:SelectBattle 'no battle is active then change to another battle
    End Select

    ' increase the playfield multiplier for 30 seconds
    If light54.State = 1 Then
        AddPlayfieldMultiplier 1
        light54.State = 0
    End If

    ' increase Jackpot
    If light53.State = 1 Then
        AddJackpot 50000
        light53.State = 0
    End If
End Sub

'****************************
'  Jackal Hole Hit & Awards
'****************************

Sub JackalHole_Hit
    EDDShake
    EDDLEGShake
    EYEDOL2Shake
    Playsound "HOOF"
    FlashEyedolWeapon 800, 100
    If NOT bMultiBallMode Then PlaySound"tEyedol_"&RndNbr(5)
    Dim Delay
    Delay = 200
    PlaySoundAt "fx_hole_enter", JackalHole
    bsJackal.AddBall Me
    If NOT Tilted Then
        ' do something
        If(bJackpot = True)AND(light40.State = 2)Then
            light40.State = 0
            AwardJackpot
            Delay = 2000
        End If
        If light38.State = 1 Then ' mystery light is lit
            light38.State = 0
            GiveRandomAward
            Delay = 3500
        End If
        If light39.State = 2 Then ' extra ball is lit
            light39.State = 0
            AwardExtraBall
            Delay = 2000
        End If
        Select Case Battle(CurrentPlayer, 0)
            Case 5
                If Light32.State = 2 Then
                    Light32.State = 0
                    Addscore 100000
                    CheckWinBattle
                    Delay = 1000
                End If
            Case 6
                If Light32.State = 2 Then
                    Light36.State = 2
                    Light32.State = 0
                    Addscore 100000
                    CheckWinBattle
                    Delay = 1000
                End If
            Case 9
                If Light32.State = 2 Then
                    AddScore 100000
                    FlashEffect 3
                    DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
                End If
            Case 11
                If Light32.State = 2 Then
                    AddScore 120000
                    FlashEffect 3
                    DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
                End If
        End Select
    End If

	JackalHole.timerinterval = Delay
	JackalHole.timerenabled = True
'    vpmtimer.addtimer Delay, "JackalExit '"
End Sub

Sub JackalHole_timer
	JackalHole.timerenabled = False
	JackalHole.timerinterval = 1000
    If bsJackal.Balls > 0 Then
        FlashForMs f1, 1000, 50, 0
        PlaySoundAt SoundFXDOF("fx_kicker", 119, DOFPulse, DOFContactors), JackalHole
        DOF 121, DOFPulse
        
        'add a small delay before actually kicking the ball
		JackalHole.timerenabled = True
'       vpmtimer.addtimer 500, "bsJackal.ExitSol_On '"
		Jackalkick.enabled = True
    End If
	'kick out all the balls
'    If bsJackal.Balls > 0 Then
'		JackalHole.timerinterval = 500
'		JackalHole.timerenabled = True
''       vpmtimer.Addtimer 500, "JackalExit '"
'    End If
End Sub

Sub Jackalkick_Timer
	Jackalkick.enabled = False
	bsJackal.ExitSol_On	
End Sub


Sub GiveRandomAward() 'from the Jackal Sphynx
    StartEDL
    Dim tmp, tmp2

    ' show some random values on the dmd
    DMD CL("KILLER AWARD"), "", "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("PLAYFIELD X"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BUMPER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA BALL"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BONUS X"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("SPINNER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BUMPER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("PLAYFIELD X"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BUMPER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA BALL"), "", eNone, eNone, eNone, 50, False, "fx_spinner"

    tmp = INT(RND(1) * 80)
    Select Case tmp
        Case 1, 2, 3, 4, 5, 6 'Lit Extra Ball
            DMD "", CL("EXTRA BALL IS LIT"), "", eNone, eBlink, eNone, 1500, True, "fx_fanfare1"
            light39.State = 2
        Case 7, 8, 13, 14, 15 '100,000 points
            DMD CL("BIG POINTS"), CL("100000"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            AddScore 100000
        Case 9, 10, 11, 12 'Hold Bonus
            DMD CL("BONUS HELD"), CL("ACTIVATED"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            bBonusHeld = True
        Case 16, 17, 18 'Increase Bonus Multiplier
            DMD CL("INCREASED"), CL("BONUS X"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            AddBonusMultiplier 1
        Case 19, 20, 21 'Complete Battle
            If Battle(CurrentPlayer, 0) > 0 AND Battle(CurrentPlayer, 0) < 13 Then
                DMD CL("BATTLE"), CL("COMPLETED"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
                WinBattle
            Else
                DMD CL("BIG POINTS"), CL("100000"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
                AddScore 100000
            End If
        Case 22, 23, 36, 37, 38 'PlayField multiplier
            DMD CL("INCREASED"), CL("PLAYFIELD X"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            AddPlayfieldMultiplier 1
        Case 24, 25, 26, 27, 28 '100,000 points
            DMD CL("BIG POINTS"), CL("100000"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            AddScore 100000
        Case 29, 30, 31, 32, 33, 34, 35 'Increase Bumper value
            BumperValue(CurrentPlayer) = BumperValue(CurrentPlayer) + 500
            DMD CL("BUMPER VALUE"), CL(BumperValue(CurrentPlayer)), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
        Case 39, 40, 43, 44 'extra multiball
            DMD CL("EXTRA"), CL("MULTIBALL"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            AddMultiball 1
        Case 45, 46, 47, 48 ' Ball Save
            DMD CL("BALL SAVE"), CL("ACTIVATED"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            EnableBallSaver 20
        Case 49, 50, 51, 52
            DMD CL("OUTLANE SAVER"), CL("ACTIVATED"), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            Light13.State = 1
        Case ELSE 'Add a Random score from 10.000 to 100,000 points
            tmp2 = INT((RND) * 9) * 10000 + 10000
            DMD CL("EXTRA POINTS"), CL(tmp2), "", eBlink, eBlink, eNone, 1500, True, "fx_fanfare1"
            AddScore tmp2
    End Select
End Sub

Dim fDuration
Dim fPeriod

Sub FlashEyedolWeapon (fDuration, fPeriod)
	FlashEyedolTimer.interval = fPeriod
	FlashEyedolTimer.enabled = True
	StopEyedolTimer.interval = fDuration
	StopEyedolTimer.Enabled = True
End Sub

Sub FlashEyedolTimer_Timer
	if ed2.visible = false Then
		ed2.visible = True
	Else
		ed2.visible = False
	End If
End Sub

Sub StopEyedolTimer_Timer
	StopEyedolTimer.enabled = False
	FlashEyedolTimer.enabled = False
	ed2.visible = False
End Sub

Sub FlashJAGOWeapon (fDuration, fPeriod)
	FlashJAGOSWORDTimer.interval = fPeriod
	FlashJAGOSWORDTimer.enabled = True
	StopJAGOSWORDTimer.interval = fDuration
	StopJAGOSWORDTimer.Enabled = True
End Sub

Sub FlashJAGOSWORDTimer_Timer
	if JAGSWORD.visible = false Then
		JAGSWORD.visible = True
	Else
		JAGSWORD.visible = False
	End If
End Sub

Sub StopJAGOSWORDTimer_Timer
	StopJAGOSWORDTimer.enabled = False
	FlashJAGOSWORDTimer.enabled = False
	JAGSWORD.visible = False
End Sub

Sub FlashORCHWeapon (fDuration, fPeriod)
	FlashORCHBATTimer.interval = fPeriod
	FlashORCHBATTimer.enabled = True
	StopORCHBATTimer.interval = fDuration
	StopORCHBATTimer.Enabled = True
End Sub

Sub FlashORCHBATTimer_Timer
	if BATONS.visible = false Then
		BATONS.visible = True
	Else
		BATONS.visible = False
	End If
End Sub

Sub StopORCHBATTimer_Timer
	StopORCHBATTimer.enabled = False
	FlashORCHBATTimer.enabled = False
	BATONS.visible = False
End Sub



'*******************
'   The Orbit lanes
'*******************

Sub sw8_Hit
    ShakeARM
    SPINALShake
    If FlippersBlood Then LeftSplat
    Playsound "target_Spinal"
    Playsound "Slice2"
    DOF 130, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    If(bJackpot = True)AND(light41.State = 2)Then
        light41.State = 0
        AwardJackpot
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 4:OrbitHits = OrbitHits + 1:Addscore 70000:CheckWinBattle
        Case 5
            If Light33.State = 2 Then
                Light33.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light33.State = 2 Then
                Light32.State = 2
                Light33.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light33.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 10
            If LastSwitchHit = "sw7" Then
                LastSwitchHit = ""
                loopCount = loopCount + 1
                Addscore 140000
                CheckWinBattle
            End If
        Case 11
            If Light33.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light33.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 2
                Light36.State = 2
                Light33.State = 0
                Light35.State = 0
                Addscore 100000
                CheckWinBattle
            End If
    End Select
    LastSwitchHit = "sw8"
End Sub

Sub sw7_Hit
    DOF 131, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    If(bJackpot = True)AND(light51.State = 2)Then
        light51.State = 0
        AwardJackpot
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 4:OrbitHits = OrbitHits + 1:Addscore 70000:CheckWinBattle
        Case 5
            If Light35.State = 2 Then
                Light35.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light35.State = 2 Then
                Light29.State = 2
                Light35.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light35.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 10
            If LastSwitchHit = "sw8" Then
                LastSwitchHit = ""
                loopCount = loopCount + 1
                Addscore 140000
                CheckWinBattle
            End If
        Case 11
            If Light35.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light35.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 2
                Light36.State = 2
                Light33.State = 0
                Light35.State = 0
                Addscore 100000
                CheckWinBattle
            End If
    End Select
    LastSwitchHit = "sw7"
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'****************
'     Ramps
'****************

Sub LeftRampDone_Hit
    Playsound "Slice2"
    If FlippersBlood Then LeftSplat
    ShakeARM
    SPINALShake
    If NOT bMultiBallMode Then PlaySound"Spinal_"&RndNbr(2)
    Dim tmp
	DOF 311, DOFPulse
    If Tilted Then Exit Sub
    'increase the ramp bonus
    RampBonus = RampBonus + 1
    If(bJackpot = True)AND(light44.State = 2)Then
        light44.State = 0
        AwardJackpot
    End If
    'PowerUp - left ramp only counts the variable
    PowerupHits = PowerupHits + 1
    CheckPowerup
    'Battles
    Select Case Battle(CurrentPlayer, 0)
        Case 3:RampHits3 = RampHits3 + 1:Addscore 100000:CheckWinBattle
        Case 5
            If Light36.State = 2 Then
                Light36.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light36.State = 2 Then
                Light37.State = 2
                Light36.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light36.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light36.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light36.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 0
                Light36.State = 0
                Light33.State = 2
                Light35.State = 2
                Addscore 100000
                CheckWinBattle
            End If
        Case else
            ' play ss quote
            
    End Select
    'check for combos
    if LastSwitchHit = "RightRampDone" OR LastSwitchHit = "LeftRampDone" Then
        AwardCombo
    End If
    LastSwitchHit = "LeftRampDone"
	
End Sub

Sub RightRampDone_Hit
    CINDERShake
    CINDER2Shake
    FlashForMs f23, 1000, 50, 0
    If NOT bMultiBallMode Then PlaySound"tCind_"&RndNbr(4)
    Dim tmp
	DOF 312, DOFPulse
    If Tilted Then Exit Sub
    'increase the ramp bonus
    RampBonus = RampBonus + 1
    If(bJackpot = True)AND(light49.State = 2)Then
        light49.State = 0
        AwardJackpot
    End If
    'Powerup - rightt ramp counts the variable and give the jackpot if light31 is lit
    If light50.State = 2 Then
        DMD CL("POWERUP AWARD"), CL(jackpot(CurrentPlayer)), "_", eNone, eBlinkFast, eNone, 1000, True, "vo_Jackpot"
        AddScore Jackpot(CurrentPlayer)
        LightEffect 2
        FlashEffect 2
    Else
        PowerupHits = PowerupHits + 1
        CheckPowerup
    End If
    'Battles
    Select Case Battle(CurrentPlayer, 0)
        Case 3:RampHits3 = RampHits3 + 1:Addscore 100000:CheckWinBattle
        Case 5
            If Light34.State = 2 Then
                Light34.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light34.State = 2 Then
                Light35.State = 2
                Light34.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light34.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light34.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light34.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 0
                Light36.State = 0
                Light33.State = 2
                Light35.State = 2
                Addscore 100000
                CheckWinBattle
            End If
        Case else
            ' play ss quote
            
    End Select

    'check for combos
    if LastSwitchHit = "RightRampDone" OR LastSwitchHit = "LeftRampDone" Then
        AwardCombo
    End If
    LastSwitchHit = "RightRampDone"
End Sub

'******************
' Left Sphynx : Ram
'******************

Sub Target12_Hit
    GLACIUSShake
    GLACIUSARMShake
    StartGlac
    PlaySoundAtBall "fx_target"
	DOF 112, DOFPulse
    If NOT bMultiBallMode Then PlaySound"glac"
    If Tilted Then Exit Sub
    If(bJackpot = True)AND(light42.State = 2)Then
        light42.State = 0
        AwardJackpot
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Light30.State = 2 Then
                Light30.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light30.State = 2 Then
                Light31.State = 2
                Light30.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light30.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light30.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    LastSwitchHit = "Target12"
End Sub

'************************
'       Battles
'************************

' This table has 12 main battles, and a final battle
' you may choose any the 12 main battles you want to play
' After completing all 12 battles you play the final battle

' current active battle number is stored in Battle(CurrentPlayer,0)

Sub SelectBattle 'select a new random battle if none is active
    Dim i
    If Battle(CurrentPlayer, 0) = 0 Then
        ' reset the battles that are not finished
        For i = 1 to 12
            If Battle(CurrentPlayer, i) = 2 Then Battle(CurrentPlayer, i) = 0
        Next
        If BattlesWon(CurrentPlayer) = 12 tHEN
            NewBattle = 13:Battle(CurrentPlayer, NewBattle) = 2:UpdateBattleLights:StartBattle '13 battle is the wizard
        Else
            NewBattle = INT(RND * 12 + 1)
            do while Battle(CurrentPlayer, NewBattle) <> 0
                NewBattle = INT(RND * 11 + 1)
            loop
            Battle(CurrentPlayer, NewBattle) = 2
            Light47.State = 2
            UpdateBattleLights
        End iF
    'debug.print "newbatle " & newbattle
    End If
End Sub

' Update the lights according to the battle's state
Sub UpdateBattleLights
    Light11.State = Battle(CurrentPlayer, 1)
    Light12.State = Battle(CurrentPlayer, 2)
    Light9.State = Battle(CurrentPlayer, 3)
    Light10.State = Battle(CurrentPlayer, 4)
    Light7.State = Battle(CurrentPlayer, 5)
    Light8.State = Battle(CurrentPlayer, 6)
    Light5.State = Battle(CurrentPlayer, 7)
    Light6.State = Battle(CurrentPlayer, 8)
    Light1.State = Battle(CurrentPlayer, 9)
    Light2.State = Battle(CurrentPlayer, 10)
    Light4.State = Battle(CurrentPlayer, 11)
    Light3.State = Battle(CurrentPlayer, 12)
End Sub

' Starting a battle means to setup some lights and variables, maybe timers
' Battle lights will always blink during an active battle
Sub StartBattle
    Battle(CurrentPlayer, 0) = NewBattle
    Light47.State = 0
    ChangeSong
    PlaySound "dooropen_READY"
    EnableBallSaver 15 'start a 15 seconds ball save
    Select Case NewBattle
        Case 1         'Bast = Super Spinners
            DMD CL("THUNDER"), CL("SHOOT THE SPINNERS"), "", eNone, eNone, eNone, 1500, True, ""
            StartTHU
            THUNDERShake
            FlashForMs f22, 1000, 50, 0
            Light33.State = 2
            Light35.State = 2
            SpinCount = 0
            ChangeGi orange
        Case 2 'Isis = Super Pop Bumpers
            DMD CL("RIPTOR"), CL("HIT THE POP BUMPERS"), "", eNone, eNone, eNone, 1500, True, ""
            Playsound "BRiptor"
            RIPTORShake
            RIPTORARMShake
            Light55.State = 2
            LightSeqBumpers.Play SeqRandom, 10, , 1000
            SuperBumperHits = 0
            ChangeGi green
        Case 3 'Amon = Ramps
            DMD CL("CINDER"), CL("SHOOT THE RAMPS"), "", eNone, eNone, eNone, 1500, True, ""
            StartIron
            Light36.State = 2
            Light34.State = 2
            RampHits3 = 0
            ChangeGi yellow
        Case 4 'Shu = Orbits
            DMD CL("SABREWULF"), CL("SHOOT THE ORBITS"), "", eNone, eNone, eNone, 1500, True, ""
            Playsound "Saberhowl"
            Playsound "Slice2"
            SABREWULF1Shake
            SABREARMShake
            Light33.State = 2
            Light35.State = 2
            ChangeGi Red
            OrbitHits = 0
        Case 5 'Set = Shoot the lights 2
            DMD CL("SPINAL"), CL("SHOOT THE LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            Playsound "Slice2"
            Playsound "Spinal_1"
            StartSPIN
            If FlippersBlood Then LeftSplat
            ShakeARM
            SPINALShake
            Light29.State = 2
            Light30.State = 2
            Light31.State = 2
            Light32.State = 2
            Light33.State = 2
            Light34.State = 2
            Light35.State = 2
            Light36.State = 2
            Light37.State = 2
            ChangeGi amber
        Case 6 'Nepthys= Shoot the lights 1
            DMD CL("FULGORE"), CL("SHOOT THE LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            StartGun
            Light30.State = 2
            ChangeGi teal
        Case 7 'Bes =  Blue Target Frenzy
            DMD CL("TJ COMBO"), CL("SHOOT THE TARGETS"), "", eNone, eNone, eNone, 1500, True, ""
            Playsound "target_TJCombo"
            ShakeRightCat
            ShakeRight2Cat
            LightSeqBlueTargets.Play SeqRandom, 10, , 1000
            TargetHits7 = 0
            ChangeGi red
        Case 8 'Serget = Left & Right Targets
            DMD CL("JAGO"), CL("SHOOT THE TARGETS"), "", eNone, eNone, eNone, 1500, True, ""
            StartJAG
            FlashJAGOWeapon 800, 100
            Light31.State = 2
            Light29.State = 2
            ChangeGi blue
            TargetHits8 = 0
        Case 9 'Anubis = Follow the Lights 1
            DMD CL("EYEDOL"), CL("SHOOT LIT LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            EDDShake
            EDDLEGShake
            EYEDOL2Shake
            Playsound "HOOF"
            Playsound "Eyedol_award"
            FlashEyedolWeapon 800, 100
            FollowTheLights.Enabled = 1
            LightHits9 = 0
            ChangeGi purple
        Case 10 'Osiris = Super Loops
            DMD CL("ORCHID"), CL("SHOOT THE LOOPS"), "", eNone, eNone, eNone, 1500, True, ""
            Playsound "target_Orchid"
            FlashForMs flash2, 1000, 50, 0
            StartOrc
            Light33.State = 2
            Light35.State = 2
            Gate2.Open = 1
            Gate3.Open = 1
            loopCount = 0
            ChangeGi darkgreen
        Case 11 'Horus = Follow the Lights 2
            DMD CL("GLACIUS"), CL("SHOOT LIT LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            GLACIUSShake
            GLACIUSARMShake
            StartGlac
            Playsound "target_Glacius"
            FollowTheLights.Enabled = 1
            LightHits11 = 0
            ChangeGi blue
        Case 12 'Ra = Ramps and Orbits
            'uses the ramphits12 to count the hits
            DMD CL("ULTRA TECH"), CL("SHOOT RAMPS ORBITS"), "", eNone, eNone, eNone, 1500, True, ""
            Light36.State = 2
            Light34.State = 2
            RampHits12 = 0
            ChangeGi red
        Case 13 'Mordekai the Summoner - the final battle
            DMD CL("KILLER INSTINCT"), CL("SHOOT THE JACKPOTS"), "", eNone, eNone, eNone, 1500, True, ""
            AddMultiball 4
            StartJackpots
            ChangeGi blue
            '20 second ball saver
            EnableBallSaver 20
    End Select
End Sub

' check if the battle is completed
Sub CheckWinBattle
    dim tmp
    tmp = INT(RND * 7) + 1
    PlaySound "fx_thunder" & tmp
    FlashForMs f10, 1000, 50, 0
    DOF 126, DOFPulse
    LightSeqInserts.StopPlay 'stop the light effects before starting again so they don't play too long.
    LightEffect 3
    FlashEffect 3
    Select Case NewBattle
        Case 1
            If SpinCount = 100 Then WinBattle:End if
        Case 2
            If SuperBumperHits = 25 Then WinBattle:End if
        Case 3
            If RampHits3 = 6 Then WinBattle:End if
        Case 4
            If OrbitHits = 6 Then WinBattle:End if
        Case 5
            If Light29.State + Light30.State + Light31.State + Light32.State + Light33.State + Light34.State + Light35.State + Light36.State + Light37.State = 0 Then WinBattle:End if
        Case 6 'the last light win the battle
        Case 7
            If TargetHits7 = 20 Then WinBattle:End if
        Case 8
            If TargetHits8 = 6 Then WinBattle:End if
        Case 9
            If LightHits9 = 8 Then WinBattle:End if
        Case 10:
            If loopCount = 6 Then WinBattle
        Case 11
            If LightHits11 = 8 Then WinBattle:End if
        Case 12
            If RampHits12 = 6 Then WinBattle:End if
    End Select
End Sub

Sub StopBattle 'called at the end of a ball
    Dim i
    Battle(CurrentPlayer, 0) = 0
    For i = 0 to 15
        If Battle(CurrentPlayer, i) = 2 Then Battle(CurrentPlayer, i) = 0
    Next
    UpdateBattleLights
    StopBattle2
    NewBattle = 0
End Sub

'called after completing a battle
Sub WinBattle
    Dim tmp
    BattlesWon(CurrentPlayer) = BattlesWon(CurrentPlayer) + 1
    Battle(CurrentPlayer, 0) = 0
    Battle(CurrentPlayer, NewBattle) = 1
    UpdateBattleLights
    FlashEffect 2
    LightEffect 2
    GiEffect 2
    DMD "", CL("BATTLE COMPLETED"), "_", eNone, eBlinkFast, eNone, 1000, True, "fx_Explosion01"
    DOF 139, DOFPulse
    tmp = INT(RND * 4)
    Select Case tmp
        Case 0:vpmtimer.addtimer 1500, "PlaySound ""vo_excelent"" '"
        Case 1:vpmtimer.addtimer 1500, "PlaySound ""vo_impressive"" '"
        Case 2:vpmtimer.addtimer 1500, "PlaySound ""vo_welldone"" '"
        Case 3:vpmtimer.addtimer 1500, "PlaySound ""vo_YouWon"" '"
    End Select
    StopBattle2
    NewBattle = 0
    SelectBattle 'automatically select a new battle
    'add a multiball after each 2 won battles
    Select Case BattlesWon(CurrentPlayer)
        Case 3, 6, 9:AddMultiball 2: EnableBallSaver 20
    End Select
    ChangeSong
End Sub

Sub StopBattle2
    'Turn off the bomb lights
    Light29.State = 0
    Light30.State = 0
    Light31.State = 0
    Light32.State = 0
    Light33.State = 0
    Light34.State = 0
    Light35.State = 0
    Light36.State = 0
    Light37.State = 0
    ' stop some timers or reset battle variables
    Select Case NewBattle
        Case 1:SpinCount = 0
        Case 2:Light55.State = 0:LightSeqBumpers.StopPlay:SuperBumperHits = 0
        Case 3, 12:
        Case 4:OrbitHits = 0
        Case 7:LightSeqBlueTargets.StopPlay
        Case 8:
        Case 9, 11:FollowTheLights.Enabled = 0
        Case 10:LoopCount = 0:Gate2.Open = 0:Gate3.Open = 0
        Case 13:ResetBattles:SelectBattle
    End Select
End Sub

Sub ResetBattles
    Dim i, j
    For j = 0 to 4
        BattlesWon(j) = 0
        For i = 0 to 12
            Battle(CurrentPlayer, i) = 0
        Next
    Next
    NewBattle = 0
End Sub

'Extra subs for the battles

Sub LightSeqAllTargets_PlayDone()
    LightSeqAllTargets.Play SeqRandom, 10, , 1000
End Sub

Sub LightSeqBumpers_PlayDone()
    LightSeqBumpers.Play SeqRandom, 10, , 1000
End Sub

Sub LightSeqBlueTargets_PlayDone()
    LightSeqBlueTargets.Play SeqRandom, 10, , 1000
End Sub

' Wizards modes timer
Dim FTLstep:FTLstep = 0

Sub FollowTheLights_Timer
    Light29.State = 0
    Light30.State = 0
    Light31.State = 0
    Light32.State = 0
    Light33.State = 0
    Light34.State = 0
    Light35.State = 0
    Light36.State = 0
    Light37.State = 0
    Select Case Battle(CurrentPlayer, 0)
        Case 9
            Select case FTLstep
                Case 0:FTLstep = 1:Light29.State = 2
                Case 1:FTLstep = 2:Light30.State = 2
                Case 2:FTLstep = 3:Light31.State = 2
                Case 3:FTLstep = 4:Light32.State = 2
                Case 4:FTLstep = 5:Light33.State = 2
                Case 5:FTLstep = 6:Light34.State = 2
                Case 6:FTLstep = 7:Light35.State = 2
                Case 7:FTLstep = 8:Light36.State = 2
                Case 8:FTLstep = 0:Light37.State = 2
            End Select
        Case 11
            FTLstep = INT(RND * 9)
            Select case FTLstep
                Case 0:Light29.State = 2
                Case 1:Light30.State = 2
                Case 2:Light31.State = 2
                Case 3:Light32.State = 2
                Case 4:Light33.State = 2
                Case 5:Light34.State = 2
                Case 6:Light35.State = 2
                Case 7:Light36.State = 2
                Case 8:Light37.State = 2
            End Select
    End Select
End Sub

'**********************
' Power up Jackpot
'**********************
' 30 seconds hurry up with jackpots on the right ramp
' uses variable PowerupHits and the light50

Sub CheckPowerup
    If light50.State = 0 Then
        If PowerupHits MOD 10 = 0 Then
            EnablePowerup
        End If
    End If
End Sub

Sub EnablePowerup
    ' start the timers
    PowerupTimerExpired.Enabled = True
    PowerupSpeedUpTimer.Enabled = True
    ' turn on the light
    Light50.BlinkInterval = 160
    Light50.State = 2
End Sub

Sub PowerupTimerExpired_Timer()
    PowerupTimerExpired.Enabled = False
    ' turn off the light
    Light50.State = 0
End Sub

Sub PowerupSpeedUpTimer_Timer()
    PowerupSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    Light50.BlinkInterval = 80
    Light50.State = 2
End Sub

'***********************
' Shake Models
'***********************
Dim CatLPos,CATL2Pos, CatRPos, CatR2Pos, BatonsPOS

Sub ShakeLeftCat
    CatLPos = 8
    CatLTimer.Enabled = 1
End Sub

Sub CatLTimer_Timer
    CatL.TransY = CatLPos
    If CatLPos = 0 Then Me.Enabled = 0:Exit Sub
    If CatLPos < 0 Then
        CatLPos = ABS(CatLPos)- 1
    Else
        CatLPos = - CatLPos + 1
    End If
End Sub

Sub ShakeBatons
    BatonsPos = 8
    BatonsTimer.Enabled = 1
End Sub

Sub BatonsTimer_Timer
    Batons.TransY = BatonsPos
    If BatonsPos = 0 Then Me.Enabled = 0:Exit Sub
    If BatonsPos < 0 Then
        BatonsPos = ABS(BatonsPos)- 1
    Else
        BatonsPos = - BatonsPos + 1
    End If
End Sub

Sub ShakeLeft2Cat
    CatL2Pos = 8
    CatL2Timer.Enabled = 1
End Sub

Sub CatL2Timer_Timer
    CatL2.TransY = CatL2Pos
    If CatL2Pos = 0 Then Me.Enabled = 0:Exit Sub
    If CatL2Pos < 0 Then
        CatL2Pos = ABS(CatL2Pos)- 1
    Else
        CatL2Pos = - CatL2Pos + 1
    End If
End Sub

Sub ShakeRightCat
    CatRPos = -15
    CatRTimer.Enabled = 1
End Sub

Sub CatRTimer_Timer
    CatR.RotZ = CatRPos * -90
    If CatRPos = 0 Then Me.Enabled = 0:Exit Sub
    If CatRPos < 0 Then
        CatRPos = ABS(CatRPos)- 1
    Else
        CatRPos = - CatRPos + 1
    End If
End Sub

Sub ShakeRight2Cat
    CatR2Pos = 5
    CatR2Timer.Enabled = 1
End Sub

Sub CatR2Timer_Timer
    CatR2.TransY = CatR2Pos
    If CatR2Pos = 0 Then Me.Enabled = 0:Exit Sub
    If CatR2Pos < 0 Then
        CatR2Pos = ABS(CatR2Pos)- 1
    Else
        CatR2Pos = - CatR2Pos + 1
    End If
End Sub

'****Rotate Arm*****

Dim ARMPos

Sub ShakeARM
    ARMPos = 4
    ARMTimer.Enabled = 1
End Sub

Sub ARMTimer_Timer
    ARM.RotX = ARMPos * -90
    If ARMPos = 0 Then Me.Enabled = 0:Exit Sub
    If ARMPos < 0 Then
        ARMPos = ABS(ARMPos)- 1
    Else
        ARMPos = - ARMPos + 1
    End If
End Sub

Dim JAGOPos

Sub JAGOShake
    JAGOPos = 4
    JAGOTimer.Enabled = 1
End Sub

Sub JAGOTimer_Timer
    JAGO.RotX = JAGOPos * -20
    If JAGOPos = 0 Then Me.Enabled = 0:Exit Sub
    If JAGOPos < 0 Then
        JAGOPos = ABS(JAGOPos)- 1
    Else
        JAGOPos = - JAGOPos + 1
    End If
End Sub

Dim EDDLEGPos

Sub EDDLEGShake
    EDDLEGPos = 8
    EDDLEGTimer.Enabled = 1
End Sub

Sub EDDLEGTimer_Timer
    EDDLEG.RotX = EDDLEGPos * -3
    If EDDLEGPos = 0 Then Me.Enabled = 0:Exit Sub
    If EDDLEGPos < 0 Then
        EDDLEGPos = ABS(EDDLEGPos)- 1
    Else
        EDDLEGPos = - EDDLEGPos + 1
    End If
End Sub

Dim RIPTORARMPos

Sub RIPTORARMShake
    RIPTORARMPos = 6
    RIPTORARMTimer.Enabled = 1
End Sub

Sub RIPTORARMTimer_Timer
    RIPTORARM.RotX = RIPTORARMPos * 20
    If RIPTORARMPos = 0 Then Me.Enabled = 0:Exit Sub
    If RIPTORARMPos < 0 Then
        RIPTORARMPos = ABS(RIPTORARMPos)- 1
    Else
        RIPTORARMPos = - RIPTORARMPos + 1
    End If
End Sub

Dim SABREARMPos

Sub SABREARMShake
    SABREARMPos = 4
    SABREARMTimer.Enabled = 1
End Sub

Sub SABREARMTimer_Timer
    SABREARM.RotX = SABREARMPos * -20
    If SABREARMPos = 0 Then Me.Enabled = 0:Exit Sub
    If SABREARMPos < 0 Then
        SABREARMPos = ABS(SABREARMPos)- 1
    Else
        SABREARMPos = - SABREARMPos + 1
    End If
End Sub

Dim GLACIUSARMPos

Sub GLACIUSARMShake
    GLACIUSARMPos = 4
    GLACIUSARMTimer.Enabled = 1
End Sub

Sub GLACIUSARMTimer_Timer
    GLACIUSARM.RotX = GLACIUSARMPos * -20
    If GLACIUSARMPos = 0 Then Me.Enabled = 0:Exit Sub
    If GLACIUSARMPos < 0 Then
        GLACIUSARMPos = ABS(GLACIUSARMPos)- 1
    Else
        GLACIUSARMPos = - GLACIUSARMPos + 1
    End If
End Sub



' Fire lamps
Dim Fire1Pos, Fire2Pos, Fire3Pos, Fire4Pos, Fire5Pos, Flames
Flames = Array("fire01", "fire02", "fire03", "fire04", "fire05", "fire06", "fire07", "fire08", "fire09", _
    "fire10", "fire11", "fire12", "fire13", "fire14", "fire15", "fire16")

Sub StartFire
    Fire1Pos = 0
    Fire2Pos = 2
    Fire3Pos = 5
    Fire4Pos = 8
    Fire5Pos = 11
   
    FireTimer.Enabled = 1
End Sub

Sub FireTimer_Timer
    'debug.print fire1pos
    Fire1.ImageA = Flames(Fire1Pos)
    Fire2.ImageA = Flames(Fire2Pos)
    Fire3.ImageA = Flames(Fire3Pos)
    Fire4.ImageA = Flames(Fire4Pos)
    Fire5.ImageA = Flames(Fire5Pos)
    
    Fire1Pos = (Fire1Pos + 1) MOD 16
    Fire2Pos = (Fire2Pos + 1) MOD 16
    Fire3Pos = (Fire3Pos + 1) MOD 16
    Fire4Pos = (Fire4Pos + 1) MOD 16
    Fire5Pos = (Fire5Pos + 1) MOD 16
    
End Sub

Sub SphereTimer_Timer
   spherelord.rotY = spherelord.rotY + 2

End Sub

'*****************
' Animated Parts
'*****************

Dim GLACShake, THUNShake, RIPTShake, SABRE1Shake, JAGShake, ED1Shake, ED2Shake, SPINShake, FULGOSHAKE, FULGO2SHAKE, CINDShake, CIND2Shake 

Sub GLACIUSShake()
    GLACShake = 3
    GLACShakeTimer.Enabled = True
End Sub

Sub GLACShakeTimer_Timer()
    GLACIUS.Transz = GLACShake / 2
    If GLACShake = 0 Then Me.Enabled = False:Exit Sub
    If GLACShake < 0 Then
        GLACShake = ABS(GLACShake)- 0.1
    Else
        GLACShake = - GLACShake + 0.1
    End If
End Sub

Sub THUNDERShake()
    THUNShake = 6
    THUNShakeTimer.Enabled = True
End Sub

Sub THUNShakeTimer_Timer()
    THUNDER.Transz = THUNShake / 2
    If THUNShake = 0 Then Me.Enabled = False:Exit Sub
    If THUNShake < 0 Then
        THUNShake = ABS(THUNShake)- 0.1
    Else
        THUNShake = - THUNShake + 0.1
    End If
End Sub

Sub RIPTORShake()
    RIPTShake = 3
    RIPTShakeTimer.Enabled = True
End Sub

Sub RIPTShakeTimer_Timer()
    RIPTOR.Transy = RIPTShake / 2
    If RIPTShake = 0 Then Me.Enabled = False:Exit Sub
    If RIPTShake < 0 Then
        RIPTShake = ABS(RIPTShake)- 0.1
    Else
        RIPTShake = - RIPTShake + 0.1
    End If
End Sub

Sub SABREWULF1Shake()
    SABRE1Shake = 2
    SABRE1ShakeTimer.Enabled = True
End Sub

Sub SABRE1ShakeTimer_Timer()
    SABREWULF1.Transy = SABRE1Shake / 2
    If SABRE1Shake = 0 Then Me.Enabled = False:Exit Sub
    If SABRE1Shake < 0 Then
        SABRE1Shake = ABS(SABRE1Shake)- 0.1
    Else
        SABRE1Shake = - SABRE1Shake + 0.1
    End If
End Sub

Sub EDDShake()
    ED1Shake = 6
    EYE1ShakeTimer.Enabled = True
End Sub

Sub EYE1ShakeTimer_Timer()
    EDD.Transy = ED1Shake / 2
    If ED1Shake = 0 Then Me.Enabled = False:Exit Sub
    If ED1Shake < 0 Then
        ED1Shake = ABS(ED1Shake)- 0.1
    Else
        ED1Shake = - ED1Shake + 0.1
    End If
End Sub

Sub EYEDOL2Shake()
    ED2Shake = 6
    ED2ShakeTimer.Enabled = True
End Sub

Sub ED2ShakeTimer_Timer()
    ED2.Transy = ED2Shake / 2
    If ED2Shake = 0 Then Me.Enabled = False:Exit Sub
    If ED2Shake < 0 Then
        ED2Shake = ABS(ED2Shake)- 0.1
    Else
        ED2Shake = - ED2Shake + 0.1
    End If
End Sub

Sub FULGOREShake()
    FULGOShake = 6
    FULGORE1ShakeTimer.Enabled = True
End Sub

Sub FULGORE1ShakeTimer_Timer()
    FULG.Transy = FULGOShake / 2
    If FULGOShake = 0 Then Me.Enabled = False:Exit Sub
    If FULGOShake < 0 Then
        FULGOShake = ABS(FULGOShake)- 0.1
    Else
        FULGOShake = - FULGOShake + 0.1
    End If
End Sub

Sub FULGORE2Shake()
    FULGO2Shake = 6
    FULGORE2ShakeTimer.Enabled = True
End Sub

Sub FULGORE2ShakeTimer_Timer()
    FULG2.Transy = FULGO2Shake / 2
    If FULGO2Shake = 0 Then Me.Enabled = False:Exit Sub
    If FULGO2Shake < 0 Then
        FULGO2Shake = ABS(FULGO2Shake)- 0.1
    Else
        FULGO2Shake = - FULGO2Shake + 0.1
    End If
End Sub

Sub CINDERShake()
    CINDShake = 6
    CINDERShakeTimer.Enabled = True
End Sub

Sub CINDERShakeTimer_Timer()
    CINDER.Transy = CINDShake / 2
    If CINDShake = 0 Then Me.Enabled = False:Exit Sub
    If CINDShake < 0 Then
        CINDShake = ABS(CINDShake)- 0.1
    Else
        CINDShake = - CINDShake + 0.1
    End If
End Sub

Sub CINDER2Shake()
    CIND2Shake = 6
    CINDER2ShakeTimer.Enabled = True
End Sub

Sub CINDER2ShakeTimer_Timer()
    CINDER2.Transy = CIND2Shake / 2
    If CIND2Shake = 0 Then Me.Enabled = False:Exit Sub
    If CIND2Shake < 0 Then
        CIND2Shake = ABS(CIND2Shake)- 0.1
    Else
        CIND2Shake = - CIND2Shake + 0.1
    End If
End Sub

Sub SPINALShake()
    SPINShake = 2
    SPINALShakeTimer.Enabled = True
End Sub

Sub SPINALShakeTimer_Timer()
    SPINAL.Transy = SPINShake / 2
    If SPINShake = 0 Then Me.Enabled = False:Exit Sub
    If SPINShake < 0 Then
        SPINShake = ABS(SPINShake)- 0.1
    Else
        SPINShake = - SPINShake + 0.1
    End If
End Sub

'**************************
' LAVA lamps
'**************************

Dim LAVA1Pos, LAVAFLOW
LAVAFLOW = Array("f01", "f02", "f03", "f04", "f05", "f06", "f07", "f08", "f09", _
    "f10", "f11", "f12", "f13", "f14", "f15", "f16", "F17", "F18", "F19", "F20",_
    "F21", "F22", "F23", "F24", "F25", "F26", "F27", "F28", "F29", "F30", "F31",_
    "F32", "F33", "F34", "F35", "F36", "F37", "F38", "F39", "F40", "F41", "F42", "F43", "F44", "F45", "F46", "F47", "F48")

Sub StartLAVA
    LAVA1Pos = 0
   
    LAVATimer.Enabled = 1
End Sub

Sub LAVATimer_Timer
    'debug.print fire1pos
    LAVA1.ImageA = LAVAFLOW(LAVA1Pos)
    
    LAVA1Pos = (LAVA1Pos + 1) MOD 48    
    
End Sub

'***************************
' ARCADE SCREEN
'***************************
Dim ARCA1Pos, ARCADE
ARCADE = Array("KI-0", "KI-1", "KI-2", "KI-3", "KI-4", "KI-5", "KI-6", "KI-7", "KI-8", "KI-9", _
    "KI-10", "KI-11", "KI-12", "KI-13", "KI-14", "KI-15", "KI-16", "KI-17", "KI-18", "KI-19", "KI-20",_
    "KI-21", "KI-22", "KI-23", "KI-24", "KI-25", "KI-26", "KI-27", "KI-28", "KI-29", "KI-30", "KI-31",_
    "KI-32", "KI-33", "KI-34", "KI-35", "KI-36", "KI-37", "KI-38", "KI-39", "KI-40", "KI-41", "KI-42", "KI-43")

Sub StartARCA
    ARCA1Pos = 0
   
    ARCATimer.Enabled = 1
End Sub

Sub ARCATimer_Timer
    'debug.print fire1pos
    ARCA1.ImageA = ARCADE(ARCA1Pos)
    
    ARCA1Pos = (ARCA1Pos + 1) MOD 44    
    
End Sub
'******************************
'Animations
'******************************

 '*********Orchid*********

Dim	Orc1Pos, OrcFrames
OrcFrames = Array("OR-0", "OR-1", "OR-2", "OR-3", "OR-4", "OR-5", "OR-6", "OR-7", "OR-8", "OR-9", "OR-10", "OR-11", "OR-12", "OR-13", "OR-14", "OR-15", "OR-16", "OR-17", "OR-18", "OR-19", "OR-20", _
"OR-21", "OR-22", "OR-23", "OR-24")

Sub StartOrc
    OrcFlash.visible = True
    Orc1Pos = 0
    OrcTimer.Enabled = True
    StopOrcTimer.interval = 1000
    StopOrcTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub OrcTimer_timer
    OrcFlash.ImageA = OrcFrames(Orc1Pos)
    Orc1Pos = (Orc1Pos + 1) MOD 25
End Sub
 
Sub StopOrc
    OrcFlash.visible = False
    OrcTimer.Enabled = False
End Sub

Sub StopOrcTimer_Timer
    StopOrc
    Me.Enabled = False
    OrcTimer.enabled = False
End Sub
 '*********Glacius*********

Dim	Glac1Pos, GlacFrames
GlacFrames = Array("GL-0", "GL-1", "GL-2", "GL-3", "GL-4", "GL-5", "GL-6", "GL-7", "GL-8", "GL-9", "GL-10", "GL-11", "GL-12", "GL-13", "GL-14", "GL-15", "GL-16", "GL-17", "GL-18", "GL-19", "GL-20", _
"GL-21", "GL-22", "GL-23", "GL-24")

Sub StartGlac
    GlacFlash.visible = True
    Glac1Pos = 0
    GlacTimer.Enabled = True
    StopGlacTimer.interval = 1200
    StopGlacTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub GlacTimer_timer
    GlacFlash.ImageA = GlacFrames(Glac1Pos)
    Glac1Pos = (Glac1Pos + 1) MOD 25
End Sub
 
Sub StopGlac
    GlacFlash.visible = False
    GlacTimer.Enabled = False
End Sub

Sub StopGlacTimer_Timer
    StopGlac
    Me.Enabled = False
    GlacTimer.enabled = False
End Sub
 '*********Fulgore*********

Dim	Gun1Pos, Frames
Frames = Array("MS-000", "MS-001", "MS-002", "MS-003", "MS-004", "MS-005", "MS-006", "MS-007", "MS-008", "MS-009", "MS-0010", "MS-0011", "MS-0012", "MS-0013", "MS-0014", "MS-0015", "MS-0016", "MS-0017", "MS-0018", "MS-0019", "MS-0020", _
"MS-0021", "MS-0022", "MS-0023", "MS-0024", "MS-0025", "MS-0026", "MS-0027", "MS-0028", "MS-0029")

Sub StartGun
    GunFlash.visible = True
    Gun1Pos = 0
    GunTimer.Enabled = True
    StopGunTimer.interval = 1500
    StopGunTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub GunTimer_timer
    GunFlash.ImageA = Frames(Gun1Pos)
    Gun1Pos = (Gun1Pos + 1) MOD 25
End Sub
 
Sub StopGun
    GunFlash.visible = False
    GunTimer.Enabled = False
End Sub

Sub StopGunTimer_Timer
    StopGun
    Me.Enabled = False
    GunTimer.enabled = False
End Sub


 '*********Eyedol*********

Dim	EDL1Pos, EDLFrames
EDLFrames = Array("EY-0", "EY-1", "EY-2", "EY-3", "EY-4", "EY-5", "EY-6", "EY-7", "EY-8", "EY-9", "EY-10", "EY-11", "EY-12", "EY-13", "EY-14", "EY-15", "EY-16", "EY-17", "EY-18", "EY-19", "EY-20", "EY-21", _
"EY-22", "EY-23", "EY-24")


Sub StartEDL
    EDLFlash.visible = True
    EDL1Pos = 0
    EDLTimer.Enabled = True
    StopEDLTimer.interval = 1200
    StopEDLTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub EDLTimer_timer
    EDLFlash.ImageA = EDLFrames(EDL1Pos)
    EDL1Pos = (EDL1Pos + 1) MOD 25
End Sub
 
Sub StopEDL
    EDLFlash.visible = False
    EDLTimer.Enabled = False
End Sub

Sub StopEDLTimer_Timer
    StopEDL
    Me.Enabled = False
    EDLTimer.enabled = False
End Sub

'******Inferno******

Dim	Iron1Pos, IronFrames 
IronFrames = Array("ML-00", "ML-01", "ML-02", "ML-03", "ML-04", "ML-05", "ML-06", "ML-07", "ML-08", "ML-09", "ML-010", "ML-011", "ML-012", "ML-013", "ML-014", "ML-015", "ML-016", "ML-017", "ML-018", "ML-019", "ML-020", _
"ML-021", "ML-022", "ML-023", "ML-024")

Sub StartIron
    IronFlash.visible = True
    Iron1Pos = 0
    IronTimer.Enabled = True
    StopIronTimer.interval = 1500
    StopIronTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub IronTimer_timer
    IronFlash.ImageA = IronFrames(Iron1Pos)
    Iron1Pos = (Iron1Pos + 1) MOD 25
End Sub
 
Sub StopIron
    IronFlash.visible = False
    IronTimer.Enabled = False
End Sub

Sub StopIronTimer_Timer
    StopIron
    Me.Enabled = False
    IronTimer.enabled = False
End Sub

 '*********JAGO*********

Dim	JAG1Pos, JAGFrames
JAGFrames = Array("JA-0", "JA-1", "JA-2", "JA-3", "JA-4", "JA-5", "JA-6", "JA-7", "JA-8", "JA-9", "JA-10", "JA-11", "JA-12", "JA-13", "JA-14", "JA-15", "JA-16", "JA-17", "JA-18", "JA-19", "JA-20", "JA-21", _
"JA-22", "JA-23", "JA-24")

Sub StartJAG
    JAGFlash.visible = True
    JAG1Pos = 0
    JAGTimer.Enabled = True
    StopJAGTimer.interval = 1000
    StopJAGTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub JAGTimer_timer
    JAGFlash.ImageA = JAGFrames(JAG1Pos)
    JAG1Pos = (JAG1Pos + 1) MOD 25
End Sub
 
Sub StopJAG
    JAGFlash.visible = False
    JAGTimer.Enabled = False
End Sub

Sub StopJAGTimer_Timer
    StopJAG
    Me.Enabled = False
    JAGTimer.enabled = False
End Sub
'*********THUNDER***********

Dim	THU1Pos, THUFrames
THUFrames = Array("TH-0", "TH-1", "TH-2", "TH-3", "TH-4", "TH-5", "TH-6", "TH-7", "TH-8", "TH-9", "TH-10", "TH-11", "TH-12", "TH-13", "TH-14", "TH-15", "TH-16", "TH-17", "TH-18", "TH-19", "TH-20", "TH-21", _
"TH-22", "TH-23", "TH-24")

Sub StartTHU
    THUFlash.visible = True
    THU1Pos = 0
    THUTimer.Enabled = True
    StopTHUTimer.interval = 1000
    StopTHUTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub THUTimer_timer
    THUFlash.ImageA = THUFrames(THU1Pos)
    THU1Pos = (THU1Pos + 1) MOD 25
End Sub
 
Sub StopTHU
    THUFlash.visible = False
    THUTimer.Enabled = False
End Sub

Sub StopTHUTimer_Timer
    StopTHU
    Me.Enabled = False
    THUTimer.enabled = False
End Sub


 '*********Spinal*********

Dim	SPIN1Pos, SPINFrames
SPINFrames = Array("SP-0", "SP-1", "SP-2", "SP-3", "SP-4", "SP-5", "SP-6", "SP-7", "SP-8", "SP-9", "SP-10", "SP-11", "SP-12", "SP-13", "SP-14", "SP-15", "SP-16", "SP-17", "SP-18", "SP-19", "SP-20", _
"SP-21", "SP-22", "SP-23", "SP-24")

Sub StartSPIN
    SPINFlash.visible = True
    SPIN1Pos = 0
    SPINTimer.Enabled = True
    StopSPINTimer.interval = 1500
    StopSPINTimer.Enabled = True
    PlaySound ""
End Sub
 
Sub SPINTimer_timer
    SPINFlash.ImageA = SPINFrames(SPIN1Pos)
    SPIN1Pos = (SPIN1Pos + 1) MOD 25
End Sub
 
Sub StopSPIN
    SPINFlash.visible = False
    SPINTimer.Enabled = False
End Sub

Sub StopSPINTimer_Timer
    StopSPIN
    Me.Enabled = False
    SPINTimer.enabled = False
End Sub

'******************************************************
'	ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

'*******************************************
' Early 90's and after

Sub InitPolarity()
	Dim x, a
	a = Array(LF, RF)
	For Each x In a
		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 60
		x.DebugOn=False ' prints some info in debugger
		
		x.AddPt "Polarity", 0, 0, 0
		x.AddPt "Polarity", 1, 0.05, -5.5
		x.AddPt "Polarity", 2, 0.4, -5.5
		x.AddPt "Polarity", 3, 0.6, -5.0
		x.AddPt "Polarity", 4, 0.65, -4.5
		x.AddPt "Polarity", 5, 0.7, -4.0
		x.AddPt "Polarity", 6, 0.75, -3.5
		x.AddPt "Polarity", 7, 0.8, -3.0
		x.AddPt "Polarity", 8, 0.85, -2.5
		x.AddPt "Polarity", 9, 0.9,-2.0
		x.AddPt "Polarity", 10, 0.95, -1.5
		x.AddPt "Polarity", 11, 1, -1.0
		x.AddPt "Polarity", 12, 1.05, -0.5
		x.AddPt "Polarity", 13, 1.1, 0
		x.AddPt "Polarity", 14, 1.3, 0
		
		x.AddPt "Velocity", 0, 0,	   1
		x.AddPt "Velocity", 1, 0.160, 1.06
		x.AddPt "Velocity", 2, 0.410, 1.05
		x.AddPt "Velocity", 3, 0.530, 1'0.982
		x.AddPt "Velocity", 4, 0.702, 0.968
		x.AddPt "Velocity", 5, 0.95,  0.968
		x.AddPt "Velocity", 6, 1.03,  0.945
	Next
	
	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
	LF.SetObjects "LF", LeftFlipper, TriggerLF
	RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt		'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay		'delay before trigger turns off and polarity is disabled
	Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)
	Private Name
	
	Dim PolarityIn, PolarityOut
	Dim VelocityIn, VelocityOut
	Dim YcoefIn, YcoefOut
	Public Sub Class_Initialize
		ReDim PolarityIn(0)
		ReDim PolarityOut(0)
		ReDim VelocityIn(0)
		ReDim VelocityOut(0)
		ReDim YcoefIn(0)
		ReDim YcoefOut(0)
		Enabled = True
		TimeDelay = 50
		LR = 1
		Dim x
		For x = 0 To UBound(balls)
			balls(x) = Empty
			Set Balldata(x) = new SpoofBall
		Next
	End Sub
	
	Public Sub SetObjects(aName, aFlipper, aTrigger)
		
		If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
		If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
		If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
		If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
		Name = aName
		Set Flipper = aFlipper
		FlipperStart = aFlipper.x
		FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
		FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y
		
		Dim str
		str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
		ExecuteGlobal(str)
		str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
		ExecuteGlobal(str)
		
	End Sub
	
	' Legacy: just no op
	Public Property Let EndPoint(aInput)
		
	End Property
	
	Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
		Select Case aChooseArray
			Case "Polarity"
				ShuffleArrays PolarityIn, PolarityOut, 1
				PolarityIn(aIDX) = aX
				PolarityOut(aIDX) = aY
				ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity"
				ShuffleArrays VelocityIn, VelocityOut, 1
				VelocityIn(aIDX) = aX
				VelocityOut(aIDX) = aY
				ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef"
				ShuffleArrays YcoefIn, YcoefOut, 1
				YcoefIn(aIDX) = aX
				YcoefOut(aIDX) = aY
				ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
	End Sub
	
	Public Sub AddBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If IsEmpty(balls(x)) Then
				Set balls(x) = aBall
				Exit Sub
			End If
		Next
	End Sub
	
	Private Sub RemoveBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If TypeName(balls(x) ) = "IBall" Then
				If aBall.ID = Balls(x).ID Then
					balls(x) = Empty
					Balldata(x).Reset
				End If
			End If
		Next
	End Sub
	
	Public Sub Fire()
		Flipper.RotateToEnd
		processballs
	End Sub
	
	Public Property Get Pos 'returns % position a ball. For debug stuff.
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next
	End Property
	
	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub
	'Timer shutoff for polaritycorrect
	Private Function FlipperOn()
		If GameTime < FlipAt+TimeDelay Then
			FlipperOn = True
		End If
	End Function
	
	Public Sub PolarityCorrect(aBall)
		If FlipperOn() Then
			Dim tmp, BallPos, x, IDX, Ycoef
			Ycoef = 1
			
			'y safety Exit
			If aBall.VelY > -8 Then 'ball going down
				RemoveBall aBall
				Exit Sub
			End If
			
			'Find balldata. BallPos = % on Flipper
			For x = 0 To UBound(Balls)
				If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)								'find safety coefficient 'ycoef' data
				End If
			Next
			
			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)												'find safety coefficient 'ycoef' data
			End If
			
			'Velocity correction
			If Not IsEmpty(VelocityIn(0) ) Then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
				
				If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
				
				If Enabled Then aBall.Velx = aBall.Velx*VelCoef
				If Enabled Then aBall.Vely = aBall.Vely*VelCoef
			End If
			
			'Polarity Correction (optional now)
			If Not IsEmpty(PolarityIn(0) ) Then
				Dim AddX
				AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				
				If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
			End If
			If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	Dim x, aCount
	aCount = 0
	ReDim a(UBound(aArray) )
	For x = 0 To UBound(aArray)		'Shuffle objects in a temp array
		If Not IsEmpty(aArray(x) ) Then
			If IsObject(aArray(x)) Then
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	If offset < 0 Then offset = 0
	ReDim aArray(aCount-1+offset)		'Resize original array
	For x = 0 To aCount-1				'set objects back into original array
		If IsObject(a(x)) Then
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
	BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)		'Set up line via two points, no clamping. Input X, output Y
	Dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
	Public Property Let Data(aBall)
		With aBall
			x = .x
			y = .y
			z = .z
			velx = .velx
			vely = .vely
			velz = .velz
			id = .ID
			mass = .mass
			radius = .radius
		End With
	End Property
	Public Sub Reset()
		x = Empty
		y = Empty
		z = Empty
		velx = Empty
		vely = Empty
		velz = Empty
		id = Empty
		mass = Empty
		radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	Dim y 'Y output
	Dim L 'Line
	'find active line
	Dim ii
	For ii = 1 To UBound(xKeyFrame)
		If xInput <= xKeyFrame(ii) Then
			L = ii
			Exit For
		End If
	Next
	If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)		'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )
	
	If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )		 'Clamp lower
	If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )		'Clamp upper
	
	LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b
	Dim gBOT
	gBOT = GetBalls
	
	If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					Exit Sub
				End If
			Next
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
					gBOT(b).velx = gBOT(b).velx / 1.3
					gBOT(b).vely = gBOT(b).vely - 0.5
				End If
			Next
		End If
	Else
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
	End If
End Sub

'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
	dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
	dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		End If
	ElseIf dx = 0 Then
		If dy = 0 Then
			Atn2 = 0
		Else
			Atn2 = Sgn(dy) * pi / 2
		End If
	End If
End Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
	DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
	Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
	
	If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
	Case 0
		SOSRampup = 2.5
	Case 1
		SOSRampup = 6
	Case 2
		SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
'	Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity
	
	Flipper.eostorque = EOST
	Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST * EOSReturn / FReturn
	
	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b, gBOT
		gBOT = GetBalls
		
		For b = 0 To UBound(gBOT)
			If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper
	
	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup
			Flipper.endangle = FEndAngle - 3 * Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
		If FCount = 0 Then FCount = GameTime
		
		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
		If FState <> 3 Then
			Flipper.eostorque = EOST
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If
	End If
End Sub

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle)	'-1 for Right Flipper
	Dim LiveCatchBounce																														'If live catch is not perfect, it won't freeze ball totally
	Dim CatchTime
	CatchTime = GameTime - FCount

	If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
		If CatchTime <= LiveCatch * 0.5 Then												'Perfect catch only when catch time happens in the beginning of the window
			LiveCatchBounce = 0
		Else
			LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)		'Partial catch when catch happens a bit late
		End If
		
		If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
		ball.angmomx = 0
		ball.angmomy = 0
		ball.angmomz = 0
	Else
		If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
	End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************

'******************************************************
' 	ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
	RubbersD.dampen ActiveBall
	TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
	SleevesD.Dampen ActiveBall
	TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD				'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False	'shows info in textbox "TBPout"
RubbersD.Print = False	  'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1		 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967	'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64	   'there's clamping so interpolate up to 56 at least

Dim SleevesD	'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False	'shows info in textbox "TBPout"
SleevesD.Print = False	  'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn   'tbpOut.text
	Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
	End Sub
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If GameTime > 100 Then Report
	End Sub
	
	Public Sub Dampen(aBall)
		If threshold Then
			If BallSpeed(aBall) < threshold Then Exit Sub
		End If
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
		"actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
		If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)
		
		aBall.velx = aBall.velx * coef
		aBall.vely = aBall.vely * coef
		If debugOn Then TBPout.text = str
	End Sub
	
	Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
			aBall.velx = aBall.velx * coef
			aBall.vely = aBall.vely * coef
		End If
	End Sub
	
	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		Dim x
		For x = 0 To UBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
		Next
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
	Public ballvel, ballvelx, ballvely
	
	Private Sub Class_Initialize
		ReDim ballvel(0)
		ReDim ballvelx(0)
		ReDim ballvely(0)
	End Sub
	
	Public Sub Update()	'tracks in-ball-velocity
		Dim str, b, AllBalls, highestID
		allBalls = GetBalls
		
		For Each b In allballs
			If b.id >= HighestID Then highestID = b.id
		Next
		
		If UBound(ballvel) < highestID Then ReDim ballvel(highestID)	'set bounds
		If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)	'set bounds
		If UBound(ballvely) < highestID Then ReDim ballvely(highestID)	'set bounds
		
		For Each b In allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
Sub RDampen_Timer
	Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************

'******************************************************
' 	ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1	  '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7	 'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
	Dim zMultiplier, vel, vratio
	If TargetBouncerEnabled = 1 And aball.z < 30 Then
		'   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		vel = BallSpeed(aBall)
		If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
		Select Case Int(Rnd * 6) + 1
			Case 1
				zMultiplier = 0.2 * defvalue
			Case 2
				zMultiplier = 0.25 * defvalue
			Case 3
				zMultiplier = 0.3 * defvalue
			Case 4
				zMultiplier = 0.4 * defvalue
			Case 5
				zMultiplier = 0.45 * defvalue
			Case 6
				zMultiplier = 0.5 * defvalue
		End Select
		aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
		aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
		aBall.vely = aBall.velx * vratio
		'   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		'   debug.print "conservation check: " & BallSpeed(aBall)/vel
	End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
	TargetBouncer ActiveBall, 1
End Sub

'******************************************************
' 	ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'	 Metals (all metal objects, metal walls, metal posts, metal wire guides)
'	 Apron (the apron walls and plunger wall)
'	 Walls (all wood or plastic walls)
'	 Rollovers (wire rollover triggers, star triggers, or button triggers)
'	 Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'	 Gates (plate gates)
'	 GatesWire (wire gates)
'	 Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
'	- Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
'	- Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
'	- Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
'	- Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1				https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2				https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3				https://youtu.be/ESXWGJZY_EI

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1					  'volume level; range [0, 1]
NudgeLeftSoundLevel = 1				 'volume level; range [0, 1]
NudgeRightSoundLevel = 1				'volume level; range [0, 1]
NudgeCenterSoundLevel = 1			   'volume level; range [0, 1]
StartButtonSoundLevel = 0.1			 'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1			   'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010		'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635		'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0					   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45					'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel		'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel	   'sound helper; not configurable
SlingshotSoundLevel = 0.95					  'volume level; range [0, 1]
BumperSoundFactor = 4.25						'volume multiplier; must not be zero
KnockerSoundLevel = 1						   'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2		  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5			 'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5			   'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.375 / 5			'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025		   'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025		   'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8	  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075				   'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5			'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10	 'volume multiplier; must not be zero
DTSoundLevel = 0.25				 'volume multiplier; must not be zero
RolloverSoundLevel = 0.25		   'volume level; range [0, 1]
SpinnerSoundLevel = 0.5			 'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8				   'volume level; range [0, 1]
BallReleaseSoundLevel = 1			   'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2	'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015	 'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5			 'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
	PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
	PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
	PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
	PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
	PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.y * 2 / tableheight - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 10)
	Else
		AudioFade = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.x * 2 / tablewidth - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 10)
	Else
		AudioPan = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
	Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
	Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
	BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
	PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
	RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
	RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
	PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
	PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
	PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
	PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
	FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
	FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
	FlipperLeftHitParm = parm / 10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
	FlipperRightHitParm = parm / 10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
	PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
	PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
	RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 5 Then
		RandomSoundRubberStrong 1
	End If
	If finalspeed <= 5 Then
		RandomSoundRubberWeak()
	End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
	Select Case Int(Rnd * 10) + 1
		Case 1
			PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 2
			PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 3
			PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 4
			PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 5
			PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 6
			PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 7
			PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 8
			PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 9
			PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 10
			PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
	PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
	RandomSoundWall()
End Sub

Sub RandomSoundWall()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 5) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 5
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 4) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 3) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
	PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
	RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
	RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
	PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
	If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
		RandomSoundBottomArchBallGuideHardHit()
	Else
		RandomSoundBottomArchBallGuide
	End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
	If finalspeed < 6 Then
		PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 10 Then
		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft ActiveBall
	Else
		RandomSoundTargetHitWeak()
	End If
End Sub

Sub Targets_Hit (idx)
	PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
	Select Case Int(Rnd * 9) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 6
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 7
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 8
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 9
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
	PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
	Select Case Int(Rnd * 5) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
	End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
	SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
	SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
	If ActiveBall.velx > 1 Then SoundPlayfieldGate
	StopSound "Arch_L1"
	StopSound "Arch_L2"
	StopSound "Arch_L3"
	StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
	If ActiveBall.velx <  - 8 Then
		RandomSoundRightArch
	End If
End Sub

Sub Arch2_hit()
	If ActiveBall.velx < 1 Then SoundPlayfieldGate
	StopSound "Arch_R1"
	StopSound "Arch_R2"
	StopSound "Arch_R3"
	StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
	If ActiveBall.velx > 10 Then
		RandomSoundLeftArch
	End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
	Select Case scenario
		Case 0
			PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
		Case 1
			PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
	End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim snd
	Select Case Int(Rnd * 7) + 1
		Case 1
			snd = "Ball_Collide_1"
		Case 2
			snd = "Ball_Collide_2"
		Case 3
			snd = "Ball_Collide_3"
		Case 4
			snd = "Ball_Collide_4"
		Case 5
			snd = "Ball_Collide_5"
		Case 6
			snd = "Ball_Collide_6"
		Case 7
			snd = "Ball_Collide_7"
	End Select
	
	PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
	PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
	PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05	  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
	End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
	End Select
End Sub

'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
'	ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
	Dim i
	For i = 0 To tnob
		rolling(i) = False
	Next
End Sub

Sub RollingUpdate()
	Dim b
	Dim gBOT
	gBOT = GetBalls
	
	' stop the sound of deleted balls
	For b = UBound(gBOT) + 1 To tnob - 1
		' Comment the next line if you are not implementing Dyanmic Ball Shadows
		'If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
		rolling(b) = False
		StopSound("BallRoll_" & b)
	Next
	
	' exit the sub if no balls on the table
	If UBound(gBOT) =  - 1 Then Exit Sub
	
	' play the rolling sound for each ball
	For b = 0 To UBound(gBOT)
		If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If
		
		' Ball Drop Sounds
		If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If gBOT(b).velz >  - 7 Then
					RandomSoundBallBouncePlayfieldSoft gBOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard gBOT(b)
				End If
			End If
		End If
		
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If
		
		' "Static" Ball Shadows
		' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
		'If AmbientBallShadowOn = 0 Then
		'	If gBOT(b).Z > 30 Then
		'		BallShadowA(b).height = gBOT(b).z - BallSize / 4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
		'	Else
		'		BallShadowA(b).height = 0.1
		'	End If
		'	BallShadowA(b).Y = gBOT(b).Y + offsetY
		'	BallShadowA(b).X = gBOT(b).X + offsetX
		'	BallShadowA(b).visible = 1
		'End If
	Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
' 	ZRRL: RAMP ROLLING SFX
'******************************************************

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
	Waddball ActiveBall, input
	RampRollUpdate
End Sub

Sub WireRampOff()
	WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
	' This will loop through the RampBalls array checking each element of the array x, position 1
	' To see if the the ball was already added to the array.
	' If the ball is found then exit the subroutine
	Dim x
	For x = 1 To UBound(RampBalls)	'Check, don't add balls twice
		If RampBalls(x, 1) = input.id Then
			If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub	'Frustating issue with BallId 0. Empty variable = 0
		End If
	Next
	
	' This will itterate through the RampBalls Array.
	' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
	' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
	' The RampType(BallId) is set to RampInput
	' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
	For x = 1 To UBound(RampBalls)
		If IsEmpty(RampBalls(x, 1)) Then
			Set RampBalls(x, 0) = input
			RampBalls(x, 1) = input.ID
			RampType(x) = RampInput
			RampBalls(x, 2) = 0
			'exit For
			RampBalls(0,0) = True
			RampRoll.Enabled = 1	 'Turn on timer
			'RampRoll.Interval = RampRoll.Interval 'reset timer
			Exit Sub
		End If
		If x = UBound(RampBalls) Then	 'debug
			Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
			RampBalls(0, 0) & vbNewLine & _
			TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
			TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
			TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
			TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
			TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
			" "
		End If
	Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
	'   Debug.Print "In WRemoveBall() + Remove ball from loop array"
	Dim ballcount
	ballcount = 0
	Dim x
	For x = 1 To UBound(RampBalls)
		If ID = RampBalls(x, 1) Then 'remove ball
			Set RampBalls(x, 0) = Nothing
			RampBalls(x, 1) = Empty
			RampType(x) = Empty
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
		'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
		If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
	Next
	If BallCount = 0 Then RampBalls(0,0) = False	'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
	RampRollUpdate
End Sub

Sub RampRollUpdate()	'Timer update
	Dim x
	For x = 1 To UBound(RampBalls)
		If Not IsEmpty(RampBalls(x,1) ) Then
			If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
				If RampType(x) Then
					PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
					StopSound("wireloop" & x)
				Else
					StopSound("RampLoop" & x)
					PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
				End If
				RampBalls(x, 2) = RampBalls(x, 2) + 1
			Else
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
			End If
			If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then	'if ball is on the PF, remove  it
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
				Wremoveball RampBalls(x,1)
			End If
		Else
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
	Next
	If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()	'debug textbox
	Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
	"1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
	"2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
	"3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
	"4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
	"5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
	"6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
	" "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
	BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
	BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

' Ramp Sounds
Sub MiddleRampStart_Hit
	'ActiveBall.VelY = ActiveBall.VelY * 1.05 ' Ramp Helper for balance
	WireRampOn True
End Sub

Sub LeftRampStart_Hit
	'ActiveBall.VelY = ActiveBall.VelY * 1.05 ' Ramp Helper for balance
	WireRampOn False
End Sub

'************Primitive Swap*******************

CatL.visible = True
CatL2.visible = False

' Initialize the timer
TimerJAG.Interval = 5
TimerJAG.Enabled = False
Dim Timer1Count: Timer1Count = 0

Sub TimerJAG_Timer
	Timer1Count = Timer1Count + 1
	Select Case Timer1Count
		Case 1:  
			CatL.visible = False
			CatL2.visible = True
		Case 100:  ' number of timer intervals to wait before swapping back
			CatL.visible = True
			CatL2.visible = False
			Timer1Count = 0
			TimerJAG.Enabled = False
	End Select
End Sub


Sub Leftslingshot_Hit()
	
End Sub

'***************************************************************************
' VR Plunger Code
'***************************************************************************
Sub TimerVRPlunger_Timer
	If PinCab_Shooter.Y < 100 then
		PinCab_Shooter.Y = PinCab_Shooter.Y + 5
	End If
End Sub

Sub TimerVRPlunger1_Timer
	PinCab_Shooter.Y = 0 + (5* Plunger.Position) - 20
End Sub
