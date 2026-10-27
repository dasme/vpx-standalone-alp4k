' ****************************************************************
'     The Office by Nesta78 based on married with children by phenom mod of:JPSalas
'         Script Básico para juegos EM Script hasta 4 players
'		     usa el core.vbs para funciones extras
'               Apache! - Taito do Brasil 1978
' ****************************************************************

Option Explicit
Randomize

Dim UltraDMD
Const ScoreType = 1 		' 1=UltraDMD   2=PuPlayer DMD     3=PuPlayer PostDMD, if you want to use different dmd's down the road)


' Valores Constantes de las físicas de los flippers - se usan en la creación de las bolas, tienen que cargarse antes del core.vbs
Const BallSize = 50 ' el diametro de la bola. el tamaño normal es 50 unidades de VP
Const BallMass = 1.7  ' la pesadez de la bola, este valor va de acuerdo a la fuerza de los flippers y el plunger

Const UltraDMD_VideoMode_Stretch = 0
Const UltraDMD_VideoMode_Top = 1
Const UltraDMD_VideoMode_Middle = 2
Const UltraDMD_VideoMode_Bottom = 3

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  Player Options
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  

	turnonultradmd = 1 ' change to 1 to turn on ultradmd

' Carga el core.vbs para poder usas sus funciones, sobre todo el vpintimer.addtimer
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Resume Next
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
End Sub

' Valores Constants
Const TableName = "TheOffice"
Const cGameName = "TheOffice"
Const MaxPlayers = 4    ' de 1 a 4
Const BallsPerGame = 5  ' normalmente 3 ó 5
Const Special1 = 100000 ' puntuación a obtener para partida extra
Const Special2 = 500000 ' puntuación a obtener para partida extra
Const Special3 = 960000 ' puntuación a obtener para partida extra

' Variables Globales
Dim turnonultradmd
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim BonusMultiplier
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Special1Awarded(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Score(4)
Dim HighScore
Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add10
Dim Add100
Dim Add1000
Dim Add10000

' Variables de control
Dim BallsOnPlayfield

' Variables de tipo Boolean (verdadero ó falso, True ó False)
Dim bAttractMode
Dim bromconfig
Dim bInstantInfo
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bBallInPlungerLane
Dim bBallSaverActive

Dim x
'****** PuP Variables ******

Dim usePUP: Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!

'*************************** PuP Settings for this table ********************************

usePUP   = true               ' enable Pinup Player functions for this table
cPuPPack = "TheOffice"    ' name of the PuP-Pack / PuPVideos folder for this table

'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

' This is used for the startup and control of Pinup Player

Sub PuPStart(cPuPPack)
    If PUPStatus=true then Exit Sub
    If usePUP=true then
        Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
        If PuPlayer is Nothing Then
            usePUP=false
            PUPStatus=false
        Else
            PuPlayer.B2SInit "",cPuPPack 'start the Pup-Pack
            PUPStatus=true
        End If
    End If
End Sub

Sub pupevent(EventNum)
    if (usePUP=false or PUPStatus=false) then Exit Sub
    PuPlayer.B2SData "E"&EventNum,1  'send event to Pup-Pack
End Sub

' ******* How to use PUPEvent to trigger / control a PuP-Pack *******

' Usage: pupevent(EventNum)

' EventNum = PuP Exxx trigger from the PuP-Pack

' Example: pupevent 102

' This will trigger E102 from the table's PuP-Pack

' DO NOT use any Exxx triggers already used for DOF (if used) to avoid any possible confusion

'************ PuP-Pack Startup **************

PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   ULTRADMD SCRIPTING
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


'UDMD Initialization
Sub LoadUltraDMD 

			If turnonultradmd = 0 then exit sub
			Set UltraDMD = CreateObject("UltraDMD.DMDObject")
			If UltraDMD is Nothing Then
				MsgBox "No UltraDMD found.  This table will NOT run without it."
				Exit Sub
			End If 

            Set UltraDMD = CreateObject("UltraDMD.DMDObject")
            UltraDMD.Init

            ' A Major version change indicates the version is no longer backward compatible
			If turnonultradmd = 0 then exit sub
            If Not UltraDMD.GetMajorVersion = 1 Then
                MsgBox "Incompatible Version of UltraDMD found."
                Exit Sub
            End If

            'A Minor version change indicates new features that are all backward compatible
            If UltraDMD.GetMinorVersion < 0 Then
                MsgBox "Incompatible Version of UltraDMD found.  Please update to version 1.0 or newer."
                Exit Sub
            End If

			Dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
			Dim curDir:curDir = fso.GetAbsolutePathName(".")

			Dim DirName
			DirName = curDir& "\" &TableName& ".UltraDMD"

			If Not fso.FolderExists(DirName) Then _
					Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
			UltraDMD.SetProjectFolder DirName

			' wait for the animation to end
			While UltraDMD.IsRendering = True
			WEnd

            UltraDMD.SetVideoStretchMode UltraDMD_VideoMode_Middle   

End Sub         

Sub TheOfficeIntro

	UltraDMD.DisplayScene00 "Scene1.png", "", 15, "", -1, 14, 5000, 14
	UltraDMD.DisplayScene00 "Scene2.png", "", 15, "", -1, 14, 5000, 14
	UltraDMD.DisplayScene00 "Scene3.png", "", 15, "", -1, 14, 5000, 14
	UltraDMD.DisplayScene00 "Scene4.png", "", 15, "", -1, 14, 5000, 14
	UltraDMD.DisplayScene00 "Scene5.png", "", 15, "", -1, 14, 5000, 14
	UltraDMD.DisplayScene00 "Scene3.5.png", "", 15, "", -1, 14, 5000, 14
	UltraDMD.DisplayScene00 "Scene4.png", "", 15, "", -1, 14, 0, 14

End Sub

Sub Table1_InitSleepWhileRendering 
		
			Do 
				If Not UltraDMD.IsRendering Then 
					Exit Sub 
				End If 
			WScript.Sleep 500 
			Loop Until False
			
End Sub 

'---------------------------ULTRA DMD FUNCTIONS-------------------------------------------

	Sub DMD(background, toptext, bottomtext, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.SetScoreboardBackgroundImage "Main1.jpg", 15, 7
		UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
		UltraDMDTimer.Enabled = 1 'to show the score after the animation/message
	End Sub

	Sub DMDScore
		If turnonultradmd = 0 then exit sub
		UltraDMD.SetScoreboardBackgroundImage "Main1.jpg", 15, 7
		UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
	End Sub

	Sub DMDScoreNow
		If turnonultradmd = 0 then exit sub
		DMDFlush
		DMDScore
	End Sub

	Sub DMDFLush
		If turnonultradmd = 0 then exit sub
		UltraDMDTimer.Enabled = 0
		UltraDMD.CancelRendering
	End Sub

	Sub DMDScrollCredits(background, text, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.ScrollingCredits background, text, 15, 14, duration, 14
	End Sub

	Sub DMDId(id, background, toptext, bottomtext, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.DisplayScene00ExwithID id, False, background, toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
	End Sub

	Sub DMDMod(id, toptext, bottomtext, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
	End Sub

	Sub UltraDMDTimer_Timer 'used for the attrackmode and the instant info.
		If turnonultradmd = 0 then exit sub
		If bInstantInfo Then
			InstantInfo
		ElseIf bAttractMode Then
		ElseIf NOT UltraDMD.IsRendering Then
			DMDScoreNow
		ElseIf bromconfig Then
			romconfig
		End If
	End Sub

Sub DMD_DisplayScene(toptext,bottomtext,animateIn,pauseTime,animateOut)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.DisplayScene00 "", toptext, 13, bottomtext, 15, animateIn, pauseTime, animateOut
	ScoreTimer.enabled = true
End Sub

Sub DMD_DisplaySceneExWithId(id,toptext,bottomtext,animateIn,pauseTime,animateOut)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.DisplayScene00ExWithId id, 0, "", toptext, 13, 9, bottomtext, 15, 9, animateIn, pauseTime, animateOut
	ScoreTimer.enabled = true
End Sub

Sub DMD_ScrollingCredits(text,animateIn,pauseTime,animateOut)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.ScrollingCredits "", text, 15, animateIn, pauseTime, animateOut
	ScoreTimer.enabled = true
End Sub

Sub DMD_ModifyScene(id,toptext,bottomtext)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.ModifyScene00 id, toptext, bottomtext
End Sub

Sub DMD_ModifySceneEx(id,toptext,bottomtext,pauseTime)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.ModifyScene00Ex id, toptext, bottomtext,pauseTime
End Sub

Sub DMDId(id, toptext, bottomtext, duration) 'used in the highscore entry routine
	If not ScoreType = 1 Then Exit Sub: End If
    UltraDMD.DisplayScene00ExwithID id, False, "", toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
End Sub

Sub DMDMod(id, toptext, bottomtext, duration) 'used in the highscore entry routine
	If not ScoreType = 1 Then Exit Sub: End If
    UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
End Sub

Sub DMD_CancelScene(id)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.CancelRenderingWithId id
End Sub

Sub DMD_DisplayhighScene(toptext,bottomtext,animateIn,pauseTime,animateOut)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.DisplayScene00 "", toptext, 13, bottomtext, 15, animateIn, pauseTime, animateOut
	'ScoreTimer.enabled = true
End Sub

Sub DMD_CancelRendering()
	scoreupdate = true
	addscore(0)
	If not ScoreType = 1 Then Exit Sub: End If
	UltraDMD.CancelRendering
End Sub
'------------------------------------------------------------------------------------------------

Function CL(id, NumString) 'center line
    Dim Temp, TempStr
	NumString = LEFT(NumString, dCharsPerLine(id))
    Temp = (dCharsPerLine(id)- Len(NumString)) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function


Sub Table1_Init()
	LoadUltraDMD
	TheOfficeIntro
	'UltraDMDIntro
	UltraDMD.DisplayScene00 "Main1.png", "", 15, "", -1, 14, 0, 14

    ' Inicializar diversos objetos de la mesa, como droptargets, animations...
    VPObjects_Init
    LoadEM

    ' Carga los valores grabados highscore y créditos
    Loadhs
    'HighscoreReel.SetValue Highscore
    ' también pon el highscore en el reel del primer jugador
    ' ScoreReel1.SetValue Highscore
    ' en esta mesa el attractmode se encarga de mostrar el highscore

    UpdateCredits

    ' Juego libre o con monedas: si es True entonces no se usarán monedas
    bFreePlay = True 'queremos monedas

    ' Inicialiar las variables globales de la mesa
    bAttractMode = False
    bOnTheFirstBall = False
    bGameInPlay = False
    bBallInPlungerLane = False
    BallsOnPlayfield = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bJustStarted = True
    Add10 = 0
    Add100 = 0
    Add1000 = 0

    ' pone la mesa en modo de espera
    EndOfGame

    'Enciende las luces GI despues de un segundo
    vpmtimer.addtimer 1500, "GiOn '"

    ' quita reels
    If Table1.ShowDT = true then
        For each x in aReels
            x.Visible = 1
        Next
    else
        For each x in aReels
            x.Visible = 0
        Next
    end if

    ' Arranca el Game Timer de las animaciones, sonido de las bolas rodando, etc
    GameTimer.Enabled = 1

	If turnonultradmd > 0 then 
		UltraDMD.SetScoreboardBackgroundImage "Main1.jpg", 15, 12
	end if

End Sub

' Start the random music 
    MusicOn

Sub MusicOn
    Dim x
    x = INT(6 * RND(1) )
    Select Case x
       Case 0:PlayMusic "theoffice1.mp3"
       Case 1:PlayMusic "theoffice2.mp3"
       Case 2:PlayMusic "theoffice3.mp3"
       Case 3:PlayMusic "theoffice4.mp3"
       Case 4:PlayMusic "theoffice5.mp3"
       Case 5:PlayMusic "theoffice6.mp3"
    End Select
 End Sub
 
 Sub Table1_MusicDone()
    MusicOn
 End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    ' añade monedas
    If Keycode = AddCreditKey Then
        If(Tilted = False)Then
            AddCredits 1
            PlaySound "fx_coin"
            PlaySound "coin"
        End If
    End If

    ' el plunger
    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

' iluminación
		If keycode = LeftMagnaSave Then bLutActive = True
	If keycode = RightMagnaSave Then 
		If bLutActive Then NextLUT: End If
	End If

    ' teclas de la falta
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

    ' Funcionamiento normal de los flipers y otras teclas durante el juego

    If bGameInPlay AND NOT Tilted Then

        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        ' tecla de empezar el juego
        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    PlaySound "start" 
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
                        PlaySound "start" 
                    Else
                        ' no hay suficientes créditos para empezar el juego.
                        PlaySound "10000"
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                If(bFreePlay = True)Then
                    If(BallsOnPlayfield = 0)Then
                        ResetScores
                        ResetForNewGame()
                        PlaySound "start"
                    End If
                Else
                    If(Credits > 0)Then
                        If(BallsOnPlayfield = 0)Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
                            PlaySound "start" :pupevent 800
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        PlaySound "10000"
                    End If
                End If
            End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = LeftMagnaSave Then bLutActive = False

    If bGameInPlay AND NOT Tilted Then
        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 0
        If keycode = RightFlipperKey Then SolRFlipper 0
    End If

    If keycode = PlungerKey Then
        Plunger.Fire
        If bBallInPlungerLane Then
            PlaySoundAt "fx_plunger", plunger
        Else
            PlaySoundAt "fx_plunger_empty", plunger
        End If
    End If
End Sub

' LUT

Dim bLutActive, LUTImage
Sub LoadLUT
	bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
	UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 9: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: table1.ColorGradeImage = "LUT0"
Case 1: table1.ColorGradeImage = "LUT1"
Case 2: table1.ColorGradeImage = "LUT2"
Case 3: table1.ColorGradeImage = "LUT3"
Case 4: table1.ColorGradeImage = "LUT4"
Case 5: table1.ColorGradeImage = "LUT5"
Case 6: table1.ColorGradeImage = "LUT6"
Case 7: table1.ColorGradeImage = "LUT7"
Case 8: table1.ColorGradeImage = "LUT8"
End Select
End Sub

'*************
' Para la mesa
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
    Savehs
    If B2SOn Then Controller.Stop
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt "fx_flipperup", LeftFlipper
        LeftFlipper.EOSTorque = 0.75:LeftFlipper.RotateToEnd
    Else
        PlaySoundAt "fx_flipperdown", LeftFlipper
        LeftFlipper.EOSTorque = 0.1:LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt "fx_flipperup", RightFlipper
        RightFlipper.EOSTorque = 0.75:RightFlipper.RotateToEnd
    Else
        PlaySoundAt "fx_flipperdown", RightFlipper
        RightFlipper.EOSTorque = 0.1:RightFlipper.RotateToStart
    End If
End Sub


'*******************
' Luces GI
'*******************

Sub GiOn 'enciende las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    If B2SOn Then Controller.B2SSetData 60, 1
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    If B2SOn Then Controller.B2SSetData 60, 0
End Sub

'**************
' TILT - Falta
'**************

'el "timer" TiltDecreaseTimer resta .01 de la variable "Tilt" cada ronda

Sub CheckTilt                     'esta rutina se llama cada vez que das un golpe a la mesa
    Tilt = Tilt + TiltSensitivity 'añade un valor al contador "Tilt"
    TiltDecreaseTimer.Enabled = True
    If Tilt > 15 Then             'Si la variable "Tilt" es más de 15 entonces haz falta
        Tilted = True
        TiltReel.SetValue 1       'muestra Tilt en la pantalla
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'empieza una pausa a fin de que todas las bolas se cuelen
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
        'Apaga todas las luces Gi de la mesa
        GiOff

        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        Bumper1.Force = 0
        Bumper2.Force = 0
        Bumper3.Force = 0
    Else
        'enciende de nuevo todas las luces GI
        GiOn
        Bumper1.Force = 10
        Bumper2.Force = 10
        Bumper3.Force = 10
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' si todas las bolas se han colado, entonces ..
    If(BallsOnPlayfield = 0)Then
        '... haz el fin de bola normal
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' de lo contrario la rutina mirará si todavía hay bolas en la mesa
End Sub



'***************************************
' Sonidos de las colecciones de objetos
' como metales, gomas, plásticos, etc
'***************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

'************************************************************************************************************************
' Solo para VPX 10.2 y posteriores.
' FlashForMs hará parpadear una luz o un flash por unos milisegundos "TotalPeriod" cada tantos milisegundos "BlinkPeriod"
' Cuando el "TotalPeriod" haya terminado, la luz o el flasher se pondrá en el estado especificado por el valor "FinalState"
' El valor de "FinalState" puede ser: 0=apagado, 1=encendido, 2=regreso al estado anterior
'************************************************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState)

    If TypeName(MyLight) = "Light" Then ' la luz es del tipo "light"

        If FinalState = 2 Then
            FinalState = MyLight.State  'guarda el estado actual de la luz
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then ' la luz es del tipo "flash"
        Dim steps
        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'número de encendidos y apagados que hay que ejecutar
        If FinalState = 2 Then                      'guarda el estado actual del flash
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'guarda el número de parpadeos

        ' empieza los parpadeos y crea la rutina que se va a ejecutar como un timer que se va a ejecutar los parpadeos
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

'****************************************
' Inicializa la mesa para un juego nuevo
'****************************************

Sub ResetForNewGame()
    'debug.print "ResetForNewGame"
    Dim i

    bGameInPLay = True
    bBallSaverActive = False

    'pone a cero los marcadores y apaga las luces de espera.
    StopAttractMode
    If B2SOn then
        Controller.B2SSetGameOver 0
    end if
    ' enciende las luces GI si estuvieran apagadas
    GiOn

    MatchWheel.Enabled = true

    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        ExtraBallsAwards(i) = 0
        Special1Awarded(i) = False
        Special2Awarded(i) = False
        Special3Awarded(i) = False
        BallsRemaining(i) = BallsPerGame
    Next
    Bonus = 0
    BonusMultiplier = 1
    UpdateBallInPlay

    Clear_Match

    ' inicializa otras variables
    Tilt = 0

    ' inicializa las variables del juego
    Game_Init()

    ' ahora puedes empezar una música si quieres
    ' empieza la rutina "Firstball" despues de una pequeña pausa
    vpmtimer.addtimer 2000, "FirstBall '"
End Sub

' esta pausa es para que la mesa tenga tiempo de poner los marcadores a cero y actualizar las luces

Sub FirstBall
    'debug.print "FirstBall"
    ' ajusta la mesa para una bola nueva, sube las dianas abatibles, etc
    ResetForNewPlayerBall()
    ' crea una bola nueva en la zona del plunger
    CreateNewBall()
End Sub

' (Re-)inicializa la mesa para una bola nueva, tanto si has perdido la bola, oe le toca el turno al otro jugador

Sub ResetForNewPlayerBall()
    'debug.print "ResetForNewPlayerBall"
    ' Se asegura que los marcadores están activados para el jugador de turno
    AddScore 0

    ' ajusta el multiplicador del bonus multiplier a 1X (si hubiese multiplicador en la mesa)

    ' enciende las luces, reinicializa las variables del juego, etc
    bExtraBallWonThisBall = False
    ResetNewBallLights
    ResetNewBallVariables
End Sub

' Crea una bola nueva en la mesa

Sub CreateNewBall()
    ' crea una bola nueva basada en el tamaño y la masa de la bola especificados al principio del script
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' incrementa el número de bolas en el tablero, ya que hay que contarlas
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' actualiza las luces del backdrop
    UpdateBallInPlay

    ' y expulsa la bola
    PlaySoundAt "fx_Ballrel", BallRelease :MusicOn
    BallRelease.Kick 90, 4
End Sub

' El jugador ha perdido su bola, y ya no hay más bolas en juego
' Empieza a contar los bonos

Sub EndOfBall()
    'debug.print "EndOfBall"
    Dim AwardPoints, ii
    AwardPoints = 0
    ' La primera se ha perdido. Desde aquí ya no se puede aceptar más jugadores
    bOnTheFirstBall = False

    ' solo recoge los bonos si no hay falta
    ' el sistema del la falta se encargará de nuevas bolas o del fin de la partida

    If NOT Tilted Then
        BonusCountTimer.Interval = 200
        CreditLight.State = 2 'taito tables
        BonusCountTimer.Enabled = 1
    Else                      'Si hay falta simplemente espera un momento y va directo a la segunta parte después de perder la bola
        vpmtimer.addtimer 400, "EndOfBall2 '"
    End If
End Sub

Sub BonusCountTimer_Timer 'añade los bonos y actualiza las luces
    'debug.print "BonusCount_Timer"
    If Bonus > 0 Then
        Bonus = Bonus -1
        AddScore 1000
        UpdateBonusLights
    Else
        ' termina la cuenta de los bonos y continúa con el fin de bola
        CreditLight.State = 0 'Taito tables
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 1000, "EndOfBall2 '"
    End If
End Sub

' La cuenta de los bonos ha terminado. Mira si el jugador ha ganado bolas extras
' y si no mira si es el último jugador o la última bola
'
Sub EndOfBall2()
    'debug.print "EndOfBall2"
    ' si hubiese falta, quítala, y pon la cuenta a cero de la falta para el próximo jugador, ó bola

    Tilted = False
    Tilt = 0
    TiltReel.SetValue 0
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False 'activa de nuevo los bumpers y los slingshots

    ' ¿ha ganado el jugador una bola extra?
    If(ExtraBallsAwards(CurrentPlayer) > 0)Then
        'debug.print "Extra Ball"

        ' sí? entonces se la das al jugador
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' si no hay más bolas apaga la luz de jugar de nuevo
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            LightShootAgain.State = 0
            If B2SOn then
                Controller.B2SSetShootAgain 0
            end if
        End If

' aquí se podría poner algún sonido de bola extra o alguna luz que parpadee

' En esta mesa hacemos la bola extra igual como si fuese la siguente bola, haciendo un reset de las variables y dianas
        ResetForNewPlayerBall()

        ' creamos una bola nueva en el pasillo de disparo
        CreateNewBall()
    Else ' no hay bolas extras

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' ¿Es ésta la última bola?
        If(BallsRemaining(CurrentPlayer) <= 0)Then

            ' miramos si la puntuación clasifica como el Highscore
            CheckHighScore()
        End If

        ' ésta no es la última bola para éste jugador
        ' y si hay más de un jugador continúa con el siguente
        EndOfBallComplete()
    End If
End Sub

' Esta rutina se llama al final de la cuenta del bonus
' y pasa a la siguente bola o al siguente jugador
'
Sub EndOfBallComplete()
    'debug.print "EndOfBallComplete"
    Dim NextPlayer

    'debug.print "EndOfBall - Complete"

    ' ¿hay otros jugadores?
    If(PlayersPlayingGame > 1)Then
        ' entonces pasa al siguente jugador
        NextPlayer = CurrentPlayer + 1
        ' ¿vamos a pasar del último jugador al primero?
        ' (por ejemplo del jugador 4 al no. 1)
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' ¿Hemos llegado al final del juego? (todas las bolas se han jugado de todos los jugadores)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then

        ' aquí se empieza la lotería, normalmente cuando se juega con monedas
        If bFreePlay = False Then
            Verification_Match
        End If

        ' ahora se pone la mesa en el modo de final de juego
        EndOfGame()
    Else
        ' pasamos al siguente jugador
        CurrentPlayer = NextPlayer

        ' nos aseguramos de que el backdrop muestra el jugador actual
        AddScore 0

        ' hacemos un reset del la mesa para el siguente jugador (ó bola)
        ResetForNewPlayerBall()

        ' y sacamos una bola
        CreateNewBall()
    End If
End Sub

' Esta función se llama al final del juego

Sub EndOfGame()
    'debug.print "EndOfGame"
    bGameInPLay = False
    bJustStarted = False
    If B2SOn then
        Controller.B2SSetGameOver 1 
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetCanPlay 0
    end if
    ' asegúrate de que los flippers están en modo de reposo
    SolLFlipper 0
    SolRFlipper 0

    ' pon las luces en el modo de fin de juego
    StartAttractMode
End Sub

' Esta función calcula el no de bolas que quedan
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' Esta función calcula el Highscore y te da una partida gratis si has conseguido el Highscore
Sub CheckHighscore :pupevent 803
    For x = 1 to PlayersPlayingGame
        If Score(x) > Highscore Then 
            Highscore = Score(x) :pupevent 804
            PlaySound "fx_knocker"
            'HighscoreReel.SetValue Highscore
            AddCredits 1
        End If
    Next
End Sub

'******************
'  Match - Loteria
'******************

Dim MatchCounter
MatchCounter = 0

sub MatchWheel_timer
    MatchCounter = (MatchCounter + 1)MOD 10

    If B2SOn then
        If MatchCounter = 0 Then
            Controller.B2SSetMatch 100
        Else
            Controller.B2SSetMatch MatchCounter * 10
        End If
    end if

    For each x in aMatchLights
        x.state = 0
    next
    aMatchLights(MatchCounter).state = 1
End Sub

Sub Verification_Match()
    MatchWheel.Enabled = False
    'PlaySound "fx_match"
    Match = MatchCounter * 10
    debug.print match
    If(Score(CurrentPlayer)MOD 100) = Match Then
        PlaySound "fx_knocker"
        AddCredits 1
    End If
End Sub

Sub Clear_Match()
    For each x in aMatchLights
        x.State = 0
    Next
    If B2SOn then
        Controller.B2SSetMatch 0
    end if
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' has perdido la bola ;-( mira cuantas bolas hay en el tablero.
' si solamente hay una entonces reduce el número de bola y mira si es la última para finalizar el juego
' si hay más de una, significa que hay multiball, entonces continua con la partida
'
Sub Drain_Hit() :pupevent 801
    ' destruye la bola
    Drain.DestroyBall

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' haz sonar el ruido de la bola
    PlaySoundAt "fx_drain", Drain :EndMusic

    'si hay falta el systema de tilt se encargará de continuar con la siguente bola/jugador
    If Tilted Then
        Exit Sub
    End If

    ' si estás jugando y no hay falta
    If(bGameInPLay = True)AND(Tilted = False)Then

        ' ¿está el salva bolas activado?
        If(bBallSaverActive = True)Then

            ' ¿sí?, pues creamos una bola
            CreateNewBall()
        Else
            ' ¿es ésta la última bola en juego?
            If(BallsOnPlayfield = 0)Then
                vpmtimer.addtimer 1000, "EndOfBall '" 'hacemos una pequeña pausa anter de continuar con el fin de bola
                Exit Sub
            End If
        End If
    End If
End Sub

Sub swPlungerRest_Hit()
    bBallInPlungerLane = True
End Sub

' La bola ha sido disparada, así que cambiamos la variable, que en esta mesa se usa solo para que el sonido del disparador cambie según hay allí una bola o no
' En otras mesas podrá usarse para poner en marcha un contador para salvar la bola

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
End Sub

' *********************************************************************
'               Funciones para la cuenta de los puntos
' *********************************************************************

' Añade puntos al jugador, hace sonar las campanas y actualiza el backdrop

Sub AddScore(Points)
    If Tilted Then Exit Sub
            ' añade los puntos a la variable del actual jugador
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
			DMDScore
            ' actualiza los contadores
            UpdateScore Points

            ' hace sonar las campanillas de acuerdo a los puntos obtenidos
    Select Case Points
        Case 10, 100, 1000, 10000

            If Points = 1000 AND(Score(CurrentPlayer)MOD 10000) \ 1000 = 0 Then  'nuevo reel de 10000
                PlaySound "10000"
            ElseIf Points = 100 AND(Score(CurrentPlayer)MOD 1000) \ 100 = 0 Then 'nuevo reel de 1000
                PlaySound "1000"
            ElseIf Points = 10 AND(Score(CurrentPlayer)MOD 100) \ 10 = 0 Then    'nuevo reel de 100
                PlaySound "100"
            Else
                PlaySound Points
            End If
        Case 20, 30, 40, 50, 60, 70, 80, 90
            Add10 = Add10 + Points \ 10
            AddScore10Timer.Enabled = TRUE
        Case 200, 300, 400, 500, 600, 700, 800, 900
            Add100 = Add100 + Points \ 100
            AddScore100Timer.Enabled = TRUE
        Case 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000
            Add1000 = Add1000 + Points \ 1000
            AddScore1000Timer.Enabled = TRUE
        Case 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000
            Add10000 = Add10000 + Points \ 10000
            AddScore10000Timer.Enabled = TRUE
    End Select

    ' ' aquí se puede hacer un chequeo si el jugador ha ganado alguna puntuación alta y darle un crédito ó bola extra
    If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special1Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special2Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special3Awarded(CurrentPlayer) = True
    End If
End Sub

'************************************
'TIMER DE 10, 100, 1000 y 1000 PUNTOS
'************************************

' hace sonar las campanillas según los puntos

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddScore 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddScore 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000 > 0 then
        AddScore 1000
        Add1000 = Add1000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore10000Timer_Timer()
    if Add10000 > 0 then
        AddScore 10000
        Add10000 = Add10000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'*******************
'     BONOS
'*******************

Sub AddBonus(bonuspoints)
    If(Tilted = False)Then
        ' añade los bonos al jugador actual
        Bonus = Bonus + bonuspoints
        If Bonus > 19 Then
            Bonus = 19
        End If
        ' actualiza las luces
        UpdateBonusLights
    End if
End Sub

Sub UpdateBonusLights 'enciende o apaga las luces de los bonos según la variable "Bonus"
    Select Case Bonus
        Case 0:litotem.State = 0:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 0:li5.State = 0:li10.State = 0:lifull.State = 0
        Case 1:litotem.State = 1:li1.State = 1:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 0:li5.State = 0:li10.State = 0:lifull.State = 0
        Case 2:litotem.State = 1:li1.State = 0:li2.State = 1:li3.State = 0:li4.State = 0:liplus.State = 0:li5.State = 0:li10.State = 0:lifull.State = 0
        Case 3:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 1:li4.State = 0:liplus.State = 0:li5.State = 0:li10.State = 0:lifull.State = 0
        Case 4:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 1:liplus.State = 0:li5.State = 0:li10.State = 0:lifull.State = 0
        Case 5:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 0:li5.State = 1:li10.State = 0:lifull.State = 0
        Case 6:litotem.State = 1:li1.State = 1:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 0:li5.State = 1:li10.State = 0:lifull.State = 0
        Case 7:litotem.State = 1:li1.State = 0:li2.State = 1:li3.State = 0:li4.State = 0:liplus.State = 1:li5.State = 1:li10.State = 0:lifull.State = 0
        Case 8:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 1:li4.State = 0:liplus.State = 1:li5.State = 1:li10.State = 0:lifull.State = 0
        Case 9:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 1:liplus.State = 1:li5.State = 1:li10.State = 0:lifull.State = 0
        Case 10:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 0:li5.State = 0:li10.State = 1:lifull.State = 0
        Case 11:litotem.State = 1:li1.State = 1:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 1:li5.State = 0:li10.State = 1:lifull.State = 0
        Case 12:litotem.State = 1:li1.State = 0:li2.State = 1:li3.State = 0:li4.State = 0:liplus.State = 1:li5.State = 0:li10.State = 1:lifull.State = 0
        Case 13:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 1:li4.State = 0:liplus.State = 1:li5.State = 0:li10.State = 1:lifull.State = 0
        Case 14:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 1:liplus.State = 1:li5.State = 0:li10.State = 1:lifull.State = 0
        Case 15:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 1:li5.State = 1:li10.State = 1:lifull.State = 0
        Case 16:litotem.State = 1:li1.State = 1:li2.State = 0:li3.State = 0:li4.State = 0:liplus.State = 1:li5.State = 1:li10.State = 1:lifull.State = 0
        Case 17:litotem.State = 1:li1.State = 0:li2.State = 1:li3.State = 0:li4.State = 0:liplus.State = 1:li5.State = 1:li10.State = 1:lifull.State = 0
        Case 18:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 1:li4.State = 0:liplus.State = 1:li5.State = 1:li10.State = 1:lifull.State = 0
        Case 19:litotem.State = 1:li1.State = 0:li2.State = 0:li3.State = 0:li4.State = 1:liplus.State = 1:li5.State = 1:li10.State = 1:lifull.State = 1
    End Select
End Sub

'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuación del jugador

Sub UpdateScore(playerpoints)
    Select Case CurrentPlayer
        Case 1:ScoreReel1.Addvalue playerpoints
        Case 2:ScoreReel2.Addvalue playerpoints
        Case 3:ScoreReel3.Addvalue playerpoints
        Case 4:ScoreReel4.Addvalue playerpoints
    End Select
    If B2SOn then
        Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
        If Score(CurrentPlayer) >= 100000 then
            Controller.B2SSetScoreRollover 24 + CurrentPlayer, 1
        end if
    end if
End Sub

' pone todos los marcadores a 0
Sub ResetScores
    ScoreReel1.ResetToZero
    ScoreReel2.ResetToZero
    ScoreReel3.ResetToZero
    ScoreReel4.ResetToZero
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScorePlayer2 0
        Controller.B2SSetScorePlayer3 0
        Controller.B2SSetScorePlayer4 0
        Controller.B2SSetScoreRolloverPlayer1 0
        Controller.B2SSetScoreRolloverPlayer2 0
        Controller.B2SSetScoreRolloverPlayer3 0
        Controller.B2SSetScoreRolloverPlayer4 0
    end if
End Sub

Sub AddCredits(value)
    If Credits < 9 Then
        Credits = Credits + value
        CreditReel.SetValue Credits
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits > 0 Then 'this is for Bally tables
    'CreditLight.State = 1
    Else
    'CreditLight.State = 0
    End If
    CreditReel.SetValue credits
    If B2SOn then
        If Credits = 0 then
            Controller.B2SSetData 29, 10
        Else
            Controller.B2SSetData 29, Credits
        End If
    End If
End Sub

Sub UpdateBallInPlay 'actualiza los marcadores de las bolas, el número de jugador y el número total de jugadores
    Select Case Balls
        Case 1:bip1.State = 1:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0
        Case 2:bip1.State = 0:bip2.State = 1:bip3.State = 0:bip4.State = 0:bip5.State = 0
        Case 3:bip1.State = 0:bip2.State = 0:bip3.State = 1:bip4.State = 0:bip5.State = 0
        Case 4:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 1:bip5.State = 0
        Case 5:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetBallInPlay Balls
    end if
    Select Case CurrentPlayer
        Case 1:pl1.State = 1:pl2.State = 0:pl3.State = 0:pl4.State = 0
        Case 2:pl1.State = 0:pl2.State = 1:pl3.State = 0:pl4.State = 0
        Case 3:pl1.State = 0:pl2.State = 0:pl3.State = 1:pl4.State = 0
        Case 4:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetPlayerUp CurrentPlayer
    end if
    Select Case PlayersPlayingGame
        Case 1:cp1.State = 1:cp2.State = 0:cp3.State = 0:cp4.State = 0
        Case 2:cp1.State = 0:cp2.State = 1:cp3.State = 0:cp4.State = 0
        Case 3:cp1.State = 0:cp2.State = 0:cp3.State = 1:cp4.State = 0
        Case 4:cp1.State = 0:cp2.State = 0:cp3.State = 0:cp4.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetCanPlay PlayersPlayingGame
        Controller.B2SSetCanPlay PlayersPlayingGame
    end if
End Sub

'*************************
' Partidas y bolas extras
'*************************

Sub AwardExtraBall() :pupevent 802
    If NOT bExtraBallWonThisBall Then
        PlaySound "fx_knocker"
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1
        If B2SOn then
            Controller.B2SSetShootAgain 1
        end if
    END If
End Sub

Sub AwardSpecial()
    PlaySound "fx_knocker"
    AddCredits 1
End Sub

' ********************************
'        Attract Mode
' ********************************
' las luces simplemente parpadean de acuerdo a los valores que hemos puesto en el "Blink Pattern" de cada luz

Sub StartAttractMode()
    bAttractMode = True
    For each x in aLights
        x.State = 2
    Next

    ' enciente la luz de fin de partida
	UltraDMDTimer.Enabled = 1
    GameOverR.SetValue 1
    ' empieza un timer
    AttractTimer.Enabled = 1
End Sub

Sub StopAttractMode()
    bAttractMode = False
    ResetScores
    For each x in aLights
        x.State = 0
    Next

    ' apaga la luz de fin de partida
    GameOverR.SetValue 0
    ' apaga timer
    AttractTimer.Enabled = 0
End Sub

Dim AttractStep
AttractStep = 0

Sub AttractTimer_Timer
    Select Case AttractStep
        Case 0
            AttractStep = 1
            ScoreReel1.SetValue Score(1)
            ScoreReel2.SetValue Score(2)
            ScoreReel3.SetValue Score(3)
            ScoreReel4.SetValue Score(4)
            If B2SOn then
                Controller.B2SSetScorePlayer1 Score(1)
                Controller.B2SSetScorePlayer2 Score(2)
                Controller.B2SSetScorePlayer3 Score(3)
                Controller.B2SSetScorePlayer4 Score(4)
            end if
        Case 1
            AttractStep = 0
            ScoreReel1.SetValue Highscore
            ScoreReel2.SetValue Highscore
            ScoreReel3.SetValue Highscore
            ScoreReel4.SetValue Highscore
            If B2SOn then
                Controller.B2SSetScorePlayer1 Highscore
                Controller.B2SSetScorePlayer2 Highscore
                Controller.B2SSetScorePlayer3 Highscore
                Controller.B2SSetScorePlayer4 Highscore
            end if
    End Select
End Sub

'************************************************
'    Load (cargar) / Save (guardar)/ Highscore
'************************************************

' solamente guardamos el número de créditos y la puntuación más alta

Sub Loadhs
    x = LoadValue(TableName, "HighScore")
    If(x <> "")Then HighScore = CDbl(x)Else HighScore = 0 End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If
End Sub

Sub Savehs
    SaveValue TableName, "HighScore", HighScore
    SaveValue TableName, "Credits", Credits
End Sub

' por si se necesitara quitar la actual puntuación más alta, se le puede poner a una tecla,
' o simplemente abres la ventana de debug y escribes Reseths y le das al enter
Sub Reseths
    HighScore = 0
    Savehs
End Sub

'****************************************
' Actualizaciones en tiempo real
'****************************************
' se usa sobre todo para hacer animaciones o sonidos que cambian en tiempo real
' como por ejemplo para sincronizar los flipers, puertas ó molinillos con primitivas

Sub GameTimer_Timer
'    RollingUpdate 'actualiza el sonido de la bola rodando
    ' y también algunas animaciones, sobre todo de primitivas
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
End Sub

'***********************************************************************
' *********************************************************************
'            Aquí empieza el código particular a la mesa
' (hasta ahora todas las rutinas han sido muy generales para todas las mesas)
' (y hay muy pocas rutinas que necesitan cambiar de mesa a mesa)
' *********************************************************************
'***********************************************************************

' se inicia las dianas abatibles, primitivas, etc.
' aunque en el VPX no hay muchos objetos que necesitan ser iniciados

Sub VPObjects_Init
End Sub

' variables de la mesa

Sub Game_Init() 'esta rutina se llama al principio de un nuevo juego
    Dim i

    'Empezar alguna música, si hubiera música en esta mesa

    'iniciar variables, en esta mesa hay muy pocas variables ya que usamos las luces, y el UserValue de las dianas

    'iniciar algún timer

    'Iniciar algunas luces, en esta mesa son las mismas luces que las de una bola nueva
    TurnOffPlayfieldLights()
End Sub

Sub ResetNewBallVariables() 'inicia las variable para una bola ó jugador nuevo
    ResetLeftDroptargets
    vpmtimer.addtimer 300, "ResetRightDropTargets '"
End Sub

Sub ResetNewBallLights() 'enciende ó apaga las luces para una bola nueva
    TurnOffPlayfieldLights()
    UpperLightA.State = 1
    UpperLightB.State = 1
    UpperLightC.State = 1
    LeftInlaneLightA.State = 1
    LeftInlaneLightB.State = 1
    RightInlaneLightB.State = 1
    RightInlaneLightC.State = 1
    TargetLeftLight1.State = 1
    TargetRightLight1.State = 1
    LeftTarget500.State = 1
    RightTarget500.State = 1
End Sub

Sub TurnOffPlayfieldLights()
    For each x in aLights
        x.State = 0
    Next
End Sub

' *********************************************************************
'       Eventos de la mesa - choque de la bola contra dianas
'
' En cada diana u objeto con el que la bola choque habrá que hacer:
' - sonar un sonido físico
' - hacer algún movimiento, si es necesario
' - añadir alguna puntuación
' - encender/apagar una luz
' - hacer algún chequeo para ver si el jugador ha completado algo
' *********************************************************************

'**************************
'   Gomas de 10 puntos
'**************************

' animation rubbers
Dim Rub1, Rub2

Sub rlband001_Hit
    If NOT Tilted Then
        AddScore 10
    End If
    Rub1 = 1:rlband001_Timer
End Sub

Sub rlband001_Timer
    Select Case Rub1
        Case 1:r009.Visible = 0:r022.Visible = 1:rlband001.TimerEnabled = 1
        Case 2:r022.Visible = 0:r023.Visible = 1
        Case 3:r023.Visible = 0:r009.Visible = 1:rlband001.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub rlband002_Hit
    If NOT Tilted Then
        AddScore 10
    End If
    Rub2 = 1:rlband002_Timer
End Sub

Sub rlband002_Timer
    Select Case Rub2
        Case 1:r010.Visible = 0:r024.Visible = 1:rlband002.TimerEnabled = 1
        Case 2:r024.Visible = 0:r025.Visible = 1
        Case 3:r025.Visible = 0:r010.Visible = 1:rlband002.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub rlband003_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

Sub rlband004_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

'*********
' Bumpers
'*********

Sub Bumper1_Hit :pupevent 805
    If Tilted Then Exit Sub
    PlaySoundAt "fx_Bumper", bumper1
    ' añade algunos puntos
    AddScore 100 + 900 * LiBumper1a.State
End Sub

Sub LitBumper1:LiBumper1a.State = 1:liBumper1b.State = 1:End Sub

Sub Bumper2_Hit :pupevent 806
    If Tilted Then Exit Sub
    PlaySoundAt "fx_Bumper", bumper2
    ' añade algunos puntos
    AddScore 100 + 900 * LiBumper2a.State
End Sub

Sub LitBumper2:LiBumper2a.State = 1:liBumper2b.State = 1:End Sub

Sub Bumper3_Hit :pupevent 807
    dim tmp
    If Tilted Then Exit Sub
    PlaySoundAt "fx_Bumper", bumper3
    ' añade algunos puntos
    AddScore 100 + 900 * LiBumper3a.State
    ' rotates lights
    tmp = TargetLeftLight1.State
    TargetLeftLight1.State = TargetLeftLight2.State
    TargetLeftLight2.State = tmp

    tmp = TargetRightLight1.State
    TargetRightLight1.State = TargetRightLight2.State
    TargetRightLight2.State = tmp

    tmp = LeftSpecialLight.State
    LeftSpecialLight.State = RightSpecialLight.State
    RightSpecialLight.State = tmp
End Sub

Sub LitBumper3:LiBumper3a.State = 1:liBumper3b.State = 1:End Sub

'******************************
'     Pasillos inferiores
'******************************

Sub sw001_Hit
    PlaySoundAt "fx_sensor", sw001
    If Tilted Then Exit Sub
    AddScore 5000
End Sub

Sub sw002_Hit
    PlaySoundAt "fx_sensor", sw002
    If Tilted Then Exit Sub
    AddScore 200
    LeftInlaneLightA.State = 0
    UpperLightA.State = 0
    CenterLightA.State = 1
    LitBumper1
    CheckABC
End Sub

Sub sw003_Hit
    PlaySoundAt "fx_sensor", sw003
    If Tilted Then Exit Sub
    AddScore 500
    LeftInlaneLightB.State = 0
    RightInlaneLightB.State = 0
    UpperLightB.State = 0
    CenterLightB.State = 1
    LitBumper3
    CheckABC
End Sub

Sub sw004_Hit
    PlaySoundAt "fx_sensor", sw004
    If Tilted Then Exit Sub
    AddScore 500
    LeftInlaneLightB.State = 0
    RightInlaneLightB.State = 0
    UpperLightB.State = 0
    CenterLightB.State = 1
    LitBumper3
    CheckABC
End Sub

Sub sw005_Hit
    PlaySoundAt "fx_sensor", sw005
    If Tilted Then Exit Sub
    AddScore 500
    RightInlaneLightC.State = 0
    UpperLightC.State = 0
    CenterLightC.State = 1
    LitBumper2
    CheckABC
End Sub

Sub sw006_Hit
    PlaySoundAt "fx_sensor", sw006
    If Tilted Then Exit Sub
    AddScore 5000
End Sub

'******************************
'     Pasillos superiores
'******************************

Sub sw007_Hit
    PlaySoundAt "fx_sensor", sw007
    If Tilted Then Exit Sub
    AddScore 200
    LeftInlaneLightA.State = 0
    UpperLightA.State = 0
    CenterLightA.State = 1
    LitBumper1
    CheckABC
End Sub

Sub sw008_Hit
    PlaySoundAt "fx_sensor", sw008
    If Tilted Then Exit Sub
    AddScore 500
    LeftInlaneLightB.State = 0
    RightInlaneLightB.State = 0
    UpperLightB.State = 0
    CenterLightB.State = 1
    LitBumper3
    CheckABC
End Sub

Sub sw009_Hit
    PlaySoundAt "fx_sensor", sw009
    If Tilted Then Exit Sub
    AddScore 500
    RightInlaneLightC.State = 0
    UpperLightC.State = 0
    CenterLightC.State = 1
    LitBumper2
    CheckABC
End Sub

Sub CheckABC
    If CenterLightA.State + CenterLightB.State + CenterLightC.State = 3 Then
        DoubleBonus.State = 1
        BonusMultiplier = 2
        If BallsRemaining(CurrentPlayer) = 1 Then
            BonusMultiplier = 3
        End If
    End If
End Sub

Function ABCComplete
    If CenterLightA.State + CenterLightB.State + CenterLightC.State = 3 Then
        ABCComplete = True
    Else
        ABCComplete = False
    End If
End Function

'******************************
'         Targets
'******************************

Sub Target001_hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 5000
    If TargetLeftLight1.State Then AddBonus 1
End Sub

Sub Target002_hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 5000
    If TargetLeftLight2.State Then AddBonus 1
End Sub

Sub Target003_hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 5000
    If TargetRightLight1.State Then AddBonus 1
End Sub

Sub Target004_hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 5000
    If TargetRightLight2.State Then AddBonus 1
End Sub

'******************************
'       Drop Targets
'******************************

' left bank

Sub dt001_Hit
    PlaySoundAt "fx_droptarget", dt001
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt001.UserValue = 1
        CheckLeftBank
    End If
End Sub

Sub dt002_Hit
    PlaySoundAt "fx_droptarget", dt002
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt002.UserValue = 1
        CheckLeftBank
    End If
End Sub

Sub dt003_Hit
    PlaySoundAt "fx_droptarget", dt003
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt003.UserValue = 1
        CheckLeftBank
    End If
    If LeftTarget5000.State Then
        Addscore 50000
        dt003.UserValue = 2
        If dt008.UserValue = 2 Then
			If ABCComplete Then
            LeftSpecialLight.State = 1
			End If
            vpmtimer.addtimer 1000, "ResetWhiteTargets '"
        End If
    End If
End Sub

Sub dt004_Hit
    PlaySoundAt "fx_droptarget", dt004
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt004.UserValue = 1
        CheckLeftBank
    End If
End Sub

Sub dt005_Hit
    PlaySoundAt "fx_droptarget", dt005
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt005.UserValue = 1
        CheckLeftBank
    End If
End Sub

Sub CheckLeftBank
    If dt001.UserValue + dt002.UserValue + dt003.UserValue + dt004.UserValue + dt005.UserValue = 5 Then
        If ABCComplete Then
            LeftSpecialLight.State = 1
        End If
        CheckAllDown
    End If
End Sub

' right bank

Sub dt006_Hit
    PlaySoundAt "fx_droptarget", dt006
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt006.UserValue = 1
        CheckRightBank
    End If
End Sub

Sub dt007_Hit
    PlaySoundAt "fx_droptarget", dt007
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt007.UserValue = 1
        CheckRightBank
    End If
End Sub

Sub dt008_Hit
    PlaySoundAt "fx_droptarget", dt008
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt008.UserValue = 1
        CheckRightBank
    End If
    If RightTarget5000.State Then
        Addscore 50000
        dt008.UserValue = 2
        If dt003.UserValue = 2 Then
			If ABCComplete Then
            RightSpecialLight.State = 1
			End If
            vpmtimer.addtimer 1000, "ResetWhiteTargets '"
        End If
    End If
End Sub

Sub dt009_Hit
    PlaySoundAt "fx_droptarget", dt009
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt009.UserValue = 1
        CheckRightBank
    End If
End Sub

Sub dt010_Hit
    PlaySoundAt "fx_droptarget", dt010
    If Tilted Then Exit Sub
    If LeftTarget500.State Then
        AddBonus 1
        Addscore 5000
        dt010.UserValue = 1
        CheckRightBank
    End If
End Sub

Sub CheckRightBank
    If dt006.UserValue + dt007.UserValue + dt008.UserValue + dt009.UserValue + dt010.UserValue = 5 Then
        If ABCComplete Then
            RightSpecialLight.State = 1
        End If
        CheckAllDown
    End If
End Sub

Sub CheckAllDown
    If dt001.UserValue + dt002.UserValue + dt003.UserValue + dt004.UserValue + dt005.UserValue + dt006.UserValue + dt007.UserValue + dt008.UserValue + dt009.UserValue + dt010.UserValue = 10 Then
        LeftTarget500.State = 0
        RightTarget500.State = 0
        LeftTarget5000.State = 1
        RightTarget5000.State = 1
		vpmtimer.addtimer 1000, "ResetWhiteTargets '"
		
    End If
End Sub

Sub ResetWhiteTargets
    PlaySoundAt "fx_resetdrop", lifull
    dt003.IsDropped = 0
    dt003.UserValue = 0
    dt008.IsDropped = 0
    dt008.UserValue = 0
End Sub

Sub ResetLeftDroptargets
    Dim i
    PlaySoundAt "fx_resetdrop", dt003
    For each i in aLeftDroptargets
        i.IsDropped = 0
        i.UserValue = 0
    Next
End Sub

Sub ResetRightDroptargets
    Dim i
    PlaySoundAt "fx_resetdrop", dt008
    For each i in aRightDroptargets
        i.IsDropped = 0
        i.UserValue = 0
    Next
End Sub

'******************************
'          Holes
'******************************

Sub kicker001_Hit :pupevent 808
    PlaySoundAt "fx_kicker_enter", kicker001
    If Not Tilted Then
        AddScore 100
        If LeftSpecialLight.State Then
            LeftSpecialLight.State = 0
            AwardSpecial
        End If
    End If
    vpmtimer.addtimer 1500, "PlaySoundAt""fx_kicker"", kicker001: kicker001.kick 38,13 + RND(1) * 4 '"
End Sub

Sub kicker002_Hit :pupevent 809 'center
    PlaySoundAt "fx_kicker_enter", kicker002
    If Not Tilted Then
        AddScore 100
        If RightSpecialLight.State Then
            RightSpecialLight.State = 0
            AwardSpecial
        End If
    End If
    vpmtimer.addtimer 1500, "PlaySoundAt""fx_kicker"", kicker002: kicker002.kick 322,13 + RND(1) * 4 '"
End Sub


'-----------------------------
'-----  FS Display Code  -----
'-----------------------------

'If You want to hide a display, set the reel value of every reel to 44. This picture is transparent
'This is best done using collection:
'
'	If HideDisplay then 
'		For Each obj in ReelsCollection:obj.setvalue(44):next
'	end if
 

 Dim Char(32),TempText                    'increase dimension if You need larger displays



'-----------------------------------------------
'-----  B2S section, not used in the demo  -----
'-----------------------------------------------

Sub DisplayB2SText(TextPar)							'Procedure to display Text on a 28 digit B2S LED reel. Assuming that it is display 1 with internal digit numbers 1-32
  If B2SOn Then
	TempText = TextPar		
	for i = 1 to 32
		if i <= len(TextPar) then
			Char(i) = left(TempText,1)
			TempText = right(Temptext,len(TempText)-1)		
		else
			Char(i) = " "
		end if
	next
	if B2SOn Then
	for i = 1 to 32
		controller.B2SSetLED i,B2SLEDValue(Char(i))
	next
	end if
  End If
End Sub

Dim i

Sub DisplayB2SText2(TextPar)							'Procedure to display Text on a 28 digit B2S LED reel. Assuming that it is display 1 with internal digit numbers 1-32
 If B2SOn Then
	TempText = TextPar		
	for i = 1 to 32
		if i <= len(TextPar) then
			Char(i) = left(TempText,1)
			TempText = right(Temptext,len(TempText)-1)		
		else
			Char(i) = " "
		end if
	next

	for i = 1 to 32
		controller.B2SSetLED i,B2SLEDValue(Char(i))

	next
End If

DMDUpdate.interval = 2000
DMDUpdate.enabled = 1

End Sub



Function B2SLEDValue(CharPar)						'to be used with dB2S 15-segments-LED used in Herweh's Designer
	B2SLEDValue = 0									'default for unknown characters
	select case CharPar
		Case "","":	B2SLEDValue = 0
		Case "0":	B2SLEDValue = 63	
		Case "1":	B2SLEDValue = 8704
		Case "2":	B2SLEDValue = 2139
		Case "3":	B2SLEDValue = 2127	
		Case "4":	B2SLEDValue = 2150
		Case "5":	B2SLEDValue = 2157
		Case "6":	B2SLEDValue = 2172
		Case "7":	B2SLEDValue = 7
		Case "8":	B2SLEDValue = 2175
		Case "9":	B2SLEDValue = 2159
		Case "A":	B2SLEDValue = 2167
		Case "B":	B2SLEDValue = 10767
		Case "C":	B2SLEDValue = 57
		Case "D":	B2SLEDValue = 8719
		Case "E":	B2SLEDValue = 121
		Case "F":	B2SLEDValue = 2161
		Case "G":	B2SLEDValue = 2109
		Case "H":	B2SLEDValue = 2166
		Case "I":	B2SLEDValue = 8713
		Case "J":	B2SLEDValue = 31
		Case "K":	B2SLEDValue = 5232
		Case "L":	B2SLEDValue = 56
		Case "M":	B2SLEDValue = 1334
		Case "N":	B2SLEDValue = 4406
		Case "O":	B2SLEDValue = 63
		Case "P":	B2SLEDValue = 2163
		Case "Q":	B2SLEDValue = 4287
		Case "R":	B2SLEDValue = 6259
		Case "S":	B2SLEDValue = 2157
		Case "T":	B2SLEDValue = 8705
		Case "U":	B2SLEDValue = 62
		Case "V":	B2SLEDValue = 17456
		Case "W":	B2SLEDValue = 20534
		Case "X":	B2SLEDValue = 21760
		Case "Y":	B2SLEDValue = 9472
		Case "Z":	B2SLEDValue = 17417
		Case "<":	B2SLEDValue = 5120
		Case ">":	B2SLEDValue = 16640
		Case "^":	B2SLEDValue = 17414
		Case ".":	B2SLEDValue = 8
		Case "!":	B2SLEDValue = 0
		Case ".":	B2SLEDValue = 128
		Case "*":	B2SLEDValue = 32576
		Case "/":	B2SLEDValue = 17408
		Case "\":	B2SLEDValue = 4352
		Case "|":	B2SLEDValue = 8704
		Case "=":	B2SLEDValue = 2120
		Case "+":	B2SLEDValue = 10816
		Case "-":	B2SLEDValue = 2112
	end select			
	B2SLEDValue = cint(B2SLEDValue)
End Function





Dim Ball


Sub DMDUpdate_Timer
    DMDUpdate.enabled = 0
 If PlayersPlayingGame = 1 And BallsOnPlayfield >= 1 Then
    DMDScore 
    AddScore (0)
    DMDUpdate.interval = 1000	
    DMDUpdate.enabled = 1
  Else
    DMDUpdate.interval = 1500	
    DMDUpdate.enabled = 1
 End If
End Sub

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / table1.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / table1.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub




'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and 
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they 
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

