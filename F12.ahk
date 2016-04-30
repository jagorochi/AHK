#SingleInstance force

SetDefaultMouseSpeed, 0
SetMouseDelay, 15
#include va.ahk
;~ #include SoundHandle.ahk

;~ ; Add resources to the compiled script
;~ ;@Ahk2Exe-AddResource bulle.png, bulle.png
;~ ;@Ahk2Exe-AddResource sonnerie.wav, sonnerie.wav



;~ #Persistent
SetBatchLines, -1
Process, Priority,, High

checkForIncomingCallEvent()





^F11::
SoundPlay, %A_ScriptDir%\sonnerie.wav
return

F12::
decrochage_semi_automatique() ; décrochage de l'appel sans avoir a utiliser la souris
return



;--------------------------------------------------------------------------------------------------------
;    fonctions 
;--------------------------------------------------------------------------------------------------------
decrochage_semi_automatique(){
  ;~ ToolTip, F12, 5, 5
  ;~ SetTimer, EndToolTip, 500
  global gInteractionClient
  xFound :=0
  yFound :=0
  ;~ static strImageBullePath := "\\europe.bmw.corp\WINFS\NSC-Europe\France\SF-proj\SF_PRO_001_B2S\Finance\AHK\bulle.png"
  static strImageBullePath := A_ScriptDir . "\bulle.png"
  static gnNewInt_xPos := 0
  static gnNewInt_yPos := 0
  WinGet, id, list,,, Program Manager
  Loop, %id%
  {
    this_id := id%A_Index%
    WinGetTitle, this_title, ahk_id %this_id%
    if InStr(this_title,"Nouvelle interaction"){
      ;~ ToolTip, Fenetre Nouvelle interaction trouvée !
      ;~ SetTimer, EndToolTip, 1000
      WinActivate, ahk_id %this_id%
      WinGetClass, this_class, ahk_id %this_id%
      ;xFound := 0
      ;yFound := 0
      if (gnNewInt_xPos > 0) {
        Click %gnNewInt_xPos%, %gnNewInt_yPos%, 2
		sleep, 1000
		winClose, ahk_id %this_id%
        break
      }else{
        xPos := 18 - 2 ; 14
        yPos := 64 - 2 ; 58
        xPosL := xPos + 16 + 2
        yPosH := yPos + 16 + 100 + 2
        ImageSearch , xFound, yFound, xPos, yPos, xPosL, yPosH, *50 %strImageBullePath%
        if (xFound > 0){
          gnNewInt_xPos := xFound + 40
          gnNewInt_yPos := yFound + 15
          Click %gnNewInt_xPos%, %gnNewInt_yPos%, 2
          ;msgbox, image trouvé aux coordonnées : xFound = %xFound%,  yFound = %yFound%
          WinGet, ININ_ID, ID,Interaction Client
          GroupAdd, gInteractionClient,  ahk_id %ININ_ID%
		  sleep, 1000
		  winClose, ahk_id %this_id%
          break
          
        }else if (ErrorLevel  = 2){
          ToolTip, ErrorLevel  = 2 `n gstrImageBullePath = %strImageBullePath%
          SetTimer, EndToolTip, 1000
          ;msgbox, ErrorLevel  = 2 !
        }else if (ErrorLevel  = 1) {
          ToolTip, ErrorLevel  = 1
          SetTimer, EndToolTip, 1000
          ;msgbox, image introuvable !
        }
      }
    }
  }
}

authorized_user(){
	;----------------------------------------------- B2B
	if (A_Username <> "qxe1373"){ ; william
		return true
	}
	if (A_Username <> "qxb2961"){ ; karin
		return true
	}
	if (A_Username <> "qxb2962"){ ; sabrina
		return true
	}
	if (A_Username <> "qxj0667"){ ; pauline
		return true
	}
	
	;----------------------------------------------- B2C
	if (A_Username <> "qx94523"){ ; martin
		return true
	}
	if (A_Username <> "qxb2980"){ ; stéphane
		return true
	}
	if (A_Username <> "qxb2965"){ ; jérémy
		return true
	}
	if (A_Username <> "qxe1870"){ ; aurélie François
		return true
	}
	if (A_Username <> "qxe1962"){ ; widaad
		return true
	}
	if (A_Username <> "qxg5761"){ ; keltoum
		return true
	}
	if (A_Username <> "qxj9145"){ ; lova
		return true
	}
	if (A_Username <> "qxk1393"){ ; wallya
		return true
	}
	if (A_Username <> "qxk7497"){ ; abdel akim
		return true
	}
	if (A_Username <> "qxe8498"){ ; juliette
		return true
	}
	if (A_Username <> "qxk7498"){ ; momo
		return true
	}
	if (A_Username <> "qxh8279"){ ; jara
		return true
	}
	
	return false
}

checkForIncomingCallEvent(){
	global gstrINFO1
	global gbWin
	;--------------------------------------------------------------------

	
	if (authorized_user() = false) {
		msgBox % "Cette application est en phase de test et son accès est restreint..."
		ExitApp
	}
	;~ FileSelectFile, file, 1,, Pick a sound file
	;~ if file =
		;~ ExitApp
	;~ msgbox % "file = " . file
	;~ hSound := Sound_Open(file)
	;~ If (Not hSound){
		
		;~ msgbox % "test.wav" not found
	;~ }
	
	gbWin := false
	while true
	{
		WinGet, id, list,,, Program Manager
		bFound := false
		Loop, %id%
		{
			this_id := id%A_Index%
			WinGetTitle, this_title, ahk_id %this_id%
			if InStr(this_title,"Nouvelle interaction"){
				bFound := true
				if (gbWin = false){
					; couper le son de l'appli
					; jouer un autre son "plus doux"
					gbWin := true
					
					
					WinGet, ActivePid, PID, Interaction Client
					if (Volume := GetVolumeObject(ActivePid)){
						;~ MsgBox, There was a problem retrieving the application volume interface
						VA_ISimpleAudioVolume_GetMute(Volume, Mute)  ;Get mute state
						;Msgbox % "Application " ActivePID " is currently " (mute ? "muted" : "not muted")
						if (Mute=false){
							VA_ISimpleAudioVolume_SetMute(Volume, true) ;mute !
						}
						ObjRelease(Volume)
					}
					SoundPlay, %A_ScriptDir%\sonnerie.wav
					;~ WinGet, winInIn, ID,Interaction Client
					;~ GuiControl, WinCtrlTest:Text, gstrINFO1, "Nouvelle interaction !"
					sleep, 1000
					break
				}
			}
			
		;~ GuiControl, WinCtrlTest:Text, gstrINFO1, ...
		
		}
		
		if (bFound = false){
			gbWin := false
		}
		sleep, 10
		
	}

}


createCtrlWindows(){
	global gstrINFO1
	global gstrINFO2
	global gstrINFO3
	Gui, WinCtrlTest: New 
	;~ Gui, WinCtrlTest:Color, AAAAAA
	Gui, WinCtrlTest:font, s9, Verdana 
	;Gui, WinCtrlTest: Add, Picture, x0 y0 w100 h100 vgLightImage_Ref, %strDir%status_question.png

	Gui, WinCtrlTest: Add, Text, +BackgroundTrans x0 y0 w120 h20  +Left vgstrINFO1, StartUp 1
	Gui, WinCtrlTest: Add, Text, +BackgroundTrans x0 y12 w120 h20  +left vgstrINFO2, StartUp 2
	Gui, WinCtrlTest: Add, Text, +BackgroundTrans x0 y24 w120 h20  +left vgstrINFO3, StartUp 3
	;~ Gui, WinCtrlTest: Add, Text, x0 y100 w78 h14 +Center vgUserStatus, Unknow
	
	Gui, WinCtrlTest:+AlwaysOnTop -MinimizeBox -caption +Border ; +0x40000 ; -Resize ;   
	Gui, WinCtrlTest:+Owner
	Gui, WinCtrlTest:Show, x500 y10 h40 w120, WinCtrlTest
	OnMessage(0x201, "WM_LBUTTONDOWN")
	
}


EndToolTip:
  setTimer, EndToolTip, off
  ToolTip
return





GetVolumeObject(Param = 0)
{
    static IID_IASM2 := "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}"
    , IID_IASC2 := "{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}"
    , IID_ISAV := "{87CE5498-68D6-44E5-9215-6DA47EF883D8}"
    
    ; Get PID from process name
    if Param is not Integer
    {
        Process, Exist, %Param%
        Param := ErrorLevel
    }
    
    ; GetDefaultAudioEndpoint
    DAE := VA_GetDevice()
    
    ; activate the session manager
    VA_IMMDevice_Activate(DAE, IID_IASM2, 0, 0, IASM2)
    
    ; enumerate sessions for on this device
    VA_IAudioSessionManager2_GetSessionEnumerator(IASM2, IASE)
    VA_IAudioSessionEnumerator_GetCount(IASE, Count)
    
    ; search for an audio session with the required name
    Loop, % Count
    {
        ; Get the IAudioSessionControl object
        VA_IAudioSessionEnumerator_GetSession(IASE, A_Index-1, IASC)
        
        ; Query the IAudioSessionControl for an IAudioSessionControl2 object
        IASC2 := ComObjQuery(IASC, IID_IASC2)
        ObjRelease(IASC)
        
        ; Get the session's process ID
        VA_IAudioSessionControl2_GetProcessID(IASC2, SPID)
        
        ; If the process name is the one we are looking for
        if (SPID == Param)
        {
            ; Query for the ISimpleAudioVolume
            ISAV := ComObjQuery(IASC2, IID_ISAV)
            
            ObjRelease(IASC2)
            break
        }
        ObjRelease(IASC2)
    }
    ObjRelease(IASE)
    ObjRelease(IASM2)
    ObjRelease(DAE)
    return ISAV
}
 
;
; ISimpleAudioVolume : {87CE5498-68D6-44E5-9215-6DA47EF883D8}
;
VA_ISimpleAudioVolume_SetMasterVolume(this, ByRef fLevel, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "float", fLevel, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMasterVolume(this, ByRef fLevel) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "float*", fLevel)
}
VA_ISimpleAudioVolume_SetMute(this, ByRef Muted, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "int", Muted, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMute(this, ByRef Muted) {
    return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "int*", Muted)
}

