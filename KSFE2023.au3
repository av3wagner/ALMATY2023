;-*- coding: utf-8 -*
;https://www.autoitscript.com/forum/topic/142569-solved-listview-with-transparent-background-problem/
;!!! https://github.com/TheDcoder/Process-UDF-for-AutoIt

  ;;*****************************************************************;;;
  ;;*****************************************************************;;;
  ;;;****************************************************************;;;
  ;;;***  FIRMA          : PARADOX GmbH                           ***;;;
  ;;;***  Autor          : ALEXANDer WAGNER                       ***;;;
  ;;;***  STUDIEN-NAME   : VaS-Bericht                            ***;;;
  ;;;***  STUDIEN-NUMMER :                                        ***;;;
  ;;;***  SPONSOR        :                                        ***;;;
  ;;;***  ARBEITSBEGIN   : 15.11.2022                             ***;;;
  ;;;****************************************************************;;;
  ;;;*--------------------------------------------------------------*;;;
  ;;;*---  PROGRAMM      : KSFE2023.au3                          ---*;;;
  ;;;*---  Parent        : cMAIN_VAS2023m5.au3                   ---*;;;
  ;;;*---  BESCHREIBUNG  : System fÃ¼r KI auf Disk C:             ---*;;;
  ;;;*---                :                                       ---*;;;
  ;;;*---                :                                       ---*;;;
  ;;;*---  VERSION   VOM : 07.11.2022                            ---*;;;
  ;;;*--   KORREKTUR VOM : 06.12.2022, 16.12.2022, 06.01.2023    ---*;;;
  ;;;*--                 :                                       ---*;;;
  ;;;*---  INPUT         :.INI                                   ---*;;;
  ;;;*---  OUTPUT        :                                       ---*;;;
  ;;;*--------------------------------------------------------------*;;;
  ;;;************************ Ã„nderung ******************************;;;
  ;;;****************************************************************;;;
  ;;;  Wann              :               Was                        *;;;
  ;;;*--------------------------------------------------------------*;;;
  ;;;* 07.12.2022        : C:\ Disc                                 *;;;
  ;;;* 16.12.2022        : C:\ Disc, Restore                        *;;;
  ;;;* 06.01.2023        : C:\ Disc, Restore                        *;;;
  ;;;*---                :                                       ---*;;;
  ;;;****************************************************************;;;


#include <IE.au3>
#include <WinAPI.au3>
#include <GUIConstants.au3>
#include <Windowsconstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <Constants.au3>
#include <array.au3>
#include <Math.au3>
#include <Misc.au3>
#include <Timers.au3>
#include <Date.au3>
#Include <File.au3>
#include <String.au3>
#include <GUIMenu.au3>
#include <GuiRichEdit.au3>
#include <GuiEdit.au3>
#include <GuiImageList.au3>
#include <ButtonConstants.au3>
#include <FontConstants.au3>
#include <GuiListView.au3>
#include <ListviewConstants.au3>
#include <GUIListBox.au3>
#include <GDIPlus.au3>
#include <TabConstants.au3>
#include <EditConstants.au3>
#include <Timers.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <GuiTab.au3>
#include <GuiListView.au3>
#include <GuiImageList.au3>
#include <ButtonConstants.au3>
#include <GUIListBox.au3>
#include <Word.au3>
#include <Excel.au3>
#include <GUIRichLabel.au3>
#include <GuiRichEdit.au3>
#include <WindowsConstants.au3>
#include <GuiTab.au3>
#include <ColorConstants.au3>
#include <GuiImageList.au3>
#include <GuiListView.au3>

HotKeySet("!{ESC}", "Terminate")
HotKeySet("^!a",    "Terminate")

Opt('MustDeclareVars', 0)
Opt("WinWaitDelay", 250)
Opt("GUIOnEventMode", 1)
Opt ("MouseCoordMode", 0)
Opt ("WinTitleMatchMode", 4)
Global $Land = @ScriptDir & "\KZ.htm", $Word
Global $DOS, $Message = ''
Global $oIE, $INTERNET, $Dash="http://127.0.0.1:8055/"
Global $GUI, $MenueDatei, $DateiContext, $HelpAbout, $HelpContext, $PicAW, $APP=0, $WB=@DesktopWidth-800, $HB=7*(@DesktopWidth-800)/9
Global $Fdoc, $PY

Global $Land = @ScriptDir & "\KZ.htm"
Global $oIE, $INTERNET, $Dash="http://127.0.0.1:8085/", $Land=@ScriptDir & "\OUTPUT\MAPKZ-00.HTM", $RLand=@ScriptDir & "\OUTPUT\RVBARKZ-11.htm", $VBAR=@ScriptDir & "\OUTPUT\RVBARKZ-63.HTM"
Global $GUI, $MenueDatei, $DateiContext, $HelpAbout, $HelpContext, $MenueWagner, $DateiContext2, $PicAW, $APP=0, $WB=@DesktopWidth-800, $HB=7*(@DesktopWidth-800)/9
Global $oExcel, $ws, $Bhandle, $OBL="KZ-19", $START=0, $ExitPr, $Home, $oIE2, $oIE3, $Tab, $INTERNET3, $iCheckSum, $iCheckSum2
Global $oIE, $handle

WinMinimizeAll ( )
WinSetState("[Class:Shell_TrayWnd]","", @SW_SHOW)
WinSetState("Start","", @SW_SHOW )
Sleep(100)


While ProcessExists("Winword.exe")
	ProcessClose("Winword.exe")
	Sleep(500)
WEnd

While ProcessExists("msedge.exe")
	ProcessClose("msedge.exe")
	Sleep(500)
WEnd


WinSetState("[Class:Shell_TrayWnd]","", @SW_HIDE)
Sleep(100)

STARTGUI()
GUISetFont(10, 800)

GUICtrlSetState(-1, $GUI_DISABLE)
GUISetState ()

While 1

WEnd

Func STARTGUI()
	$RAND=0.999*@DesktopWidth-24
	Local $Factor=95, $Step=30
	Local $Breite=$Factor-5
	$Unten=42
    $LR=15

    $Title="Â© PARADOX GmbH. VaS-System Version 1.0 (Dezember 2022), entwickelt von Dr. Alexander Wagner"
	$GUI = GUICreate (" ", 0.999*@DesktopWidth, 0.999*@DesktopHeight, 0, 0, $WS_POPUP, BitOR($SS_NOTIFY,$WS_GROUP))
	$NOTE= _GUICtrlRichLabel_Create($GUI, '', 1, 0.999*@DesktopHeight-25, 890) ;0.965*@DesktopHeight-30
    GLOBAL $NOTEText='<font attrib="normal" size="9" name="Tahoma" color="WHITE" align="L"> ' & $Title & ' </font>'
	_GUICtrlRichLabel_SetData($NOTE, $NOTEText & @LF)

	$PicAW=GUICtrlCreatePic(@ScriptDir & "\KI.jpg", (@DesktopWidth-76), 1, 70, 38)
	$hRichLabel1 = _GUICtrlRichLabel_Create($GUI, '', 5, 1, @DesktopWidth-77, 40)
	GLOBAL $LabelTitle1='<font attrib="bold" size="20" name="Tahoma" color="YELLOW" align="C">Web- und Cloudbasiertes Eco-System fÃ¼r die automatisierte Erstellung des Validierung-Berichts</font>'
    _GUICtrlRichLabel_SetData($hRichLabel1, $LabelTitle1 & @LF)

	;$hRichLabel2 = _GUICtrlRichLabel_Create($GUI, '', 105, 36, @DesktopWidth-105, 28)
	$hRichLabel2 = _GUICtrlRichLabel_Create($GUI, '', 105, 64, @DesktopWidth-105, 28)
  	GLOBAL $LabelTitle2='<font attrib="bold" size="14" name="Tahoma" color="darkgreen" align="C">PARADOX GmbH KÃ¼nstliche Intelligenz Dashboard. Demo-Version 3.0, Stand: 30. November 2022</font>'
    _GUICtrlRichLabel_SetData($hRichLabel2, $LabelTitle2 & @LF)

	Global $ExitPr  = GUICtrlCreateIcon("C:\WINDOWS\system32\shell32.dll",   -28, 0.999*@DesktopWidth-28, $Unten+2, 20, 20, BitOR($SS_NOTIFY,$WS_GROUP))
	GUICtrlSetOnEvent($ExitPr, "Terminate")

	Global $PIPE =  GUICtrlCreateButton("WEB APP", $Rand-1*$Factor-3,   $Unten, $Breite, 25, 0)
	GUICtrlSetBkColor(-1, 0x000000)
	GUICtrlSetColor($PIPE, 0xFFFF00)
	GUICtrlSetOnEvent($PIPE, "APP")

	;Global $ToeXCEL  = GUICtrlCreateButton("Back-End",        $Rand-2*$Factor+$Step-100, $Unten, $Breite, 25, 0)
	;GUICtrlSetBkColor(-1, 0x000000)
	;GUICtrlSetColor($ToeXCEL, 0xFFFF00)

    Global $Script  = GUICtrlCreateButton("Open Py-Script",         3*$Factor+$Step+5, $Unten, $Breite+7, 25, 0)
	GUICtrlSetBkColor(-1, 0x000000)
	GUICtrlSetColor($Script, 0xFFFF00)
	GUICtrlSetOnEvent($Script, "OpenScript")
    ;notepad++.exe

	Global $ToWord  = GUICtrlCreateButton("Open Bericht",         2*$Factor+$Step+5, $Unten, $Breite, 25, 0)
	GUICtrlSetBkColor(-1, 0x000000)
	GUICtrlSetColor($ToWord, 0xFFFF00)
	GUICtrlSetOnEvent($ToWord, "xOpenWord")

	Global $BERICHT  = GUICtrlCreateButton("Start Bericht",         1*$Factor+$Step+5, $Unten, $Breite, 25, 0)
	GUICtrlSetBkColor(-1, 0x000000)
	GUICtrlSetColor($BERICHT, 0xFFFF00)
	GUICtrlSetOnEvent($BERICHT, "BERICHT")

 	Global $ReStart  = GUICtrlCreateButton("System-Konzept",         $Step-15, $Unten, $Breite+15, 25, 0)
	GUICtrlSetBkColor(-1, 0x000000)
	GUICtrlSetColor($ReStart, 0xFFFF00)
	;GUICtrlSetOnEvent($ReStart, "ReAPP")

	Global $LEER  = GUICtrlCreateButton("", 1, 1, 1, 1, 0)
    GUISetBkColor(0x000000)
	;$RC=Run(@ProgramFilesDir & "\Internet Explorer\Iexplore.exe -k " & "https://medium.com/@itsdaniyalm/automate-data-pipeline-for-front-end-dashboards-in-python-321d9c0db098", @WorkingDir)
	Sleep(50)
	;$RC=Run(@ProgramFilesDir & "\Mozilla Firefox\firefox.exe -k " & "https://av3wagner-streamlit-example-streamlit-app-05799z.streamlit.app", @WorkingDir)
EndFunc

Func APP()
	;$PY=@ProgramFilesDir & "\Mozilla Firefox\firefox.exe -k " & "https://av3wagner-streamlit-example-streamlit-app-05799z.streamlit.app" ;, @WorkingDir)'
	;$PY="https://av3wagner-streamlit-example-streamlit-app-05799z.streamlit.app"

    $PY="http://127.0.0.1:8085/"
	Send("#r")
	Sleep(50)
	Send($PY)
	Sleep(50)
	Send("{Enter}")
	Sleep(5000)

	Local $hFireFoxWin=0,$aWinList=WinList("[REGEXPCLASS:Mozilla(UI)?WindowClass]")
	For $i=1 To $aWinList[0][0]
		If BitAND(_WinAPI_GetWindowLong($aWinList[$i][1],$GWL_STYLE),$WS_POPUP)=0 Then
			$hFireFoxWin=$aWinList[$i][1]
			ExitLoop
		EndIf
	Next
	If $hFireFoxWin Then WinActivate($hFireFoxWin)

	Opt("WinTitleMatchMode", 4)
	$winMatchFirefox = "[CLASS:MozillaWindowClass]"
	Local $Bhandle = WinActivate($winMatchFirefox)

	WinActivate ($Bhandle)
	ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

	Sleep(100)
	DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
	DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
	Sleep(300)
	WinMove($Bhandle, "", -1, @DesktopHeight/2-90, @DesktopWidth+2, @DesktopHeight/2+90)
	;WinMove($Bhandle, "", -1, 90, @DesktopWidth+2, @DesktopHeight-115)

    #CS
	Sleep(3000)
    ;Local $aWinList = WinList("[REGEXPTITLE:(?i)(.*Streamlit â€“ Mozilla Firefox.*)]")
	Local $aWinList = WinList("[REGEXPTITLE:(?i)(.*Kazakhstan Dashboard â€“ Mozilla Firefox.*)]")
	_ArrayDisplay($aWinList)

	$Bhandle=$aWinList[1][1]
	WinActivate ($Bhandle)
	ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

	Sleep(100)
	DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
	DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
	;Sleep(300)
 	WinMove($Bhandle, "", -1, @DesktopHeight/2-90, @DesktopWidth+2, @DesktopHeight/2+90)
	WinMove($Bhandle, "", -1, 90, @DesktopWidth+2, @DesktopHeight-115)
	#CE
EndFunc

Func BERICHT()
	;PRIVAT NB
	$Fdoc=@ScriptDir & '\OUTPUT\AVaS2023Finish.docx'
	$PY='Python ' & @ScriptDir & '\KSFE2023.py'
	$CONDA='%windir%\System32\cmd.exe "/K" C:\Users\satur\anaconda3\Scripts\activate.bat'

	While (FileExists($Fdoc))
		FileDelete($Fdoc)
		Sleep(50)
	WEnd

	Send("#r")
	Sleep(50)
	Send($CONDA)
	Sleep(500)
	Send("{Enter}")
	Sleep(500)
	Send($PY)
	Sleep(50)
	Send("{Enter}")

	While NOT (FileExists($Fdoc))
		Sleep(50)
	WEnd
	$Word=1
	Sleep(12000)

	Local $aWinList = WinList("[REGEXPTITLE:(?i)(.*AVaS2023Finish.*)]")
	$strSearch="AVaS2023Finish"

	$Bhandle=$aWinList[1][1]
	WinActivate ($Bhandle)
	ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

	If IsObj($WORD)=True Then
		$Word.ActiveDocument.ActiveWindow.View.Type = 3
		$Word.ActiveDocument.ActiveWindow.View.Zoom.Percentage = 110
		$Word.ActiveDocument.ActiveWindow.ActivePane.View.ShowAll = 0
		Sleep(50)
		$oWordRC=1
	EndIf

	Sleep(100)
	DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
	DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
	Sleep(300)
	WinMove($Bhandle, "", @DesktopWidth/8, 90, 3*@DesktopWidth/4, @DesktopHeight-115)


	While ProcessExists("CMD.EXE")
		  ProcessClose("CMD.EXE")
		  Sleep(50)
	WEnd
EndFunc

Func _MenuPressed()
    Switch @GUI_CtrlId
        Case $MenueDatei
            ShowMenu($GUI, $MenueDatei, $DateiContext)
    Case $HelpAbout
            ShowMenu($Gui, $HelpAbout, $HelpContext)
    EndSwitch
EndFunc

Func ShowMenu($hWnd, $CtrlID, $nContextID)
    Local $arPos, $x, $y
    Local $hMenu = GUICtrlGetHandle($nContextID)

    $arPos = ControlGetPos($hWnd, "", $CtrlID)

    $x = $arPos[0]
    $y = $arPos[1] + $arPos[3]

    ClientToScreen($hWnd, $x, $y)
    TrackPopupMenu($hWnd, $hMenu, $x, $y)
EndFunc   ;==>ShowMenu

; Convert the client (GUI) coordinates to screen (desktop) coordinates
Func ClientToScreen($hWnd, ByRef $x, ByRef $y)
    Local $stPoint = DllStructCreate("int;int")

    DllStructSetData($stPoint, 1, $x)
    DllStructSetData($stPoint, 2, $y)

    DllCall("user32.dll", "int", "ClientToScreen", "hwnd", $hWnd, "ptr", DllStructGetPtr($stPoint))

    $x = DllStructGetData($stPoint, 1)
    $y = DllStructGetData($stPoint, 2)
    ; release Struct not really needed as it is a local
    $stPoint = 0
EndFunc   ;==>ClientToScreen

; Show at the given coordinates (x, y) the popup menu (hMenu) which belongs to a given GUI window (hWnd)
Func TrackPopupMenu($hWnd, $hMenu, $x, $y)
    DllCall("user32.dll", "int", "TrackPopupMenuEx", "hwnd", $hMenu, "int", 0, "int", $x, "int", $y, "hwnd", $hWnd, "ptr", 0)
EndFunc   ;==>TrackPopupMenu

Func AUTOR()
	;MsgBox(0, "IE Quit", "IE Stopp")
    GUICtrlSetPos($INTERNET, 1, 1, 1, 1)
	$PicAW=GUICtrlCreatePic(@ScriptDir & "\DoktorAW.jpg", (@DesktopWidth-$WB)/2, 75, $WB, $HB)
EndFunc

Func INIBACK()

    MsgBox(64, "Update...", "Es fehlt im Moment Zugriff zu Datenbanken")
EndFunc

Func _About()
    MsgBox(64, "About...", "Beispiel fÃ¼r ein eigenes MenÃ¼")
EndFunc

Func _Website()
    MsgBox(64, "Website...", "www.autoit.de")
EndFunc

Func ReAPP()
	;$Dash="http://127.0.0.1:8055/"
	;GUICtrlSetPos($PicAW, 1, 1, 1, 1)
	GUICtrlSetPos($INTERNET, 0, 65, 0.999*@DesktopWidth, 0.915*@DesktopHeight-65)
	;$oIE.navigate($Dash)
	$RC=Run(@ProgramFilesDir & "\Internet Explorer\Iexplore.exe -k " & "https://medium.com/@itsdaniyalm/automate-data-pipeline-for-front-end-dashboards-in-python-321d9c0db098", @WorkingDir)
EndFunc

Func xAPP($Dash)
    ;$Dash="http://127.0.0.1:8055/"
	If $APP = 0 Then
		$oIE = ObjCreate("Shell.Explorer.2")
		GUICtrlSetPos($PicAW, 1, 1, 1, 1)
		If @Compiled Then
			RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", @ScriptName, "REG_DWORD", 11001)
		Else
			RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", "AutoIt3.exe", "REG_DWORD", 11001)
			RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", "autoit3_x64.exe", "REG_DWORD", 11001)
		EndIf

		Local $iOffsetLaengeKompatibilitaet = 5
		Global $INTERNET = GUICtrlCreateObj($oIE, 0, 65, 0.999*@DesktopWidth, 0.915*@DesktopHeight-65)
		Global $IOhandle=WinGetHandle("[ACTIVE]")
		ConsoleWrite("ACTIVE: " & $IOhandle & @CR)
		$oIE.navigate ("about:blank")
		$oIE.navigate($Dash)

		$oIE2 = ObjCreate("Shell.Explorer.2")
		$INTERNET2 = GUICtrlCreateObj($oIE2, @DesktopWidth, @DesktopHeight, 1, 1)
		_IEPropertySet($oIE2, 'LEFT', 0)
		_IEPropertySet($oIE2, 'TOP', 48)
		_IEPropertySet($oIE2, 'HEIGHT', 380)
		_IEPropertySet($oIE2, 'WIDTH',  700)
		_IEPropertySet($oIE2, "resizable",  True)

		Global $IOhandle2=WinGetHandle("[ACTIVE]")
		ConsoleWrite("ACTIVE  $IOhandle2: " & $IOhandle2 & @CR)
		$oIE2.navigate ("about:blank")
		$oIE2.navigate($Land)
		_IEAction ($oIE2, "refresh")
		$APP +=1
	Else
	  ;$oIE.navigate("https://medium.com/@itsdaniyalm/automate-data-pipeline-for-front-end-dashboards-in-python-321d9c0db098") ;$Dash)
	  $oIE.navigate($Dash)
	EndIf
EndFunc

Func OpenWord($Sdoc)
	Local $ARA = StringSplit($Sdoc, '\')
	$VORLAGE_DOC = $ARA[$ARA[0]]
	ConsoleWrite("VORLAGE_DOC: " & $VORLAGE_DOC & @CR)
	If $Word=False Then
		MsgBox(1, "Achtung!", "Word-Dokument nicht geÃ¶ffnet!")
	EndIf

    $RCA=WinActivate($VORLAGE_DOC)
	Local $aWinList = WinList("[REGEXPTITLE:(?i)(.*WORD.*)]")
	$strSearch = "Microsoft Word"
	$intIndex = -1
	$intIndex = _ArraySearch($aWinList, $strSearch)
	If @error Then
		ConsoleWrite("_ArraySearch() ERR: " & @error & @CRLF)
	Else
		ConsoleWrite("Index: " & $intIndex & @CRLF)
		$Bhandle=$aWinList[$intIndex][1]
		WinActivate ($Bhandle)
		ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

		If IsObj($WORD)=True Then
			$Word.ActiveDocument.ActiveWindow.View.Type = 3
			$Word.ActiveDocument.ActiveWindow.View.Zoom.Percentage = 130
			$Word.ActiveDocument.ActiveWindow.ActivePane.View.ShowAll = 0
			Sleep(50)
			$oWordRC=1
		EndIf

		Sleep(100)
		DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
		DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
		Sleep(300)
		WinMove($Bhandle, "", @DesktopWidth/8, 80, 3*@DesktopWidth/4, @DesktopHeight-105)
    EndIf
EndFunc

Func yOpenWord()
	#cs
	If BitAND(GUICtrlRead($idRadio1), $GUI_CHECKED)=$GUI_CHECKED Then
		$Fdoc="C:\IPYNB\KSFE2023\OUTPUT\AVaS2023Finish.docx"
		$PY='Python C:\IPYNB\KSFE2023\PROGRAMME\C2AVaS2023.py'
	ElseIf BitAND(GUICtrlRead($idRadio2), $GUI_CHECKED)=$GUI_CHECKED Then
		$Fdoc="C:\IPYNB\KSFE2023\OUTPUT\AVaS2023Finish.docx"
		$PY='Python C:\IPYNB\KSFE2023\PROGRAMME\C2AVaS2023.py'
    EndIf
    #ce

	;WinMinimizeAll( )
	$oWordRC=0
	Local $sDocument = $Fdoc
	$VORLAGE_DOC = "AVaS2023Finish.docx"
	$strSearch = "AVaS2023Finish"
	If IsObj($Word)=False Then $Word = _Word_Create()
	$Word.Visible=False

	Sleep(50)
	Send("#r")
	Sleep(50)
	Send($sDocument)
	Send("{Enter}")
	Sleep(500)
	WinWait($strSearch)
	Local $aWinList = WinList("[REGEXPTITLE:(?i)(.*WORD.*)]")
	;_ArrayDisplay($aWinList)
	$intIndex = -1
	$intIndex = _ArraySearch($aWinList, $strSearch)

	If @error Then
		ConsoleWrite("_ArraySearch() ERR: " & @error & @CRLF)
	Else
		ConsoleWrite("Index: " & $intIndex & @CRLF)
	EndIf

    $Bhandle=$aWinList[1][1]
	WinActivate ($Bhandle)
	ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

    If IsObj($WORD)=True Then
		$Word.ActiveDocument.ActiveWindow.View.Type = 3
		$Word.ActiveDocument.ActiveWindow.View.Zoom.Percentage = 130
		$Word.ActiveDocument.ActiveWindow.ActivePane.View.ShowAll = 0
		Sleep(50)
		$oWordRC=1
	EndIf

  $Word.Visible=True
  Sleep(500)
  DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
  DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
  Sleep(300)
  WinMove($Bhandle, "", @DesktopWidth/8, 80, 3*@DesktopWidth/4, @DesktopHeight-105)
EndFunc

Func OpenScript()
	$notepad_path = 'C:\Program Files\Notepad++\notepad++.exe'
	$file_path = 'C:\IPYNB\KSFE2023\PROGRAMME\C2AVaS2023.py'
	ConsoleWrite("File_path: " & $file_path & @CR)
	Run('"' & $notepad_path & '" "' & $file_path & '"')
EndFunc

Func zOpenScript()
	$notepad_path = '%windir%\System32\cmd.exe'
	$file_path = 'C:\Anaconda3\Scripts\activate.bat'
	Run('"' & $notepad_path & '" "' & $file_path & '"')
EndFunc
;%windir%\System32\cmd.exe "/K" C:\Anaconda3\Scripts\activate.bat

Func xOpenWord()
	#cs
	If BitAND(GUICtrlRead($idRadio1), $GUI_CHECKED)=$GUI_CHECKED Then
		$Fdoc='C:\IPYNB\KSFE2023\OUTPUT\AVaS2023Finish.docx'
		$PY='Python C:\IPYNB\KSFE2023\PROGRAMME\C2AVaS2023.py'
	ElseIf BitAND(GUICtrlRead($idRadio2), $GUI_CHECKED)=$GUI_CHECKED Then
		$Fdoc='C:\IPYNB\KSFE2023\OUTPUT\AVaS2023Finish.docx'
		$PY='Python C:\IPYNB\KSFE2023\PROGRAMME\C2AVaS2023.py'
    EndIf
    #ce

	Local $ARA = StringSplit($Fdoc, '\')
	$VORLAGE_DOC = $ARA[$ARA[0]]
  	$strSearch = "AVaS2023Finish" ; - KompatibilitÃ¤tsmodus"
	;ZI NB
	;$notepad_path = 'C:\Program Files\Microsoft Office\Office14\WINWORD.EXE'
	;PRIVAT NB
	$notepad_path = 'C:\Program Files\Microsoft Office\Office14\WINWORD.EXE'
	;$notepad_path = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
	Run('"' & $notepad_path & '" "' & $Fdoc & '"')
	WinWait($strSearch)

	Local $Bhandle = WinGetHandle($strSearch)
	WinActivate ($Bhandle)
	ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

	Sleep(100)
	DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
	DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
	Sleep(300)
	WinMove($Bhandle, "", @DesktopWidth/8, 90, 3*@DesktopWidth/4, @DesktopHeight-115)
EndFunc

Func pOpenWord()
	$Fdoc='C:\IPYNB\KSFE2023\OUTPUT\Pipeline.PPTX'
	;$Fdoc='C:\IPYNB\KSFE2023\OUTPUT\
	Local $ARA = StringSplit($Fdoc, '\')
	$VORLAGE_DOC = $ARA[$ARA[0]]
  	$strSearch = "Pipeline"
	$notepad_path = 'C:\Program Files\Microsoft Office\Office14\POWERPNT.EXE'
	;$notepad_path = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
	$Word=Run('"' & $notepad_path & '" "' & $Fdoc & '"')
	WinWait($strSearch)

	Local $Bhandle = WinGetHandle($strSearch)
	WinActivate ($Bhandle)
	ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

	WinMove($Bhandle, "", 5, 90, @DesktopWidth-10, @DesktopHeight-115)

	If IsObj($WORD)=True Then
		$Word.ActiveDocument.ActiveWindow.View.Type = 3
		$Word.ActiveDocument.ActiveWindow.View.Zoom.Percentage = 130
		$Word.ActiveDocument.ActiveWindow.ActivePane.View.ShowAll = 0
		Sleep(50)
		$oWordRC=1
	EndIf

    #cs
	Sleep(100)
	DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
	DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
	Sleep(300)

	Sleep(500)
	MouseClick("Left", @DesktopWidth-20, @DesktopHeight-120)
	Sleep(50)
	Send("F10")
	Sleep(50)
	Send("R")
	Sleep(50)
	Send("ZR")
	Sleep(50)
	Send("V")
	#ce
EndFunc


Func wOpenWord()
	$Fdoc='C:\IPYNB\KSFE2023\OUTPUT\Pipeline.docx'
	Local $ARA = StringSplit($Fdoc, '\')
	$VORLAGE_DOC = $ARA[$ARA[0]]
  	$strSearch = "Pipeline"
	$notepad_path = 'C:\Program Files\Microsoft Office\Office14\WINWORD.EXE'
	;$notepad_path = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
	$Word=Run('"' & $notepad_path & '" "' & $Fdoc & '"')
	WinWait($strSearch)

	Local $Bhandle = WinGetHandle($strSearch)
	WinActivate ($Bhandle)
	ConsoleWrite("Bhandle1=" & $Bhandle & @CR)

	WinMove($Bhandle, "", 5, 78, @DesktopWidth-10, @DesktopHeight-105)

	If IsObj($WORD)=True Then
		$Word.ActiveDocument.ActiveWindow.View.Type = 3
		$Word.ActiveDocument.ActiveWindow.View.Zoom.Percentage = 130
		$Word.ActiveDocument.ActiveWindow.ActivePane.View.ShowAll = 0
		Sleep(50)
		$oWordRC=1
	EndIf

    #cs
	Sleep(100)
	DllCall("user32.dll", "int", "SetParent", "hwnd", $Bhandle, "hwnd", WinGetHandle($GUI))
	DllCall("user32.dll", "long", "SetWindowLong", "hwnd", $Bhandle, "int", -16, "long", BitOR($WS_POPUP, $WS_CHILD, $WS_VISIBLE, $WS_CLIPSIBLINGS))
	Sleep(300)

	Sleep(500)
	MouseClick("Left", @DesktopWidth-20, @DesktopHeight-120)
	Sleep(50)
	Send("F10")
	Sleep(50)
	Send("R")
	Sleep(50)
	Send("ZR")
	Sleep(50)
	Send("V")
	#ce
EndFunc

Func Terminate()
   $RC=1
   WinSetState("[Class:Shell_TrayWnd]","", @SW_SHOW)
   Sleep(100)

    ;While (ProcessExists("firefox.exe"))
	;   ProcessClose("firefox.exe")
	;   Sleep(50)
    ;WEnd
   Exit
EndFunc


