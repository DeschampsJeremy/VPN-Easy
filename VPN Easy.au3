#include <Array.au3>
#Include <WinAPI.au3>

Local $progName = "VPN Easy"

Local $absolutePathAutoITFolder = "C:\AutoIT\"
Local $absolutePathAutoITUsernameFile = "USERNAME.txt"
Local $absolutePathAutoITIVANTIFile = "IVANTI NVP.txt"
Local $absolutePathAutoITENTRUSTFile = "ENTRUST PIN.txt"

Local $IVANTIwindowName = "Ivanti Secure Access Client"
Local $IVANTI2windowName = "Connectez-vous à : ERAS Connect"
Local $IVANTIlocalPath = "C:\Program Files (x86)\Common Files\Pulse Secure\JamUI\"
Local $IVANTIlocalProcess = "Pulse.exe"

Local $ENTRUSTwindowName = "Entrust IdentityGuard Token"
Local $ENTRUSTlocalPath = "C:\Program Files (x86)\Entrust\IdentityGuard Soft Token\"
Local $ENTRUSTlocalProcess = "SoftToken.exe"

Local $info1 = "Pour modifier vos codes allez dans le répertoire : " & $absolutePathAutoITFolder
Local $onRun = "L'exécution est en cours ...  " & @CRLF & @CRLF &  "Ne touchez pas votre clavier ni votre souris durant l'exécution !" & @CRLF & @CRLF & $info1
Local $networkDetected = "Vous êtes actuellement sur un réseau CGI." & @CRLF & @CRLF &  "Lancement du VPN annulé !" & @CRLF & @CRLF &  $info1

; CONFIG
FolderExist($absolutePathAutoITFolder)
CheckConfig($IVANTIlocalPath, $IVANTIlocalProcess)
CheckConfig($ENTRUSTlocalPath, $ENTRUSTlocalProcess)

FileExist($absolutePathAutoITFolder & $absolutePathAutoITUsernameFile, $info1 & @CRLF & @CRLF & "Entrer votre nom d'utilisateur IVANTI (prenom.nom) :")
$IVANTIUsername = ReadCodeContentFile($absolutePathAutoITFolder & $absolutePathAutoITUsernameFile)

FileExist($absolutePathAutoITFolder & $absolutePathAutoITIVANTIFile, $info1 & @CRLF & @CRLF & "Entrer votre code NVP IVANTI (8 chiffres) :")
$IVANTI_NVP = ReadCodeContentFile($absolutePathAutoITFolder & $absolutePathAutoITIVANTIFile)

FileExist($absolutePathAutoITFolder & $absolutePathAutoITENTRUSTFile, $info1 & @CRLF & @CRLF & "Entrer votre code PIN (ou NIP) ENTRUST (4 chiffres) :")
$ENTRUST_PIN = ReadCodeContentFile($absolutePathAutoITFolder & $absolutePathAutoITENTRUSTFile)

If IsCGINetwork() Then 
	SplashTextOn($progName, $networkDetected, 500, 130, -1, -1, 16)
	Sleep(5000)
	SplashOff()
Else
	SplashTextOn($progName, $onRun, 500, 130, -1, -1, 16)
	WinSetOnTop($progName, "", 1)

	; ENTRUST
	CloseApp($ENTRUSTwindowName, $ENTRUSTlocalProcess)
	OpenApp($ENTRUSTwindowName, $ENTRUSTlocalPath, $ENTRUSTlocalProcess)
	SetPinENTRUST($ENTRUSTwindowName, $ENTRUST_PIN)
	$ENTRUST_Code = GetCodeENTRUST($ENTRUSTwindowName)

	; IVANTI
	WinClose($IVANTI2windowName)
	CloseApp($IVANTIwindowName, $IVANTIlocalProcess)
	OpenApp($IVANTIwindowName, $IVANTIlocalPath, $IVANTIlocalProcess)
	ClickToConnexionIVANTI($IVANTIwindowName)
	SetFormIVANTI($IVANTI2windowName, $IVANTIUsername, $IVANTI_NVP & $ENTRUST_Code)
	ClickToFinalConnexionIVANTI($IVANTI2windowName)

	; CLOSE
	WinClose($ENTRUSTwindowName)
	WinClose($IVANTIwindowName)
	Sleep(8000)

	SplashOff()
EndIf
;///////////////////////////////////////////////////////////////////////////////////////////////////////////
; CONFIGURATION
;///////////////////////////////////////////////////////////////////////////////////////////////////////////

; Check if AutoIT INVANTI file exist and create it
Func CheckConfig($path, $process)
	If Not FileExists($path & $process) Then
		MsgBox(16, "Erreur", "Programme introuvable :  " & $process & @CRLF & "Répertoire d'installation: " & $path & @CRLF & @CRLF & "Installer le programme dans le bon répertoire avant de continuer.")
		Exit
	EndIf
EndFunc

; Get input user
Func InputUser($message)
	Local $valeur = InputBox("Saisie", $message)
	If @error Then
	    MsgBox(16, "Erreur", "Vous devez entrer une valeur pour continuer.")
	    Exit
	EndIf
	Return $valeur
EndFunc

; Check if AutoIT folder exist and create it
Func FolderExist($absolutePath)
	If Not FileExists($absolutePath) Then
		If Not DirCreate($absolutePath) Then
			MsgBox(16, "Erreur", "Impossible de créer le dossier : " & $absolutePath & ", les droits en écritures sont restreints. Exécuter ce programme en tant qu'administrateur pour obtenir les droits requis.")
		EndIf
	EndIf		
EndFunc

; Check if AutoIT INVANTI file exist and create it
Func FileExist($absolutePath, $message)
	If Not FileExists($absolutePath) Then
		Local $fileHandle = FileOpen($absolutePath, 2)
		If $fileHandle = -1 Then
			MsgBox(16, "Erreur", "Impossible de créer le fichier : " & $absolutePath & ", les droits en écritures sont restreints. Exécuter ce programme en tant qu'administrateur pour obtenir les droits requis.")
		Else
			$codeUser = InputUser($message)
			FileWrite($fileHandle, $codeUser)
			FileClose($fileHandle)
		EndIf
	EndIf
EndFunc

; Read content file
Func ReadCodeContentFile($absolutePath)
	Local $contentFile = FileRead($absolutePath)
	If @error Then
		MsgBox(16, "Erreur", "Impossible de trouver le fichier : " & $absolutePath & ", les droits en écritures sont restreints. Exécuter ce programme en tant qu'administrateur pour obtenir les droits requis.")
		Exit
	EndIf
	Return $contentFile
EndFunc

; Add value on a file
Func AddOnFile($absolutePath, $value)
	If Not FileWrite($absolutePath, $codeUser) Then
		MsgBox(16, "Erreur", "Impossible d'écrire dans le fichier :" & $absolutePath & ", les droits en écritures sont restreints. Exécuter ce programme en tant qu'administrateur pour obtenir les droits requis.")
	EndIf
EndFunc

;///////////////////////////////////////////////////////////////////////////////////////////////////////////
; CGI Network detection
;///////////////////////////////////////////////////////////////////////////////////////////////////////////

Func IsStringFindedInCLICommand($stringToFind, $CLICommand)
	Local $ipconfigOutput = Run(@ComSpec & ' /c '&$CLICommand, "", @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD)
	Local $output = ""
	While 1
		$line = StdoutRead($ipconfigOutput)
		If @error Then ExitLoop
		$output &= $line
	WEnd
	ProcessClose($ipconfigOutput)
	;ConsoleWrite($output) ;DEBUG !!!
	If StringInStr($output, $stringToFind) Then
		return true
	Else
		return false
	EndIf
EndFunc

Func GetWIFINetworkName()
	Local $networkName = "";
	Local $ipconfigOutput = Run(@ComSpec & ' /c netsh wlan show interfaces', "", @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD)
	Local $output = ""
	While 1
		$line = StdoutRead($ipconfigOutput)
		If @error Then ExitLoop
		If StringInStr($line, "SSID")  Then 
			Local $subString = "SSID                  ÿ: "
			Local $iPosition = StringInStr($line, $subString)
			If $iPosition > 0 Then
			    Local $textFinal = StringMid($line, $iPosition + StringLen($subString))
				Local $aLines = StringSplit($textFinal, @CRLF, 1)
				$networkName = $aLines[1]
				;ConsoleWrite($networkName) ;DEBUG !!!		   
			EndIf
		EndIf
		$output &= $line
	WEnd
	;ConsoleWrite($output) ;DEBUG !!!
	ProcessClose($ipconfigOutput)
	Return $networkName
EndFunc

Func IsConnectedToEWNET()
	Return GetWIFINetworkName() == "EWNET";
EndFunc

Func IsConnectedToVPN()
	Return IsStringFindedInCLICommand("Adresse IPv4. . . . . . . . . . . . . .: 10.209.", "ipconfig")
EndFunc

Func IsConnectedToEthernet()
	Return IsStringFindedInCLICommand("Passerelle par d‚faut. . . .ÿ. . . . . : 10.82.", "ipconfig")
EndFunc

Func IsCGINetwork()
	Return IsConnectedToEWNET() Or IsConnectedToVPN() Or IsConnectedToEthernet()
EndFunc

;///////////////////////////////////////////////////////////////////////////////////////////////////////////
; LAUNCH APP
;///////////////////////////////////////////////////////////////////////////////////////////////////////////

; Close App
Func CloseApp($windowName, $localProcess)
	; Check if window open's and close it
	If WinExists($windowName) Then
		WinClose($windowName)
	EndIf
	; Check if process open's and close it
	Local $pid = ProcessExists($localProcess)
	If $pid Then
		ProcessClose($pid)
		Sleep(500)
	EndIf
EndFunc

; Open App
Func OpenApp($windowName, $localPath, $localProcess)
	ShellExecute($localPath & $localProcess, "","","",@SW_HIDE)
	WinSetOnTop($progName, "", 1)
	Sleep(500)
EndFunc

; Find all fileds names availables
Func GetFieldNames($windowName)
	Local $listFields[0]
	Local $hWnd = WinGetHandle($windowName)
	If IsHWnd($hWnd) Then
		Local $sClassList = WinGetClassList($hWnd)
		If $sClassList <> "" Then
			Local $aClasses = StringSplit($sClassList, @LF)
			For $i = 1 To $aClasses[0]
				_ArrayAdd($listFields, $aClasses[$i])
			Next
		EndIf
	EndIf
	Return $listFields
EndFunc

; Find position of control
Func GetControlPosition($windowName, $sClassNN)
    Local $hWnd = WinGetHandle($windowName)
    If Not $hWnd Then
		;ConsoleWrite("La fenêtre avec le titre '" & $windowName & "' n'a pas été trouvée.")
        Return ""
    EndIf
    Local $hCtrl = ControlGetHandle($hWnd, "", "[CLASSNN:" & $sClassNN & "]")
    If $hCtrl Then
        Local $aPos = ControlGetPos($hWnd, "", $hCtrl)
        Return $aPos[0] & "-" & $aPos[1]
    Else
		;ConsoleWrite("Le champ avec la classe '" & $sClassNN & "' n'a pas été trouvée.")
        Return ""
    EndIf
EndFunc

; Set PIN code ENTRUST
Func SetPinENTRUST($ENTRUSTwindowName, $ENTRUST_PIN)
	WinWaitActive($ENTRUSTwindowName)
	Send($ENTRUST_PIN)
	WinSetOnTop($progName, "", 1)
EndFunc

; Get the final code ENTRUST
Func GetCodeENTRUST($ENTRUSTwindowName)
	WinActivate($ENTRUSTwindowName)
	WinSetOnTop($progName, "", 1)
	Sleep(500)
	ControlFocus($ENTRUSTwindowName, "", "WindowsForms10.STATIC.app.0.bb8560_r9_ad16")
	Local $value = ControlGetText($ENTRUSTwindowName, "", "WindowsForms10.STATIC.app.0.bb8560_r9_ad16")
	$value = StringReplace($value, " ", "")
	Return $value
EndFunc

; Click on the button connexion
Func ClickToConnexionIVANTI($IVANTIwindowName)
	WinActivate($IVANTIwindowName)
	WinSetOnTop($progName, "", 1)
	Sleep(500)
	Local $btnName = ""
	If (StringInStr(GetControlPosition($IVANTIwindowName, "JAM_BitmapButton8"), "-160")) Then
		$btnName = "JAM_BitmapButton8"
	EndIf
	If (StringInStr(GetControlPosition($IVANTIwindowName, "JAM_BitmapButton10"), "-160")) Then
		$btnName = "JAM_BitmapButton10"
	EndIf
	ControlFocus($IVANTIwindowName, "", $btnName)
	ControlClick($IVANTIwindowName, "", $btnName)
	;NOTE : Changement de nom du bouton aléatoire selon les utilisateurs !!!
	;JAM_BitmapButton8 => 12/12/23
	;JAM_BitmapButton10 => 13/12/23 Correctif fait, ajout d'une vérif sur la position, en test
EndFunc

; Set IVANTI form
Func SetFormIVANTI($IVANTI2windowName, $IVANTIUsername, $password)
	WinWaitActive($IVANTI2windowName)
	WinSetOnTop($progName, "", 1)
	Sleep(500)
	$fieldNames = GetFieldNames($IVANTI2windowName)
	For $i = 0 To UBound($fieldNames) - 1
		;ConsoleWrite($fieldNames[$i] & @CRLF) ;DEBUG !!!
		If StringInStr($fieldNames[$i], "ATL") Then
			;ConsoleWrite($fieldNames[$i] & @CRLF) ;DEBUG !!!
			Local $aPosUsername = GetControlPosition($IVANTI2windowName, $fieldNames[$i] & "1")
			Local $aPosPassword = GetControlPosition($IVANTI2windowName, $fieldNames[$i] & "2")
			;ConsoleWrite($aPosUsername & @CRLF) ;DEBUG !!!
			;ConsoleWrite($aPosPassword & @CRLF) ;DEBUG !!!
			If StringInStr($aPosUsername, "-134") And StringInStr($aPosPassword, "-202") Then
				ControlSetText($IVANTI2windowName, "", $fieldNames[$i] & "1", $IVANTIUsername)
				ControlSetText($IVANTI2windowName, "", $fieldNames[$i] & "2", $password)				
			EndIf
		EndIf
	Next
		; NOTE : Changement de nom du champ aléatoire selon les utilisateurs !!!
		;ATL:00F3"B4201 => 07/10/23
		;ATL:009DB4201 => 09/10/23 Correctif fait
		;ATL:00??E6841 => 13/12/23 Correctif fait, ajout d'une vérif sur la position, en test
EndFunc

; Click on the final button connexion 
Func ClickToFinalConnexionIVANTI($IVANTI2windowName)
	WinActivate($IVANTI2windowName)
	WinSetOnTop($progName, "", 1)
	Sleep(500)
	ControlFocus($IVANTI2windowName, "", "JAM_BitmapButton1")
	ControlClick($IVANTI2windowName, "", "JAM_BitmapButton1")
EndFunc
;///////////////////////////////////////////////////////////////////////////////////////////////////////////