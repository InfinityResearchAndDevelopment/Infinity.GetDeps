#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SystemRound.256.ico
#AutoIt3Wrapper_Outfile=GetDeps.32.exe
#AutoIt3Wrapper_Outfile_x64=GetDeps.64.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Description=Shows missing/failed File/Registry calls
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=InfinityResearchAndDevelopment 2017
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         BiatuAutMiahn[@outlook.com]


-If Exec, and deps needed
#ce ----------------------------------------------------------------------------
#Include <WinAPIProc.au3>
#Include <GuiListView.au3>
#Include <Array.au3>
#include <WinAPIReg.au3>
;#Include <Includes\Reg.au3>

; Script Start - Add your code below here
Global $sExec, $aDeps[0][2], $aRegDeps[0], $iPID
Global $sData=@ScriptDir&"\Data"; Data dir for procmon.exe
Global $iReg, $aFilter[0],$aDeDupEvt[0],$iDeDupEvt=0
If Not FileExists($sData) Then DirCreate($sData)

If @OSArch="X64" Then
    If Not @AutoItX64 Then
        MsgBox(16,"Error","You must use the 64 Bit Version on a 64 Bit OS.")
        Exit 1
    EndIf
    $sExec="ProcMon64.exe"
Else
    If @AutoItX64 Then
        MsgBox(16,"Error","You must use the 32 Bit Version on a 32 Bit OS."); This should not happen, but eh, whatever.
        Exit 1
    EndIf
    $sExec="ProcMon.exe"
EndIf

;If Exec not exist, extract it.
If Not FileExists($sData&"\"&$sExec) Then
    If @OSArch="X64" Then
        MsgBox(16,"Error","Procmon64.exe is required to run this program."&@CRLF&"    -Procmon64.exe can be obtained by downloading, and"&@CRLF&"        running ProcMon.exe from Microsoft TechNet."&@CRLF&"    -While running ProcMon.exe, goto your TEMP directory"&@CRLF&"        and copy Procmon64.exe to "&$sData)
    Else
        MsgBox(16,"Error","Procmon.exe is required to run this program."&@CRLF&"    -Procmon.exe can be obtained by from Microsoft TechNet."&@CRLF&"    -ProcMon.exe must be placed in "&$sData)
    EndIf
    Exit 1
EndIf
ConsoleWrite("Copyright InfinityResearchAndDevelopment 2017"&@CRLF)
ConsoleWrite("Website: InfinityCommunicationsGateway.net"&@CRLF)
ConsoleWrite("Git: https://github.com/InfinityResearchAndDevelopment/Infinity.GetDeps"&@CRLF)
ConsoleWrite("Contact: BiatuAutMiahn@outlook.com"&@CRLF&@CRLF)
;Load up Filter
$hEvtFilter=FileOpen("EvtFilter.ini")
$aEvtFilter=FileReadToArray($hEvtFilter)
$iEvtFilter=UBound($aEvtFilter,1)
;If ProcMon Running, Terminate it.
While ProcessExists($sExec)
    RunWait($sData&"\"&$sExec&" /Terminate /AcceptEula")
WEnd

;Load SourceHives
;~ ConsoleWrite("Loading Registry Hives..."&@CRLF)
;~ $sHivesRoot=$sData&"\SourceHives"
;~ Local $aHives[]=["SOFTWARE","SYSTEM","DEFAULT"]
;~ For $i=0 To UBound($aHives)-1
;~     ConsoleWrite($sHivesRoot&"...")
;~     ConsoleWrite(RunWait("Reg Load HKLM\Win_"&$aHives[$i]&' "'&$sHivesRoot&"\"&$aHives[$i]&'"', @SystemDir,@SW_HIDE,65536)&@CRLF);
;~ Next
;~ ConsoleWrite(@CRLF)
;~ ConsoleWrite("Starting Caputre..."&@CRLF)
;Run ProcMon.
$iPID=Run($sData&"\"&$sExec&" /acceptEula /LoadConfig Conf.pmc /Quiet",$sData,@SW_HIDE)
ProcessWait($iPID)

;Get Window Handle
Global $aWnd=_WinAPI_EnumProcessWindows($iPID,False)
If Not IsArray($aWnd) Then _Exit(); If we cant get it, exit.
Global $hWnd
For $i=1 to $aWnd[0][0]
	If $aWnd[$i][1]="PROCMON_WINDOW_CLASS" Then
		$hWnd=$aWnd[$i][0]
	EndIf
Next
Global $hListView=ControlGetHandle($hWnd,"","SysListView321"); Get ListView Handle.
Local $iMax=0; Max Entries in ListView.
Local $iLast=0; Number of Entries we've captured.
;AdlibRegister("_EventWatch",250) ;Check for various conditions.
;_InitReg()
Local $sFile
Local $sFilePath
While Sleep(500); every 1/4th second
    $iEvents=_GetEventCount()
	If $iEvents>=50000 Then
		_Clear()
        ;Dim $aDeps[0][2]
        $iMax=0
        $iLast=0
	EndIf
	$iMax=_GUICtrlListView_GetItemCount($hListView)
	If $iMax>$iLast And $iMax>0 Then; if there are more items than read, and more than 0 items...
		For $i=$iLast to $iMax; Parse from the last read to the last item.
			$sExec=_GUICtrlListView_GetItemText($hListView,$i,1); Process Column.
			$sPath=_GUICtrlListView_GetItemText($hListView,$i,4); Path Column.
			If $sExec=@AutoItExe Then ContinueLoop; If The Process if our self...ignore it or else infinite loop.
			;_AddDep($sPath)
            If $sPath="" Then ContinueLoop
            If IsArray($aEvtFilter) Then
                For $f=0 To $iEvtFilter-1
                    If StringRegExp($sPath,$aEvtFilter[$f]) Then ContinueLoop 2
                Next
            EndIf
            $iDeDupEvt=UBound($aDeDupEvt,1)
            For $d=0 To $iDeDupEvt-1
                If $sPath=$aDeDupEvt[$d] Then ContinueLoop 2
            Next
            If StringRegExp(StringLeft($sPath,4),"HKCR|HKLM|HKCU|HKCC|HKU\\") Then
                RegEnumKey($sPath,1)
                If @error<>0 Then ContinueLoop
            ElseIf StringLeft($sPath,2)="X:" Then
                    If FileExists($sPath) Then ContinueLoop
;~                     $sFile=StringTrimLeft($sPath,StringInStr($sPath,"\",0,-1))
;~                     $sFilePath=StringLeft($sPath,StringInStr($sPath,"\",0,-1)-1)
;~                     _WinAPI_PathFindOnPath($sFile)
;~                     If Not @error Then ContinueLoop
            EndIf
            ReDim $aDeDupEvt[$iDeDupEvt+1]
            $aDeDupEvt[$iDeDupEvt]=$sPath
            ConsoleWrite($sPath&@CRLF)
            FileWriteLine("Log.txt",$sPath)
;~             If StringLeft($aDeDupEvt[$iDeDupEvt],5)="HKCR\" Then
;~                 ;_RegImport($aDeDupEvt[$iDeDupEvt])
;~             EndIf
		Next
		$iLast=$iMax; We have caught up

    ElseIf $iMax<$iLast Then
        $iMax=0
        $iLast=0
	EndIf
;~     If UBound($aRegDeps,1)>0 Then
;~         For $i=0 To UBound($aRegDeps,1)-1
;~             If $aRegDeps[$i]="" Then ContinueLoop
;~             ;ConsoleWrite("[RDL] "&$aRegDeps[$i]&"...")
;~             ;$iReg=_RegDll($aRegDeps[$i])
;~             If $iReg=5 Or $iReg=0 Then
;~                 $aRegDeps[$i]=""
;~             EndIf
;~             ;ConsoleWrite(@CRLF)
;~         Next
;~         Dim $aRegDeps[0]
;~     EndIf
	If Not ProcessExists($iPID) Then _Exit()
WEnd

Func _RegImport($sKey)
    $sKey=StringReplace($sKey,"HKCR\","Classes\")
    Local $sKeyPath=StringLeft($sKey,StringInStr($sKey,"\",0,-1)-1)
    Local $sKeyName=StringTrimLeft($sKey,StringInStr($sKey,"\",0,-1))
    ;ConsoleWrite('  Importing HKLM\Win_SOFTWARE\'&$sKeyPath&'...'&@CRLF)
    ;ConsoleWrite("  "&$sKeyName&@CRLF)
    ;ConsoleWrite("  "&$sKeyPath&@CRLF)
    ;Return
    Local $hSrcKey=_WinAPI_RegOpenKey($HKEY_LOCAL_MACHINE,'Win_SOFTWARE\'&$sKeyPath,$KEY_READ)
    If @error Then Return SetError(1,1,0)
    Local $hDestKey=_WinAPI_RegCreateKey($HKEY_LOCAL_MACHINE,'SOFTWARE\'&$sKeyPath,$KEY_ALL_ACCESS)
    If @error Then Return SetError(2,1,0)
    ConsoleWrite('  Importing HKLM\Win_SOFTWARE\'&$sKeyPath&'\'&$sKeyName&'...')
    ConsoleWrite(_WinAPI_RegCopyTreeEx($hSrcKey,$sKeyName,$hDestKey)&@CRLF)
EndFunc

;
;~ Func _ProcMon()
;~ EndFunc

;~ Func _AddDep($sPath)
;~     If $sPath="" Then Return; We can ignore empty data.
;~     If UBound($aRegDeps,1)>0 Then
;~         For $i=0 To UBound($aRegDeps,1)-1
;~             If $aRegDeps[$i]="" Then ContinueLoop
;~             If $aRegDeps[$i]=$sPath Then Return
;~         Next
;~     EndIf
;~     If StringLeft($sPath,2)="HK" Then
;~         Return
;~         ;$iReg=1
;~     Else
;~         ;$iReg=0
;~         $sPath=StringTrimLeft($sPath,3); Strip the "X:" we don't care about the drive letter, just the realative path.
;~     EndIf
;~     $iMax=UBound($aDeps,1); Get the size of the Dep array/list.
;~ 	For $i=0 to $iMax-1
;~ 		If $aDeps[$i][0]=$sPath Then Return; If it already exists...ignore it.
;~     Next
;~     ReDim $aDeps[$iMax+1][2]; Extend array by 1 to fit another dep.
;~     $aDeps[$iMax][0]=$sPath

;~     ; Apply Filter.
;~     ;$sPath=_Filter($sPath)
;~     If $sPath="" Then Return
;~     If $iReg Then
;~         ;Return
;~         If StringLeft($sPath,4)="HKU\" Then Return
;~         If StringLeft($sPath,17)="HKLM\New_Software" Then Return
;~         If StringLeft($sPath,15)="HKLM\New_System" Then Return
;~         ;RegEnumKey($sPath,1)
;~         ;If Not @error Then Return
;~         ;ReWrite Reg Paths for Import.
;~         Local $sSrcRegPath
;~         If StringLeft($sPath,13)="HKLM\Software" Then $sSrcRegPath="HKLM\New_Software"&StringTrimLeft($sPath,13)
;~         If StringLeft($sPath,11)="HKLM\System" Then $sSrcRegPath="HKLM\New_System"&StringTrimLeft($sPath,11)
;~         RegEnumKey($sSrcRegPath,1)
;~         If @error Then
;~             ;_AddFilter($sPath)
;~             Return
;~         EndIf
;~         ;ConsoleWrite("[Ign] "&StringTrimLeft($sSrcRegPath,5)&@CRLF)
;~         ConsoleWrite("[Cpy] "&$sPath&"...")
;~         $sPath=StringTrimLeft($sPath,5)
;~         ;Copy Reg
;~         $hSrcReg=_WinAPI_RegOpenKey($HKEY_LOCAL_MACHINE,StringTrimLeft($sSrcRegPath,5),$KEY_READ)
;~         ConsoleWrite(Int(@Error<>1))
;~         $hDstReg=_WinAPI_RegCreateKey($HKEY_LOCAL_MACHINE,$sPath)
;~         ConsoleWrite(Int(@Error<>1))

;~         _WinAPI_RegCopyTree($HKEY_LOCAL_MACHINE,StringTrimLeft($sSrcRegPath,5),$hDstReg)
;~         ConsoleWrite(Int(@Error<>1))
;~         _WinAPI_RegCloseKey($hSrcReg,False)
;~         ConsoleWrite(Int(@Error<>1))
;~         _WinAPI_RegCloseKey($hDstReg,True)
;~         ConsoleWrite(Int(@Error<>1)&@CRLF)
;~     Else
;~         ;Ignore if is already exists, or not available from source. Can modify later for ini exclues.
;~         If FileExists("X:\"&$sPath) Then Return
;~         If StringInStr(FileGetAttrib($sSrc&"\"&$sPath),"D") Then Return
;~         If Not FileExists($sSrc&"\"&$sPath) Then
;~             ;_AddFilter($sPath)
;~             ;ConsoleWrite("[Ign] "&$sPath&@CRLF)
;~             Return
;~         EndIf
;~         $sFile=StringTrimLeft($sPath,StringInStr($sPath,"\",0,-1)); Get FileName.
;~         $sPath=StringTrimRight($sPath,StringLen($sPath)-StringInStr($sPath,"\",0,-1)+1); Get path to file.
;~         ConsoleWrite("[Cpy] "&$sPath&"\"&$sFile&"...");[Cpy] <...>...0101, Binary, 1=file copied, 0=file not copied.
;~                                                       ; 1=Src->X:   2=Src.mui->X:   3=Src->Dep   4=Src.mui->Dep

;~         _CopyDep($sSrc&"\"&$sPath&"\"&$sFile,"X:\"&$sPath&"\"&$sFile);Copy to X:\...
;~         _CopyDep($sSrc&"\"&$sPath&"\en-US\"&$sFile&".mui","X:\"&$sPath&"\en-US\"&$sFile&".mui");Copy mui
;~         ConsoleWrite(@CRLF)
;~         ConsoleWrite("[Reg] "&$sPath&"\"&$sFile&"...")
;~         _RegDll($sPath&'\'&$sFile)
;~         ConsoleWrite(@CRLF)

;~     EndIf
;~ EndFunc

Func _RegDll($sPath)
    If StringRight($sPath,4)=".msc" Then
        ConsoleWrite(0)
        Return False
    EndIf
    $iMax=UBound($aRegDeps,1)
    Local $iTimer=TimerInit()
    $iRet=RunWait('Regsvr32.exe /s "X:\'&$sPath&'"')
    If $iRet=4 Then; DllRegisterServer Not found.
        $iRet=RunWait('Regsvr32.exe /s /i "X:\'&$sPath&'"')
        If $iRet=4 Then; DllInstall not Found.
            ConsoleWrite(0)
            Return False
        ElseIf $iRet=3 Then; File Not Found. unmet Deps.
            ReDim $aRegDeps[$iMax+1]
            $aRegDeps[$iMax]=$sPath
            ;RunWait('Regsvr32.exe /i "X:\'&$sPath&'"')
            ConsoleWrite(3)
            Return 3
;~             While Sleep(250)
;~                 If TimerDiff($iTimer)>=5000 Then
;~                     ConsoleWrite(3)
;~                     Return False
;~                 EndIf
;~                 If FileExists($sPath) Then ExitLoop
;~             WEnd
;~             If Not _RegDll($sPath) Then
;~                 ConsoleWrite(3)
;~                 Return False
;~             EndIf
        EndIf
        ConsoleWrite(2)
        Return True
    ElseIf $iRet=3 Then; File Not Found.
        ReDim $aRegDeps[$iMax+1]
        $aRegDeps[$iMax]=$sPath
        ConsoleWrite(3)
        Return 3
;~         While Sleep(250)
;~             If TimerDiff($iTimer)>=5000 Then
;~                 ConsoleWrite(3)
;~                 Return False
;~             EndIf
;~             If FileExists($sPath) Then ExitLoop
;~         WEnd
;~         If Not _RegDll($sPath) Then
;~             ConsoleWrite(3)
;~             Return False
;~         EndIf
    ElseIf $iRet<>0 Then
        ConsoleWrite($iRet)
        Return False
    EndIf
    ConsoleWrite(1)
    Return True
EndFunc

Func _CopyDep($sCpySrc,$sCpyDst)
    If FileExists($sCpySrc) Then
        If FileCopy($sCpySrc,$sCpyDst,1+8) Then
            $hFile=FileOpen($sCpyDst,0)
            FileFlush($hFile)
            FileClose($hFile)
            ConsoleWrite("1")
        Else
            ConsoleWrite("0")
        EndIf
    Else
        ConsoleWrite("0")
    EndIf
EndFunc

Func _Exit()
	ProcessClose($iPID)
    ;_CloseReg()
    Exit
EndFunc

Func _EventWatch()
	$iEvents=_GetEventCount()
	If $iEvents>=20000 Then
		Call("_Clear")
		_GetEventCount()
	EndIf
EndFunc

Func _GetEventCount(); Get number of events ProcMon Captured.
	Local $iEvents=0
	$sStatus=StatusbarGetText($hWnd,"",1)
	$iEvents=StringRegExp($sStatus,"The current filter excludes all (\d{1,3}(?:,\d{1,3})?) events",1)
	If @Error Then
		$iEvents=StringRegExp($sStatus,"Showing \d{1,3}(?:,\d{1,3})? of (\d{1,3}(?:,\d{1,3})?) events",1)
		If @Error Then
			$iEvents=0
			Return $iEvents
		EndIf
	EndIf
	Return StringReplace($iEvents[0],",","")
EndFunc

Func _Clear(); Clear ProcMon Log.
	;AdlibUnRegister("_EventWatch"); Stop EventWatcher so its not called too many times.
	ConsoleWrite("[Inf] Clearing Events...")
	ControlClick($hWnd,"","ToolbarWindow321","left",1,16*6,15)
	ControlClick($hWnd,"","ToolbarWindow321","left",1,16*8,15)
    Sleep(500)
	ControlClick($hWnd,"","ToolbarWindow321","left",1,16*6,15)
    ;Dim $aDeps[0][2]
	;AdlibRegister("_EventWatch",250)
	ConsoleWrite("Done"&@CRLF&@CRLF)
EndFunc

Func _Filter($sPath)
    Local Static $iFilterLastSize
    Local Static $iFilterSize=FileGetSize($sData&"\Filter.ini")
    If Not IsArray($aFilter) Or $iFilterSize<>$iFilterLastSize Then; Load filter.ini on first run, or when filesize of filter.ini changes.
        $aFilter=FileReadToArray($sData&"\Filter.ini")
        ;_ArrayDisplay($aFilter,$sPath)
        $iFilterLastSize=$iFilterSize
    EndIf
    For $i=0 To UBound($aFilter,1)-1
        If StringLeft($aFilter[$i],1)=";" Then ContinueLoop; Skip Comment lines.
        If $aFilter[$i]="" Then ContinueLoop
        If $aFilter[$i]=$sPath Then $sPath=""
    Next
    Return $sPath
EndFunc

Func _AddFilter($sPath)
    $iMax=UBound($aFilter,1)
    For $i=0 To $iMax-1
        If $aFilter[$i]=$sPath Then Return
    Next
    ReDim $aFilter[$iMax+1]
    $aFilter[$iMax]=$sPath
    ;FileWriteLine($sData&"\Filter.ini",$sPath)
EndFunc

;~ Func _InitReg()
;~     ConsoleWrite("[Inf] Loading Registry...")
;~     ConsoleWrite(_RegLoadHive($sSrc&"\Windows\System32\config\SOFTWARE", "HKLM\New_Software"))
;~     ConsoleWrite(_RegLoadHive($sSrc&"\Windows\System32\config\SYSTEM", "HKLM\New_System"))
;~     ConsoleWrite(_RegLoadHive($sSrc&"\Windows\System32\config\DEFAULT", "HKLM\New_DefaultUser"))
;~     ConsoleWrite(@CRLF)
;~ EndFunc

;~ Func _CloseReg()
;~     ConsoleWrite("[Inf] UnLoading Registry...")
;~     ConsoleWrite(_RegUnloadHive("HKLM\New_Software"))
;~     ConsoleWrite(_RegUnloadHive("HKLM\New_System"))
;~     ConsoleWrite(_RegUnloadHive("HKLM\New_DefaultUser"))
;~     ConsoleWrite(@CRLF)
;~ EndFunc
