#Persistent
#NoEnv
#SingleInstance Force ; 중복 실행 방지
SetBatchLines, -1 ; 스크립트 실행 속도 최대화 (대량의 텍스트 런 처리 시 필수)
CoordMode, Pixel, Screen
SendMode Input
CoordMode, ToolTip, Screen
SetWorkingDir %A_ScriptDir%
SetMouseDelay,-1 ;마우스딜레이 조절
SetDefaultMouseSpeed,0 ;마우스속도 조절
SetKeyDelay, -1
FileEncoding, UTF-8
;★★★★★★★★★파일을 메모장에서 다른이름으로 저장할땐 모든파일(*.*) UTF-8(BOM)    *.ahk으로 저장하기



;--- [2025-10-22 지침 반영: v1 문법] ---

;==============================================================================
; 0. 기본 설정 및 경로 정의 (최상단)
;==============================================================================
Username       := "UNICEF-BECLAW"
RepoName       := "UNICEF-Shortcut"
유니세프구매링크 = https://signist.gumroad.com/l/unicef-key

; ★ 최신버전 이부분을 수정하고 깃허브의 최신버전 텍스트를 이것과 같도록 수정하면 해결됨
CurrentVersion := "UNICEF-Shortcut-Key-2026-02-03-1.exe"
단축키프로그램버전=ver. %CurrentVersion%

ProjectID      := "unicef-key"

; 로컬 INI 파일 경로 정의
IniPath        := A_ScriptDir . "\UNI-Value.ini"
;★★★★★★★★★이파일은 메모장에서 인코딩 UTF-16LE으로 다른이름으로 저장후 깃허브에 직접 업로드해야 잘돌아감

; [최종 진화] 3개의 파일을 단 하나의 마스터 파일로 통일!
; ★ 수정: 깃허브 실제 파일명에 맞춰 "-ID-" 를 추가했습니다.
BaseURL          := "https://raw.githubusercontent.com/" . Username . "/" . RepoName . "/main/"
ConfigURL_Master := BaseURL . "UnicefConfig-Web-ID-Master.ini"
TempIni_Master   := A_Temp . "\RemoteUnicefConfig-Master.ini"

;==============================================================================
; 1. 서버(GitHub)에서 통합 설정 파일 1번만 다운로드 (속도 극대화)
;==============================================================================
ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
UrlDownloadToFile, %ConfigURL_Master%, %TempIni_Master%

if (ErrorLevel) {
    MsgBox, 262208, Message, %msgboxuni_0001%
    RestoreCursors()       ; 커서 복구

    ; ★오프라인상태면 실행가능하게 하기
    IniRead, SavedToken, %IniPath%, 인증정보, 토큰, ERROR
    if (SavedToken != "ERROR" && SavedToken != "") 
    {
        UserGrade=2
        goto, 초기활성인증확인
    }
    ExitApp
}

; 통합 파일에서 필요한 정보 3가지를 한 번에 싹 읽어옵니다.
IniRead, ClientID, %TempIni_Master%, GoogleAPI, ClientID, ERROR
IniRead, ClientSecret, %TempIni_Master%, GoogleAPI, ClientSecret, ERROR
IniRead, LatestExeName, %TempIni_Master%, Versiontxt, Version, ERROR

RestoreCursors()       ; 커서 복구 완료

; 보안을 위해 사용 후 즉시 임시 파일 삭제
if FileExist(TempIni_Master)
    FileDelete, %TempIni_Master%

; 정보 확인 (필수값 누락 체크)
if (ClientID = "ERROR" or ClientSecret = "ERROR") {
    MsgBox, 262208, Message, %msgboxuni_0002%
    ExitApp
}
if (LatestExeName = "ERROR" or LatestExeName = "") {
    ; 버전 정보를 제대로 읽지 못했을 때의 예외 처리 (필요시 추가)
}


;==============================================================================
; 2. 원격 INI 파일 자동 다운로드 및 설정 (최초 1회)
;==============================================================================
if !FileExist(IniPath)
{
    BaseURL2        := "https://raw.githubusercontent.com/" . Username . "/" . RepoName . "/main/"
    IniDownloadURL2 := BaseURL2 . "UNI-Value.ini"

    ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경

    ; 2) 서버에서 INI 파일 다운로드
    UrlDownloadToFile, %IniDownloadURL2%, %IniPath%
    
    ; 3) 커서 정상 복구
    RestoreCursors() 
    
    ; 4) 다운로드 성공 여부 검증 및 처리
    if (ErrorLevel)
    {
        MsgBox, 262208, Message, %msgboxuni_0003%
        return 
    }
    else
    {
        FileSetAttrib, +H, %IniPath%
        MsgBox, 262208, Message, %msgboxuni_0004%
    }
}


;==============================================================================
; 3. 로그인 및 구독 확인 로직 (개선된 구조)
;==============================================================================
AuthSuccess := false

; 기존 토큰이 있는지 확인
IniRead, SavedToken, %IniPath%, 인증정보, 토큰, ERROR

if (SavedToken != "ERROR" && SavedToken != "") 
{
    ; 토큰이 있다면 갱신 시도
    GoSub, RefreshTokenLogin
    
    if (AuthSuccess = true) {
        ; 갱신 및 인증에 성공했다면 메인 로직으로 이동
        GoSub, 업데이트체크로직
        return  ; <--- 여기서 중단되어야 최초 로그인이 뜨지 않습니다.
    }
}

; --- 위에서 인증 실패 시에만 아래 '최초 1회 원클릭 로그인' 코드가 실행됩니다 ---

if (ClientID = "ERROR" || ClientID = "") {
    MsgBox, 262208, Message, %msgboxuni_0005%
    RestoreCursors()
    ExitApp
}

; ★ 변경: 리디렉션 주소를 구글 콘솔에 등록한 로컬 주소로 지정
RedirectURI := "http://127.0.0.1:18080"

AuthURL := "https://accounts.google.com/o/oauth2/v2/auth?client_id=" . ClientID 
        . "&redirect_uri=" . RedirectURI . "&response_type=code"
        . "&scope=https://www.googleapis.com/auth/userinfo.email"
        . "&access_type=offline&prompt=consent"

RestoreCursors()       ; 커서 복구
MsgBox, 262208, Message, %msgboxuni_0006%
Run, %AuthURL%

;==============================================================================
; [핵심] PowerShell을 이용한 로컬 서버 구동 및 승인 코드 낚아채기
;==============================================================================
PSFile := A_Temp . "\auth_listener.ps1"
psCode = 
(
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:18080/")
$listener.Start()
$context = $listener.GetContext()
$code = $context.Request.QueryString["code"]
$res = $context.Response
$buffer = [System.Text.Encoding]::UTF8.GetBytes("<html><head><meta charset='utf-8'><title>Welcome to UNI-Key</title></head><body style='text-align:center; padding-top:100px; font-family:sans-serif;'><h2>Welcome to UNI-Key</h2><p>Registration has been completed.</p><script>setTimeout(function(){window.close();}, 2000);</script></body></html>")
$res.ContentLength64 = $buffer.Length
$res.OutputStream.Write($buffer, 0, $buffer.Length)
$res.Close()
$listener.Stop()
Write-Output $code
)

; 임시 파일로 만들어서 안전하게 실행
FileDelete, %PSFile%
FileAppend, %psCode%, %PSFile%, UTF-8

; PowerShell 숨김 상태로 실행
shell := ComObjCreate("WScript.Shell")
exec := shell.Exec("powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ . PSFile """")

; 응답(승인 코드)이 올 때까지 프로그램 대기 (InputBox 역할을 대신함)
While (exec.Status = 0) {
    Sleep, 100
}

; 출력된 승인 코드를 AHK 변수로 가져오고 임시 파일 삭제
AuthCode := Trim(exec.StdOut.ReadAll(), " `t`r`n")
FileDelete, %PSFile%

if (AuthCode = "") {
    MsgBox, 262208, Message, %msgboxuni_0007%
    RestoreCursors()
    ExitApp
}

;==============================================================================
; 낚아챈 코드로 실제 토큰 발급받기
;==============================================================================
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("POST", "https://oauth2.googleapis.com/token", false)
whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")

PostData := "code=" . AuthCode . "&client_id=" . ClientID . "&client_secret=" . ClientSecret 
         . "&redirect_uri=" . RedirectURI . "&grant_type=authorization_code"

whr.Send(PostData)
RestoreCursors()       ; 커서 복구

RegExMatch(whr.ResponseText, """refresh_token"":\s*""([^""]+)""", RFMatch)
if (RFMatch1 != "") {
    IniWrite, %RFMatch1%, %IniPath%, 인증정보, 토큰
    GoSub, GetEmailAndAuth
    if (AuthSuccess = true) {
        GoSub, 업데이트체크로직
        return
    }
} else {
    MsgBox, 262208, Message, %msgboxuni_0008%
    RestoreCursors()
    ExitApp
}
return


;==============================================================================
; [서브 루틴] - 인증 및 구독 확인 (토큰 갱신)
;==============================================================================
RefreshTokenLogin:
    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    whr.Open("POST", "https://oauth2.googleapis.com/token", false)
    whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    
    PostData := "client_id=" . ClientID . "&client_secret=" . ClientSecret 
             . "&refresh_token=" . SavedToken . "&grant_type=refresh_token"

    ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
    whr.Send(PostData)
    RestoreCursors()       ; 커서 복구

    ; 응답에 access_token이 포함되어 있는지 확인
    if InStr(whr.ResponseText, "access_token") {
        GoSub, GetEmailAndAuth
    } else {
        AuthSuccess := false ; 토큰 만료 등의 사유로 실패
    }
return


;==============================================================================
; [서브 루틴] - 이메일 정보 가져오기 및 등록 체크
;==============================================================================
GetEmailAndAuth:
    RegExMatch(whr.ResponseText, """access_token"":\s*""([^""]+)""", ATMatch)
    AccessToken := ATMatch1
    if (AccessToken = "")
        return

    ; 유저 이메일 가져오기
    whr.Open("GET", "https://www.googleapis.com/oauth2/v2/userinfo", false)
    whr.SetRequestHeader("Authorization", "Bearer " . AccessToken)

    ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
    whr.Send()
    RestoreCursors()       ; 커서 복구

    RegExMatch(whr.ResponseText, """email"":\s*""([^""]+)""", EmailMatch)
    UserEmail := EmailMatch1
    
    if (UserEmail != "") {
        ; 이메일을 찾았으면 바로 체크하지 말고, '없으면 등록하는' 단계로 이동
        GoSub, CreateUserIfNotExist
    }
return


;==============================================================================
; [서브 루틴] - 신규 사용자 자동 등록 (3개월 무료)
;==============================================================================
CreateUserIfNotExist:
    TargetURL := "https://firestore.googleapis.com/v1/projects/" . ProjectID . "/databases/(default)/documents/users/" . UserEmail
    
    ; 1. 먼저 사용자가 존재하는지 조회
    whr.Open("GET", TargetURL, false)

    ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
    whr.Send()
    RestoreCursors()       ; 커서 복구
    
    ; 2. 404 에러(찾을 수 없음)가 뜨면 -> 신규 가입자임! -> 데이터 생성
    if (InStr(whr.Status, "404")) {
        
        ; 깃허브에서 실시간으로 StartKey Raw 가져오기
        StartKeyURL := "https://raw.githubusercontent.com/UNICEF-BECLAW/UNICEF-Shortcut/refs/heads/main/unicefstartkey"
        
        whr.Open("GET", StartKeyURL, false)
        ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
        whr.Send()
        RestoreCursors()       ; 커서 복구
        
        GitHubStatus := whr.Status
        
        if (GitHubStatus != 200) {
            MsgBox, 262208, Message, %msgboxuni_0009% (Code: %GitHubStatus%)
            RestoreCursors()
            ExitApp
        }
        
        ; 가져온 키값의 앞뒤 공백/줄바꿈 제거
        RealStartKey := Trim(whr.ResponseText, " `t`n`r")
        
        ; 3개월 뒤 날짜 계산 (YYYY-MM-DD)
        ExpiryDate := ""
        FormatTime, Today, , yyyyMMdd
        EnvAdd, Today, 90, Days ; 90일(약 3개월) 추가
        FormatTime, ExpiryDate, %Today%, yyyy-MM-dd

        ; JSON에 적용
        CreateJson =
        (
        {
          "fields": {
            "startkey": { "stringValue": "%RealStartKey%" },
            "Notice1": { "integerValue": "1" },
            "Notice2": { "integerValue": "0" },
            "userGrade": { "integerValue": "2" },
            "isSubscribed": { "booleanValue": true },
            "expiryDate": { "stringValue": "%ExpiryDate%" }
          }
        }
        )

        ; PATCH 메서드로 데이터 생성
        whr.Open("PATCH", TargetURL, false)
        whr.SetRequestHeader("Content-Type", "application/json")

        ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
        whr.Send(CreateJson)
        RestoreCursors()       ; 커서 복구

        StatusVal := whr.Status
        
        if (StatusVal == 200) {
            ; 성공 처리 (알림창 생략)
        } else {
            ResponseMsg := whr.ResponseText
            MsgBox, 262208, Message, %msgboxuni_0010%`n%StatusVal%`n%ResponseMsg%
            RestoreCursors()
            ExitApp
        }
    }
    
    ; 3. 등록 완료 혹은 기존 유저의 경우 최종 권한 체크로 이동
    GoSub, FinalAuthCheck
return


;==============================================================================
; [서브 루틴] - 최종 권한 및 구독 확인
;==============================================================================
FinalAuthCheck:
    TargetURL := "https://firestore.googleapis.com/v1/projects/" . ProjectID . "/databases/(default)/documents/users/" . UserEmail
    whr.Open("GET", TargetURL, false)

    ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
    whr.Send()
    RestoreCursors()       ; 커서 복구
    
    ; Firestore 응답 데이터 확인
    if (InStr(whr.ResponseText, """booleanValue"": true")) {
        
        RegExMatch(whr.ResponseText, """Notice1"":\s*\{\s*""integerValue"":\s*""(\d+)""", Notice1Match)
        Notice1 := Notice1Match1 

        RegExMatch(whr.ResponseText, """Notice2"":\s*\{\s*""integerValue"":\s*""(\d+)""", Notice2Match)
        Notice2 := Notice2Match1 

        RegExMatch(whr.ResponseText, """userGrade"":\s*\{\s*""integerValue"":\s*""(\d+)""", GradeMatch)
        UserGrade := GradeMatch1 

        RegExMatch(whr.ResponseText, """expiryDate"":\s*\{\s*""stringValue"":\s*""([^""]+)""", DateMatch)
        UserExpiryDate := DateMatch1

        AuthSuccess := true
        goto, 초기활성인증확인

    } else {
        AuthSuccess := false

        MsgBox, 262208, Message, %msgboxuni_0011%
        토큰초기화=
        IniWrite, %토큰초기화%, %IniPath%, 인증정보, 토큰
        run, %유니세프구매링크%
        RestoreCursors()       ; 커서 복구
        ExitApp
    }
return

;==============================================================================
; [서브 루틴] - 업데이트 체크 로직
;==============================================================================
업데이트체크로직:
    ; 공백 제거
    LatestExeName := Trim(LatestExeName, " `t`r`n")

    ; 파일명에서 버전 숫자만 추출
    LatestVersionNum := RegExReplace(LatestExeName, "[^0-9]") 
    CurrentVersionNum := RegExReplace(CurrentVersion, "[^0-9]") 

    if (LatestVersionNum > CurrentVersionNum) {
        DownloadURL := "https://raw.githubusercontent.com/" . Username . "/" . RepoName . "/refs/heads/main/" . LatestExeName
        
        MsgBox, 4100, Message, %msgboxuni_0012%.`n`n%LatestExeName%
        IfMsgBox Yes
        {
            ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
            UrlDownloadToFile, %DownloadURL%, %A_ScriptDir%\%LatestExeName%
            RestoreCursors()       ; 커서 복구
            if (ErrorLevel) {
                MsgBox, 262208, Message, %msgboxuni_0013%
                return
            }
            MsgBox, 262208, Message, %msgboxuni_0014%
            RestoreCursors()       ; 커서 복구
            ExitApp
        }
    }
    GoSub, CheckNotices
return

;==============================================================================
; [서브 루틴] - 공지사항(Notice) 확인 및 상태 업데이트
;==============================================================================
CheckNotices:
    BaseDocURL := "https://firestore.googleapis.com/v1/projects/" . ProjectID . "/databases/(default)/documents/users/" . UserEmail

    whr.Open("GET", BaseDocURL, false)
    ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
    whr.Send()
    RestoreCursors()       ; 커서 복구

    if (whr.Status != 200)
        return

    CurrentData := whr.ResponseText

    ; [STEP 1] Notice1 처리
    Notice1_Val := 0
    if RegExMatch(CurrentData, """Notice1""\s*:\s*\{\s*""integerValue""\s*:\s*""(\d+)""", Match)
        Notice1_Val := Match1

    if (Notice1_Val == 1) {
        NoticeURL := BaseURL . "Notice1.txt"
        whr.Open("GET", NoticeURL, false)

        ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
        whr.Send()
        RestoreCursors()       ; 커서 복구
        
        if (whr.Status == 200) {
            MsgBox, 262208, Message, % whr.ResponseText
            
            PatchURL := BaseDocURL . "?updateMask.fieldPaths=Notice1"
            UpdateJson1 =
            (
            {
              "fields": {
                "Notice1": { "integerValue": "0" }
              }
            }
            )
            whr.Open("PATCH", PatchURL, false)
            whr.SetRequestHeader("Content-Type", "application/json")

            ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
            whr.Send(UpdateJson1)
            RestoreCursors()       ; 커서 복구
        }
    }

    ; [STEP 2] Notice2 처리
    Notice2_Val := 0
    if RegExMatch(CurrentData, """Notice2""\s*:\s*\{\s*""integerValue""\s*:\s*""(\d+)""", Match)
        Notice2_Val := Match1

    if (Notice2_Val == 1) {
        NoticeURL := BaseURL . "Notice2.txt"
        whr.Open("GET", NoticeURL, false)

        ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
        whr.Send()
        RestoreCursors()       ; 커서 복구

        if (whr.Status == 200) {
            MsgBox, 262208, Message, % whr.ResponseText
            
            PatchURL := BaseDocURL . "?updateMask.fieldPaths=Notice2"
            UpdateJson2 =
            (
            {
              "fields": {
                "Notice2": { "integerValue": "0" }
              }
            }
            )
            whr.Open("PATCH", PatchURL, false)
            whr.SetRequestHeader("Content-Type", "application/json")
            
            ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경
            whr.Send(UpdateJson2)
            RestoreCursors()       ; 커서 복구
        }
    }

return









초기활성인증확인:





gosub, CheckNotices


RestoreCursors()       ; 커서 복구






; 1. 기본 메뉴 삭제 및 전용 메뉴 구성
;Menu, Tray, NoStandard ; 오토핫키 기본 메뉴(Suspend, Pause 등)를 모두 제거합니다.

; 2. 항목 추가 (메뉴 이름, 클릭 시 실행될 라벨 이름)
Menu, Tray, Add, Contact (UNI-Key), OpenContact
Menu, Tray, Add, User Manual, OpenManual
Menu, Tray, Add, Key-Setting, Key-Setting

Menu, Tray, Add, LogOut (%UserEmail%), 로그아웃
Menu, Tray, Add ; 구분선(Separator) 추가
Menu, Tray, Add, Exit, ExitMenu







;--- [2025-10-22 지침 반영: v1 문법] ---

; INI 파일 경로 설정
IniFile := A_ScriptDir . "\UNI-Value.ini"

;--- [2025-10-22 지침 반영: v1 문법] ---

; ==============================================================================
; [핵심 최적화 함수] 섹션 전체를 한 번에 읽어 전역 변수로 할당 (줄바꿈 처리 추가)
; ==============================================================================
LoadSectionToVars(File, Section) {
    global ; 함수의 첫 줄에 배치하여 '전역 변수 모드'로 설정 (오류 해결 핵심)
    
    IniRead, SectionText, %File%, %Section%
    if (SectionText = "" || SectionText = "ERROR")
        return

    Loop, Parse, SectionText, `n, `r
    {
        ; 'Key=Value' 형태를 찾아 분리
        SplitPos := InStr(A_LoopField, "=")
        if (SplitPos) {
            VarName := Trim(SubStr(A_LoopField, 1, SplitPos - 1))
            VarValue := Trim(SubStr(A_LoopField, SplitPos + 1))
            
            if (VarName != "") {
                ; ★ [핵심 추가] 텍스트 "\n"을 오토핫키의 실제 엔터 기호인 "`n"으로 일괄 치환
                ; (사장님의 구버전 호환성을 위해 StringReplace 명령어 사용)
                StringReplace, VarValue, VarValue, \n, `n, All
                
                %VarName% := VarValue ; 동적 변수 생성
            }
        }
    }
}

; ==============================================================================
; [속도 향상] 각 섹션을 통째로 읽기 (사장님 코드 유지)
; ==============================================================================
LoadSectionToVars(IniFile, "기본설정")
LoadSectionToVars(IniFile, "단축키")
LoadSectionToVars(IniFile, "시스템")
LoadSectionToVars(IniFile, "디자인")
LoadSectionToVars(IniFile, "이미지수정")
LoadSectionToVars(IniFile, "본문폰트 기본설정")
LoadSectionToVars(IniFile, "타이틀본트 기본설정")





; ==============================================================================
; 핫키리스트 미리 불러오기
; ==============================================================================

; ==============================================================================
; [설정] 파일 경로 및 기초 데이터
; ==============================================================================
Global IniFile := A_ScriptDir . "\UNI-Value.ini"
Global ModList := "|Ctrl|Alt|Shift|Win"
Global KeyList := "|a|b|c|d|e|f|g|h|i|j|k|l|m|n|o|p|q|r|s|t|u|v|w|x|y|z|[|]|/|'|.|1|Numpad1|2|Numpad2|3|Numpad3|4|Numpad4|5|Numpad5|6|Numpad6|7|Numpad7|8|Numpad8|9|Numpad9|0|Numpad0|Up|Down|Left|Right|Home|End|PgUp|PgDn|Space|Enter|Tab|Esc|Backspace|Delete|Insert|F1|F2|F3|F4|F5|F6|F7|F8|F9|F10|F11|F12"


; Action_Uni0090:


Global HotkeyList := []
HotkeyList.Push("Uni0050") ; F8
HotkeyList.Push("Uni0060") ; alt1
HotkeyList.Push("Uni0181") ; 자간증가
HotkeyList.Push("Uni0182") ; 자간감소
HotkeyList.Push("Uni0187") ; 줄간격 증가
HotkeyList.Push("Uni0186") ; 줄간격 감소
HotkeyList.Push("Uni0191") ; 단락의 뒤간격 크게⭣
HotkeyList.Push("Uni0192") ; 단락의 뒤간격 작게⭡
HotkeyList.Push("Uni0201") ; 단락앞 내어쓰기 간격 크게⭢
HotkeyList.Push("Uni0202") ; 단락앞 내어쓰기 간격 작게⭠
HotkeyList.Push("Uni0203") ; 단락앞 내어쓰기 초기화(여러개 선택후 일괄적용)
HotkeyList.Push("Uni0211") ; 가로 간격 동일 ⮂
HotkeyList.Push("Uni0212") ; 세로 간격 동일 ⮃
HotkeyList.Push("Uni0221") ; 맨 앞으로 가져오기
HotkeyList.Push("Uni0222") ; 맨 뒤로 보내기
HotkeyList.Push("Uni0231") ; 도형 윤곽선 두께 증가 ⭡
HotkeyList.Push("Uni0232") ; 도형 윤곽선 두께 감소 ⭣
HotkeyList.Push("Uni0233") ; 도형 윤곽선 테두리 중심 이동
HotkeyList.Push("Uni0241") ; 도형 윤곽선 대시 종류 변경 ⭢
HotkeyList.Push("Uni0242") ; 도형 윤곽선 대시 종류 변경 ⭠
HotkeyList.Push("Uni0251") ; 선 화살표 꼬리유형 변경 ⭢
HotkeyList.Push("Uni0252") ; 선 화살표 꼬리유형 변경 ⭠
HotkeyList.Push("Uni0253") ; 선 화살표 꼬리크기 변경 ⭡
HotkeyList.Push("Uni0254") ; 선 화살표 꼬리크기 변경 ⭣
HotkeyList.Push("Uni0261") ; 선 화살표 머리유형 변경 ⭢
HotkeyList.Push("Uni0262") ; 선 화살표 머리유형 변경 ⭠
HotkeyList.Push("Uni0263") ; 선 화살표 머리크기 변경 ⭡
HotkeyList.Push("Uni0264") ; 선 화살표 머리크기 변경 ⭣
HotkeyList.Push("Uni0271") ; 도형 병합 / 교차
HotkeyList.Push("Uni0272") ; 교차된 그림을 원래대로
HotkeyList.Push("Uni0281") ; 그림 자르기
HotkeyList.Push("Uni0282") ; 그림 자르기 / 채우기
HotkeyList.Push("Uni0283") ; 첫번째 선택한 이미지(1)를 두번째 선택한 도형(2)에 끼워넣기
HotkeyList.Push("Uni0291") ; 선택한 오브젝트의 위치정보,크기정보 복사
HotkeyList.Push("Uni0292") ; 위 복사한 위치정보를 같도록
HotkeyList.Push("Uni0293") ; 위 복사한 크기정보 같도록
HotkeyList.Push("Uni0301") ; 검정바탕 흰색글씨
HotkeyList.Push("Uni0302") ; 투명바탕 검정글씨
HotkeyList.Push("Uni0311") ; 도형 체우기 없음
HotkeyList.Push("Uni0312") ; 도형 윤곽선 없음
HotkeyList.Push("Uni0313") ; 채우기 색상 투명도 증가 ( ⭢100)
HotkeyList.Push("Uni0314") ; 채우기 색상 투명도 감소  (0⭠ )
HotkeyList.Push("Uni0315") ; 글꼴 색 투명도 증가 ( ⭢100)
HotkeyList.Push("Uni0316") ; 글꼴 색 투명도 감소  (0⭠ )
HotkeyList.Push("Uni0317") ; 윤곽선 색 투명도 증가 ( ⭢100)
HotkeyList.Push("Uni0318") ; 윤곽선 색 투명도 감소  (0⭠ )
HotkeyList.Push("Uni0321") ; 색상 스포이드 / 채우기
HotkeyList.Push("Uni0322") ; 색상 스포이드 / 윤곽선
HotkeyList.Push("Uni0323") ; 색상 스포이드 / 글꼴
HotkeyList.Push("Uni0331") ; 마스터페이지로 이동
HotkeyList.Push("Uni0332") ; 기본 슬라이드로 이동
HotkeyList.Push("Uni0341") ; 반투명하게 숨기기 (잠금기능)
HotkeyList.Push("Uni0342") ; 모든 객체 다시 보이기
HotkeyList.Push("Uni0351") ; 현재 화면의 안내선 복사
HotkeyList.Push("Uni0352") ; 현재 화면의 안내선 붙여넣기
HotkeyList.Push("Uni0353") ; 현재 화면의 안내선 삭제
HotkeyList.Push("Uni0371") ; 영어 대문자로
HotkeyList.Push("Uni0372") ; 영어 각 단어를 대문자로
HotkeyList.Push("Uni0373") ; 영어 소문자로
HotkeyList.Push("Uni0381") ; 텍스트상자 / 도형의 텍스트 배치(W)
HotkeyList.Push("Uni0391") ; 재실행 (실행취소 되돌리기)
HotkeyList.Push("Uni0401") ; 사용자 지정 슬라이드 크기 / A4사이즈 입력창
HotkeyList.Push("Uni0441") ; 현재 마우스 위치의 칼라값을 클립보드로 복사합니다.



; ==============================================================================
; 1. 파워포인트 실행 전 언어 감지 로직 (최초 1회 판별 및 INI 저장)
; ==============================================================================
IniRead, LangID, %IniFile%, 시스템, LangID, %A_Space%

if (LangID="" || LangID=" ") {
    ; 1단계: COM 객체를 통한 정확한 언어 감지
    try {
        try {
            ppt := ComObjActive("PowerPoint.Application")
            isCreated := false
        } catch {
            ppt := ComObjCreate("PowerPoint.Application")
            isCreated := true
        }
        LangID := ppt.LanguageSettings.LanguageID(2) ; UI 언어 ID
        if (isCreated) {
            ppt.Quit()
        }
        ppt := "" 
    } catch {
        ; 2단계: COM 객체 에러 시 레지스트리 기반 구버전 언어 감지
        officeVersions := ["16.0", "15.0", "14.0", "12.0", "11.0"]
        for index, ver in officeVersions {
            RegRead, regLangID, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Office\%ver%\Common\LanguageResources, UILanguage
            if (!ErrorLevel && regLangID != "") {
                LangID := regLangID
                break
            }
        }
        if (LangID="")
            LangID := 1033 ; 못 찾으면 기본 영어
    }
    
    ; 3단계: 감지된 언어팩이 INI 파일에 실제로 존재하는지 검사 (완전 동적 확인)
    IniRead, CheckLangExist, %IniFile%, Lang_%LangID%, LangName, ERROR
    
    ; 만약 'ERROR'가 반환되었다면 (즉, INI 파일에 해당 언어 번역본이 없다면)
    if (CheckLangExist = "ERROR") {
        LangID := 1033 ; 기본 언어인 영어(1033)로 강제 고정
    }
    
    ; 최종 결정된 언어를 INI에 저장
    IniWrite, %LangID%, %IniFile%, 시스템, LangID
}

; ==============================================================================
; 2. ★ 선택된 언어의 다국어 텍스트 섹션을 통째로 로드! ★
; ==============================================================================
Global LangSec := "Lang_" . LangID
LoadSectionToVars(IniFile, LangSec) 
; 이제 INI에 적힌 변수가 전역 변수로 세팅됩니다!

; ==============================================================================
; 3. 동적 언어 리스트(DropDownList) 문자열 만들기 (수정 버전)
; ==============================================================================
LangList := ""
IniRead, AllSections, %IniFile%

Loop, Parse, AllSections, `n, `r
{
    ; "Lang_" 으로 시작하는 섹션만 골라냄
    if (InStr(A_LoopField, "Lang_") = 1)
    {
        ; [핵심 수정] 기존에 로드된 LangName 변수를 덮어쓰지 않도록 tempLangName 사용
        IniRead, tempLangName, %IniFile%, %A_LoopField%, LangName, %A_LoopField%
        
        ; --- 오류 수정 구간 시작 ---
        ; StrReplace() 함수 대신 StringReplace 명령어를 사용 (구버전 호환)
        StringReplace, LangCode, A_LoopField, Lang_
        ; --- 오류 수정 구간 끝 ---
        
        ; 현재 설정된 LangID면 선택 표시(||)를 달고, 아니면 파이프(|)만 달기
        if (LangCode = LangID)
            LangList .= tempLangName . "||"
        else
            LangList .= tempLangName . "|"
    }
}

; 시작과 동시에 INI 파일을 읽고 단축키를 활성화합니다.
GoSub, InitHotkeys 

;-------------------------상하부 구분










WinSet, Transparent, 255, Microsoft Visual Basic for Applications
;중간에 오류나서 창닫았을때를 대비 항상 잘보이게 설정

; ★★★★★★★★★★★★★★★★★★★ 초기 변수값 세팅 --------------------------------


StringReplace, 기본테이블라인색상박스, 기본테이블라인색상, #, c0x
StringReplace, 포인트테이블라인색상1박스, 포인트테이블라인색상1, #, c0x
StringReplace, 포인트테이블라인색상2박스, 포인트테이블라인색상2, #, c0x




포토룸로그인확인버튼가로크기=160
포토룸로그인확인버튼세로크기=50
포토룸다운로드확인버튼가로크기=70
포토룸다운로드확인버튼세로크기=50
포토룸사람확인가로크기=7
포토룸사람확인세로크기=14
; ★포토룸 다운로드버튼은 보라색 "무료로편집"을 검색후 470, 30 을 + 한 좌표로 클릭하도록

RemoveBG로그인확인버튼가로크기=155
RemoveBG로그인확인버튼세로크기=50
RemoveBG다운로드확인버튼가로크기=68
RemoveBG다운로드확인버튼세로크기=38


노트북단독=0
집컴서브모니터=0
모든사용자=0





;현재의창 값 읽어오기(한글단축키에서 사용됨)
WinGet, hWnd, ID, A
WinGetClass, className, ahk_id %hWnd%
hwpClass := className



;모니터 설정값 읽어오는 스크립트, 컴퓨터를 확인하기 위한 스크립트-----------------------------------------------------

    SysGet, MonitorName, MonitorName, 1
    SysGet, Monitor, Monitor, 1
    SysGet, MonitorWorkArea, MonitorWorkArea, 1

;확인하는 코드
;★노트북단독★
if (MonitorWorkAreaRight = "1920")
{
서치성공=0
사용자가로위치 := A_ScreenWidth / 2 - 100
사용자세로위치 := A_ScreenHeight / 2 - 50
사용자화면가로크기 := A_ScreenWidth
사용자화면세로크기 := A_ScreenHeight
이미지서치폴더 = ★단축img\노트북단독
;★노트북단독★★★★★★★★★★★★★==스샷을 찍고(드래그) 파워포인트에 넣은후 "ctrl + home" 으로 칼라값을 알아낼것

포토룸로그인확인칼라:=노트북단독_포토룸로그인확인칼라
포토룸다운로드확인칼라:=노트북단독_포토룸다운로드확인칼라
포토룸사람확인칼라:=노트북단독_포토룸사람확인칼라

RemoveBG로그인확인칼라:=노트북단독_RemoveBG로그인확인칼라
RemoveBG다운로드확인칼라:=노트북단독_RemoveBG다운로드확인칼라

노트북단독=1

;★노트북단독★★★★★★★★★★★★★==스샷을 찍고(드래그) 파워포인트에 넣은후 "ctrl + home" 으로 칼라값을 알아낼것

msgbox, 
(
%단축키프로그램버전%

E-mail : %UserEmail%
Level : %UserGrade%
Subscription date : %UserExpiryDate%

Program Language : %LangName%

★MonitorWorkAreaRight : 1920★
)

goto, 컴퓨터확인뛰어넘기
}


;★집컴서브모니터★
if (MonitorWorkAreaRight = "4480")
{
서치성공=0
사용자가로위치 := A_ScreenWidth / 2 - 100
사용자세로위치 := A_ScreenHeight / 2 - 50
사용자화면가로크기 := A_ScreenWidth + 3840
사용자화면세로크기 := A_ScreenHeight + 3840
이미지서치폴더 = ★단축img\집컴서브모니터


;★집컴서브모니터★★★★★★★★★★★★★==스샷을 찍고(드래그) 파워포인트에 넣은후 "ctrl + home" 으로 칼라값을 알아낼것


포토룸로그인확인칼라:=집컴서브모니터_포토룸로그인확인칼라
포토룸다운로드확인칼라:=집컴서브모니터_포토룸다운로드확인칼라
포토룸사람확인칼라:=집컴서브모니터_포토룸사람확인칼라

RemoveBG로그인확인칼라:=집컴서브모니터_RemoveBG로그인확인칼라
RemoveBG다운로드확인칼라:=집컴서브모니터_RemoveBG다운로드확인칼라

집컴서브모니터=1

;★집컴서브모니터★★★★★★★★★★★★★==스샷을 찍고(드래그) 파워포인트에 넣은후 "ctrl + home" 으로 칼라값을 알아낼것



msgbox, 
(
%단축키프로그램버전%

E-mail : %UserEmail%
Level : %UserGrade%
Subscription date : %UserExpiryDate%

Program Language : %LangName%

★MonitorWorkAreaRight : 4480★
)

goto, 컴퓨터확인뛰어넘기
}





;★모든사용자★
;★★★★★★★★★★★★★모든사용자
서치성공=0
사용자가로위치 := A_ScreenWidth / 2 - 100
사용자세로위치 := A_ScreenHeight / 2 - 50
사용자화면가로크기 := A_ScreenWidth + 3500
사용자화면세로크기 := A_ScreenHeight + 3500
이미지서치폴더 = ★단축img\모든사용자


;★모든사용자★★★★★★★★★★★★★==스샷을 찍고(드래그) 파워포인트에 넣은후 "ctrl + home" 으로 칼라값을 알아낼것

포토룸로그인확인칼라:=모든사용자_포토룸로그인확인칼라
포토룸다운로드확인칼라:=모든사용자_포토룸다운로드확인칼라
포토룸사람확인칼라:=모든사용자_포토룸사람확인칼라

RemoveBG로그인확인칼라:=모든사용자_RemoveBG로그인확인칼라
RemoveBG다운로드확인칼라:=모든사용자_RemoveBG다운로드확인칼라

모든사용자=1

;★모든사용자★★★★★★★★★★★★★==스샷을 찍고(드래그) 파워포인트에 넣은후 "ctrl + home" 으로 칼라값을 알아낼것


msgbox, 
(
%단축키프로그램버전%

E-mail : %UserEmail%
Level : %UserGrade%
Subscription date : %UserExpiryDate%

Program Language : %LangName%

★All users★
)

goto, 컴퓨터확인뛰어넘기

;------------------------------------------------ 컴퓨터 확인 끝



/*
Monitor:		#%A_Index%
Name:		%MonitorName%

MonitorLeft:		%MonitorLeft%
MonitorWorkAreaLeft: 	%MonitorWorkAreaLeft%
MonitorRight:		%MonitorRight%	
MonitorWorkAreaRight: 	%MonitorWorkAreaRight%
MonitorTop:		%MonitorTop%
MonitorWorkAreaTop: 	%MonitorWorkAreaTop%
MonitorBottom:		%MonitorBottom%
MonitorWorkAreaBottom: 	%MonitorWorkAreaBottom%

포토룸-color =		%포토룸로그인확인칼라%
RemoveBG-color =	%RemoveBG로그인확인칼라%

*/



컴퓨터확인뛰어넘기:
;★★★★★★★★★★★★★★★★★


gosub, 키보드올리기


Clipboard := ""





return
;★★★★★⬆오핫 실행하면 이쪽 리턴에서 마무리됨










;$ 이것은 핫키의 실행이 또 핫키가 안되도록  예를 들어 $a::    send, abc   이런것들

Action_Uni0050:
; $F8::

$^!+d::

;일러스트에서는 도구창 나오게 설정
if WinActive("ahk_exe Illustrator.exe")
{
sendinput, ^!+{d}
sleep, 10
return

}else{



;관리자모드 시작
if (UserGrade >=4)
{
goto, 회전대칭고해상도
}else{


if WinActive("ahk_exe POWERPNT.EXE")
{
goto, 회전대칭고해상도
}else{
sendinput, ^!+{d}
sleep, 10
return
}





}





}




; 회전/대칭 ====================================================================
; 정보창을 띄운다음 회전값이나 좌우상하대칭 체크를 확인후 진행


회전대칭고해상도:

객체확인 := ""






번역하기 = 0
대체텍스트변경 = 0
영상트리밍 = 0
캡쳐본생성 = 0
쪽번호전체삭제 = 0
쪽번호증가일괄변경 = 0
자동코드실행 = 0

고해상도이미지 = 0
투명배경포토룸 = 0
RemoveBG = 0












; 시스템 메시지(0x0111 = WM_COMMAND)를 가로채어 함수로 연결 (스크립트 최상단 부근에 배치)
OnMessage(0x0111, "WM_COMMAND")

Gui, 대칭: Destroy



;Gui, 대칭:New, +AlwaysOnTop +ToolWindow, UNI
Gui, 대칭:New, +AlwaysOnTop, UNI
Gui, 대칭:Font, s9, 맑은 고딕
Gui, 대칭:Add,Text, w40 h10 +0x200, 

대칭창높이 := 25



Try
	{

;파워포인트 실행중일때만 보이기
        if WinActive("ahk_exe POWERPNT.EXE")
        {




Try
	{
ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type
선택갯수:=ppt.ActiveWindow.Selection.Shaperange.Count()
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
Connector:=ppt.ActiveWindow.Selection.Shaperange.Connector
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
HasTable:=ppt.ActiveWindow.Selection.Shaperange.HasTable
HasTextFrame:=ppt.ActiveWindow.Selection.Shaperange.HasTextFrame
Id:=ppt.ActiveWindow.Selection.Shaperange.Id
Name:=ppt.ActiveWindow.Selection.Shaperange.Name
Parent:=ppt.ActiveWindow.Selection.Shaperange.Parent
Visible:=ppt.ActiveWindow.Selection.Shaperange.Visible
	}


Try
	{
SetFormat, float, 0.2
기존회전값:=ppt.ActiveWindow.Selection.Shaperange.Rotation
}


Try
	{
SetFormat, float, 0.2
기존기준점1:=ppt.ActiveWindow.Selection.Shaperange.Adjustments(1)
}

Try
	{
SetFormat, float, 0.2
기존기준점2:=ppt.ActiveWindow.Selection.Shaperange.Adjustments(2)
}

Try
	{
SetFormat, float, 0.2
기존기준점3:=ppt.ActiveWindow.Selection.Shaperange.Adjustments(3)
}

Try
	{
SetFormat, float, 0.2
기존기준점4:=ppt.ActiveWindow.Selection.Shaperange.Adjustments(4)
}



대칭창높이 := 대칭창높이 + 50
Gui, 대칭:Add,Text, Section w119 h20 +0x200, %F8EditRotae% :
Gui, 대칭:Add, Edit, x+5 w55 h20 v회전값, %기존회전값%
Gui, 대칭:Add,Text, xs w40 h5 +0x200, 



if (기존기준점1!="")
{
대칭창높이 := 대칭창높이 + 27
Gui, 대칭:Add,Text, Section w119 h20 +0x200, %F8EditAnchor1%
Gui, 대칭:Add, Edit, x+5 w55 h20 v변경기준점1, %기존기준점1%
}

if (기존기준점2!="")
{
대칭창높이 := 대칭창높이 + 27
Gui, 대칭:Add,Text, xs Section w119 h20 +0x200, %F8EditAnchor2%
Gui, 대칭:Add, Edit, x+5 w55 h20 v변경기준점2, %기존기준점2%
}

if (기존기준점3!="")
{
대칭창높이 := 대칭창높이 + 27
Gui, 대칭:Add,Text, xs Section w119 h20 +0x200, %F8EditAnchor3%
Gui, 대칭:Add, Edit, x+5 w55 h20 v변경기준점3, %기존기준점3%
}

if (기존기준점4!="")
{
대칭창높이 := 대칭창높이 + 27
Gui, 대칭:Add,Text, xs Section w119 h20 +0x200, %F8EditAnchor4%
Gui, 대칭:Add, Edit, x+5 w55 h20 v변경기준점4, %기존기준점4%
}




대칭창높이 := 대칭창높이 + 70
Gui, 대칭:Add,Text, xs w40 h1 +0x200, 
Gui, 대칭:Add, Checkbox, w150 h20 v가로대칭 g체크박스확인1,  ↔ %F8ChkHorizontal%
;Gui, 대칭:Add,Text, w40 h10 +0x200, 
Gui, 대칭:Add, Checkbox, w150 h20 v상하대칭 g체크박스확인1,  ↕ %F8ChkVertical%
;Gui, 대칭:Add,Text, w40 h10 +0x200, 




;관리자모드 시작
if (UserGrade >=4)
{



;이미지또는 그래픽이면 보이기 시작
if (객체확인=13||객체확인=28)
{
대칭창높이 := 대칭창높이 + 109
Gui, 대칭:Add,Text, w180 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Gui, 대칭:Add,Text, w40 h1 +0x200, 
Gui, 대칭:Add, Checkbox, w150 h20 v고해상도이미지 g체크박스확인1,  %F8ChkResolution%
;Gui, 대칭:Add,Text, w40 h10 +0x200, 
Gui, 대칭:Add, Checkbox, w150 h20 v투명배경포토룸 g체크박스확인1,  %F8ChkPhotoRoom%
;Gui, 대칭:Add,Text, w40 h10 +0x200, 
Gui, 대칭:Add, Checkbox, w150 h20 vRemoveBG g체크박스확인1,  %F8ChkRemovebg%
;Gui, 대칭:Add,Text, w40 h10 +0x200, 

}
;이미지또는 그래픽이면 보이기 끝


}
;관리자모드 끝




대칭창높이 := 대칭창높이 + 56
Gui, 대칭:Add,Text, xs w180 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Gui, 대칭:Font, s9, 맑은 고딕

Gui, 대칭:Add, Button, w180 h35 v기본실행버튼 g대칭진행실행, %F8BtnApply%



Gui, 대칭:Font, s9, 맑은 고딕





대칭창높이 := 대칭창높이 + 106
Gui, 대칭:Add,Text, xs w180 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; 1. 헤더 텍스트
Gui, 대칭:Add, Text, Section w88 h20 +0x200, %F8BtnExportPNG%
Gui, 대칭:Add, Edit, x+5 w38 v이미지변환커스텀, 

Gui, 대칭:Font, s7.5, 맑은 고딕
Gui, 대칭:Add, Button, x+1 yp-1 w46 h25 v커스텀실행버튼 g이미지변환하기0, Custom


Gui, 대칭:Add, Button, xs Section w40 h20 g이미지변환하기1, 480
Gui, 대칭:Add, Button, x+6 w40 h20 g이미지변환하기2, 720
Gui, 대칭:Add, Button, x+6 w40 h20 g이미지변환하기3, 1280
Gui, 대칭:Add, Button, x+6 w40 h20 g이미지변환하기4, 1920

Gui, 대칭:Add, Button, xs Section w40 h20 g이미지변환하기5, 2560
Gui, 대칭:Add, Button, x+6 w40 h20 g이미지변환하기6, 3840
Gui, 대칭:Add, Button, x+6 w40 h20 g이미지변환하기7, 4320
Gui, 대칭:Add, Button, x+6 w40 h20 g이미지변환하기8, 7680

Gui, 대칭:Font, s9, 맑은 고딕







        }
;파워포인트 실행중일때만 보이기








;관리자모드 시작
if (UserGrade >=4)
{

;이미지또는 그래픽이면 안보이게
if (객체확인=13||객체확인=28||객체확인=16)
{
;안보이게하기
}else{

대칭창높이 := 대칭창높이 + 122

Gui, 대칭:Add,Text, xs w180 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; 1. 헤더 텍스트
Gui, 대칭:Add, Text, w180 h20 +0x200, %F8ListGoogle%

; 2. 언어 선택 드롭다운 리스트 (파이프 | 로 구분, Choose1은 첫번째 항목 기본선택)
Gui, 대칭:Add, DropDownList, w88 vTargetLang Choose1, Language|English|日本|việt nam|россия|中国

; 3. 메인 번역 버튼 (간격 조정을 위해 위쪽 여백 y+5 추가 가능)
Gui, 대칭:Add, Button, x+4 w88 h23 v글로벌번역버튼 g글로벌번역, %F8BtnGoogleApply%

; 4. 하단 분할 버튼 (한글 / 영어)
; w85 두 개와 사이 간격 10을 합치면 180 (85+10+85=180)
Gui, 대칭:Add,Text, xs w40 h1 +0x200, 

Gui, 대칭:Add, Button, w88 h30 g한글번역, 한국어
Gui, 대칭:Add, Button, x+4 w88 h30 g영어번역, English ; x+10은 바로 옆에 붙이라는 명령어




}


}
;관리자모드 끝




;파워포인트 실행중일때만 보이기
        if WinActive("ahk_exe POWERPNT.EXE")
        {

대칭창높이 := 대칭창높이 + 47

Gui, 대칭:Add,Text, xs w180 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


Gui, 대칭:Add, Button, w85 h30 g특수문자입력창, %F8BtnSymbol%


; --- 고급메뉴 버튼 커스터마이징 ---
Gui, 대칭:Add, Progress, x+10 w85 h30 Disabled Background000000

Gui, 대칭:Font, cWhite bold
; [핵심 추가] hwndh고급메뉴 옵션을 통해 이 컨트롤의 고유 ID를 'h고급메뉴' 변수에 저장
;Gui, 대칭:Add, Text, hwndh고급메뉴 xp yp wp hp Center 0x200 BackgroundTrans g고급메뉴창, %F8BtnAdvanced%
Gui, 대칭:Add, Text, xp yp wp hp Center 0x200 BackgroundTrans g고급메뉴창, %F8BtnAdvanced%
Gui, 대칭:Font, cDefault norm
; -------------------------------------

}




Gui, 대칭:Show, w200 h%대칭창높이%, UNI


return









; ======================================================================
; 이벤트 감지 함수 (기존 버튼상태관리자 루프를 완벽히 대체)
; ======================================================================
WM_COMMAND(wParam, lParam) {
    NotifyCode := wParam >> 16
    
    ; 0x0100 = Edit 컨트롤 포커스, 3 = DropDownList 포커스, 0 = Button/Checkbox 클릭 등 상태 변화 감지
    if (NotifyCode = 0x0100 || NotifyCode = 3 || NotifyCode = 0) {
        
        GuiControlGet, 현재포커스, 대칭:FocusV
        
        if (현재포커스 = "회전값" || 현재포커스 = "변경기준점1" || 현재포커스 = "변경기준점2" || 현재포커스 = "변경기준점3" || 현재포커스 = "변경기준점4" || 현재포커스 = "가로대칭" || 현재포커스 = "상하대칭" || 현재포커스 = "고해상도이미지" || 현재포커스 = "투명배경포토룸" || 현재포커스 = "RemoveBG")
        {
            GuiControl, 대칭:+Default, 기본실행버튼
        }
        else if (현재포커스 = "이미지변환커스텀")
        {
            GuiControl, 대칭:+Default, 커스텀실행버튼
        }
        else if (현재포커스 = "TargetLang")
        {
            GuiControl, 대칭:+Default, 글로벌번역버튼
        }
    }
}



















sleep, 1
send, ^a
;회전정보창을 미리 선택해놓는것



}

return
;회전정보창 gui 끝













이미지변환하기0:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 이미지변환커스텀
goto, 이미지변환시작하기


이미지변환하기1:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 480
goto, 이미지변환시작하기

이미지변환하기2:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 720
goto, 이미지변환시작하기

이미지변환하기3:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 1280
goto, 이미지변환시작하기

이미지변환하기4:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 1920
goto, 이미지변환시작하기

이미지변환하기5:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 2560
goto, 이미지변환시작하기

이미지변환하기6:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 3840
goto, 이미지변환시작하기

이미지변환하기7:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 4320
goto, 이미지변환시작하기

이미지변환하기8:
Gui, 대칭:Submit,NoHide
이미지변환해상도 := 7680
goto, 이미지변환시작하기




이미지변환시작하기:
Gui, 대칭: Destroy








DesktopPath := A_Desktop
BaseFolderName := "이미지PNG변환"
FullPath := DesktopPath "\" BaseFolderName

If FileExist(FullPath)
{
    index := 1
    Loop
    {
        newFullPath := FullPath . index
        if !FileExist(newFullPath)
        {
            FileCreateDir, % newFullPath
            FullPath := newFullPath
            break
        }
        index++
    }
}
else
{
    FileCreateDir, % FullPath
}

; --------------------------------------------------------------------------
; 2. PowerPoint COM 연결 및 처리
; --------------------------------------------------------------------------
try {
    ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type
} catch {
    MsgBox, 262208, Message, %msgboxuni_0015%
    return
}

try {
    window := ppt.ActiveWindow
    sel := window.Selection
    
    ; ppSelectionShapes = 2, ppSelectionText = 3
    if (sel.Type != 2) && (sel.Type != 3) { 
        MsgBox, 262208, Message, %msgboxuni_0016%
        return
    }
    
    shpRange := sel.ShapeRange
    sld := window.View.Slide
  


    ; 2. 복사 후 임시 붙여넣기 (메타파일)
    shpRange.Copy()
    Sleep, 300 ; 클립보드 복사 대기 시간 늘림
    pastedRange := sld.Shapes.PasteSpecial(2) ; 2 = ppPasteEnhancedMetafile
    targetShape := pastedRange.Item(1)

    window := ppt.ActiveWindow
    sel := window.Selection
    shpRange := sel.ShapeRange
    sld := window.View.Slide
 



붙여넣은백터왼쪽 := shpRange.Left
붙여넣은백터위쪽 := shpRange.Top
붙여넣은백터가로 := shpRange.Width
붙여넣은백터세로 := shpRange.Height


    출력가로크기 := 이미지변환해상도
    이미지가로세로비율 := 붙여넣은백터가로 / 붙여넣은백터세로
SetFormat, float, 0.0
    출력세로크기 := 출력가로크기 / 이미지가로세로비율


    FormatTime, TimeString,, yyyyMMdd_HHmmss
    FileName := FullPath . "\Selection_" . TimeString . ".png"
    
    ; 4. 내보내기 (Export)

TargetWidthPixel  := 출력가로크기
TargetHeightPixel := 출력세로크기

NewWidthPoint  := TargetWidthPixel * 72 / 96
NewHeightPoint := TargetHeightPixel * 72 / 96

; 1. 비율 잠금 해제
; msoFalse = 0
targetShape.LockAspectRatio := 0 

; 2. 도형 크기를 실제로 변경
targetShape.Width := NewWidthPoint
sleep, 1
targetShape.Height := NewHeightPoint
sleep, 1

; 3. 내보내기 (이제 숫자는 넣지 않습니다)
targetShape.Export(FileName, 2)
    ; 임시로 만들었던 메타파일 도형 삭제 (주석 해제 권장)
targetShape.Delete() 






    ; ======================================================================
    ; [수정 핵심] 파일 생성 대기 및 확인 루틴 추가
    ; ======================================================================
    Loop
    {
        If FileExist(FileName)
            Break ; 파일이 생겼으면 루프 탈출
        Sleep, 100
    }

SetFormat, float, 0.0
newPic := sld.Shapes.AddPicture(FileName, 0, -1, 붙여넣은백터왼쪽, 붙여넣은백터위쪽, 붙여넣은백터가로, 붙여넣은백터세로)





} catch e {
    MsgBox, 262208, Message, %msgboxuni_0017%`n%e%
}


SetTimer, 정보창툴팁없애기, 600
Gui, 대칭: Destroy


작업끝표시(50, "red")


;이미지PNG변환하기 끝

return










글로벌번역:
Gui, 대칭:Submit,NoHide
Gui, 대칭: Destroy



if (TargetLang = "Language")
{
번역언어 = ko
}
if (TargetLang = "English")
{
번역언어 =  en
}
if (TargetLang = "日本")
{
번역언어 = ja
}
if (TargetLang = "việt nam")
{
번역언어 = vi
}
if (TargetLang = "россия")
{
번역언어 = ru
}
if (TargetLang = "中国")
{
번역언어 = zh-CN
}



번역하기=1
자동코드실행=1
goto, 번역하기
return





한글번역:
Gui, 대칭:Submit,NoHide
Gui, 대칭: Destroy


번역하기=1
자동코드실행=1
번역언어=ko
goto, 번역하기
return




영어번역:
Gui, 대칭:Submit,NoHide
Gui, 대칭: Destroy


번역하기=1
자동코드실행=1
번역언어=en
goto, 번역하기
return






; 회전정보창 gui 중복선택 방지
체크박스확인1:
    ; 클릭된 체크박스(A_GuiControl)만 제외하고 모두 해제
    GuiControl,, 가로대칭, 0
    GuiControl,, 상하대칭, 0
    GuiControl,, 고해상도이미지, 0
    GuiControl,, 투명배경포토룸, 0
    GuiControl,, RemoveBG, 0

    ; ※ 자동코드실행은 여기서 제외 (원하면 여기에도 0 추가 가능)

    ; 현재 클릭된 체크박스만 다시 체크
    GuiControl,, %A_GuiControl%, 1
return









대칭진행실행:

Gui, 대칭:Submit,NoHide
Gui, 대칭: Destroy



;실행하기전에 고해상도이미지인지 투명배경인지 체크하고 넘기기

if (고해상도이미지=1)
{
Gui, 대칭: Destroy


sleep, 100
goto, 고해상도이미지실행
}

if (투명배경포토룸=1)
{
Gui, 대칭: Destroy


sleep, 100
goto, 투명배경포토룸실행
}

if (RemoveBG=1)
{
Gui, 대칭: Destroy


sleep, 100
goto, RemoveBG
}



;------------------------------------------------------------------------------------


if (회전값 != "")
{
Try
	{
ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
ppt.ActiveWindow.Selection.Shaperange.Rotation:=회전값
회전대칭정보 := 회전값
        }

Try
	{
ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
ppt.ActiveWindow.Selection.Shaperange.Adjustments(1):=변경기준점1
회전대칭정보 := 변경기준점1
        }

Try
	{
ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
ppt.ActiveWindow.Selection.Shaperange.Adjustments(2):=변경기준점2
회전대칭정보 := 변경기준점2
        }
Try
	{
ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
ppt.ActiveWindow.Selection.Shaperange.Adjustments(3):=변경기준점3
회전대칭정보 := 변경기준점3
        }

Try
	{
ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
ppt.ActiveWindow.Selection.Shaperange.Adjustments(4):=변경기준점4
회전대칭정보 := 변경기준점4
        }
}




if (가로대칭 = 1)
{
Gui, 대칭: hide
sleep, 10

;비디오이면
if (객체확인=16)
{
send, {Alt down}{j}{p}{Alt up}
send, {a}{y}
send, {h}
}

;이미지이면
if (객체확인=13)
{
send, {Alt down}{j}{p}{Alt up}
send, {a}{y}
send, {h}
}

;그래픽이면
if (객체확인=28)
{
send, {Alt down}{j}{g}{Alt up}
send, {a}{y}
send, {h}
}

; 비디오도 이미지도 그래픽도 아니면
if (객체확인 != 16 && 객체확인 != 13 && 객체확인 != 28)
{
send, {Alt down}{j}{d}{Alt up}
send, {a}{y}
send, {h}
}

sleep, 10
회전대칭정보 := "좌우대칭 ↔"
}




if (상하대칭 = 1)
{
Gui, 대칭: hide
sleep, 10

;비디오이면
if (객체확인=16)
{
send, {Alt down}{j}{p}{Alt up}
send, {a}{y}
send, {v}
}

;이미지이면
if (객체확인=13)
{
send, {Alt down}{j}{p}{Alt up}
send, {a}{y}
send, {v}
}

;그래픽이면
if (객체확인=28)
{
send, {Alt down}{j}{g}{Alt up}
send, {a}{y}
send, {v}
}



; 비디오도 이미지도 그래픽도 아니면
if (객체확인 != 16 && 객체확인 != 13 && 객체확인 != 28)
{
send, {Alt down}{j}{d}{Alt up}
send, {a}{y}
send, {v}
}


sleep, 10
회전대칭정보 := "상하대칭↕"
}


기존회전값 := ""
기존기준점1 := ""
기존기준점2 := ""
기존기준점3 := ""
기존기준점4 := ""

가로대칭:=0
상하대칭:=0

SetTimer, 정보창툴팁없애기, 600
Gui, 대칭: Destroy


작업끝표시(50, "red")

return
;회전대칭창 gui마무리







; 2/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위









투명배경포토룸실행:
RemoveBG:







; ============================================
; Chrome 다운로드 저장위치 확인 옵션 상태 체크
; AutoHotkey v1
; ============================================

ChromePath := "C:\Program Files\Google\Chrome\Application\chrome.exe"
IfNotExist, %ChromePath%
{
MsgBox, 262208, Message, %msgboxuni_0018%
return
}


ChromePref := "C:\Users\" A_UserName "\AppData\Local\Google\Chrome\User Data\Default\Preferences"

if !FileExist(ChromePref)
{
    MsgBox, 262208, Message, %msgboxuni_0019%
    return
}

FileRead, prefText, %ChromePref%

if InStr(prefText, """prompt_for_download"":true")
{
    ;정상 작동 맨아래로 지나가기
}
else if InStr(prefText, """prompt_for_download"":false")
{

Run, chrome.exe
Sleep, 200

loop
{
    if WinExist("ahk_exe chrome.exe")
{
break
}
WinActivate, ahk_exe chrome.exe
sleep, 100
}

; 크롬활성화 확인
Send, ^l
sleep, 50
clipboard := "chrome://settings/downloads"
Send, ^v
sleep, 50
send, {enter}
sleep, 200

    MsgBox, 262208, Message, %msgboxuni_0020%
return

}
else
{
    MsgBox, 262208, Message, %msgboxuni_0021%
return
}














; 그림=13, 그룹=6, 표=19, 도형=1, 텍스트상자=17, 그룹과 단체=-2, 벡터=28
ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

;if (객체확인=28 || 객체확인=5 || 객체확인=1 || 객체확인=19 || 객체확인=17)
if (객체확인=5 || 객체확인=1 || 객체확인=19 || 객체확인=17)
{
;넘기기

작업끝표시(50, "red")
sleep, 1
ToolTIP, ★ Convert to bitmap image , %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000

return
}
else
{
gosub, 이미지원래대로
}



;일단 객체하나 복사하기
send, ^d
sleep, 100



;바탕화면에 폴더만들기 " 투명배경 "

DesktopPath := A_Desktop
BaseFolderName := "투명배경"
FullPath := DesktopPath "\" BaseFolderName

If FileExist(FullPath)
{
    index := 1
    Loop
    {
        newFullPath := FullPath . index
        if !FileExist(newFullPath)
        {
            FileCreateDir, % newFullPath
            FullPath := newFullPath
            break
        }
        index++
    }
}
else
{
    FileCreateDir, % FullPath
}





;위 폴더에 다른이름으로 저장하기

ppt := ComObjActive("PowerPoint.Application")

window := ppt.ActiveWindow
sel := window.Selection    
shpRange := sel.ShapeRange
sld := window.View.Slide    

이미지수정왼쪽 := shpRange.Left
이미지수정위쪽 := shpRange.Top
이미지수정가로 := shpRange.Width
이미지수정세로 := shpRange.Height

FormatTime, TimeString,, yyyyMMdd_HHmmss
FileName := FullPath . "\Selection_" . TimeString . ".png"
shpRange.Export(FileName, 2)






; ▼▼▼ 추가된 부분 ▼▼▼
; 창 제목의 일부만 일치해도 되도록 설정합니다. (1:시작부분 일치, 2:포함, 3:정확히 일치)
SetTitleMatchMode, 2 

if (투명배경포토룸=1)
{
GroupAdd, 크롬타이틀투명, 무료 배경 지우기 사이트 - 온라인에서 무료로 이미지 배경 제거 | Photoroom - Whale
GroupAdd, 크롬타이틀투명, Photoroom - Chrome
}

if (RemoveBG=1)
{
GroupAdd, 크롬타이틀투명, 이미지 배경 제거, 투명 배경 만들기 ? remove.bg - Whale
GroupAdd, 크롬타이틀투명, remove.bg - Chrome
}


WinClose, ahk_group 크롬타이틀투명
sleep, 100


; 투명배경 - 포토룸 사이트 접속
if (투명배경포토룸=1)
{
Run, chrome.exe --new-window "https://www.photoroom.com/ko/tools/background-remover"
; Run, "https://www.photoroom.com/ko/tools/background-remover"
}

if (RemoveBG=1)
{
Run, chrome.exe --new-window "https://www.remove.bg/ko"
; Run, "https://www.remove.bg/ko"
}

MouseMove, 1, 1
sleep, 10

sleep, 1000






loop
{
    if WinExist("ahk_group 크롬타이틀투명")
{
break
}
}
sleep, 1000



WinActivate, ahk_group 크롬타이틀투명
sleep, 10





/*
;WinMove 명령어의 파라미터는 순서대로 WinTitle, WinText, X, Y, Width, Height 입니다.
;WinMove, %크롬타이틀투명%,, , , 800, 600

    WinGet, windowState, MinMax, A  ; 활성 창의 상태를 가져옵니다 (A = Active Window)
    
    if (windowState = 1)  ; windowState가 1이면 '최대화'된 상태입니다.
    {
        WinRestore, A  ; 활성 창을 이전 크기(창 모드)로 복원합니다.
sleep, 200
    }
*/



    WinGet, windowState, MinMax, ahk_group 크롬타이틀투명  ; 활성 창의 상태를 가져옵니다 (A = Active Window)
    
    if (windowState = 1)  ; windowState가 1이면 '최대화'된 상태입니다.
    {
        WinRestore, ahk_group 크롬타이틀투명  ; 활성 창을 이전 크기(창 모드)로 복원합니다.
sleep, 200
    }




WinMove, ahk_group 크롬타이틀투명,, 1, 1
sleep, 200

WinActivate, ahk_group 크롬타이틀투명
sleep, 10

WinMaximize, ahk_group 크롬타이틀투명
sleep, 200

send, ^0
;크롬화면을 100%로 조절하는것!
sleep, 100

; ▼▼▼ 최대화된 창의 좌표 구하기 (v1 버전) ▼▼▼
WinGetPos, winX, winY, winWidth, winHeight, ahk_group 크롬타이틀투명

; 끝 좌표 계산
endX := winX + winWidth
endY := winY + winHeight

;여유픽셀주기


if (투명배경포토룸=1)
{
;처음검색의 픽셀구역 설정
픽셀시작x := winX + 100
픽셀시작y := winY + 150
픽셀끝x := endX - 100
픽셀끝y := endY - 150

로그인확인버튼가로크기:=포토룸로그인확인버튼가로크기
로그인확인버튼세로크기:=포토룸로그인확인버튼세로크기
다운로드확인버튼가로크기:=포토룸다운로드확인버튼가로크기
다운로드확인버튼세로크기:=포토룸다운로드확인버튼세로크기
로그인확인칼라:=포토룸로그인확인칼라
다운로드확인칼라:=포토룸다운로드확인칼라


로그인버튼다시확인포토룸:

loop
{

pixelsearch px, py, 픽셀시작x, 픽셀시작y, 픽셀끝x, 픽셀끝y, %로그인확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{
픽셀시작2x := px + 로그인확인버튼가로크기
픽셀시작2y := py + 로그인확인버튼세로크기
픽셀끝2x := px + 로그인확인버튼가로크기 + 2
픽셀끝2y := py + 로그인확인버튼세로크기 + 2

pixelsearch p2x, p2y, 픽셀시작2x, 픽셀시작2y, 픽셀끝2x, 픽셀끝2y, %로그인확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{
;두번째검색확인후 버튼찾음 ★실행
px := px + 10
py := py + 10
Mouseclick, L, %px%, %py%, 1
sleep, 300
MouseMove, 1, 1
sleep, 10
break ;★실행후 루프 빠져나가기
		}
		if ( errorlevel = 1 )
		{
;두번째검색에서 없음
픽셀시작x := 픽셀시작x + 1
픽셀시작y := 픽셀시작y + 1
goto, 로그인버튼다시확인포토룸
		}


		}
		if ( errorlevel = 1 )
		{
;첫번째검색에서 없음
goto, 로그인버튼다시확인포토룸
		}

}




}







if (RemoveBG=1)
{
;처음검색의 픽셀구역 설정
픽셀시작x := winX + 100
픽셀시작y := winY + 150
픽셀끝x := endX - 100
픽셀끝y := endY - 150

로그인확인버튼가로크기:=RemoveBG로그인확인버튼가로크기
로그인확인버튼세로크기:=RemoveBG로그인확인버튼세로크기
다운로드확인버튼가로크기:=RemoveBG다운로드확인버튼가로크기
다운로드확인버튼세로크기:=RemoveBG다운로드확인버튼세로크기
로그인확인칼라:=RemoveBG로그인확인칼라
다운로드확인칼라:=RemoveBG다운로드확인칼라




로그인버튼다시확인RemoveBG:

loop
{

pixelsearch px, py, 픽셀시작x, 픽셀시작y, 픽셀끝x, 픽셀끝y, %로그인확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{
픽셀시작2x := px + 로그인확인버튼가로크기
픽셀시작2y := py + 로그인확인버튼세로크기
픽셀끝2x := px + 로그인확인버튼가로크기 + 2
픽셀끝2y := py + 로그인확인버튼세로크기 + 2

pixelsearch p2x, p2y, 픽셀시작2x, 픽셀시작2y, 픽셀끝2x, 픽셀끝2y, %로그인확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{
;두번째검색확인후 버튼찾음 ★실행
px := px + 10
py := py + 10
Mouseclick, L, %px%, %py%, 1
sleep, 300
MouseMove, 1, 1
sleep, 10
break ;★실행후 루프 빠져나가기
		}
		if ( errorlevel = 1 )
		{
;두번째검색에서 없음
픽셀시작x := 픽셀시작x + 1
픽셀시작y := 픽셀시작y + 1
goto, 로그인버튼다시확인RemoveBG
		}


		}
		if ( errorlevel = 1 )
		{
;첫번째검색에서 없음
goto, 로그인버튼다시확인RemoveBG
		}

}


}










loop
{
if WinActive("열기")
{
sleep, 300
break
}
sleep, 100
}




Clipboard = %FileName%
ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 100

send, ^v
sleep, 100
send, {enter}
sleep, 1500













if (투명배경포토룸=1)
{

;★다운로드창은 약간 아래쪽 처음검색의 픽셀구역 설정

픽셀시작x := winX + 100
픽셀시작y := winY + 610
픽셀끝x := endX - 100
픽셀끝y := endY - 50
;리무브하고 같아도 작동하긴함

다운로드버튼다시확인포토룸:
loop
{

pixelsearch px, py, 픽셀시작x, 픽셀시작y, 픽셀끝x, 픽셀끝y, %다운로드확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{

픽셀시작2x := px + 다운로드확인버튼가로크기
픽셀시작2y := py + 다운로드확인버튼세로크기
픽셀끝2x := px + 다운로드확인버튼가로크기 + 2
픽셀끝2y := py + 다운로드확인버튼세로크기 + 2

pixelsearch p2x, p2y, 픽셀시작2x, 픽셀시작2y, 픽셀끝2x, 픽셀끝2y, %다운로드확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{
;두번째검색확인후 버튼찾음 ★실행
; ★포토룸 다운로드버튼은 보라색 "무료로편집"을 검색후 470, 30 을 + 한 좌표로 클릭하도록
sleep, 700
px := px + 470
py := py + 30
Mouseclick, L, %px%, %py%, 1
sleep, 500

/*
send, {down 1}
sleep, 300
send, {enter}
sleep, 500
*/


break ;★실행후 루프 빠져나가기
		}
		if ( errorlevel = 1 )
		{
;두번째검색에서 없음
픽셀시작x := 픽셀시작x + 1
픽셀시작y := 픽셀시작y + 1

send, {down 1}
sleep, 300
goto, 다운로드버튼다시확인포토룸
		}


		}
		if ( errorlevel = 1 )
		{
;첫번째검색에서 없음



;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★---------------------------사람확인버튼 체크 시작
pixelsearch px, py, 픽셀시작x, 픽셀시작y, 픽셀끝x, 픽셀끝y, %포토룸사람확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{

픽셀시작2x := px + 포토룸사람확인가로크기
픽셀시작2y := py + 포토룸사람확인세로크기
픽셀끝2x := px + 포토룸사람확인가로크기 + 2
픽셀끝2y := py + 포토룸사람확인세로크기 + 2

pixelsearch p2x, p2y, 픽셀시작2x, 픽셀시작2y, 픽셀끝2x, 픽셀끝2y, %포토룸사람확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{
;두번째검색확인후 버튼찾음 ★실행
; ★포토룸 다운로드버튼은 보라색 "무료로편집"을 검색후 470, 30 을 + 한 좌표로 클릭하도록
sleep, 1000
px := px - 235
py := py + 20
Mouseclick, L, %px%, %py%, 1
sleep, 500
		}
		if ( errorlevel = 1 )
		{
sleep, 100
		}


		}
		if ( errorlevel = 1 )
		{
sleep, 100
		}
;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★-------------------------사람확인버튼체크 끝





send, {down 1}
sleep, 300
goto, 다운로드버튼다시확인포토룸
		}

}






}












if (RemoveBG=1)
{
;처음검색의 픽셀구역 설정
픽셀시작x := winX + 100
픽셀시작y := winY + 150
픽셀끝x := endX - 100
픽셀끝y := endY - 150



다운로드버튼다시확인RemoveBG:
loop
{
pixelsearch px, py, 픽셀시작x, 픽셀시작y, 픽셀끝x, 픽셀끝y, %다운로드확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{

픽셀시작2x := px + 다운로드확인버튼가로크기
픽셀시작2y := py + 다운로드확인버튼세로크기
픽셀끝2x := px + 다운로드확인버튼가로크기 + 2
픽셀끝2y := py + 다운로드확인버튼세로크기 + 2

pixelsearch p2x, p2y, 픽셀시작2x, 픽셀시작2y, 픽셀끝2x, 픽셀끝2y, %다운로드확인칼라%, 10, Fast
		if ( errorlevel = 0 )
		{
sleep, 700
;두번째검색확인후 버튼찾음 ★실행
px := px + 10
py := py + 10
Mouseclick, L, %px%, %py%, 1
sleep, 500

px := px - 10
py := py - 10
px := px + 0
py := py + 70

Mouseclick, L, %px%, %py%, 1
sleep, 300

break ;★실행후 루프 빠져나가기
		}
		if ( errorlevel = 1 )
		{
;두번째검색에서 없음
픽셀시작x := 픽셀시작x + 1
픽셀시작y := 픽셀시작y + 1
goto, 다운로드버튼다시확인RemoveBG
		}


		}
		if ( errorlevel = 1 )
		{
;첫번째검색에서 없음
goto, 다운로드버튼다시확인RemoveBG
		}

}




}




















loop
{
if (WinActive("다른 이름으로 저장") || WinActive("Save As Picture"))
{
sleep, 500
break
}
sleep, 400
}





FileNamedown := FullPath . "\Selection_" . TimeString . "-투명배경.png"
Clipboard = %FileNamedown%

ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 100

send, ^v
sleep, 100



send, {enter}
sleep, 1000





WinClose, ahk_group 크롬타이틀투명
sleep, 100












; PPTFrameClass 창 활성 확인
Loop
{
    ; PPT 창이 존재하면
    if WinExist("ahk_exe POWERPNT.EXE")
    {
        ; 이미 활성화되어 있으면 루프 종료
        if WinActive("ahk_exe POWERPNT.EXE")
        {
            Sleep, 100
            break
        }
        else
        {
            ; 비활성 상태면 PPT 창 활성화 시도
            WinActivate, ahk_exe POWERPNT.EXE
            Sleep, 100
        }
    }
    else
    {
        ; PPT 창이 열려있지 않다면 대기
        Sleep, 100
    }
}





send, {del}
sleep, 100
send, {esc}
sleep, 100





SetFormat, float, 0.0
newPic := sld.Shapes.AddPicture(FileNamedown, 0, -1, 이미지수정왼쪽, 이미지수정위쪽, 이미지수정가로, 이미지수정세로)








작업끝표시(50, "red")
ToolTIP, ★ Transparent completed , %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000

return





; 3/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위










;이미지 해상도 자동2배
;^!+PrintScreen::

고해상도이미지실행:


    mousegetpos, xx1, yy1 ; 행동 후 원위치 복귀를 위한 현재 위치 저장

UpscaylPath := "C:\Program Files\Upscayl\Upscayl.exe"
IfNotExist, %UpscaylPath%
{

MsgBox, 262208, Message, %msgboxuni_0023%

    ; 설치되어 있지 않으면 바로 완료 표시로 이동
    작업끝표시(50, "red")

    return
}


Process, Close, Upscayl.exe
sleep, 100









; 그림=13, 그룹=6, 표=19, 도형=1, 텍스트상자=17, 그룹과 단체=-2, 벡터=28
ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

if (객체확인=28 || 객체확인=5 || 객체확인=1 || 객체확인=19 || 객체확인=17)
{
;넘기기

작업끝표시(50, "red")
sleep, 1
ToolTIP, ★ Convert to bitmap image , %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000

return
}
else
{
gosub, 이미지원래대로
}


;일단 객체하나 복사하기
send, ^d
sleep, 100







;바탕화면에 폴더만들기 " 해상도변환 "

DesktopPath := A_Desktop
BaseFolderName := "해상도변환"
FullPath := DesktopPath "\" BaseFolderName

If FileExist(FullPath)
{
    index := 1
    Loop
    {
        newFullPath := FullPath . index
        if !FileExist(newFullPath)
        {
            FileCreateDir, % newFullPath
            FullPath := newFullPath
            break
        }
        index++
    }
}
else
{
    FileCreateDir, % FullPath
}






;위 폴더에 다른이름으로 저장하기

ppt := ComObjActive("PowerPoint.Application")

window := ppt.ActiveWindow
sel := window.Selection    
shpRange := sel.ShapeRange
sld := window.View.Slide    

이미지수정왼쪽 := shpRange.Left
이미지수정위쪽 := shpRange.Top
이미지수정가로 := shpRange.Width
이미지수정세로 := shpRange.Height

FormatTime, TimeString,, yyyyMMdd_HHmmss
FileName := FullPath . "\Selection_" . TimeString . ".png"
shpRange.Export(FileName, 2)





run, C:\Program Files\Upscayl\Upscayl.exe
sleep, 3000

;Upscayl
;ahk_class Chrome_WidgetWin_1



loop
{
if WinActive("Upscayl")
{
sleep, 100
break
}
sleep, 100
}


sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100

send, {enter}
sleep, 100



loop
{
if WinActive("Select Image")
{
sleep, 100
break
}
sleep, 100
}



Clipboard = %FileName%
ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 100

send, ^v
sleep, 100
send, {enter}
sleep, 1000






;if WinExist("Select Image")
loop
{
if !WinActive("Select Image")
{
sleep, 100
break
}
sleep, 100
}





sleep, 100

send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100
send, {tab}
sleep, 100

send, {enter}
sleep, 1000







; [1단계] 감지할 폴더 경로 설정
TargetDir := FullPath

; [2단계] 루프 진입 전, '현재 존재하는 파일들의 이름'을 모두 저장해둠
ExistingFilesList := "|"
Loop, Files, %TargetDir%\*.*
{
    ; 파일명 앞뒤에 |를 붙여서 구분자로 활용 (예: |file1.png|file2.png|)
    ExistingFilesList .= A_LoopFileName . "|"
}

; 초기 개수 파악 (비교용)
InitialFileCount := 0
Loop, Parse, ExistingFilesList, |
{
    if (A_LoopField != "")
        InitialFileCount++
}

; [3단계] 감시 루프 시작
Loop
{
    CurrentFileCount := 0
    Loop, Files, %TargetDir%\*.*
    {
        CurrentFileCount++
    }

    ; 파일 개수가 늘어났다면 루프 탈출
    if (CurrentFileCount > InitialFileCount)
    {
        Sleep, 1500 ; 파일 쓰기 완료 대기
        break
    }
    
    Sleep, 1000
}

; [4단계] 아까 적어둔 명단(ExistingFilesList)에 '없는' 파일 찾기
NewFilePath := ""
Loop, Files, %TargetDir%\*.*
{
    ; "현재 파일명이 아까 리스트에 포함되어 있지 않다면" -> 신규 파일임
    if !InStr(ExistingFilesList, "|" . A_LoopFileName . "|")
    {
        NewFilePath := A_LoopFileFullPath
        break ; 찾았으니 종료
    }
}















; Upscayl 창 비활성 확인
Loop
{
    ; 1) Upscayl 창이 존재하지 않으면 (또는 비활성이라면) 루프 종료
    if !WinExist("Upscayl") || !WinActive("Upscayl")
    {
        Sleep, 100
        break
    }
    
    ; 창이 활성화되어 있다면 최소화 진행
    WinMinimize, Upscayl
    Sleep, 100
}








; PPTFrameClass 창 활성 확인
Loop
{
    ; PPT 창이 존재하면
    if WinExist("ahk_exe POWERPNT.EXE")
    {
        ; 이미 활성화되어 있으면 루프 종료
        if WinActive("ahk_exe POWERPNT.EXE")
        {
            Sleep, 100
            break
        }
        else
        {
            ; 비활성 상태면 PPT 창 활성화 시도
            WinActivate, ahk_exe POWERPNT.EXE
            Sleep, 100
        }
    }
    else
    {
        ; PPT 창이 열려있지 않다면 대기
        Sleep, 100
    }
}





send, {del}
sleep, 100
send, {esc}
sleep, 100





SetFormat, float, 0.0
newPic := sld.Shapes.AddPicture(NewFilePath, 0, -1, 이미지수정왼쪽, 이미지수정위쪽, 이미지수정가로, 이미지수정세로)








Process, Close, Upscayl.exe
Process, Close, Upscayl.exe
sleep, 100




작업끝표시(50, "red")
ToolTIP, ★ High resolution completed , %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000



return








; 4/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위







#If WinActive("ahk_exe POWERPNT.EXE") and GetKeyState("space", "P")

$tab::
;가이드보이기
SendInput, {Alt down}{F9}{Alt up}
return

#If ; 조건문 끝







; ★★★★★★★★★파워포인트가 활성화 되었을때만 아래의 단축키를 적용하는 스크립트 시작 ========================================

#If WinActive("ahk_exe POWERPNT.EXE")





;$^#!i::
Action_Uni0233:

    ; 파워포인트 연결
    ppt := ComObjActive("PowerPoint.Application")

    ; 활성 창이 없으면 리턴
    If (ppt.Windows.Count = 0)
        Return

    ; 선택 영역 가져오기
    sel := ppt.ActiveWindow.Selection

    ; 선택 유형이 텍스트(2)나 도형(3)인 경우
    If (sel.Type = 2 || sel.Type = 3)
    {
        ; 객체 수 저장
        shCount := sel.ShapeRange.Count
        
        If (shCount > 0)
        {
            ; [토글 로직]
            ; 1. 첫 번째 도형의 현재 상태를 확인합니다.
            ; InsetPen 속성: -1 (msoTrue, 안쪽), 0 (msoFalse, 중앙)
            Try {
                firstShapeState := sel.ShapeRange.Item(1).Line.InsetPen
            } Catch {
                firstShapeState := 0 ; 에러 시 기본값(중앙)으로 가정
            }

            ; 2. 상태에 따라 적용할 목표값을 설정합니다.
            ; 현재가 안쪽(-1)이면 -> 중앙(0)으로
            ; 현재가 중앙(0)이면 -> 안쪽(-1)으로
            If (firstShapeState = -1)
                targetState := 0
            Else
                targetState := -1

            ; 3. 모든 도형에 목표값 적용
            Loop, %shCount%
            {
                Try
                {
                    targetShape := sel.ShapeRange.Item(A_Index)
                    targetShape.Line.InsetPen := targetState
                }
            }
        }
    }
Return
















;마스터페이지 보기
;$!m::
Action_Uni0331:
Send, {Alt down}{w}{Alt up}
Send, {m}
return

;기본으로 돌아가기
;$!space::
Action_Uni0332:
Send, {Alt down}{w}{Alt up}
Send, {l}
return





$^#0::
$^#Numpad0::
;★이단축키는 인디자인이 실행되어 막아둠★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
return

$^#o::
;★이단축키는 화상키보드가 실행되어 막아둠★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
return














$tab::

; 파워포인트 애플리케이션 객체 연결
Try {
    ppt := ComObjActive("PowerPoint.Application")
}

 
    try {
        ; 현재 선택된 영역의 타입(Type)을 가져옴
        selType := ppt.ActiveWindow.Selection.Type
        
        ; -----------------------------------------------------------
        ; ppSelectionShapes (2) : 도형 자체가 선택됨 (첫 번째 이미지)
        ; -----------------------------------------------------------
        if (selType = 2) {
; 도형이 선택된 상태입니다.
선택됨=1
gosub, 탭누르기
        }
        
        ; -----------------------------------------------------------
        ; ppSelectionText (3) : 텍스트 편집 중 (두 번째 이미지)
        ; -----------------------------------------------------------
        else if (selType = 3) {
; 도형 내부에서 커서가 깜빡이거나 글자가 드래그된 상태입니다.
SendInput, {tab}
        }
        
        ; -----------------------------------------------------------
        ; 그 외 (선택 없음 등)
        ; -----------------------------------------------------------
        else if (selType = 0) {
; 아무것도 선택되지 않았습니다.
gosub, 탭누르기
        }
        else {
            MsgBox, 262208, Message, %msgboxuni_0024%`n(%selType%)
        }
    } 
    catch e {
        MsgBox, 262208, Message, %msgboxuni_0015%
    }



선택됨=0




return





탭누르기:


    ControlGetFocus, focusedCtrl, A

    ; 파워포인트 편집 화면(mdiClass1)일 때만 작동
    if (focusedCtrl = "mdiClass1") {

        ppt := ComObjActive("PowerPoint.Application")
                
        ControlGetText, 현재창상황, MsoCommandBar1, ahk_exe POWERPNT.EXE


            ; ==================================================================
            ; [상황 A] 편집 모드 (리본 메뉴가 보임) -> 안내선 켜기(ON)
            ; ==================================================================
            if (현재창상황 = "Ribbon") 
            {

                ; 1) 눈금자 끄기
                try {
                    ppt.ActiveWindow.ViewType := 9
                    if (ppt.CommandBars.GetPressedMso("ViewRulerPowerPoint")) {
                        ppt.CommandBars.ExecuteMso("ViewRulerPowerPoint")
                    }
                }



                ; 3) 리본 메뉴 펼치기
                try {
                    if (ppt.CommandBars.GetPressedMso("MinimizeRibbon")) {
                        ppt.CommandBars.ExecuteMso("MinimizeRibbon")
                    }
                }



/*

;눈금자 항상켜기
try {
ppt.ActiveWindow.ViewType := 9  ; ppViewNormal
; 현재 눈금자 상태 확인
isRulerOn := ppt.CommandBars.GetPressedMso("ViewRulerPowerPoint")
; 꺼져있으면 켜기
if (!isRulerOn) {
    ppt.CommandBars.ExecuteMso("ViewRulerPowerPoint")
}
}



;리본메뉴 항상켜기
try {
    isMin := ppt.CommandBars.GetPressedMso("MinimizeRibbon")     ; True면 최소화됨
    if (isMin) {
        ppt.CommandBars.ExecuteMso("MinimizeRibbon")             ; 펼침
    }
}


; 현재 노트 항상켜기
try {
    isNotesVisible := ppt.CommandBars.GetPressedMso("ShowNotes") ; True = 표시됨
    if (!isNotesVisible) {
        ppt.CommandBars.ExecuteMso("ShowNotes")  ; 꺼져 있으면 켜기
    }
}

*/




                ; 4) 노트 끄기
                try {
                    if (ppt.CommandBars.GetPressedMso("ShowNotes")) {
                        ppt.CommandBars.ExecuteMso("ShowNotes")
                    }
                }


if (선택됨=1) {
                ; 5) 마무리
                SendInput, {AppsKey}{o}
;                Sleep, 1
                SendInput, {esc}
} 
else {
                ; 5) 마무리
                SendInput, {AppsKey}{b}
;                Sleep, 1
                SendInput, {esc}
}





            }


            ; ==================================================================
            ; [상황 B] 깔끔 모드 (리본 메뉴가 안 보임) -> 안내선 끄기(OFF)
            ; ==================================================================
            else if (현재창상황 != "Ribbon") 
            {

                ; 1) 우측 서식 창 닫기
                MouseGetPos, xx1, yy1 
                ControlGetPos, x, y, w, h, MsoDockRight, ahk_exe POWERPNT.EXE
                if (x != "") {
                    MouseClick, left, x + w - 23, y + 23 
 ;                   Sleep, 1
                    MouseMove, %xx1%, %yy1%
                }

                ; 2) 눈금자 끄기
                try {
                    ppt.ActiveWindow.ViewType := 9 
                    if (ppt.CommandBars.GetPressedMso("ViewRulerPowerPoint")) {
                        ppt.CommandBars.ExecuteMso("ViewRulerPowerPoint")
                    }
                }




                ; 4) 노트 끄기
                try {
                    if (ppt.CommandBars.GetPressedMso("ShowNotes")) {
                        ppt.CommandBars.ExecuteMso("ShowNotes")
                    }
                }



/*
; 현재 노트 항상끄기
try {
    isNotesVisible := ppt.CommandBars.GetPressedMso("ShowNotes") ; True = 표시됨
    if (isNotesVisible) {
        ppt.CommandBars.ExecuteMso("ShowNotes")  ; 켜져 있으면 끄기
    }
}


; 리본메뉴 항상 숨기기 (최소화 상태로)
try {
    isMin := ppt.CommandBars.GetPressedMso("MinimizeRibbon")  ; True = 최소화, False = 펼침
    if (!isMin) {
        ppt.CommandBars.ExecuteMso("MinimizeRibbon")          ; 최소화로 변경
    }
}

*/









            }

}
else {
SendInput, {tab}
}





return



















;셀상단수정-포인트라인두께1
$^#!Numpad8::


targetColor := HexToBGR(포인트테이블라인색상1)
targetWeight := 포인트라인두께1

gosub, 셀상단수정

return


;셀하단수정-포인트라인두께1
$^#!Numpad2::

targetColor := HexToBGR(포인트테이블라인색상1)
targetWeight := 포인트라인두께1

gosub, 셀하단수정

return


;셀좌측수정-포인트라인두께1
$^#!Numpad4::

targetColor := HexToBGR(포인트테이블라인색상1)
targetWeight := 포인트라인두께1

gosub, 셀좌측수정

return


;셀우측수정-포인트라인두께1
$^#!Numpad6::

targetColor := HexToBGR(포인트테이블라인색상1)
targetWeight := 포인트라인두께1

gosub, 셀우측수정

return






;셀상단수정-포인트라인두께2
$#!Numpad8::

targetColor := HexToBGR(포인트테이블라인색상2)
targetWeight := 포인트라인두께2

gosub, 셀상단수정

return


;셀하단수정-포인트라인두께2
$#!Numpad2::

targetColor := HexToBGR(포인트테이블라인색상2)
targetWeight := 포인트라인두께2

gosub, 셀하단수정

return


;셀좌측수정-포인트라인두께2
$#!Numpad4::

targetColor := HexToBGR(포인트테이블라인색상2)
targetWeight := 포인트라인두께2

gosub, 셀좌측수정

return


;셀우측수정-포인트라인두께2
$#!Numpad6::

targetColor := HexToBGR(포인트테이블라인색상2)
targetWeight := 포인트라인두께2

gosub, 셀우측수정

return











;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★ 선택한 셀의 상단⭡ 바꾸는 스크립트 성공
셀상단수정:

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table

/*
;위에서 값을 받아옴
                targetWeight := 3
                targetColor := 0xE1B146 
*/
                
                ; 선택된 영역 중 가장 위쪽 행(Row) 번호를 찾기 위한 변수
                minRow := 9999 
                
                ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 작은 행(Row) 번호 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (r < minRow)
                                minRow := r
                        }
                    }
                }
                
                ; [2단계] 찾아낸 최상단 행(minRow)에 속한 선택된 셀에만 위쪽 테두리 적용
                applyCount := 0
                if (minRow != 9999) {
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        cell := tbl.Cell(minRow, c)
                        
                        if (cell.Selected) {
                            applyCount++
                            border := cell.Borders.Item(1) ; 1 = Top Border
                            border.Weight := targetWeight
                            border.ForeColor.RGB := targetColor
                            border.Visible := -1
                            border.Transparency := 0
                        }
                    }
                }
                
                if (applyCount > 0) {
                    ToolTip, ★ Top Border ⭡
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
        MsgBox, 262208, Message, %targetColor%
    }
return






;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★ 선택한 셀의 하단⭣ 바꾸는 스크립트 성공
셀하단수정:

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table
                
/*
;위에서 값을 받아옴
                targetWeight := 3
                targetColor := 0xE1B146 
*/
                
                ; 선택된 영역 중 가장 아래쪽 행(Row) 번호를 찾기 위한 변수 초기화
                maxRow := 0 
                
                ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 큰 행(Row) 번호 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (r > maxRow)
                                maxRow := r
                        }
                    }
                }
                
                ; [2단계] 찾아낸 최하단 행(maxRow)에 속한 선택된 셀에만 아래쪽 테두리 적용
                applyCount := 0
                if (maxRow != 0) {
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        cell := tbl.Cell(maxRow, c)
                        
                        if (cell.Selected) {
                            applyCount++
                            border := cell.Borders.Item(3) ; 3 = Bottom Border (아래쪽 테두리)
                            border.Weight := targetWeight
                            border.ForeColor.RGB := targetColor
                            border.Visible := -1
                            border.Transparency := 0
                        }
                    }
                }
                
                if (applyCount > 0) {
                    ToolTip, ★ Bottom Border ⭣
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return





;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★ 선택한 셀의 ⭠좌측 바꾸는 스크립트 성공
셀좌측수정:

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table
                
/*
;위에서 값을 받아옴
                targetWeight := 3
                targetColor := 0xE1B146 
*/
                
                ; 선택된 영역 중 가장 왼쪽 열(Column) 번호를 찾기 위한 변수
                minCol := 9999 
                
                ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 작은 열(Column) 번호 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (c < minCol)
                                minCol := c
                        }
                    }
                }
                
                ; [2단계] 찾아낸 최좌측 열(minCol)에 속한 선택된 셀에만 왼쪽 테두리 적용
                applyCount := 0
                if (minCol != 9999) {
                    Loop, % tbl.Rows.Count {
                        r := A_Index
                        cell := tbl.Cell(r, minCol)
                        
                        if (cell.Selected) {
                            applyCount++
                            border := cell.Borders.Item(2) ; 2 = Left Border (왼쪽 테두리)
                            border.Weight := targetWeight
                            border.ForeColor.RGB := targetColor
                            border.Visible := -1
                            border.Transparency := 0
                        }
                    }
                }
                
                if (applyCount > 0) {
                    ToolTip, ★ ⭠ Left Border
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return





;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★ 선택한 셀의 우측⭢ 바꾸는 스크립트 성공
셀우측수정:

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table
                
/*
;위에서 값을 받아옴
                targetWeight := 3
                targetColor := 0xE1B146
*/
                
                ; 선택된 영역 중 가장 오른쪽 열(Column) 번호를 찾기 위한 변수 초기화
                maxCol := 0 
                
                ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 큰 열(Column) 번호 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (c > maxCol)
                                maxCol := c
                        }
                    }
                }
                
                ; [2단계] 찾아낸 최우측 열(maxCol)에 속한 선택된 셀에만 오른쪽 테두리 적용
                applyCount := 0
                if (maxCol != 0) {
                    Loop, % tbl.Rows.Count {
                        r := A_Index
                        cell := tbl.Cell(r, maxCol)
                        
                        if (cell.Selected) {
                            applyCount++
                            border := cell.Borders.Item(4) ; 4 = Right Border (오른쪽 테두리)
                            border.Weight := targetWeight
                            border.ForeColor.RGB := targetColor
                            border.Visible := -1
                            border.Transparency := 0
                        }
                    }
                }
                
                if (applyCount > 0) {
                    ToolTip, ★ Right Border ⭢
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return












;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★ 선택한 셀의 상하좌우가운데⭠⭡⭢⭣ 모두 바꾸는 스크립트 성공

$^#!Numpad1::

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 1. 선택된 개체가 도형/표 형태인지 확인 (ppSelectionShape = 2)
        if (sel.Type = 2 || sel.Type = 3) {
            ; 2. 첫 번째로 선택된 개체가 표(Table)인지 확인
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table
                

targetColor := HexToBGR(기본테이블라인색상)
targetWeight := 기본테이블라인두께


                
                applyCount := 0
                
전체카운트 := tbl.Rows.Count*tbl.Columns.Count
증가 := 0

                ; 4. 표 안의 모든 셀을 순회하며 '선택된(Selected)' 셀만 적용
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        cell := tbl.Cell(r, c)
                        
                        ; 해당 셀이 드래그로 선택된 상태인지 확인
                        if (cell.Selected) {
                            applyCount++
                            
                            ; 1:Top, 2:Left, 3:Bottom, 4:Right 테두리에 적용
                            Loop, 4 {
                                border := cell.Borders.Item(A_Index)
                                border.Weight := targetWeight
                                border.ForeColor.RGB := targetColor
                                border.Visible := -1 ; msoTrue (테두리 선 켜기)
                                border.Transparency := 0
                            }
                        }

증가 := 증가 + 1
ToolTIP, ★ %증가%/%전체카운트%, %xx1%, %yy1%


                    }
                }
                
                ; 5. 결과 알림
                if (applyCount > 0) {

작업끝표시(50, "red")
                    ToolTip, ★ Completed %applyCount% cells!
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return






;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★ 선택한 셀의 가로 라인만 칠하는 스크립트
$^#!Numpad5::

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table
                
targetColor := HexToBGR(기본테이블라인색상)
targetWeight := 기본테이블라인두께

전체카운트 := tbl.Rows.Count*tbl.Columns.Count
증가 := 0

                applyCount := 0
               
                ; 2. 선택된 모든 셀을 순회하며 테두리 일괄 적용
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        cell := tbl.Cell(r, c)
                        
                        if (cell.Selected) {
                            applyCount++

                            ; [가로 라인] 위(1), 아래(3)는 설정된 색상과 두께로 칠함
                            cell.Borders.Item(1).Weight := targetWeight
                            cell.Borders.Item(1).ForeColor.RGB := targetColor
                            cell.Borders.Item(1).Visible := -1
                            cell.Borders.Item(1).Transparency := 0
                            
                            cell.Borders.Item(3).Weight := targetWeight
                            cell.Borders.Item(3).ForeColor.RGB := targetColor
                            cell.Borders.Item(3).Visible := -1
                            cell.Borders.Item(3).Transparency := 0
                            
                            ; [세로 라인] 왼쪽(2), 오른쪽(4)은 투명도(0) 및 숨김(0) 처리
                            cell.Borders.Item(2).Weight := 0
                            cell.Borders.Item(2).Visible := 0 
                            cell.Borders.Item(2).Transparency := 1
                            
                            cell.Borders.Item(4).Weight := 0
                            cell.Borders.Item(4).Visible := 0
                            cell.Borders.Item(4).Transparency := 1

증가 := 증가 + 1
ToolTIP, ★ %증가%/%전체카운트%, %xx1%, %yy1%


                        }
                    }
                }
                
                if (applyCount > 0) {
                    작업끝표시(50, "red")
                    ToolTip, ★ 가로선 적용 및 세로선 제거 완료
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return









;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★ 선택한 셀의 전체를 칠하고 좌우라인만 없애는 스크립트

$^#!Numpad9::

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table
                
targetColor := HexToBGR(기본테이블라인색상)
targetWeight := 기본테이블라인두께
                
                ; 선택 영역의 양 끝(좌/우) 열 번호를 찾기 위한 변수
                minCol := 9999 
                maxCol := 0



전체카운트 := tbl.Rows.Count*tbl.Columns.Count
증가 := 0

 
                ; [1단계] 표 전체를 훑어서 선택된 영역의 최소 열(좌측 끝)과 최대 열(우측 끝) 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (c < minCol)
                                minCol := c
                            if (c > maxCol)
                                maxCol := c
                        }
                    }
                }
                
                ; [2단계] 선택된 모든 셀에 테두리 적용하되 양끝만 예외 처리
                applyCount := 0
                if (minCol != 9999 && maxCol != 0) {
                    Loop, % tbl.Rows.Count {
                        r := A_Index
                        Loop, % tbl.Columns.Count {
                            c := A_Index
                            cell := tbl.Cell(r, c)
                            
                            if (cell.Selected) {
                                applyCount++
 

                                ; 위(1), 아래(3)는 무조건 칠함
                                cell.Borders.Item(1).Weight := targetWeight
                                cell.Borders.Item(1).ForeColor.RGB := targetColor
                                cell.Borders.Item(1).Visible := -1
                                cell.Borders.Item(1).Transparency := 0
                                
                                cell.Borders.Item(3).Weight := targetWeight
                                cell.Borders.Item(3).ForeColor.RGB := targetColor
                                cell.Borders.Item(3).Visible := -1
                                cell.Borders.Item(3).Transparency := 0
                                
                                ; 왼쪽(2) 테두리 로직: 가장 좌측 열이면 끄고, 아니면 칠함
                                if (c = minCol) {
                                    cell.Borders.Item(2).Weight := 0
                                    cell.Borders.Item(2).Visible := 0 ; msoFalse (선 지우기)
                                    cell.Borders.Item(2).Transparency := 1
                                } else {
                                    cell.Borders.Item(2).Weight := targetWeight
                                    cell.Borders.Item(2).ForeColor.RGB := targetColor
                                    cell.Borders.Item(2).Visible := -1
                                    cell.Borders.Item(2).Transparency := 0
                                }
                                
                                ; 오른쪽(4) 테두리 로직: 가장 우측 열이면 끄고, 아니면 칠함
                                if (c = maxCol) {
                                    cell.Borders.Item(4).Weight := 0
                                    cell.Borders.Item(4).Visible := 0 ; msoFalse (선 지우기)
                                    cell.Borders.Item(4).Transparency := 1
                                } else {
                                    cell.Borders.Item(4).Weight := targetWeight
                                    cell.Borders.Item(4).ForeColor.RGB := targetColor
                                    cell.Borders.Item(4).Visible := -1
                                    cell.Borders.Item(4).Transparency := 0
                                }

증가 := 증가 + 1
ToolTIP, ★ %증가%/%전체카운트%, %xx1%, %yy1%

                            }
                        }
                    }
                }
                
                if (applyCount > 0) {

작업끝표시(50, "red")
                    ToolTip, ★ Complete border
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return







;셀 상단 라인없애기
~x & Numpad8::

    if GetKeyState("Alt", "P") {

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection

targetColor := HexToBGR(기본테이블라인색상)
targetWeight := 기본테이블라인두께
        
        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table

                
                ; 선택된 영역 중 가장 위쪽 행(Row) 번호를 찾기 위한 변수
                minRow := 9999 


                ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 작은 행(Row) 번호 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (r < minRow)
                                minRow := r
                        }
                    }
                }
                
                ; [2단계] 찾아낸 최상단 행(minRow)에 속한 선택된 셀에만 위쪽 테두리 적용
                applyCount := 0
                if (minRow != 9999) {
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        cell := tbl.Cell(minRow, c)
                        
                        if (cell.Selected) {
                            applyCount++

                            border := cell.Borders.Item(1) ; 1 = Top Border (위쪽 테두리)
                            border.Weight := 0
                            border.Visible := 0
                            border.Transparency := 1

                        }
                    }
                }
                
                if (applyCount > 0) {
                    ToolTip, ★ Top Border ⭡
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
        MsgBox, 262208, Message, %targetColor%
    }
return


    }
return







;셀 좌측 라인없애기 (기존 구조 수정본)
~x & Numpad4::
    if GetKeyState("Alt", "P") {
        try {
            ppt := ComObjActive("PowerPoint.Application")
            sel := ppt.ActiveWindow.Selection
            
            targetColor := HexToBGR(기본테이블라인색상)
            targetWeight := 기본테이블라인두께
            
            ; 1. 도형/표 선택 상태 확인
            if (sel.Type = 2 || sel.Type = 3) {
                if (sel.ShapeRange.Item(1).HasTable) {
                    tbl := sel.ShapeRange.Item(1).Table
                    
                    ; 선택된 영역 중 가장 왼쪽 열(Column) 번호를 찾기 위한 변수
                    minCol := 9999 
                    
                    ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 작은 열(Column) 번호 찾기
                    Loop, % tbl.Rows.Count {
                        r := A_Index
                        Loop, % tbl.Columns.Count {
                            c := A_Index
                            if (tbl.Cell(r, c).Selected) {
                                if (c < minCol)
                                    minCol := c
                            }
                        }
                    }
                    
                    ; [2단계] 찾아낸 최좌측 열(minCol)에 속한 선택된 셀에만 왼쪽 테두리 제거 적용
                    applyCount := 0
                    if (minCol != 9999) {
                        Loop, % tbl.Rows.Count {
                            r := A_Index
                            cell := tbl.Cell(r, minCol)
                            
                            if (cell.Selected) {
                                applyCount++
                                border := cell.Borders.Item(2) ; 2 = Left Border (왼쪽 테두리)
                                border.Weight := 0
                                border.Visible := 0
                                border.Transparency := 1
                            }
                        }
                    }
                    
                    if (applyCount > 0) {
                        ToolTip, ★ Left Border ⭠
                        SetTimer, RemoveToolTip, -1500
                    } else {
                        MsgBox, 262208, Message, %msgboxuni_0027%
                    }
                    
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0028%
                }
            } else {
                MsgBox, 262208, Message, %msgboxuni_0027%
            }
        } catch e {
            MsgBox, 262208, Message, %msgboxuni_0029%
        }
    }
return










;셀 하단 라인없애기
~x & Numpad2::

    if GetKeyState("Alt", "P") {

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
targetColor := HexToBGR(기본테이블라인색상)
targetWeight := 기본테이블라인두께

        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table

                
                ; 선택된 영역 중 가장 아래쪽 행(Row) 번호를 찾기 위한 변수 초기화
                maxRow := 0 
                
                ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 큰 행(Row) 번호 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (r > maxRow)
                                maxRow := r
                        }
                    }
                }
                
                ; [2단계] 찾아낸 최하단 행(maxRow)에 속한 선택된 셀에만 아래쪽 테두리 적용
                applyCount := 0
                if (maxRow != 0) {
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        cell := tbl.Cell(maxRow, c)
                        
                        if (cell.Selected) {
                            applyCount++

                            border := cell.Borders.Item(3) ; 3 = Bottom Border (아래쪽 테두리)
                            border.Weight := 0
                            border.Visible := 0
                            border.Transparency := 1

                        }
                    }
                }
                
                if (applyCount > 0) {
                    ToolTip, ★ Bottom Border ⭣
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return





    }
return









;셀 우측 라인없애기
~x & Numpad6::

    if GetKeyState("Alt", "P") {


    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        

targetColor := HexToBGR(기본테이블라인색상)
targetWeight := 기본테이블라인두께

        ; 1. 도형/표 선택 상태 확인
        if (sel.Type = 2 || sel.Type = 3) {
            if (sel.ShapeRange.Item(1).HasTable) {
                tbl := sel.ShapeRange.Item(1).Table

                
                ; 선택된 영역 중 가장 오른쪽 열(Column) 번호를 찾기 위한 변수 초기화
                maxCol := 0 
                
                ; [1단계] 표 전체를 훑어서 선택된 셀 중 가장 큰 열(Column) 번호 찾기
                Loop, % tbl.Rows.Count {
                    r := A_Index
                    Loop, % tbl.Columns.Count {
                        c := A_Index
                        if (tbl.Cell(r, c).Selected) {
                            if (c > maxCol)
                                maxCol := c
                        }
                    }
                }
                
                ; [2단계] 찾아낸 최우측 열(maxCol)에 속한 선택된 셀에만 오른쪽 테두리 적용
                applyCount := 0
                if (maxCol != 0) {
                    Loop, % tbl.Rows.Count {
                        r := A_Index
                        cell := tbl.Cell(r, maxCol)
                        
                        if (cell.Selected) {
                            applyCount++
                            border := cell.Borders.Item(4) ; 4 = Right Border (오른쪽 테두리)
                            border.Weight := 0
                            border.Visible := 0
                            border.Transparency := 1
                        }
                    }
                }
                
                if (applyCount > 0) {
                    ToolTip, ★ Right Border ⭢
                    SetTimer, RemoveToolTip, -1500
                } else {
                    MsgBox, 262208, Message, %msgboxuni_0027%
                }
                
            } else {
                MsgBox, 262208, Message, %msgboxuni_0028%
            }
        } else {
            MsgBox, 262208, Message, %msgboxuni_0027%
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0029%
    }
return


    }
return



















; 기존속성복사 ========================================================
;$^+e::
Action_Uni0291:

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크


Try
{
    ppt := ComObjActive("PowerPoint.Application")
    selection := ppt.ActiveWindow.Selection

        shape := selection.ShapeRange

        ; 가로 크기, 세로 크기, 가로 위치, 세로 위치 값을 가져옵니다.
        너비 := shape.Width
        높이 := shape.Height
        가로위치 := shape.Left
        세로위치 := shape.Top
}


작업끝표시(50, "red")

ToolTIP, ★ Copy properties , %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000

gosub, 키보드올리기
return













; 기존속성붙여넣기 ========================================================


;같은위치
;$^+d::
Action_Uni0292:
    mousegetpos, xx1, yy1 ; 행동 후 원위치 복귀를 위한 현재 위치 저장

Try
	{
ppt := ComObjActive("PowerPoint.Application")
ppt.ActiveWindow.Selection.Shaperange.LockAspectRatio:=False
	}



Try
{
    ppt := ComObjActive("PowerPoint.Application")
    selection := ppt.ActiveWindow.Selection

        shape := selection.ShapeRange

        ; 가로 크기, 세로 크기, 가로 위치, 세로 위치 값을 가져옵니다.
shape.Left:=가로위치
shape.Top:=세로위치
}



작업끝표시(50, "red")

ToolTIP, ★ Paste location, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000

gosub, 키보드올리기
return










;$^+f::
Action_Uni0293:

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크



Try
	{
ppt := ComObjActive("PowerPoint.Application")
ppt.ActiveWindow.Selection.Shaperange.LockAspectRatio:=False
	}



Try
{
    ppt := ComObjActive("PowerPoint.Application")
    selection := ppt.ActiveWindow.Selection

        shape := selection.ShapeRange

        ; 가로 크기, 세로 크기, 가로 위치, 세로 위치 값을 가져옵니다.
shape.Width:=너비
shape.Height:=높이
}


작업끝표시(50, "red")

ToolTIP, ★ Paste location & size , %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000

gosub, 키보드올리기
return









;==========================================================================================
;여백조정 =======================================================================


$+F9::

ppt := ComObjActive("PowerPoint.Application")
객체확인타입:=ppt.ActiveWindow.Selection.Type



;커서가 살아있으면 다시 없애기
if (객체확인타입=3)
{
send, {esc}
객체확인타입:=ppt.ActiveWindow.Selection.Type
}




if (객체확인타입!=0)
{

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크
gosub, 핫키올림확인

/*
WinActivate, %파워포인트타이틀%
sleep, 1
*/


Try
{
ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type
}

;★테이블일때---------------------------------------------------------------------------
 if (객체확인=19)
{




;테이블일수도있으니 이것도 실행
Try
{
    ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
    ; 선택된 테이블에 대한 참조를 저장합니다.
    Table := ppt.ActiveWindow.Selection.ShapeRange.Table

    ; 테이블의 모든 행과 열에 대해 반복합니다.
    Rows := Table.Rows.Count
    Columns := Table.Columns.Count
전체카운트 := Rows*Columns
증가 := 0
    Loop, %Rows%
    {
        row := A_Index
        Loop, %Columns%
        {
            col := A_Index
            ; 각 셀을 참조합니다.
            Cell := Table.Cell(row, col)
            ; 선택된 셀인지 확인합니다. 
	    ; Cell.Selected 이값은 선택이면 1, 아니면 0으로 나옴

            If (Cell.Selected)
            {


Try
{
                ; 선택된 셀의 여백을 0으로 설정합니다.
            Cell.Shape.TextFrame.MarginLeft := 0
            Cell.Shape.TextFrame.MarginRight := 0
            Cell.Shape.TextFrame.MarginTop := 0
            Cell.Shape.TextFrame.MarginBottom := 0
}

Try
{
자간조정:=-0.4
            Cell.Shape.TextFrame2.TextRange.Font.Spacing:=자간조정
}


Try
{
행간조정:=1.09
            Cell.Shape.TextFrame2.TextRange.ParagraphFormat.LineRuleWithin:=-1
            Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin:=행간조정
}



Try
{
;내어쓰기 0으로 수정하기
            Cell.Shape.TextFrame2.TextRange.ParagraphFormat.LeftIndent := 0
}

Try
{
;첫줄들여쓰기 0으로 수정하기
            Cell.Shape.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent := 0
}

Try
{
;단락 앞뒤 간격 0으로 수정하기
            Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceAfter := 0
            Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceBefore := 0
}

Try
{
; 머리글 기호없음
            Cell.Shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible := 0
}




증가 := 증가 + 1
ToolTIP, ★ %증가%/%전체카운트%, %xx1%, %yy1%

            }
        }
    }

}











}
else
{




;★도형일때 일때---------------------------------------------------------------------------




SetFormat, float, 0.2



Try
	{
;자동 맞춤 안 함(D) : 0
;넘치면 텍스트 크기 조정(s) : 2
;도형을 텍스트 크기에 맞춤(f) : 1
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.AutoSize := 0
	}
catch e {
; 자동 맞춤 안 함(D) : 0
}


				;도형안의 텍스트 배치
Try
	{
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.WordWrap := False
	}
catch e {
; 도형안의 텍스트 배치
}

				;강제로 상하좌우여백 0 으로 조절
Try
	{
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginLeft := 0
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginRight := 0
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginTop := 0
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginBottom := 0
	}
catch e {
; 강제로 상하좌우여백 0 으로 조절
}

Try
{
자간조정:=-0.4
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.Spacing:=자간조정
}
catch e {
; 자간조정
}





Try
{
    행간조정 := 1.09
    
    ; 1. 현재 선택(Selection) 객체 가져오기
    oSel := ppt.ActiveWindow.Selection
    


    ; 1-1. (안전 장치) 선택 유형이 '도형'이 맞는지 확인 (ppSelectionShapes = 2) [1, 2, 3, 4, 5]
    If (oSel.Type!= 2)
    {
        MsgBox, 262208, Message, %msgboxuni_0016%
        Return
    }




    ; 2. 선택된 '도형 범위(ShapeRange)' 컬렉션 가져오기
    oShapeRange := oSel.ShapeRange
    
    ; 3. [핵심] For 반복문으로 모든 선택된 도형(oShape)을 순회 [2, 6]
    For oShape in oShapeRange
    {
        ; 4. [수정] 해당 도형이 텍스트를 가질 수 있는지(.HasTextFrame) 먼저 확인 [7]
        If (oShape.HasTextFrame)
        {
            ; 5. 텍스트 프레임 안에 실제 텍스트가 있는지도 확인
            If (oShape.TextFrame.HasText)
            {
                Try
                {
                    ; 6. [수정] TextFrame2가 아닌, 더 표준적인 TextFrame의 서식 객체에 접근 [7, 8]
                    oParaFormat := oShape.TextFrame.TextRange.ParagraphFormat
                    
                    ; 7. '배수'로 설정 [9, 10]
                    oParaFormat.LineRuleWithin := -1 ; (msoTrue)
                    
                    ; 8. 배수 '값' 설정 [9, 10, 11]
                    oParaFormat.SpaceWithin := 행간조정
                }
                Catch
                {
                    ; 그룹 내 객체 등 예외적인 경우 무시
                    Continue
                }
            }
        }
    }
}
catch e
{
;   행간오류
}










Try
{
;내어쓰기 0으로 수정하기
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.LeftIndent := 0
}
catch e {
; 내어쓰기 0으로 수정하기
}

Try
{
;첫줄들여쓰기 0으로 수정하기
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent := 0
}
catch e {
; 첫줄들여쓰기 0으로 수정하기
}

Try
{
;단락 앞뒤 간격 0으로 수정하기
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceAfter := 0
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceBefore := 0
}
catch e {
; 단락 앞뒤 간격 0으로 수정하기
}


Try
{
;이미지가로세로 비율유지 끄기
ppt.ActiveWindow.Selection.Shaperange.LockAspectRatio:=False
}
catch e {
; 이미지가로세로 비율유지 끄기
}


Try
{
윤곽선유무체크:=ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.Line.Visible
if (윤곽선유무체크=0)
{
            shpRange := ppt.ActiveWindow.Selection.ShapeRange
            line := shpRange.TextFrame2.TextRange.Font.Line
            line.Visible      := -1       ; msoTrue
            line.Weight       := 0.75
            line.ForeColor.RGB:= 14277081
            line.Transparency := 1.0
}
}
catch e {
; 윤곽선유무체크
}

				;머리글기호 없음
Try
	{
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible := 0
	}
catch e {
; 머리글기호 없음
}



}




}

작업끝표시(50, "red")








if (객체확인타입!=0)
{
ToolTIP, ★ No margin, %xx1%, %yy1%
}
else
{
ToolTIP, ★ Nothing selected, %xx1%, %yy1%
}

SetTimer, 정보창툴팁없애기, 800




gosub, 키보드올리기





return








; 5/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위





























;여백조정 값
; 0.05 = 1.417325
; 0.02= 0.56693


~MButton & Right::


ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

 if (객체확인=19)
{

;테이블일수도있으니 이것도 실행

Try
{
    ppt := ComObjActive("PowerPoint.Application")
    ; 선택된 테이블에 대한 참조를 저장합니다.
    Table := ppt.ActiveWindow.Selection.ShapeRange.Table

    ; 테이블의 모든 행과 열에 대해 반복합니다.
    Rows := Table.Rows.Count
    Columns := Table.Columns.Count
    Loop, %Rows%
    {
        row := A_Index
        Loop, %Columns%
        {
            col := A_Index
            ; 각 셀을 참조합니다.
            Cell := Table.Cell(row, col)
            ; 선택된 셀인지 확인합니다. 
	    ; Cell.Selected 이값은 선택이면 1, 아니면 0으로 나옴

            If (Cell.Selected)
            {
                ; 선택된 셀의 여백을 0으로 설정합니다.

정렬상태확인:=Cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment
; 좌정렬=1 가운데2 우측3
if (정렬상태확인 = 1) 
{
여백조정:=Cell.Shape.TextFrame.MarginLeft + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginLeft := 여백조정
}
if (정렬상태확인 = 3)
{
여백조정:=Cell.Shape.TextFrame.MarginRight - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginRight := 여백조정
}


; Cell.Shape.TextFrame.MarginLeft := 0
; Cell.Shape.TextFrame.MarginRight := 0
; Cell.Shape.TextFrame.MarginTop := 0
; Cell.Shape.TextFrame.MarginBottom := 0
            }
        }
    }
; 선택된 셀의 여백이 0으로 설정되었습니다.
}



}
else
{


Try
	{
ppt := ComObjActive("PowerPoint.Application")

정렬상태확인:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment
if (정렬상태확인 = 1)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginLeft + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginLeft := 여백조정
}
if (정렬상태확인 = 3)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginRight - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginRight := 여백조정
}

	}

}



gosub, 키보드올리기

return











~MButton & Left::



ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

 if (객체확인=19)
{

;테이블일수도있으니 이것도 실행

Try
{
    ppt := ComObjActive("PowerPoint.Application")
    ; 선택된 테이블에 대한 참조를 저장합니다.
    Table := ppt.ActiveWindow.Selection.ShapeRange.Table

    ; 테이블의 모든 행과 열에 대해 반복합니다.
    Rows := Table.Rows.Count
    Columns := Table.Columns.Count
    Loop, %Rows%
    {
        row := A_Index
        Loop, %Columns%
        {
            col := A_Index
            ; 각 셀을 참조합니다.
            Cell := Table.Cell(row, col)
            ; 선택된 셀인지 확인합니다. 
	    ; Cell.Selected 이값은 선택이면 1, 아니면 0으로 나옴

            If (Cell.Selected)
            {
                ; 선택된 셀의 여백을 0으로 설정합니다.

정렬상태확인:=Cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment
if (정렬상태확인 = 1)
{
여백조정:=Cell.Shape.TextFrame.MarginLeft - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginLeft := 여백조정
}
if (정렬상태확인 = 3)
{
여백조정:=Cell.Shape.TextFrame.MarginRight + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginRight := 여백조정
}


; Cell.Shape.TextFrame.MarginLeft := 0
; Cell.Shape.TextFrame.MarginRight := 0
; Cell.Shape.TextFrame.MarginTop := 0
; Cell.Shape.TextFrame.MarginBottom := 0
            }
        }
    }
; 선택된 셀의 여백이 0으로 설정되었습니다.
}



}
else
{


Try
	{
ppt := ComObjActive("PowerPoint.Application")
정렬상태확인:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment

if (정렬상태확인 = 1)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginLeft - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginLeft := 여백조정
}

if (정렬상태확인 = 3)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginRight + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginRight := 여백조정
}

	}

}


gosub, 키보드올리기

return














~MButton & Down::



ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

 if (객체확인=19)
{

;테이블일수도있으니 이것도 실행

Try
{
    ppt := ComObjActive("PowerPoint.Application")
    ; 선택된 테이블에 대한 참조를 저장합니다.
    Table := ppt.ActiveWindow.Selection.ShapeRange.Table

    ; 테이블의 모든 행과 열에 대해 반복합니다.
    Rows := Table.Rows.Count
    Columns := Table.Columns.Count
    Loop, %Rows%
    {
        row := A_Index
        Loop, %Columns%
        {
            col := A_Index
            ; 각 셀을 참조합니다.
            Cell := Table.Cell(row, col)
            ; 선택된 셀인지 확인합니다. 
	    ; Cell.Selected 이값은 선택이면 1, 아니면 0으로 나옴

            If (Cell.Selected)
            {
                ; 선택된 셀의 여백을 0으로 설정합니다.

정렬상태확인:=Cell.Shape.TextFrame.VerticalAnchor
; 좌정렬=1 가운데2 우측3
if (정렬상태확인 = 1) 
{
여백조정:=Cell.Shape.TextFrame.MarginTop + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginTop := 여백조정
}
if (정렬상태확인 = 4)
{
여백조정:=Cell.Shape.TextFrame.MarginBottom - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginBottom := 여백조정
}


; Cell.Shape.TextFrame.MarginLeft := 0
; Cell.Shape.TextFrame.MarginRight := 0
; Cell.Shape.TextFrame.MarginTop := 0
; Cell.Shape.TextFrame.MarginBottom := 0
            }
        }
    }
; 선택된 셀의 여백이 0으로 설정되었습니다.
}



}
else
{


Try
	{
ppt := ComObjActive("PowerPoint.Application")

정렬상태확인:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.VerticalAnchor
if (정렬상태확인 = 1)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginTop + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginTop := 여백조정
}
if (정렬상태확인 = 4)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginBottom - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginBottom := 여백조정
}

	}

}

gosub, 키보드올리기
return











~MButton & Up::


ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

 if (객체확인=19)
{

;테이블일수도있으니 이것도 실행

Try
{
    ppt := ComObjActive("PowerPoint.Application")
    ; 선택된 테이블에 대한 참조를 저장합니다.
    Table := ppt.ActiveWindow.Selection.ShapeRange.Table

    ; 테이블의 모든 행과 열에 대해 반복합니다.
    Rows := Table.Rows.Count
    Columns := Table.Columns.Count
    Loop, %Rows%
    {
        row := A_Index
        Loop, %Columns%
        {
            col := A_Index
            ; 각 셀을 참조합니다.
            Cell := Table.Cell(row, col)
            ; 선택된 셀인지 확인합니다. 
	    ; Cell.Selected 이값은 선택이면 1, 아니면 0으로 나옴

            If (Cell.Selected)
            {
                ; 선택된 셀의 여백을 0으로 설정합니다.

정렬상태확인:=Cell.Shape.TextFrame.VerticalAnchor
; 좌정렬=1 가운데2 우측3
if (정렬상태확인 = 1) 
{
여백조정:=Cell.Shape.TextFrame.MarginTop - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginTop := 여백조정
}
if (정렬상태확인 = 4)
{
여백조정:=Cell.Shape.TextFrame.MarginBottom + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
Cell.Shape.TextFrame.MarginBottom := 여백조정
}


; Cell.Shape.TextFrame.MarginLeft := 0
; Cell.Shape.TextFrame.MarginRight := 0
; Cell.Shape.TextFrame.MarginTop := 0
; Cell.Shape.TextFrame.MarginBottom := 0
            }
        }
    }
; 선택된 셀의 여백이 0으로 설정되었습니다.
}



}
else
{


Try
	{
ppt := ComObjActive("PowerPoint.Application")

정렬상태확인:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.VerticalAnchor
if (정렬상태확인 = 1)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginTop - 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginTop := 여백조정
}
if (정렬상태확인 = 4)
{
여백조정:=ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginBottom + 1.417325
if (여백조정<1.417325)
{
여백조정:=0
}
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.MarginBottom := 여백조정
}

	}

}
gosub, 키보드올리기

return






































				;도형안의 텍스트 배치
;$^F8::
Action_Uni0381:


Try
	{
ppt := ComObjActive("PowerPoint.Application")
정보확인 := ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.WordWrap
	}


    if (정보확인) {

Try
	{
ppt := ComObjActive("PowerPoint.Application")
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.WordWrap := False
	}
작업끝표시(50, "red")
return
    } else {
    }


Try
	{
ppt := ComObjActive("PowerPoint.Application")
ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.WordWrap := True
	}


작업끝표시(50, "red")

gosub, 키보드올리기

return











;관리자모드 시작
if (UserGrade >=4)
{


;한글 단어 잘림 허용이 비활성화 상태라면
$^F9::


send, {Alt down}{h}{p}{g}{Alt up}
sleep, 100

loop
{
if (WinActive("단락") || WinActive("Paragraph"))
{
sleep, 10
break
}
}



sleep, 1
send, {right}
sleep, 1
send, {tab 2}
sleep, 1
send, {space}
sleep, 1
send, {enter}
sleep, 10


작업끝표시(50, "red")

gosub, 키보드올리기

return

}






;폰트 교체==============================================================================







$^!9::
$^!Numpad9::

영문폰트변수:=영문폰트9
한글폰트변수:=한글폰트9


gosub, 폰트바꾸기
return

$^!8::
$^!Numpad8::

영문폰트변수:=영문폰트8
한글폰트변수:=한글폰트8


gosub, 폰트바꾸기
return


$^!7::
$^!Numpad7::

영문폰트변수:=영문폰트7
한글폰트변수:=한글폰트7


gosub, 폰트바꾸기
return

$^!6::
$^!Numpad6::

영문폰트변수:=영문폰트6
한글폰트변수:=한글폰트6


gosub, 폰트바꾸기
return


$^!5::
$^!Numpad5::

영문폰트변수:=영문폰트5
한글폰트변수:=한글폰트5


gosub, 폰트바꾸기
return



$^!4::
$^!Numpad4::

영문폰트변수:=영문폰트4
한글폰트변수:=한글폰트4


gosub, 폰트바꾸기
return



$^!3::
$^!Numpad3::

영문폰트변수:=영문폰트3
한글폰트변수:=한글폰트3


gosub, 폰트바꾸기
return



;$^!2:: 이 단축키는 모두표시 기능때문에 보류
$^!Numpad2::

영문폰트변수:=영문폰트2
한글폰트변수:=한글폰트2


gosub, 폰트바꾸기
return



$^!1::
$^!Numpad1::

영문폰트변수:=영문폰트1
한글폰트변수:=한글폰트1


gosub, 폰트바꾸기
return


;-------------------------------------------------------









$^#9::
$^#Numpad9::

영문폰트변수:=영문폰트99
한글폰트변수:=한글폰트99


gosub, 폰트바꾸기
return

$^#8::
$^#Numpad8::

영문폰트변수:=영문폰트88
한글폰트변수:=한글폰트88


gosub, 폰트바꾸기
return


$^#7::
$^#Numpad7::

영문폰트변수:=영문폰트77
한글폰트변수:=한글폰트77


gosub, 폰트바꾸기
return

$^#6::
$^#Numpad6::

영문폰트변수:=영문폰트66
한글폰트변수:=한글폰트66


gosub, 폰트바꾸기
return


$^#5::
$^#Numpad5::

영문폰트변수:=영문폰트55
한글폰트변수:=한글폰트55


gosub, 폰트바꾸기
return



$^#4::
$^#Numpad4::

영문폰트변수:=영문폰트44
한글폰트변수:=한글폰트44


gosub, 폰트바꾸기
return



$^#3::
$^#Numpad3::

영문폰트변수:=영문폰트33
한글폰트변수:=한글폰트33


gosub, 폰트바꾸기
return



$^#2::
$^#Numpad2::

영문폰트변수:=영문폰트22
한글폰트변수:=한글폰트22


gosub, 폰트바꾸기
return




$^#1::
$^#Numpad1::

영문폰트변수:=영문폰트11
한글폰트변수:=한글폰트11


gosub, 폰트바꾸기
return












; 6/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위











폰트바꾸기:
MouseGetPos, xx1, yy1 ; 행동하고 나서 다시 돌아오기 위해 현재 위치 체크

    ; 적용할 폰트명 변수 설정
    EngFont := 영문폰트변수
    KorFont := 한글폰트변수

    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        selType := sel.Type

        ; 1. 텍스트 블록 드래그 (단일 도형 및 표 내부 텍스트 드래그)
        ; ppSelectionText = 3
        if (selType = 3) {
            sel.TextRange.Font.Name := EngFont
            sel.TextRange.Font.NameFarEast := KorFont
        }
        ; 2. 상자 전체, 표 전체, 또는 표 내부 셀 다중 선택
        ; ppSelectionShape = 2
        else if (selType = 2) {
            ; 표의 여러 셀을 드래그해서 선택한 경우 (하위 형태 선택)
            if (sel.HasChildShapeRange) {
                Loop, % sel.ChildShapeRange.Count
                {
                    sh := sel.ChildShapeRange.Item(A_Index)
                    if (sh.HasTextFrame) {
                        if (sh.TextFrame.HasText) {
                            sh.TextFrame.TextRange.Font.Name := EngFont
                            sh.TextFrame.TextRange.Font.NameFarEast := KorFont
                        }
                    }
                }
            }
            ; 도형 전체 또는 표 전체가 선택된 경우
            else {
                Loop, % sel.ShapeRange.Count
                {
                    sh := sel.ShapeRange.Item(A_Index)
                    ChangeShapeFont_v1(sh, EngFont, KorFont)
                }
            }
        }
    } catch e {
        MsgBox, 262208, Message, %msgboxuni_0015%
    }

작업끝표시(50, "red")

ToolTIP, ★ %한글폰트변수%, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 1000
gosub, 키보드올리기

return

; 도형 내부 텍스트 및 표 전체/부분을 순회하는 v1 전용 함수
ChangeShapeFont_v1(sh, eng, kor) {
    ; 1) 그룹 도형인 경우 (msoGroup = 6)
    if (sh.Type = 6) {
        Loop, % sh.GroupItems.Count
        {
            ChangeShapeFont_v1(sh.GroupItems.Item(A_Index), eng, kor)
        }
    }
    ; 2) 텍스트 프레임이 있는 일반 도형인 경우
    if (sh.HasTextFrame) {
        if (sh.TextFrame.HasText) {
            sh.TextFrame.TextRange.Font.Name := eng
            sh.TextFrame.TextRange.Font.NameFarEast := kor
        }
    }
    ; 3) 표(Table)인 경우: 전체 선택 및 부분 선택(드래그) 감지
    if (sh.HasTable) {
        tbl := sh.Table
        
        ; 1단계: 셀이 부분적으로 선택되었는지 확인
        hasSelectedCells := false
        Loop, % tbl.Rows.Count
        {
            r := A_Index
            Loop, % tbl.Columns.Count
            {
                c := A_Index
                try {
                    if (tbl.Cell(r, c).Selected) {
                        hasSelectedCells := true
                        break
                    }
                }
            }
            if (hasSelectedCells)
                break
        }
        
        ; 2단계: 조건에 맞게 폰트 변경 루프 실행
        Loop, % tbl.Rows.Count
        {
            r := A_Index
            Loop, % tbl.Columns.Count
            {
                c := A_Index
                
                ; 부분 선택 모드일 경우, 현재 셀이 선택되지 않았으면 건너뜀
                if (hasSelectedCells) {
                    try {
                        if (!tbl.Cell(r, c).Selected) {
                            continue
                        }
                    } catch {
                        continue
                    }
                }
                
                ; 병합된 셀에서 발생할 수 있는 오류 방지 처리
                try {
                    cellSh := tbl.Cell(r, c).Shape
                    if (cellSh.HasTextFrame) {
                        if (cellSh.TextFrame.HasText) {
                            cellSh.TextFrame.TextRange.Font.Name := eng
                            cellSh.TextFrame.TextRange.Font.NameFarEast := kor
                        }
                    }
                }
            }
        }
    }
}
















; 자간조정===================================================================================
; 자간조정창 오픈하는 스크립트를 추가하는것도 고민해보기(느릴뿐...)



Action_Uni0181:
; $^!right::

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크
/*
WinActivate, %파워포인트타이틀%
*/




ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

 if (객체확인=19)
{

;테이블일수도있으니 이것도 실행
Try
{
    ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
    ; 선택된 테이블에 대한 참조를 저장합니다.
    Table := ppt.ActiveWindow.Selection.ShapeRange.Table

    ; 테이블의 모든 행과 열에 대해 반복합니다.
    Rows := Table.Rows.Count
    Columns := Table.Columns.Count
    Loop, %Rows%
    {
        row := A_Index
        Loop, %Columns%
        {
            col := A_Index
            ; 각 셀을 참조합니다.
            Cell := Table.Cell(row, col)
            ; 선택된 셀인지 확인합니다. 
	    ; Cell.Selected 이값은 선택이면 1, 아니면 0으로 나옴

            If (Cell.Selected)
            {
자간조정:=Cell.Shape.TextFrame2.TextRange.Font.Spacing + Uni0181_Key설정값
            Cell.Shape.TextFrame2.TextRange.Font.Spacing:=자간조정
            }
        }
    }
ToolTIP, ★ Spacing : %자간조정%, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
}



}
else
{
Try
	{
ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
자간조정:=ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.Spacing + Uni0181_Key설정값
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.Spacing:=자간조정

ToolTIP, ★ Spacing : %자간조정%, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
	}
}





return








Action_Uni0182:

; $^!left::
mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크

/*
WinActivate, %파워포인트타이틀%
*/






ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type

 if (객체확인=19)
{

;테이블일수도있으니 이것도 실행
Try
{
    ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
    ; 선택된 테이블에 대한 참조를 저장합니다.
    Table := ppt.ActiveWindow.Selection.ShapeRange.Table

    ; 테이블의 모든 행과 열에 대해 반복합니다.
    Rows := Table.Rows.Count
    Columns := Table.Columns.Count
    Loop, %Rows%
    {
        row := A_Index
        Loop, %Columns%
        {
            col := A_Index
            ; 각 셀을 참조합니다.
            Cell := Table.Cell(row, col)
            ; 선택된 셀인지 확인합니다. 
	    ; Cell.Selected 이값은 선택이면 1, 아니면 0으로 나옴

            If (Cell.Selected)
            {
자간조정:=Cell.Shape.TextFrame2.TextRange.Font.Spacing - Uni0182_Key설정값
            Cell.Shape.TextFrame2.TextRange.Font.Spacing:=자간조정
            }
        }
    }
ToolTIP, ★ Spacing : %자간조정%, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
}



}
else
{
Try
	{
ppt := ComObjActive("PowerPoint.Application")
SetFormat, float, 0.2
자간조정:=ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.Spacing - Uni0182_Key설정값
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.Spacing:=자간조정

ToolTIP, ★ Spacing : %자간조정%, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
	}
}




return
















; 도형 라인두께 조절하기 =======================================================================================


; Ctrl + Alt + Win + Up (선 생성 및 두께 0.25pt 증가)

;^#!up::
Action_Uni0231:

    MouseGetPos, xx1, yy1 ; 현재 마우스 위치 가져오기
try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    
    ; 도형(2) 또는 텍스트 범위(3)가 선택되었는지 확인
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        
        ; 선이 없는 상태(0: msoFalse)인 경우
        if (shape.Line.Visible = 0) {
            shape.Line.Visible := -1  ; 선을 보이게 처리(-1: msoTrue)
            shape.Line.Weight := 0.25 ; 기본 두께 0.25pt 할당
        } else {
            ; 이미 선이 있는 경우 0.25pt 증가
            shape.Line.Weight := shape.Line.Weight + 0.25
        }
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0025%
ToolTIP, ★ Weight : %newWeight% pt, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
return
}


newWeight := shape.Line.Weight
newWeight := Round(newWeight, 2)

ToolTIP, ★ Weight : %newWeight% pt, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
return





; Ctrl + Alt + Win + Down (두께 0.25pt 감소 및 선 없애기)
;^#!down::
Action_Uni0232:

    MouseGetPos, xx1, yy1 ; 현재 마우스 위치 가져오기
try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    
    ; 도형(2) 또는 텍스트 범위(3)가 선택되었는지 확인
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        
        ; 선이 있는 상태에서만 감소 로직 실행
        if (shape.Line.Visible != 0) {
            newWeight := shape.Line.Weight - 0.25
            
            ; 감소된 두께가 0 이하가 되면 선을 숨김(선 없음 처리)
            ; (부동소수점 연산 오차를 고려하여 0.01 이하로 기준점 설정)
            if (newWeight <= 0.01) {
                shape.Line.Visible := 0
            } else {
                shape.Line.Weight := newWeight
            }
        }
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0025%
ToolTIP, ★ Weight : %newWeight% pt, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
return
}

newWeight := shape.Line.Weight
newWeight := Round(newWeight, 2)

ToolTIP, ★ Weight : %newWeight% pt, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800
return








; 단축키: Ctrl + Alt + Win + Right (오른쪽 방향키)
;^#!right::
Action_Uni0241:

try {
    ; 실행 중인 파워포인트 객체 연결
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    
    ; 도형(2) 또는 텍스트 범위(3)가 선택되었는지 확인
    if (sel.Type = 2 || sel.Type = 3) { 
        shape := sel.ShapeRange
        currentStyle := shape.Line.DashStyle
        
        ; 파워포인트 UI 메뉴 순서에 맞춘 대시 스타일 고유번호 배열
        ; 1: 실선, 3: 둥근점선, 2: 사각점선, 4: 파선, 5: 점선-파선, 7: 긴파선, 8: 긴파선-점선, 9: 긴파선-점선-점선
        styles := [1, 11, 10, 4, 5, 7, 8, 9, 3, 2]
        
        ; 현재 스타일의 배열 위치(Index) 탐색
        currentIndex := 1
        for index, style in styles {
            if (style = currentStyle) {
                currentIndex := index
                break
            }
        }
        
        ; 다음 스타일 계산 (마지막 8번째면 다시 1번째로 순환)
        nextIndex := (currentIndex >= 10) ? 1 : currentIndex + 1
        
        ; 새로운 대시 스타일 즉각 적용
        shape.Line.DashStyle := styles[nextIndex]
    }
} catch {
    ; 파워포인트가 실행 중이 아니거나 도형이 선택되지 않은 경우 예외 처리
    MsgBox, 262208, Message, %msgboxuni_0026%
}
return







; 단축키: Ctrl + Alt + Win + Left (왼쪽 방향키)
;^#!left::
Action_Uni0242:

try {
    ; 실행 중인 파워포인트 객체 연결
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    
    ; 도형(2) 또는 텍스트 범위(3)가 선택되었는지 확인
    if (sel.Type = 2 || sel.Type = 3) { 
        shape := sel.ShapeRange
        currentStyle := shape.Line.DashStyle
        
        ; 파워포인트 UI 메뉴 순서에 맞춘 대시 스타일 고유번호 배열
        styles := [1, 11, 10, 4, 5, 7, 8, 9, 3, 2]
        
        ; 현재 스타일의 배열 위치(Index) 탐색
        currentIndex := 1
        for index, style in styles {
            if (style = currentStyle) {
                currentIndex := index
                break
            }
        }
        
        ; 이전 스타일 계산 (첫 번째 1번이면 마지막 8번으로 역순환)
        prevIndex := (currentIndex <= 1) ? 10 : currentIndex - 1
        
        ; 새로운 대시 스타일 즉각 적용
        shape.Line.DashStyle := styles[prevIndex]
    }
} catch {
    ; 파워포인트가 실행 중이 아니거나 도형이 선택되지 않은 경우 예외 처리
    MsgBox, 262208, Message, %msgboxuni_0026%
}
return














; ------------------------------------------------------------------
; 1. 화살표 머리 유형 순환 (Ctrl + Win + Right / Left)
; ------------------------------------------------------------------
; Ctrl + Win + Right (유형 다음 단계로)
;^#right::
Action_Uni0261:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        ; 1:없음, 2:삼각형, 3:V자형, 4:스텔스, 5:다이아몬드, 6:타원
        currentStyle := shape.Line.BeginArrowheadStyle
        nextStyle := (currentStyle >= 6) ? 1 : currentStyle + 1
        shape.Line.BeginArrowheadStyle := nextStyle
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return

; Ctrl + Win + Left (유형 이전 단계로 역순환)
;^#left::
Action_Uni0262:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        currentStyle := shape.Line.BeginArrowheadStyle
        prevStyle := (currentStyle <= 1) ? 6 : currentStyle - 1
        shape.Line.BeginArrowheadStyle := prevStyle
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return


; ------------------------------------------------------------------
; 2. 화살표 머리 크기 순환 (Ctrl + Win + Up / Down)
; ------------------------------------------------------------------
; Ctrl + Win + Up (크기 다음 단계로, 1~9 순환)
;^#up::
Action_Uni0263:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        
        ; 화살표가 '없음(1)' 상태일 때 크기를 변경하면 시각적으로 보이지 않으므로, 
        ; 기본 화살표(2)로 자동 변경하여 크기 변화를 즉시 확인할 수 있도록 처리
        if (shape.Line.BeginArrowheadStyle = 1)
            shape.Line.BeginArrowheadStyle := 2
            
        W := shape.Line.BeginArrowheadWidth
        L := shape.Line.BeginArrowheadLength
        
        ; 너비(W)와 길이(L) 조합을 1~9의 인덱스로 변환
        currentIndex := (L - 1) * 3 + W
        nextIndex := (currentIndex >= 9) ? 1 : currentIndex + 1
        
        ; 1~9 인덱스를 다시 너비와 길이 값(1~3)으로 분리하여 적용
        shape.Line.BeginArrowheadLength := Ceil(nextIndex / 3)
        shape.Line.BeginArrowheadWidth := Mod(nextIndex - 1, 3) + 1
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return

; Ctrl + Win + Down (크기 이전 단계로 역순환)
;^#down::
Action_Uni0264:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        
        if (shape.Line.BeginArrowheadStyle = 1)
            shape.Line.BeginArrowheadStyle := 2
            
        W := shape.Line.BeginArrowheadWidth
        L := shape.Line.BeginArrowheadLength
        
        currentIndex := (L - 1) * 3 + W
        prevIndex := (currentIndex <= 1) ? 9 : currentIndex - 1
        
        shape.Line.BeginArrowheadLength := Ceil(prevIndex / 3)
        shape.Line.BeginArrowheadWidth := Mod(prevIndex - 1, 3) + 1
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return








; ------------------------------------------------------------------
; 1. 화살표 꼬리 유형 순환 (Alt + Win + Right / Left)
; ------------------------------------------------------------------
; Alt + Win + Right (유형 다음 단계로)
;#!right::
Action_Uni0251:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        ; 1:없음, 2:삼각형, 3:V자형, 4:스텔스, 5:다이아몬드, 6:타원
        currentStyle := shape.Line.EndArrowheadStyle
        nextStyle := (currentStyle >= 6) ? 1 : currentStyle + 1
        shape.Line.EndArrowheadStyle := nextStyle
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return




; Alt + Win + Left (유형 이전 단계로 역순환)
;#!left::
Action_Uni0252:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        currentStyle := shape.Line.EndArrowheadStyle
        prevStyle := (currentStyle <= 1) ? 6 : currentStyle - 1
        shape.Line.EndArrowheadStyle := prevStyle
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return





; ------------------------------------------------------------------
; 2. 화살표 꼬리 크기 순환 (Alt + Win + Up / Down)
; ------------------------------------------------------------------
; Alt + Win + Up (크기 다음 단계로, 1~9 순환)
;#!up::
Action_Uni0253:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        
        ; 화살표 꼬리가 '없음(1)' 상태일 때 자동 기본 화살표(2) 생성
        if (shape.Line.EndArrowheadStyle = 1)
            shape.Line.EndArrowheadStyle := 2
            
        W := shape.Line.EndArrowheadWidth
        L := shape.Line.EndArrowheadLength
        
        ; 너비와 길이 조합을 1~9 인덱스로 변환
        currentIndex := (L - 1) * 3 + W
        nextIndex := (currentIndex >= 9) ? 1 : currentIndex + 1
        
        ; 크기 적용
        shape.Line.EndArrowheadLength := Ceil(nextIndex / 3)
        shape.Line.EndArrowheadWidth := Mod(nextIndex - 1, 3) + 1
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return



; Alt + Win + Down (크기 이전 단계로 역순환)
;#!down::
Action_Uni0254:

try {
    ppt := ComObjActive("PowerPoint.Application")
    sel := ppt.ActiveWindow.Selection
    if (sel.Type = 2 || sel.Type = 3) {
        shape := sel.ShapeRange
        
        if (shape.Line.EndArrowheadStyle = 1)
            shape.Line.EndArrowheadStyle := 2
            
        W := shape.Line.EndArrowheadWidth
        L := shape.Line.EndArrowheadLength
        
        currentIndex := (L - 1) * 3 + W
        prevIndex := (currentIndex <= 1) ? 9 : currentIndex - 1
        
        shape.Line.EndArrowheadLength := Ceil(prevIndex / 3)
        shape.Line.EndArrowheadWidth := Mod(prevIndex - 1, 3) + 1
    }
} catch {
    MsgBox, 262208, Message, %msgboxuni_0030%
}
return










; 7/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위













; 행간조정=======================================================================================
; [수정] 툴팁 기능이 추가되었습니다.
; [수정] 스크립트 로드 오류를 유발할 수 있는 불필요한 공백 문자를 제거했습니다.
; [수정] '배수'가 아닌 경우(예: 단일) '배수 1.0'을 기준으로 조정을 시작하도록 로직 변경
;=======================================================================================




Action_Uni0187:
; --- 기능 3: 줄 간격 0.03 '늘리기' ---
;$^!down::
    adjustment := Uni0187_Key설정값 ; 증가값
    GoSub, AdjustSpacing_SharedLogic
    
    ; [추가] 툴팁 표시 로직
    MouseGetPos, xx1, yy1 ; 현재 마우스 위치 가져오기
    ToolTip, ★ Line Spacing (Multiple) : %g_CurrentSpacing%, %xx1%, %yy1%
    SetTimer, 정보창툴팁없애기, -800 ; 800ms(0.8초) 후에 1회 실행
Return





Action_Uni0186:
; --- 기능 2: 줄 간격 0.03 '줄이기' ---
;$^!up::
    adjustment := -Uni0186_Key설정값 ; 감소값
    GoSub, AdjustSpacing_SharedLogic
    
    ; [추가] 툴팁 표시 로직
    MouseGetPos, xx1, yy1
    ToolTip, ★ Line Spacing (Multiple) : %g_CurrentSpacing%, %xx1%, %yy1%
    SetTimer, 정보창툴팁없애기, -800
Return






; --- '줄이기'/'늘리기' 공통 로직 ---
AdjustSpacing_SharedLogic:
    ; [수정] 툴팁에 사용할 전역 변수를 매 핫키 실행 시마다 초기화
    global g_CurrentSpacing := "N/A" 
    
    ; --- VBA 상수(Constants) 정의 ---
    msoTrue := -1           ; '배수' 줄 간격 규칙(LineRuleWithin) 활성화
    ppSelectionShapes := 2  ; 선택 유형이 '도형'임을 의미 [1]
    msoTable := 19          ; 객체 유형 '표' [2, 3, 4]
    msoGroup := 6           ; 객체 유형 '그룹' [2, 4]

    ; --- 1. PowerPoint COM 객체 연결 ---
    Try ppt := ComObjActive("PowerPoint.Application")
    Catch
    {
        MsgBox, 262208, Message, %msgboxuni_0015%
        Return
    }

    ; --- 2. 현재 선택 객체 가져오기 ---
    Try oSel := ppt.ActiveWindow.Selection
    Catch
    {
        MsgBox, 262208, Message, %msgboxuni_0032%
        Return
    }

    ; --- 3. 선택 유형 확인 ---
    If (oSel.Type!= ppSelectionShapes)
    {
        MsgBox, 262208, Message, %msgboxuni_0032%
        Return
    }

    ; --- 4. 선택된 각 도형 순회 ---
    oShapeRange := oSel.ShapeRange
    For oShape in oShapeRange
    {
        ; 미세 조정 재귀 함수 호출
        ProcessShape_Adjust(oShape, msoTable, msoGroup, msoTrue, adjustment)
    }
    
    oShapeRange := "", oSel := "", ppt := ""
Return ; GoSub의 끝

















;단락 뒤간격 조정

;$^!+Down::
Action_Uni0191:

ppt := ComObjActive("PowerPoint.Application") ; 파워포인트 활성화
선택갯수 := ppt.ActiveWindow.Selection.Shaperange.Count()

if (선택갯수 = 1) {
    객체확인 := ppt.ActiveWindow.Selection.Shaperange.Type
    if (객체확인 = 19) {
        Try {
            Table := ppt.ActiveWindow.Selection.Shaperange.Table
            Rows := Table.Rows.Count
            Columns := Table.Columns.Count
            Loop, %Rows% {
                row := A_Index
                Loop, %Columns% {
                    col := A_Index
                    Cell := Table.Cell(row, col)
                    if (Cell.Selected) {
                        Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceBefore := 0
                        단락뒤값 := Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceAfter + Uni0191_Key설정값
                        Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceAfter := 단락뒤값
                    }
                }
            }
            ToolTIP, ★Paragraph Spacing : %단락뒤값% pt
            SetTimer, 정보창툴팁없애기, 800
        }
    } else {
        Try {
            ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceBefore := 0
            단락뒤값 := ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceAfter + Uni0191_Key설정값
            ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceAfter := 단락뒤값
            ToolTIP, ★Paragraph Spacing : %단락뒤값% pt
            SetTimer, 정보창툴팁없애기, 800
        }
    }
} else {
    ToolTIP, ★ Just one Selected!
    SetTimer, 정보창툴팁없애기, 800
}


gosub, 키보드올리기
return




;$^!+Up::
Action_Uni0192:

mousegetpos, xx1, yy1 ;행동하고나서 다시돌아오기위해 현재위치 체크


ppt := ComObjActive("PowerPoint.Application") ; 파워포인트 활성화
선택갯수 := ppt.ActiveWindow.Selection.Shaperange.Count()

if (선택갯수 = 1) {
    객체확인 := ppt.ActiveWindow.Selection.Shaperange.Type
    if (객체확인 = 19) {
        Try {
            Table := ppt.ActiveWindow.Selection.Shaperange.Table
            Rows := Table.Rows.Count
            Columns := Table.Columns.Count
            Loop, %Rows% {
                row := A_Index
                Loop, %Columns% {
                    col := A_Index
                    Cell := Table.Cell(row, col)
                    if (Cell.Selected) {
                        Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceBefore := 0
                        단락뒤값 := Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceAfter - Uni0192_Key설정값
                        Cell.Shape.TextFrame2.TextRange.ParagraphFormat.SpaceAfter := 단락뒤값
                    }
                }
            }
            ToolTIP, ★Paragraph Spacing : %단락뒤값% pt
            SetTimer, 정보창툴팁없애기, 800
        }
    } else {
        Try {
            ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceBefore := 0
            단락뒤값 := ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceAfter - Uni0192_Key설정값
            ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.ParagraphFormat.SpaceAfter := 단락뒤값
            ToolTIP, ★Paragraph Spacing : %단락뒤값% pt
            SetTimer, 정보창툴팁없애기, 800
        }
    }
} else {
    ToolTIP, ★ Just one Selected!
    SetTimer, 정보창툴팁없애기, 800
}
return

















; 확대축소 ====================================================================================


$^0::

; 화면에 꽉차게 창조절
send, {Alt down}{w}{f}{Alt up}
sleep, 1
return



$^=::
Try
	{
zoom := ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom +20
if (zoom > 390)
zoom := 400
ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := zoom
	}
return



$^-::
Try
	{
zoom := ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom - 20
if (zoom < 20)
zoom := 15
ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := zoom
	}

return







$^+WheelUp::
Try
	{
zoom := ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom + 2
if (zoom > 390)
zoom := 400
ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := zoom
	}

return



$^+WheelDown::
Try
	{
zoom := ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom - 2
if (zoom < 20)
zoom := 15
ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := zoom
	}

return







; 8/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위










;vba컨트롤 ========================================================================
;클립보드에만 복사해서 사용하게 수정됨


Action_Uni0060:
; $!1::

WinClose, Microsoft Visual Basic for Applications



영상트리밍 = 0
캡쳐본생성 = 0
쪽번호전체삭제 = 0
쪽번호증가일괄변경 = 0
번역하기 = 0
대체텍스트변경=1
자동코드실행=1
goto, 자동재생대체텍스트만


return







고급메뉴창:
$^!+F11::

; Gui, 대칭:Destroy

WinClose, Microsoft Visual Basic for Applications


번역하기 = 0
대체텍스트변경 = 0
영상트리밍 = 0
캡쳐본생성 = 0
쪽번호전체삭제 = 0
쪽번호증가일괄변경 = 0
자동코드실행 = 0

Gui, VBA입력:Destroy


; 시스템 메시지(0x0111)에 VBA 전용 이벤트 함수 연결
; (대칭 창과 함께 쓴다면 최상단에 OnMessage(0x0111, "WM_COMMAND_대칭") 도 같이 적어두시면 됩니다)
OnMessage(0x0111, "WM_COMMAND_VBA")

; ==============================================================================
; 4. GUI 생성
; ==============================================================================
Gui, VBA입력:New, +AlwaysOnTop, Advanced ToolKit
;Gui, VBA입력:New, +AlwaysOnTop +ToolWindow, Advanced ToolKit


; [1. Preferences]
Gui, VBA입력:Font, s9 bold, 맑은 고딕
Gui, VBA입력:Add, Button, x15 y+10 w205 h30 gKey-Setting, %F11BtnF9%
Gui, VBA입력:Font, s9 norm, 맑은 고딕
Gui, VBA입력:Add, Button, x+10 yp w80 h30 g핫키설정변경, %F11BtnCustom%


Gui, VBA입력:Add, Text, x15 y+15 W300 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Gui, VBA입력:Font, s8, 맑은 고딕
Gui, VBA입력:Add, Text, x15 y+10 W300 c0x5F5F5F, %F11FontDesc1%
Gui, VBA입력:Add, Text, x15 y+6 W300 c0x5F5F5F, %F11FontDesc2%
Gui, VBA입력:Font, s9, 맑은 고딕
Gui, VBA입력:Add, Button, x15 y+15 w300 h25 g폰트위치찾기실행, %F11BtnFontFind%

Gui, VBA입력:Add, Text, x15 y+15 W300 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; [2. Export Images]
Gui, VBA입력:Font, s8, 맑은 고딕
Gui, VBA입력:Add, Text, x15 y+10 W300 c0x5F5F5F, %F11TxtExportDesc1%
Gui, VBA입력:Add, Text, x15 y+6 W300 c0x5F5F5F, %F11TxtExportDesc2%
Gui, VBA입력:Font, s9, 맑은 고딕
Gui, VBA입력:Add, Edit, x15 y+10 w50 h25 v캡쳐본해상도 Number Center, %캡쳐본해상도%
Gui, VBA입력:Add, Text, x+5 yp+4 w25 c0x5F5F5F, dpi
Gui, VBA입력:Add, Button, x+10 yp-4 w205 h25 v캡쳐본생성실행버튼 g캡쳐본생성실행, %F11BtnExport%
Gui, VBA입력:Add, Text, x15 y+15 W300 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; [3. Trim Videos]
Gui, VBA입력:Font, s8, 맑은 고딕
Gui, VBA입력:Add, Text, x15 y+10 W300 c0x5F5F5F, %F11TxtTrimDesc1%
Gui, VBA입력:Add, Text, x15 y+6 W300 c0x5F5F5F, %F11TxtTrimDesc2%
Gui, VBA입력:Font, s9, 맑은 고딕
Gui, VBA입력:Add, Edit, x15 y+10 w50 h25 v영상트리밍시간 Number Center, %영상트리밍시간%
Gui, VBA입력:Add, Text, x+5 yp+4 w30 c0x5F5F5F, sec
Gui, VBA입력:Add, Button, x+5 yp-4 w205 h25 v영상트리밍실행버튼 g영상트리밍실행, %F11BtnTrim%
Gui, VBA입력:Add, Text, x15 y+15 W300 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; [4. Slide Numbers]
Gui, VBA입력:Font, s8, 맑은 고딕
Gui, VBA입력:Add, Text, x15 y+10 W300 cblue, %F11TxtNumDesc1%
Gui, VBA입력:Add, Text, x15 y+6 W300 c0x5F5F5F, %F11TxtNumDesc2%
Gui, VBA입력:Add, Text, x15 y+6 W300 cblack, %F11TxtNumDesc3%
Gui, VBA입력:Add, Text, x15 y+6 W300 c0x5F5F5F, %F11TxtNumDesc4%
Gui, VBA입력:Font, s9, 맑은 고딕
Gui, VBA입력:Add, Checkbox, x15 y+15 w130 h20 v쪽번호증가일괄변경 gExclusiveCheck, %F11ChkAddNum%
Gui, VBA입력:Add, Checkbox, x+5 yp w140 h20 v쪽번호전체삭제 gExclusiveCheck, %F11ChkRemoveNum%
Gui, VBA입력:Add, Button, x15 y+15 w300 h25 v쪽번호변경버튼 gBtnOkVBA, %F11BtnApplyNum%
Gui, VBA입력:Add, Text, x15 y+15 W300 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Gui, VBA입력:Show, AutoSize, Advanced ToolKit
return


; ======================================================================
; 이벤트 감지 함수 (VBA입력 창 전용, 함수명 분리)
; ======================================================================
WM_COMMAND_VBA(wParam, lParam) {
    ; "Advanced ToolKit" 창이 활성화 상태일 때만 아래 로직 수행 (대칭 창과 완벽한 간섭 차단)
    IfWinNotActive, Advanced ToolKit
        return

    NotifyCode := wParam >> 16
    
    ; 0x0100 = Edit 컨트롤 포커스, 3 = DropDownList 포커스, 0 = Button/Checkbox 클릭 등 감지
    if (NotifyCode = 0x0100 || NotifyCode = 3 || NotifyCode = 0) {
        
        ; VBA입력 창 내부에서 현재 커서가 위치한 컨트롤의 변수명을 가져옴
        GuiControlGet, 현재포커스, VBA입력:FocusV
        
        if (현재포커스 = "캡쳐본해상도") {
            GuiControl, VBA입력:+Default, 캡쳐본생성실행버튼
        } else if (현재포커스 = "영상트리밍시간") {
            GuiControl, VBA입력:+Default, 영상트리밍실행버튼
        } else if (현재포커스 = "쪽번호증가일괄변경" || 현재포커스 = "쪽번호전체삭제") {
            GuiControl, VBA입력:+Default, 쪽번호변경버튼
        }
    }
}

; ------------------------------------------------------------------
; 단일선택(라디오 버튼처럼 동작) 서브루틴
; ------------------------------------------------------------------
ExclusiveCheck:

    ; 클릭된 체크박스(A_GuiControl)만 제외하고 모두 해제
    GuiControl,, 쪽번호증가일괄변경, 0
    GuiControl,, 쪽번호전체삭제, 0

    ; 현재 클릭된 체크박스만 다시 체크
    GuiControl,, %A_GuiControl%, 1
return















캡쳐본생성실행:






ExclusiveCheck캡쳐:

    try {
        ; 실행 중인 파워포인트 객체 가져오기
        PptApp := ComObjActive("PowerPoint.Application")
        
        ; 현재 활성화된 프레젠테이션의 경로 확인
        FilePath := PptApp.ActivePresentation.Path
        FileName := PptApp.ActivePresentation.Name

        if (FilePath = "") {

    Gui, VBA입력:Destroy

            MsgBox, 262208, Message, %msgboxuni_0052%

return

        } else {

;지나가기
        }
    } catch {
        MsgBox, 262208, Message, %msgboxuni_0015%
return
    }






캡쳐본생성 = 1
    Gui, VBA입력:Submit, NoHide
    IniWrite, %캡쳐본해상도%, %IniFile%, 기본설정, 캡쳐본해상도
    IniRead, 캡쳐본해상도, %IniFile%, 기본설정, 캡쳐본해상도


goto, BtnOkVBA


영상트리밍실행:
영상트리밍 = 1

    Gui, VBA입력:Submit, NoHide
    IniWrite, %영상트리밍시간%, %IniFile%, 기본설정, 영상트리밍시간
    IniRead, 영상트리밍시간, %IniFile%, 기본설정, 영상트리밍시간
goto, BtnOkVBA



BtnOkVBA:
Gui, VBA입력:Submit,NoHide
Gui, VBA입력:Destroy

자동코드실행 = 1






자동재생대체텍스트만:

Clipboard :=""
Clipboard :=""
sleep, 10










;파워포인트 실행중일때만 보이기
        if WinActive("ahk_exe POWERPNT.EXE")
        {


;관리자모드 시작
if (UserGrade >=4)
{


;--- [2025-10-22 지침 반영: v1 문법] ---

; ★★★★★★★★★★★★★★★★★★★★★ VBA 개체 액세스 확인 (안전한 COM 예외 처리 방식)

try {
    ; 1. 파워포인트 연결 먼저 시도
    ppt := ComObjActive("PowerPoint.Application")
} catch {
    ; 파워포인트 자체가 안 켜져 있으면 안내하고 종료
    MsgBox, 262208, Message, %msgboxuni_0015%
    return
}

AccessVBOM := true

try {
    ; 2. VBA 프로젝트 객체(VBE)에 직접 접근 시도 (핵심 포인트 ★)
    ; 보안 설정의 'VBA 개체 모델 액세스'가 꺼져있다면, 여기서 강제로 에러가 발생하여 catch로 넘어갑니다.
    dummy := ppt.VBE
} catch {
    ; 에러가 났다는 것은 권한이 없다는 뜻 (체크가 안 되어 있음)
    AccessVBOM := false
}

; 3. 권한이 없는 것으로 확인된 경우
if (AccessVBOM = false)
{
    ; 사용자에게 알림 메시지 표시
    MsgBox, 262208, Message, %msgboxuni_0053%
    
    try {
        ; 파워포인트 내부 명령어(Mso)를 이용해 '보안 센터-매크로 설정' 창 바로 열기
        ppt.CommandBars.ExecuteMso("MacroSecurity")
    }
    return
}
else
{
    ; 이미 체크되어 있는 경우 (정상 통과)
    ; 파워포인트 보안 설정이 정상입니다.
}










}

;관리자모드 끝



;파워포인트 실행중일때만 보이기
}


















;-----------------------------------------------------------------------------------------------------------------------
if (대체텍스트변경 = 1)
{

;(다음 바로Sub가 나와야 바로실행됨)
VBA내용_대체텍스트 =
(
Sub SetVideosToAutoPlayAndAddMarginShapes()
    ' ==========================================
    ' [1] 기존 변수 선언 (미디어/애니메이션용)
    ' ==========================================
    Dim sld As Slide
    Dim shp As Shape
    Dim eff As Effect
    Dim vbProj As Object
    Dim i As Long, j As Long
    Dim seq As Sequence
    Dim tempStack As Collection
    Dim innerShp As Shape
    Dim altText As String
    
    ' [핵심 변수] 기존에 '재생' 애니메이션이 있는지 확인하는 플래그
    Dim bFoundPlayEffect As Boolean
    
    ' 혹시나 모를 볼륨 저장을 위한 변수 (보험용)
    Dim currentVol As Single
    Dim currentMute As Boolean

    ' ==========================================
    ' [2] 추가 변수 선언 (슬라이드 마스터 여백 도형용)
    ' ==========================================
    Dim pDesign As Design
    Dim pMaster As Master
    Dim mShp As Shape
    Dim bMarginExists As Boolean
    Dim sWidth As Single, sHeight As Single
    Dim marginDistX As Single, marginDistY As Single
    Dim boxSize As Single


    ' ==================================================================================
    ' PART 1. 미디어 자동재생 및 볼륨 유지 로직 (기존 스크립트)
    ' ==================================================================================
    
    altText = "%대체텍스트내용1%"

    ' 0) 슬라이드 요소 이름 고유화 (기존 유지)
    For Each sld In ActivePresentation.Slides
        For i = 1 To sld.Shapes.Count
            If InStr(1, sld.Shapes(i).Name, "page_no_", vbTextCompare) = 0 Then
                sld.Shapes(i).Name = "Slide" & sld.SlideIndex & "_Shape" & i
            End If
        Next i
    Next sld




    ' 1) 모든 슬라이드 순회
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ' 미디어 객체 확인
            If shp.Type = msoMedia Then
                If shp.MediaType = ppMediaTypeMovie Or shp.MediaType = ppMediaTypeSound Then
                    
                    ' 초기화
                    bFoundPlayEffect = False
                    
                    ' ★★★ [핵심 로직 변경] ★★★
                    ' 무조건 삭제하고 다시 만드는 것이 아니라,
                    ' 타임라인을 뒤져서 "이미 이 파일에 걸려있는 재생(Play) 효과"가 있는지 찾습니다.
                    ' 찾으면 -> 그 효과의 속성(이전 효과와 함께 시작)만 바꿉니다. (볼륨 유지됨)
                    ' 못 찾으면 -> 그때만 새로 만듭니다.
                    
                    For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                        Set eff = sld.TimeLine.MainSequence(i)
                        
                        ' 현재 순회 중인 애니메이션이 '이 쉐이프(shp)'의 것인지 확인
                        If eff.Shape.Name = shp.Name Then
                            ' 그것이 '재생(MediaPlay)' 효과인가?
                            If eff.EffectType = msoAnimEffectMediaPlay Then
                                ' 찾았다! 기존 효과를 재활용합니다.
                                eff.Timing.TriggerType = msoAnimTriggerWithPrevious
                                eff.Timing.TriggerDelayTime = 0.0
                                bFoundPlayEffect = True
                            Else
                                ' 재생 효과가 아닌 다른 잡동사니 효과(나타나기 등)는 방해되므로 삭제
                                ' (원하시면 이 부분 주석 처리하여 다른 애니메이션 유지 가능)
                                ' ★ 기존애니메이션 유지
                                ' eff.Delete
                            End If
                        End If
                    Next i
                    




                    ' ★★★ [신규 생성] ★★★
                    ' 기존에 재생 효과가 아예 없었던 경우에만 새로 만듭니다.
                    ' 이때만 볼륨이 초기화될 위험이 있으므로, 볼륨 백업/복구 로직을 사용합니다.
                    If bFoundPlayEffect = False Then
                        ' 볼륨 백업
                        currentVol = shp.MediaFormat.Volume
                        currentMute = shp.MediaFormat.Muted
                        
                        ' 효과 추가
                        Set eff = sld.TimeLine.MainSequence.AddEffect( _
                                    Shape:=shp, _
                                    effectId:=ppEffectMediaPlay, _
                                    trigger:=msoAnimTriggerWithPrevious)
                        eff.Timing.TriggerDelayTime = 0.0
                        
                        ' 볼륨 복구
                        shp.MediaFormat.Volume = currentVol
                        shp.MediaFormat.Muted = currentMute
                    End If
                    
                    ' [재생 옵션 설정] - 반복 재생, 쇼 실행 중 숨기기 등
                    ' 이 설정들은 볼륨에 영향을 주지 않습니다.
                    With shp.AnimationSettings.PlaySettings
                       .LoopUntilStopped = True    ' 반복 재생
                       .PlayOnEntry = True         ' 실행 시 자동 재생
                       .RewindMovie = True         ' 재생 후 되감기
                       .HideWhileNotPlaying = False ' 재생 안 할 때 숨기기 (필요시 True)
                    End With
                    
                End If
            End If
        Next shp
    Next sld








    ' 2) ★★★★★★★★★★★★★★★★★자동재생, 이전효과 함께 위에 있지만 한번더 돌면서 확시히 해주는 스크립트
    For Each sld In ActivePresentation.Slides
        For Each eff In sld.TimeLine.MainSequence
            If eff.Shape.Type = msoMedia Then
                If eff.Shape.MediaType = ppMediaTypeMovie Then
                    If eff.EffectType = msoAnimEffectMediaPlay Then
                        eff.Timing.TriggerType = msoAnimTriggerWithPrevious
                        eff.Timing.TriggerDelayTime = 0.0
                    End If
                End If
            End If
        Next eff
    Next sld



    ' 2) ★★★★★★★★★★★★★★★★★자동재생, 이전효과 함께 위에 있지만 한번더 돌면서 확시히 해주는 스크립트
    For Each sld In ActivePresentation.Slides
        For Each eff In sld.TimeLine.MainSequence
            If eff.Shape.Type = msoMedia Then
                If eff.Shape.MediaType = ppMediaTypeMovie Then
                    If eff.EffectType = msoAnimEffectMediaPlay Then
                        eff.Timing.TriggerType = msoAnimTriggerWithPrevious
                        eff.Timing.TriggerDelayTime = 0.0
                    End If
                End If
            End If
        Next eff
    Next sld









    ' 4) '나타내기'(Appear) 등 불필요한 트리거 애니메이션 제거 (안전장치)
    ' (위의 루프에서 이미 정리했지만, 상호작용 시퀀스 등에 남아있을 수 있으므로 유지)
    For Each sld In ActivePresentation.Slides
        For Each seq In sld.TimeLine.InteractiveSequences
            For j = seq.Count To 1 Step -1
                If seq(j).Shape.Type = msoMedia Then
                    If seq(j).EffectType = msoAnimEffectAppear Then
                        ' ★ 기존애니메이션 유지
                        ' seq(j).Delete
                    End If
                End If
            Next j
        Next seq
    Next sld





    ' 5) 대체 텍스트 제거 (기존 유지)
    For Each sld In ActivePresentation.Slides
        Set tempStack = New Collection
        For Each shp In sld.Shapes
            tempStack.Add shp
        Next shp
        Do While tempStack.Count > 0
            Set shp = tempStack(1)
            tempStack.Remove 1
            If shp.Type = msoGroup Then
                For i = 1 To shp.GroupItems.Count
                    tempStack.Add shp.GroupItems(i)
                Next i
            ElseIf shp.Type = msoPicture Or shp.Type = msoMedia Then
                shp.Title = "%대체텍스트내용1%"
                shp.AlternativeText = "%대체텍스트내용2%"
            End If
        Loop
    Next sld


    ' ==================================================================================
    ' PART 2. 슬라이드 마스터 여백 도형 추가 (추가된 스크립트)
    ' ==================================================================================
    
    sWidth = ActivePresentation.PageSetup.SlideWidth
    sHeight = ActivePresentation.PageSetup.SlideHeight
    
    ' 설정: 도형 크기
    boxSize = 50       ' 도형 크기 (50x50)
    
    ' [수정됨] 가로는 너비만큼, 세로는 높이만큼 띄우기
    marginDistX = sWidth   
    marginDistY = sHeight

    ' 모든 디자인(마스터) 순회
    For Each pDesign In ActivePresentation.Designs
        Set pMaster = pDesign.SlideMaster
        bMarginExists = False
        
        ' 1. 이미 '여백자리잡기' 도형이 있는지 확인
        For Each mShp In pMaster.Shapes
            If mShp.Name = "여백자리잡기" Then
                bMarginExists = True
                Exit For
            End If
        Next mShp
        
        ' 2. 없다면 4개 추가 (채우기 없음, 검정 테두리)
        If Not bMarginExists Then
            ' (1) 좌상단 (Top-Left)
            ' X: 왼쪽으로 너비만큼, Y: 위쪽으로 높이만큼 이동
            Set mShp = pMaster.Shapes.AddShape(msoShapeRectangle, -marginDistX - boxSize, -marginDistY - boxSize, boxSize, boxSize)
            GoSub StyleMarginShape 
            
            ' (2) 우상단 (Top-Right)
            ' X: 오른쪽 끝 + 너비만큼, Y: 위쪽으로 높이만큼 이동
            Set mShp = pMaster.Shapes.AddShape(msoShapeRectangle, sWidth + marginDistX, -marginDistY - boxSize, boxSize, boxSize)
            GoSub StyleMarginShape 
            
            ' (3) 좌하단 (Bottom-Left)
            ' X: 왼쪽으로 너비만큼, Y: 아래쪽 끝 + 높이만큼 이동
            Set mShp = pMaster.Shapes.AddShape(msoShapeRectangle, -marginDistX - boxSize, sHeight + marginDistY, boxSize, boxSize)
            GoSub StyleMarginShape 
            
            ' (4) 우하단 (Bottom-Right)
            ' X: 오른쪽 끝 + 너비만큼, Y: 아래쪽 끝 + 높이만큼 이동
            Set mShp = pMaster.Shapes.AddShape(msoShapeRectangle, sWidth + marginDistX, sHeight + marginDistY, boxSize, boxSize)
            GoSub StyleMarginShape 
        End If
    Next pDesign
    
    ' (스타일 적용 서브루틴 - 실행 흐름 건너뛰기)
    GoTo SkipSubroutine

StyleMarginShape:
    With mShp
        .Name = "여백자리잡기"
        .Fill.Visible = msoFalse       ' 채우기 없음
        .Line.Visible = msoTrue        ' 테두리 있음
        .Line.ForeColor.RGB = 0        ' 검정색 (vbBlack)
        .Line.Weight = 1               ' 두께 1pt
    End With
    Return

SkipSubroutine:
    ' ==================================================================================
    ' PART 3. 완료 메시지 및 모듈 삭제 (마무리)
    ' ==================================================================================

    MsgBox "%msgboxuni_0068%", vbInformation

    ' 6) 모듈 삭제
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    vbProj.VBComponents.Remove vbProj.VBComponents("Module1")
    On Error GoTo 0

End Sub





)

Clipboard = %Clipboard%%VBA내용_대체텍스트%
ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 20
}
if (대체텍스트변경 = 0)
{
;그냥넘기기
}










;-----------------------------------------------------------------------------------------------------------------------
if (영상트리밍 = 1)
{

VBA내용_영상트리밍 =
(
Sub TrimAllVideosTo10Seconds()
    Dim pptSlide As Slide
    Dim pptShape As Shape
    Dim videoLength As Single
    Dim vbProj As Object
    
    ' 프레젠테이션의 각 슬라이드를 순회
    For Each pptSlide In ActivePresentation.Slides
        ' 슬라이드 내의 각 도형을 순회
        For Each pptShape In pptSlide.Shapes
            ' 도형이 미디어 개체이고 비디오인지 확인
            If pptShape.Type = msoMedia Then
                If pptShape.MediaType = ppMediaTypeMovie Then
                    ' 비디오의 길이를 밀리초 단위로 확인
                    videoLength = pptShape.MediaFormat.Length
                    
                    ' 시작점을 0으로 설정하고 끝점을 %영상트리밍시간%000밀리초(%영상트리밍시간%초)로 설정
                    ' 비디오가 최소 %영상트리밍시간%초 이상인지 확인
                    If videoLength > %영상트리밍시간%000 Then
                        pptShape.MediaFormat.StartPoint = 0
                        pptShape.MediaFormat.EndPoint = %영상트리밍시간%000
                    Else
                        ' 비디오 길이가 %영상트리밍시간%초보다 짧으면 비디오 길이에 맞게 자르기
                        pptShape.MediaFormat.StartPoint = 0
                        pptShape.MediaFormat.EndPoint = videoLength
                    End If
                End If
            End If
        Next pptShape
    Next pptSlide
    
    MsgBox "%msgboxuni_0069% : %영상트리밍시간%", vbInformation


' 현재 모듈 삭제
On Error Resume Next
Set vbProj = Application.VBE.ActiveVBProject
vbProj.VBComponents.Remove vbProj.VBComponents("Module1") ' Module1 이름 확인
On Error GoTo 0



End Sub

)




Clipboard = %Clipboard%%VBA내용_영상트리밍%
ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 20
}
if (영상트리밍 = 0)
{
;그냥넘기기
}








; 9/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위






;-----------------------------------------------------------------------------------------------------------------------
if (캡쳐본생성 = 1)
{

VBA내용_캡쳐본생성 =
(
Sub ConvertSlidesToImagesOnlyImproved()
    Dim sld As Slide
    Dim j As Integer, i As Integer
    Dim exportFolder As String, folderName As String, folderNumber As Integer
    Dim currentFileName As String, newFileName As String
    Dim exportWidth As Long, exportHeight As Long
    Dim slideWidth As Single, slideHeight As Single
    Dim imgShape As Shape

    On Error Resume Next

    ' 해상도 및 파일 정보 설정
    currentFileName = Left(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1)
    slideWidth = ActivePresentation.PageSetup.SlideWidth
    slideHeight = ActivePresentation.PageSetup.SlideHeight
    
    ' DPI 기반 크기 계산 (화질 핵심)
    exportWidth = slideWidth * (%캡쳐본해상도% / 72)
    exportHeight = slideHeight * (%캡쳐본해상도% / 72)

    ' 폴더 생성
    exportFolder = ActivePresentation.Path & "\★PNG%캡쳐본해상도%-" & currentFileName
    folderName = exportFolder
    folderNumber = 0
    Do While Dir(folderName, vbDirectory) <> ""
        folderNumber = folderNumber + 1
        folderName = exportFolder & "(" & folderNumber & ")"
    Loop
    MkDir folderName

    ' 슬라이드 루프
    For Each sld In ActivePresentation.Slides
        ' 이름 고유화
        For i = 1 To sld.Shapes.Count
            If InStr(1, sld.Shapes(i).Name, "page_no_", vbTextCompare) = 0 Then
                sld.Shapes(i).Name = "Slide" & sld.SlideIndex & "_Shape" & i
            End If
        Next i

        ' 고해상도 추출
        Dim imgPath As String
        imgPath = folderName & "\%캡쳐본해상도%-Slide" & sld.SlideIndex & ".png"
        sld.Export imgPath, "PNG", exportWidth, exportHeight

        ' 기존 객체 삭제 (동영상 제외)
        For j = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(j).Type <> 16 Then sld.Shapes(j).Delete ' 16 = msoMedia
        Next j

        ' 이미지 삽입 (슬라이드 크기에 딱 맞게)
        Set imgShape = sld.Shapes.AddPicture(imgPath, 0, -1, 0, 0, slideWidth, slideHeight)
        imgShape.ZOrder 1 ' 뒤로 보내기
    Next sld

    ' 새 파일 저장
    newFileName = ActivePresentation.Path & "\" & currentFileName & "★-%캡쳐본해상도%.pptx"
    ActivePresentation.SaveAs newFileName
    
    MsgBox "%msgboxuni_0070%: %캡쳐본해상도%dpi", vbInformation


    ' [4] 현재 모듈 삭제 (Module1 이름 확인)
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    vbProj.VBComponents.Remove vbProj.VBComponents("Module1")
    On Error GoTo 0
    Exit Sub



End Sub


)






Clipboard = %Clipboard%%VBA내용_캡쳐본생성%
ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 20
}
if (캡쳐본생성 = 0)
{
;그냥넘기기
}
















;-----------------------------------------------------------------------------------------------------------------------
if (쪽번호전체삭제 = 1)
{

VBA내용_쪽번호전체삭제 =
(
Sub DeletePageNumberShapes()
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Long
    
    ' 모든 슬라이드를 순회
    For Each sld In ActivePresentation.Slides
        ' 도형 삭제 시 인덱스 오류를 방지하기 위해 역순으로 진행
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If InStr(1, shp.Name, "DualPageNum_", vbTextCompare) > 0 Then
                shp.Delete
            End If
        Next i
    Next sld
    
    MsgBox "%msgboxuni_0071%", vbInformation


    ' [4] 현재 모듈 삭제 (Module1 이름 확인)
    On Error Resume Next
    Set vbProj = Application.VBE.ActiveVBProject
    vbProj.VBComponents.Remove vbProj.VBComponents("Module1")
    On Error GoTo 0
    Exit Sub



End Sub


)


Clipboard = %Clipboard%%VBA내용_쪽번호전체삭제%
ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 20
}
if (쪽번호전체삭제 = 0)
{
;그냥넘기기
}










;-----------------------------------------------------------------------------------------------------------------------
if (쪽번호증가일괄변경 = 1)
{

VBA내용_쪽번호증가일괄변경 =
(
Sub InsertPageNumbers_Universal()
    ' [!!! 실행 방법 !!!]
    ' 1. (1쪽 모드) 첫 슬라이드에 쪽번호 샘플 도형 1개를 선택하고 실행하세요.
    ' 2. (2쪽 모드) 첫 슬라이드에 좌/우 쪽번호 샘플 도형 2개를 모두 선택하고 실행하세요.
    '    -> 2개 선택 시 자동으로 좌/우를 구분하여 작동합니다.
    ' 3. 이 함수 내부를 클릭하고 F5를 눌러 실행하세요.

    On Error GoTo ERR_HANDLER

    Dim sld As Slide
    Dim shp As Shape
    Dim source1 As Shape, source2 As Shape ' source1: 단일/좌측, source2: 우측
    Dim isDualMode As Boolean

    Dim currentNum As Long
    Dim startNum As Long

    ' 정규식 관련 변수
    Dim regex As Object
    Dim match As Object
    Dim txt1 As String, txt2 As String
    Dim pre1 As String, suf1 As String
    Dim pre2 As String, suf2 As String

    ' 화면 중앙 기준 (2쪽 모드 판별용)
    Dim slideCenter As Single
    slideCenter = ActivePresentation.PageSetup.SlideWidth / 2

    ' 템플릿 슬라이드 인덱스
    Dim templateSlideIdx As Long
    templateSlideIdx = ActiveWindow.View.Slide.SlideIndex

    ' ========================================================
    ' [1] 선택 개수 확인 및 모드 설정
    ' ========================================================
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "%msgboxuni_0072%", vbExclamation, "ScriptRunError"
        GoTo DELETE_MODULE
    End If

    Dim selCount As Long
    selCount = ActiveWindow.Selection.ShapeRange.Count

    If selCount = 1 Then
        ' [1쪽 모드]
        isDualMode = False
        Set source1 = ActiveWindow.Selection.ShapeRange(1)
    ElseIf selCount = 2 Then
        ' [2쪽 모드]
        isDualMode = True
        Dim s1 As Shape, s2 As Shape
        Set s1 = ActiveWindow.Selection.ShapeRange(1)
        Set s2 = ActiveWindow.Selection.ShapeRange(2)
        ' 좌표 비교로 좌/우 구분
        If s1.Left < s2.Left Then
            Set source1 = s1 ' 좌측
            Set source2 = s2 ' 우측
        Else
            Set source1 = s2
            Set source2 = s1
        End If
    Else
        MsgBox "%msgboxuni_0073%", vbExclamation, "ScriptRunError"
        GoTo DELETE_MODULE
    End If

    ' ========================================================
    ' [2] 정규식 설정 및 시작 번호 추출
    ' ========================================================
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^(.*?)(\d+)(.*?)$"
    regex.IgnoreCase = True
    regex.Global = False

    ' 첫 번째 도형(source1) 분석 (공통)
    txt1 = source1.TextFrame.TextRange.Text
    If regex.Test(txt1) Then
        Set match = regex.Execute(txt1)(0)
        pre1 = match.SubMatches(0)
        startNum = CLng(match.SubMatches(1)) ' 여기에 적힌 숫자를 기준으로 함
        suf1 = match.SubMatches(2)
    Else
        MsgBox "%msgboxuni_0074%", vbExclamation, "ScriptRunError"
        GoTo DELETE_MODULE
    End If

    ' 2쪽 모드일 경우 두 번째 도형(source2) 분석
    If isDualMode Then
        txt2 = source2.TextFrame.TextRange.Text
        If regex.Test(txt2) Then
            Set match = regex.Execute(txt2)(0)
            pre2 = match.SubMatches(0)
            suf2 = match.SubMatches(2)
        Else
            MsgBox "%msgboxuni_0075%", vbExclamation, "ScriptRunError"
            GoTo DELETE_MODULE
        End If
    End If

    ' 시작 번호 초기화 (템플릿 숫자 + 1)
    currentNum = startNum + 1

    ' ========================================================
    ' [3] 기존 쪽번호 삭제
    ' ========================================================
    Dim i As Long
    For Each sld In ActivePresentation.Slides
        If sld.SlideIndex <> templateSlideIdx Then
            For i = sld.Shapes.Count To 1 Step -1
                If sld.Shapes(i).Name Like "DualPageNum_*" Then
                    sld.Shapes(i).Delete
                End If
            Next i
        End If
    Next sld

    ' ========================================================
    ' [4] 쪽번호 생성 루프
    ' ========================================================
    Dim blockerName As String
    blockerName = "NoNum"
    Dim isBlocked As Boolean
    Dim isBlockedL As Boolean, isBlockedR As Boolean

    For i = 1 To ActivePresentation.Slides.Count
        Set sld = ActivePresentation.Slides(i)

        ' (조건) 숨김 슬라이드 & 템플릿 슬라이드 건너뛰기
        If sld.SlideShowTransition.Hidden = msoFalse And sld.SlideIndex <> templateSlideIdx Then

            ' === [1쪽 모드 로직] ===
            If Not isDualMode Then
                isBlocked = False
                ' 슬라이드 내에 가림막이 하나라도 있으면 건너뜀
                For Each shp In sld.Shapes
                    If InStr(1, shp.Name, blockerName, vbTextCompare) > 0 Then
                        isBlocked = True
                        Exit For
                    End If
                Next shp

                If Not isBlocked Then
                    Call CreatePageNumShape(sld, source1, pre1 & currentNum & suf1, "DualPageNum_S_" & i)
                    currentNum = currentNum + 1
                End If

            ' === [2쪽 모드 로직] ===
            Else
                isBlockedL = False
                isBlockedR = False

                ' 가림막 위치(좌/우) 확인
                For Each shp In sld.Shapes
                    If InStr(1, shp.Name, blockerName, vbTextCompare) > 0 Then
                        If (shp.Left + shp.Width / 2) < slideCenter Then
                            isBlockedL = True
                        Else
                            isBlockedR = True
                        End If
                    End If
                Next shp

                ' 좌측 생성
                If Not isBlockedL Then
                    Call CreatePageNumShape(sld, source1, pre1 & currentNum & suf1, "DualPageNum_L_" & i)
                    currentNum = currentNum + 1
                End If

                ' 우측 생성
                If Not isBlockedR Then
                    Call CreatePageNumShape(sld, source2, pre2 & currentNum & suf2, "DualPageNum_R_" & i)
                    currentNum = currentNum + 1
                End If

            End If ' End Mode Check
        End If
    Next i

    ' ========================================================
    ' [5] 작업 완료 후 'NoNum' 도형 일괄 삭제
    ' ========================================================
    Dim k As Long
    For Each sld In ActivePresentation.Slides
        For k = sld.Shapes.Count To 1 Step -1
            If InStr(1, sld.Shapes(k).Name, blockerName, vbTextCompare) > 0 Then
                sld.Shapes(k).Delete
            End If
        Next k
    Next sld

    Dim modeMsg As String
    If isDualMode Then modeMsg = "[2쪽(좌우) 모드]" Else modeMsg = "[1쪽(단면) 모드]"

    MsgBox "%msgboxuni_0076%", vbInformation


' ========================================================
' [6] 현재 모듈 삭제 로직 (종료 전 필수 실행 영역)
' ========================================================
DELETE_MODULE:
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    vbProj.VBComponents.Remove vbProj.VBComponents("Module1")
    On Error GoTo 0
    Exit Sub

ERR_HANDLER:
    MsgBox "%msgboxuni_0077%" & Err.Description, vbExclamation, "ScriptRunError"
            GoTo DELETE_MODULE
End Sub


' ========================================================
' [보조 함수] 도형 복제 및 상세 속성 설정
' ========================================================
Private Sub CreatePageNumShape(targetSlide As Slide, srcShape As Shape, txtContent As String, newName As String)
    Dim newShape As Shape

    ' 1. 기본 도형 생성 (위치, 크기)
    Set newShape = targetSlide.Shapes.AddTextbox( _
        srcShape.TextFrame.Orientation, _
        srcShape.Left, srcShape.Top, _
        srcShape.Width, srcShape.Height)

    ' 이름 설정
    newShape.Name = newName

    On Error Resume Next ' 일부 속성 미지원 에러 방지

    ' 2. 텍스트 프레임 및 폰트 상세 복사
    With newShape.TextFrame
        .TextRange.Text = txtContent
        .TextRange.Font.Name = srcShape.TextFrame.TextRange.Font.Name
        .TextRange.Font.Size = srcShape.TextFrame.TextRange.Font.Size
        .TextRange.Font.Bold = srcShape.TextFrame.TextRange.Font.Bold
        .TextRange.Font.Italic = srcShape.TextFrame.TextRange.Font.Italic
        .TextRange.Font.Color.RGB = srcShape.TextFrame.TextRange.Font.Color.RGB
        .TextRange.ParagraphFormat.Alignment = srcShape.TextFrame.TextRange.ParagraphFormat.Alignment

        ' 상세 여백 및 앵커 설정
        .MarginTop = srcShape.TextFrame.MarginTop
        .MarginBottom = srcShape.TextFrame.MarginBottom
        .MarginLeft = srcShape.TextFrame.MarginLeft
        .MarginRight = srcShape.TextFrame.MarginRight
        .AutoSize = srcShape.TextFrame.AutoSize
        .Orientation = srcShape.TextFrame.Orientation
        .VerticalAnchor = srcShape.TextFrame.VerticalAnchor
        .HorizontalAnchor = srcShape.TextFrame.HorizontalAnchor
        .WordWrap = srcShape.TextFrame.WordWrap
    End With

With newShape.TextFrame2.TextRange.Font.Line
        .Visible = msoTrue
        .Weight = 0.25
        .Transparency = 1
    End With


    ' 3. 도형 모양 및 채우기/선 속성 상세 복사
    With newShape
        .Left = srcShape.Left
        .Top = srcShape.Top
        .Width = srcShape.Width
        .Height = srcShape.Height
        .Rotation = srcShape.Rotation
        .AutoShapeType = srcShape.AutoShapeType

        ' 채우기(Fill)
        .Fill.Visible = srcShape.Fill.Visible
        If .Fill.Visible = msoTrue Then
            .Fill.ForeColor.RGB = srcShape.Fill.ForeColor.RGB
            .Fill.BackColor.RGB = srcShape.Fill.BackColor.RGB
            .Fill.Transparency = srcShape.Fill.Transparency
        End If

        ' 선(Line)
        .Line.Visible = srcShape.Line.Visible
        If .Line.Visible = msoTrue Then
            .Line.ForeColor.RGB = srcShape.Line.ForeColor.RGB
            .Line.BackColor.RGB = srcShape.Line.BackColor.RGB
            .Line.Transparency = srcShape.Line.Transparency
            .Line.Weight = srcShape.Line.Weight
            .Line.Style = srcShape.Line.Style
            .Line.DashStyle = srcShape.Line.DashStyle
        End If

        .LockAspectRatio = srcShape.LockAspectRatio
    End With
    On Error GoTo 0
End Sub




)


Clipboard = %Clipboard%%VBA내용_쪽번호증가일괄변경%
ClipWait, 2
if ErrorLevel
{
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}
sleep, 20
}
if (쪽번호증가일괄변경 = 0)
{
;그냥넘기기
}



















; 10/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위







Gui, VBA입력: Hide
Gui, 대칭: Hide

WinClose, Microsoft Visual Basic for Applications




loop
{
if WinActive("Microsoft Visual Basic for Applications")
{
WinSet, Transparent, 0, Microsoft Visual Basic for Applications
sleep, 10
break
}
send, !{f11}
WinSet, Transparent, 0, Microsoft Visual Basic for Applications
sleep, 50
WinSet, Transparent, 0, Microsoft Visual Basic for Applications
sleep, 50
WinSet, Transparent, 0, Microsoft Visual Basic for Applications
sleep, 50
WinSet, Transparent, 0, Microsoft Visual Basic for Applications
sleep, 50
WinSet, Transparent, 0, Microsoft Visual Basic for Applications
sleep, 50
WinSet, Transparent, 0, Microsoft Visual Basic for Applications
sleep, 50
WinActivate, Microsoft Visual Basic for Applications
}






send, !im
sleep, 100


send, ^v
sleep, 100

send, {PgUp 10}
sleep, 100


sleep, 10


if (자동코드실행=1)
{
send, {f5}
sleep, 200

Loop
{




    ; 활성 창 제목 가져오기
    WinGetTitle, title, A
    
    ; 활성 창 제목에 "Microsoft PowerPoint"라는 단어가 포함되어 있는지 확인
    if InStr(title, "Microsoft PowerPoint") {

SetTimer, 정보창툴팁없애기, 1

        ; 메시지 박스에서 Enter 키를 사용해 기본 버튼 클릭
        sleep, 200
        Send, {Enter}
        sleep, 1

break
    }
    




    WinGetTitle, title, A    
    ; 활성 창 제목에 "ScriptRunError"라는 단어가 포함되어 있는지 확인
    if InStr(title, "ScriptRunError") {

; 3. 창을 다시 완전히 보이게 설정
WinSet, Transparent, 255, Microsoft Visual Basic for Applications

WinClose, Microsoft Visual Basic for Applications
WinClose, Microsoft Visual Basic for Applications


번역하기 = 0
대체텍스트변경 = 0
영상트리밍 = 0
캡쳐본생성 = 0
쪽번호전체삭제 = 0
쪽번호증가일괄변경 = 0
자동코드실행 = 0

return
    }






    Sleep, 20 ; CPU 점유율을 낮추기 위해 100ms 대기
}


}
 ;자동코드실행 끝


; 3. 창을 다시 완전히 보이게 설정
WinSet, Transparent, 255, Microsoft Visual Basic for Applications

WinClose, Microsoft Visual Basic for Applications
WinClose, Microsoft Visual Basic for Applications




SetTimer, 정보창툴팁없애기, 1

Gui, VBA입력: Destroy


Gui, 대칭: Destroy


작업끝표시(50, "red")

Clipboard := ""



번역하기 = 0
대체텍스트변경 = 0
영상트리밍 = 0
캡쳐본생성 = 0
쪽번호전체삭제 = 0
쪽번호증가일괄변경 = 0
자동코드실행 = 0




return









번역하기:

; ==========================================================
; [오토핫키 v1] PPT 다중 라인 번역 완벽 해결 스크립트
; ==========================================================

    ; 1. 클립보드 복사 및 초기화
    Clipboard := ""
    Send, ^c
    ClipWait, 1
    If (ErrorLevel) {
        MsgBox, 262208, Message, %msgboxuni_0054%
        Return
    }

    SourceText := Clipboard
    TargetLang := 번역언어

    ; 2. 번역 API 호출
    TranslatedText := GoogleTranslate(SourceText, "auto", TargetLang)

    ; 3. 오류 방지 및 결과 붙여넣기
    if (TranslatedText != "") {
        Send, {Del}
        Sleep, 50
        
        Clipboard := ""
        Clipboard := TranslatedText
        ClipWait, 2
        
        Send, ^v
        Sleep, 200
    } else {
        MsgBox, 262208, Message, %msgboxuni_0001%
    }



번역하기 = 0
대체텍스트변경 = 0
영상트리밍 = 0
캡쳐본생성 = 0
쪽번호전체삭제 = 0
쪽번호증가일괄변경 = 0
자동코드실행 = 0


Return

; ==========================================================
; [핵심 함수] 구글 API 호출 및 안정적인 텍스트 파싱
; ==========================================================
GoogleTranslate(str, from := "auto", to := "ko") {
    EncodedStr := UriEncode_Fix(str)
    URL := "https://translate.googleapis.com/translate_a/single?client=gtx&sl=" . from . "&tl=" . to . "&dt=t&q=" . EncodedStr
    
    try {
        WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        WebRequest.Open("GET", URL, false)
        WebRequest.SetRequestHeader("User-Agent", "Mozilla/5.0")
        WebRequest.Send()
        Response := WebRequest.ResponseText
    } catch {
        return ""
    }

    ; 자바스크립트 엔진을 활용해 다중 라인 배열을 한 번에 조립
    try {
        doc := ComObjCreate("HTMLFile")
        doc.write("<meta http-equiv='X-UA-Compatible' content='IE=edge'>")
        JS := doc.parentWindow
        JS.eval("var data = " . Response . "; var res = ''; for(var i=0; i<data[0].length; i++) { res += data[0][i][0]; }")
        return JS.res
    } catch {
        return ""
    }
}




/*
; ==========================================================
; [보조 함수] 줄바꿈(Enter) 인코딩 버그 완벽 해결
; ==========================================================
UriEncode_Fix(Uri) {
    VarSetCapacity(Var, StrPut(Uri, "UTF-8"), 0)
    StrPut(Uri, &Var, "UTF-8")
    Res := ""
    Loop, % StrPut(Uri, "UTF-8") - 1 {
        byte := NumGet(Var, A_Index-1, "UChar")
        if ((byte >= 0x30 && byte <= 0x39) || (byte >= 0x41 && byte <= 0x5A) || (byte >= 0x61 && byte <= 0x7A) || byte == 0x2D || byte == 0x2E || byte == 0x5F || byte == 0x7E) {
            Res .= Chr(byte)
        } else {
            ; 기존 버그(%A)를 수정하여 무조건 2자리 헥사코드(%0A)로 강제 출력
            Res .= Format("%{:02X}", byte) 
        }
    }
    return Res
}
*/



; ==========================================================
; [보조 함수] 줄바꿈(Enter) 인코딩 버그 완벽 해결 (버전 호환성 강화)
; ==========================================================
UriEncode_Fix(Uri) {
    VarSetCapacity(Var, StrPut(Uri, "UTF-8"), 0)
    StrPut(Uri, &Var, "UTF-8")
    Res := ""
    hex_chars := "0123456789ABCDEF" ; 16진수 변환용 참조 문자열
    
    Loop, % StrPut(Uri, "UTF-8") - 1 {
        byte := NumGet(Var, A_Index-1, "UChar")
        if ((byte >= 0x30 && byte <= 0x39) || (byte >= 0x41 && byte <= 0x5A) || (byte >= 0x61 && byte <= 0x7A) || byte == 0x2D || byte == 0x2E || byte == 0x5F || byte == 0x7E) {
            Res .= Chr(byte)
        } else {
            ; Format() 대신 비트 연산을 활용하여 무조건 2자리 헥사코드(%0A 등)로 출력
            hi := (byte >> 4) & 0xF
            lo := byte & 0xF
            Res .= "%" . SubStr(hex_chars, hi + 1, 1) . SubStr(hex_chars, lo + 1, 1)
        }
    }
    return Res
}












;기타 단축키======================================================================







				;이미지 원래대로
;$^!+7::
Action_Uni0272:

이미지원래대로:

Try
	{
ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type
선택갯수:=ppt.ActiveWindow.Selection.Shaperange.Count()
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
Connector:=ppt.ActiveWindow.Selection.Shaperange.Connector
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
HasTable:=ppt.ActiveWindow.Selection.Shaperange.HasTable
HasTextFrame:=ppt.ActiveWindow.Selection.Shaperange.HasTextFrame
Id:=ppt.ActiveWindow.Selection.Shaperange.Id
Name:=ppt.ActiveWindow.Selection.Shaperange.Name
Parent:=ppt.ActiveWindow.Selection.Shaperange.Parent
Visible:=ppt.ActiveWindow.Selection.Shaperange.Visible
	}



if (객체확인=16)
{
send, {Alt down}{j}{p}{d}{s}{Alt up}
sleep, 10


객체확인:=""
선택갯수:=""
AlternativeText:=""
Connector:=""
AlternativeText:=""
HasTable:=""
HasTextFrame:=""
Id:=""
Name:=""
Parent:=""
Visible:=""

return
}





if (객체확인=28)
{

객체확인:=""
선택갯수:=""
AlternativeText:=""
Connector:=""
AlternativeText:=""
HasTable:=""
HasTextFrame:=""
Id:=""
Name:=""
Parent:=""
Visible:=""

return
}





;객체확인이 16과 28이 아니면 아래를 실행~

send, {Alt down}{j}{p}{q}{s}{Alt up}
sleep, 10



객체확인:=""
선택갯수:=""
AlternativeText:=""
Connector:=""
AlternativeText:=""
HasTable:=""
HasTextFrame:=""
Id:=""
Name:=""
Parent:=""
Visible:=""



return








				;교차
;$^7::
Action_Uni0271:

send, {Alt down}{h}{Alt up}
sleep, 1
send, %단축키도형병합영어%
sleep, 1
send, %단축키도형병합숫자%
sleep, 1
send, {i}
sleep, 1
return













				;재실행
;$^+z::
Action_Uni0391:

send, ^y




return

				;맨뒤로
;$+F8::
$^+[::
Action_Uni0222:

send, {Alt down}{h}{Alt up}
sleep, 1
send, {g}
sleep, 1
send, {k}





return






				;맨위로
;$+F7::
$^+]::
Action_Uni0221:

send, {Alt down}{h}{Alt up}
sleep, 1
send, {g}
sleep, 1
send, {r}




return












				;위로정렬
$F9::

send, {Alt down}{h}{Alt up}

send, {g}

send, {a}

send, {t}



return



				;아래정렬
$F10::

send, {Alt down}{h}{Alt up}

send, {g}

send, {a}

send, {b}


return





				;좌정렬
$F11::

send, {Alt down}{h}{Alt up}

send, {g}

send, {a}

send, {l}


return






				;우정렬
$F12::

send, {Alt down}{h}{Alt up}

send, {g}

send, {a}

send, {r}


return







				;좌우 센터
$^F10::

send, {Alt down}{h}{Alt up}
sleep, 1
send, {g}
sleep, 1
send, {a}
sleep, 1
send, {c}


return






				;상하 센터
$^F12::

send, {Alt down}{h}{Alt up}
sleep, 1
send, {g}
sleep, 1
send, {a}
sleep, 1
send, {m}


return









~MButton & F9::

    ; 툴팁 위치용 마우스 좌표 저장
    mousegetpos, xx1, yy1




    If GetKeyState("Shift", "P")  ; Ctrl 키가 눌려 있는지 확인
    {

;윤곽선투명하게 하기

    ; 파워포인트 창 활성화
    WinActivate, %파워포인트타이틀%



    ; PowerPoint COM 객체 가져오기

Try
{
    ppt := ComObjActive("PowerPoint.Application")
    객체확인 := ppt.ActiveWindow.Selection.ShapeRange.Type


    if (객체확인 = 19)  ; 표가 선택된 경우
    {




send, {Alt down}{j}{t}{Alt up}
sleep, 10
send, {t}
sleep, 10
send, {o}
sleep, 10
send, {w}
sleep, 10
send, {m}
sleep, 500
send, +{tab}
sleep, 10
send, {s}
sleep, 10
send, {s}
sleep, 10
send, {space}
sleep, 500
send, {w}
sleep, 100
send, +{tab}
sleep, 100
send, 100
sleep, 10
send, {enter}
sleep, 10
send, {esc}
sleep, 10


작업끝표시(50, "red")

return




;------------------------------------------------------------------------------------------------------------------------------
; 윤곽선 투명하게하기 테이블일때 테스트중
;------------------------------------------------------------------------------------------------------------------------------
/*
        Try
        {
            ; 선택된 표 Shape & 슬라이드 참조
            sel := ppt.ActiveWindow.Selection.ShapeRange(1)
            sld := sel.Parent

        ; --- 기존에 남아있을 수 있는 임시박스 삭제 ---
           Try sld.Shapes("textlinedell").Delete()
           Try sld.Shapes("textlinedell").Delete()
           Try sld.Shapes("textlinedell").Delete()
;Try tshp.Delete()


            ; 1) 임시 텍스트박스 추가 (셀 서식 복사용)
            tshp := sld.Shapes.AddTextbox(1, 0, 0, 10, 10)  ; msoTextOrientationHorizontal = 1
            tshp.Name := "textlinedell"
            tshp.Visible := False

            ; 2) 표 객체 참조
            Table := sel.Table
            Rows := Table.Rows.Count

            Columns := Table.Columns.Count

            전체카운트 := Rows * Columns
            증가 := 0

            ; 3) 각 셀 순회
            Loop, %Rows%
            {
                row := A_Index
                Loop, %Columns%
                {
                    col := A_Index
                    Cell := Table.Cell(row, col)
                    if ( Cell.Selected )  ; 선택된 셀만 처리
                    {

; HasText 속성은 텍스트가 있으면 -1 (True), 없으면 0 (False)을 반환합니다.
If (Cell.Shape.TextFrame.HasText = 0) 
{

Clipboard := ""
Clipboard := ""
Clipboard := ""

;클립보드 비우기 대기
Loop
{
    ; 클립보드에 내용이 있다면 (성공)
    if (Clipboard = "")
        break ; 반복문 탈출

    ; 클립보드가 여전히 비어있다면 (실패)
;    Sleep, 1 ; 0.2초 대기 (너무 짧으면 CPU 부담, 길면 속도 저하)
Clipboard := ""
}



goto, 윤곽선테이블넘기기
}


Clipboard := ""
Clipboard := ""
Clipboard := ""

;클립보드 비우기 대기
Loop
{
    ; 클립보드에 내용이 있다면 (성공)
    if (Clipboard = "")
        break ; 반복문 탈출

    ; 클립보드가 여전히 비어있다면 (실패)
;    Sleep, 1 ; 0.2초 대기 (너무 짧으면 CPU 부담, 길면 속도 저하)
Clipboard := ""
}





                        ; (1) 셀 텍스트 → 임시박스 복사
                        Cell.Shape.TextFrame2.TextRange.Copy()
                        Cell.Shape.TextFrame2.TextRange.Copy()
                        Cell.Shape.TextFrame2.TextRange.Copy()

;클립보드에 내용 들어왔는지 확인
Loop
{
    ; 클립보드에 내용이 있다면 (성공)
    if (Clipboard != "")
        break ; 반복문 탈출

    ; 클립보드가 여전히 비어있다면 (실패)
;    Sleep, 1 ; 0.2초 대기 (너무 짧으면 CPU 부담, 길면 속도 저하)
    Cell.Shape.TextFrame2.TextRange.Copy() ; 다시 복사 시도
}

                        tshp.TextFrame2.TextRange.Paste()
;sleep, 1
Clipboard := ""
Clipboard := ""
Clipboard := ""


;클립보드 비우기 대기
Loop
{
    ; 클립보드에 내용이 있다면 (성공)
    if (Clipboard = "")
        break ; 반복문 탈출

    ; 클립보드가 여전히 비어있다면 (실패)
;    Sleep, 1 ; 0.2초 대기 (너무 짧으면 CPU 부담, 길면 속도 저하)
Clipboard := ""
}






                        ; (2) 임시박스에 윤곽선 스타일 적용
;                        line := tshp.TextFrame2.TextRange.Font.Line
                        tshp.TextFrame2.TextRange.Font.Line.Visible      := -1       ; msoTrue
;                        tshp.TextFrame2.TextRange.Font.Line.Weight       := 0.75
;                        tshp.TextFrame2.TextRange.Font.Line.ForeColor.RGB:= 14277081
                        tshp.TextFrame2.TextRange.Font.Line.Transparency := 1.0
;sleep, 1


                        ; (3) 임시박스 → 다시 셀에 붙여넣기
                        tshp.TextFrame2.TextRange.Copy()
                        tshp.TextFrame2.TextRange.Copy()
                        tshp.TextFrame2.TextRange.Copy()


;클립보드에 내용 들어왔는지 확인
Loop
{
    ; 클립보드에 내용이 있다면 (성공)
    if (Clipboard != "")
        break ; 반복문 탈출

    ; 클립보드가 여전히 비어있다면 (실패)
;    Sleep, 1 ; 0.2초 대기 (너무 짧으면 CPU 부담, 길면 속도 저하)
    Cell.Shape.TextFrame2.TextRange.Copy() ; 다시 복사 시도
}



                        Cell.Shape.TextFrame2.TextRange.Paste()
;sleep, 1
Clipboard := ""
Clipboard := ""
Clipboard := ""


;클립보드 비우기 대기
Loop
{
    ; 클립보드에 내용이 있다면 (성공)
    if (Clipboard = "")
        break ; 반복문 탈출

    ; 클립보드가 여전히 비어있다면 (실패)
;    Sleep, 1 ; 0.2초 대기 (너무 짧으면 CPU 부담, 길면 속도 저하)
Clipboard := ""
}






윤곽선테이블넘기기:

                    }

                    증가++
                    ToolTIP, ★ Text line Transparent : %증가%/%전체카운트%, %xx1%, %yy1%
                }
            }

            ; 4) 임시 텍스트박스 제거
            Try sld.Shapes("textlinedell").Delete()
            Try sld.Shapes("textlinedell").Delete()
            Try sld.Shapes("textlinedell").Delete()
; Try tshp.Delete()
        }
        Catch
        {
            ; 에러 무시
        }

*/
;------------------------------------------------------------------------------------------------------------------------------
; 윤곽선 투명하게하기 테이블일때 테스트중
;------------------------------------------------------------------------------------------------------------------------------






    }
    else  ; 도형(표 외)이 선택된 경우
    {
        Try
        {
            ; 전체 텍스트에 윤곽선 스타일 적용
            shpRange := ppt.ActiveWindow.Selection.ShapeRange
            line := shpRange.TextFrame2.TextRange.Font.Line
            line.Visible      := -1       ; msoTrue
            line.Weight       := 0.75
            line.ForeColor.RGB:= 14277081
            line.Transparency := 1.0
sleep, 1

            ToolTIP, ★ Text line Transparent, %xx1%, %yy1%
        }
        Catch
        {
            ; 에러 무시
        }
    }



Clipboard := ""
Clipboard := ""
Clipboard := ""

작업끝표시(50, "red")
SetTimer, 정보창툴팁없애기, 800


gosub, 키보드올리기

return



}






    }










/*
            Cell.Shape.TextFrame2.TextRange.Font.line.Visible:=0
            Cell.Shape.TextFrame2.TextRange.Font.line.Weight:=0.750000
            Cell.Shape.TextFrame2.TextRange.Font.line.ForeColor.RGB:=14277081
            Cell.Shape.TextFrame2.TextRange.Font.line.Transparency:=1.000000

            Cell.Shape.TextFrame2.TextRange.Font.line.Visible:=-1
            Cell.Shape.TextFrame2.TextRange.Font.line.Weight:=0.750000
            Cell.Shape.TextFrame2.TextRange.Font.line.ForeColor.RGB:=14277081
            Cell.Shape.TextFrame2.TextRange.Font.line.Transparency:=1.000000

ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.Visible:=0
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.Weight:=0.750000
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.ForeColor.RGB:=14277081
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.Transparency:=1.000000

ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.Visible:=-1
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.Weight:=0.750000
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.ForeColor.RGB:=14277081
ppt.ActiveWindow.Selection.Shaperange.TextFrame2.TextRange.Font.line.Transparency:=1.000000

*/









;shift 없이 휠마우스 f9하면 내부정렬 상단정렬 진행

send, {Alt down}{h}{Alt up}
sleep, 1
send, {a}
sleep, 1
send, {t}
sleep, 1
send, {t}
sleep, 1

return







; 11/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위









~MButton & F10::
    If GetKeyState("Ctrl", "P")  ; Ctrl 키가 눌려 있는지 확인
    {
send, {Alt down}{h}{Alt up}
send, {a}
send, {c}
return
    }

send, {Alt down}{h}{Alt up}

send, {a}

send, {t}

send, {b}

return





~MButton & F11::
send, {Alt down}{h}{Alt up}

send, {a}

send, {l}

return



~MButton & F12::

    If GetKeyState("Ctrl", "P")  ; Ctrl 키가 눌려 있는지 확인
    {
send, {Alt down}{h}{Alt up}

send, {a}

send, {t}

send, {m}

return
    }


send, {Alt down}{h}{Alt up}

send, {a}

send, {r}


return







$!+up::
send, {Alt down}{h}{f}{g}{Alt up}
sleep, 1

return



$!+down::
send, {Alt down}{h}{f}{k}{Alt up}
sleep, 1

return









				;상하 간격 동일
;$+F12::
Action_Uni0212:

send, {Alt down}{h}{Alt up}
sleep, 1
send, {g}
sleep, 1
send, {a}
sleep, 1
send, {v}
sleep, 1

return



				;좌우 간격 동일
;$+F10::
Action_Uni0211:

send, {Alt down}{h}{Alt up}
sleep, 1
send, {g}
sleep, 1
send, {a}
sleep, 1
send, {h}
sleep, 1



return









;문서크기 A4지정

;$^+a::
Action_Uni0401:

gosub, 핫키올림확인
send, {Alt down}{g}{Alt up}
sleep, 1
send, {s}
sleep, 1
send, {c}
sleep, 1
send, {tab}
sleep, 1
send, 21
sleep, 1
send, {tab}
sleep, 1
send, 29.7
sleep, 1
return






;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
;아래4개스크립트중 2개 수정요망


				;그림맞춤
;$!a::
Action_Uni0282:
gosub, 핫키올림확인
send, {Alt down}{h}{Alt up}
sleep, 1
send, %자르기영어%
sleep, 1
send, %자르기숫자%
sleep, 1
send, {l}
sleep, 1
send, {enter}
sleep, 1
return







				;이미지 비디오 크랍

;$!c::
Action_Uni0281:

send, {Alt down}{h}{Alt up}
sleep, 10
send, %자르기영어%
sleep, 10
send, %자르기숫자%
sleep, 10
send, {c}
sleep, 30


return








				;필 스포이드
;$!i::
Action_Uni0321:


Try
	{
ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type
선택갯수:=ppt.ActiveWindow.Selection.Shaperange.Count()
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
Connector:=ppt.ActiveWindow.Selection.Shaperange.Connector
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
HasTable:=ppt.ActiveWindow.Selection.Shaperange.HasTable
HasTextFrame:=ppt.ActiveWindow.Selection.Shaperange.HasTextFrame
Id:=ppt.ActiveWindow.Selection.Shaperange.Id
Name:=ppt.ActiveWindow.Selection.Shaperange.Name
Parent:=ppt.ActiveWindow.Selection.Shaperange.Parent
Visible:=ppt.ActiveWindow.Selection.Shaperange.Visible
	}

gosub, 핫키올림확인


send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{f}
sleep, 10
send, {e}


return






				;텍스트 스포이드
;$!o::
Action_Uni0323:
gosub, 핫키올림확인
send, {Alt down}{h}{Alt up}
sleep, 10
send, {f}{c}
sleep, 10
send, {e}

return




				;스트록 스포이드
;$!p::
Action_Uni0322:

Try
	{
ppt := ComObjActive("PowerPoint.Application")
객체확인:=ppt.ActiveWindow.Selection.Shaperange.type
선택갯수:=ppt.ActiveWindow.Selection.Shaperange.Count()
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
Connector:=ppt.ActiveWindow.Selection.Shaperange.Connector
AlternativeText:=ppt.ActiveWindow.Selection.Shaperange.AlternativeText
HasTable:=ppt.ActiveWindow.Selection.Shaperange.HasTable
HasTextFrame:=ppt.ActiveWindow.Selection.Shaperange.HasTextFrame
Id:=ppt.ActiveWindow.Selection.Shaperange.Id
Name:=ppt.ActiveWindow.Selection.Shaperange.Name
Parent:=ppt.ActiveWindow.Selection.Shaperange.Parent
Visible:=ppt.ActiveWindow.Selection.Shaperange.Visible
	}


gosub, 핫키올림확인


send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{o}
sleep, 10
send, {e}




return









;$!w::
Action_Uni0283:

gosub, 핫키올림확인

				;위로정렬 F9::
send, {Alt down}{h}{Alt up}
sleep, 10
send, {g}
sleep, 10
send, {a}
sleep, 10
send, {t}
sleep, 10

				;좌정렬 F11::
send, {Alt down}{h}{Alt up}
sleep, 10
send, {g}
sleep, 10
send, {a}
sleep, 10
send, {l}
sleep, 10
				;교차 ^7::
send, {Alt down}{h}{Alt up}
sleep, 10
send, %단축키도형병합영어%
sleep, 10
send, %단축키도형병합숫자%
sleep, 10
send, {i}
sleep, 100
				;그림맞춤 !a::
send, {Alt down}{h}{Alt up}
sleep, 10
send, %자르기영어%
sleep, 10
send, %자르기숫자%
sleep, 10
send, {l}
sleep, 10
send, {enter}
sleep, 10




작업끝표시(50, "red")


gosub, 키보드올리기
return



;★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★




;$^!+Backspace::
Action_Uni0203:

Try {
    ppt := ComObjActive("PowerPoint.Application")
    객체확인타입 := ppt.ActiveWindow.Selection.Type
    ; 0: 없음, 2: 객체(도형/표), 3: 텍스트(커서)

    SetFormat, float, 0.2

    ; [Case 1] 텍스트 편집 모드 (커서 활성화 또는 드래그)
    If (객체확인타입 = 3) {
        SelectedTR := ppt.ActiveWindow.Selection.TextRange2
        반복 := SelectedTR.Paragraphs.Count
        Loop, %반복% {
            Para := SelectedTR.Paragraphs(A_Index).ParagraphFormat
            Para.LeftIndent := 10
            Para.FirstLineIndent := -10
            텍스트앞값 := Para.LeftIndent
            내어쓰기값 := Para.FirstLineIndent

        }
    }
    ; [Case 2] 객체 선택 모드 (도형 전체 또는 테이블 셀)
    Else If (객체확인타입 = 2) {
        ; 선택된 객체가 테이블(Type 19)인지 확인
        ShapeType := ppt.ActiveWindow.Selection.ShapeRange.Type
        
        If (ShapeType = 19) { ; 테이블인 경우
            Table := ppt.ActiveWindow.Selection.ShapeRange.Table
            Loop, % Table.Rows.Count {
                r := A_Index
                Loop, % Table.Columns.Count {
                    c := A_Index
                    Cell := Table.Cell(r, c)
                    If (Cell.Selected) {
                        ; 셀 내부의 모든 문단 조정
                        ParaFmt := Cell.Shape.TextFrame2.TextRange.ParagraphFormat
                        ParaFmt.LeftIndent  := 10
                        ParaFmt.FirstLineIndent := -10
                        텍스트앞값 := ParaFmt.LeftIndent
                        내어쓰기값 := ParaFmt.FirstLineIndent
                    }
                }
            }
        } 
        Else { ; 일반 도형인 경우
            ParaFmt := ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat
            ParaFmt.LeftIndent := 10
            ParaFmt.FirstLineIndent := -10
            텍스트앞값 := ParaFmt.LeftIndent
            내어쓰기값 := ParaFmt.FirstLineIndent
        }
    }
    ; [Case 3] 선택 안됨
    Else {
        MsgBox, 262208, Message, %msgboxuni_0032%
        Return
    }

    ; 결과 표시
    GoSub, 결과툴팁표시
} Catch e {
}
Return









;단락 내어쓰기 간격조정

; =======================================================
; 왼쪽 정렬 및 첫 줄 내어쓰기 증가 (내어쓰기 강화)
; =======================================================

;$^!+left::

Action_Uni0202:

Try {
    ppt := ComObjActive("PowerPoint.Application")
    객체확인타입 := ppt.ActiveWindow.Selection.Type
    ; 0: 없음, 2: 객체(도형/표), 3: 텍스트(커서)

    SetFormat, float, 0.2

    ; [Case 1] 텍스트 편집 모드 (커서 활성화 또는 드래그)
    If (객체확인타입 = 3) {
        SelectedTR := ppt.ActiveWindow.Selection.TextRange2
        반복 := SelectedTR.Paragraphs.Count
        Loop, %반복% {
            Para := SelectedTR.Paragraphs(A_Index).ParagraphFormat
            Para.LeftIndent := Para.LeftIndent  - Uni0202_Key설정값
            Para.FirstLineIndent := Para.FirstLineIndent + Uni0202_Key설정값
            텍스트앞값 := Para.LeftIndent
            내어쓰기값 := Para.FirstLineIndent
        }
    }
    ; [Case 2] 객체 선택 모드 (도형 전체 또는 테이블 셀)
    Else If (객체확인타입 = 2) {
        ; 선택된 객체가 테이블(Type 19)인지 확인
        ShapeType := ppt.ActiveWindow.Selection.ShapeRange.Type
        
        If (ShapeType = 19) { ; 테이블인 경우
            Table := ppt.ActiveWindow.Selection.ShapeRange.Table
            Loop, % Table.Rows.Count {
                r := A_Index
                Loop, % Table.Columns.Count {
                    c := A_Index
                    Cell := Table.Cell(r, c)
                    If (Cell.Selected) {
                        ; 셀 내부의 모든 문단 조정
                        ParaFmt := Cell.Shape.TextFrame2.TextRange.ParagraphFormat
                        ParaFmt.LeftIndent  := ParaFmt.LeftIndent - Uni0202_Key설정값
                        ParaFmt.FirstLineIndent := ParaFmt.FirstLineIndent + Uni0202_Key설정값
                        텍스트앞값 := ParaFmt.LeftIndent
                        내어쓰기값 := ParaFmt.FirstLineIndent
                    }
                }
            }
        } 
        Else { ; 일반 도형인 경우
            ParaFmt := ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat
            ParaFmt.LeftIndent := ParaFmt.LeftIndent - Uni0202_Key설정값
            ParaFmt.FirstLineIndent := ParaFmt.FirstLineIndent + Uni0202_Key설정값
            텍스트앞값 := ParaFmt.LeftIndent
            내어쓰기값 := ParaFmt.FirstLineIndent
        }
    }
    ; [Case 3] 선택 안됨
    Else {
        MsgBox, 262208, Message, %msgboxuni_0032%
        Return
    }

    ; 결과 표시
    GoSub, 결과툴팁표시
} Catch e {
}
Return





; =======================================================
; 왼쪽 정렬 및 첫 줄 내어쓰기 감소 (내어쓰기 약화)
; =======================================================


;$^!+right::
Action_Uni0201:

Try {
    ppt := ComObjActive("PowerPoint.Application")
    객체확인타입 := ppt.ActiveWindow.Selection.Type

    SetFormat, float, 0.2

    If (객체확인타입 = 3) {
        SelectedTR := ppt.ActiveWindow.Selection.TextRange2
        반복 := SelectedTR.Paragraphs.Count
        Loop, %반복% {
            Para := SelectedTR.Paragraphs(A_Index).ParagraphFormat
            Para.LeftIndent := Para.LeftIndent  + Uni0201_Key설정값
            Para.FirstLineIndent := Para.FirstLineIndent - Uni0201_Key설정값
            텍스트앞값 := Para.LeftIndent
            내어쓰기값 := Para.FirstLineIndent
        }
    }
    Else If (객체확인타입 = 2) {
        ShapeType := ppt.ActiveWindow.Selection.ShapeRange.Type
        
        If (ShapeType = 19) { ; 테이블
            Table := ppt.ActiveWindow.Selection.ShapeRange.Table
            Loop, % Table.Rows.Count {
                r := A_Index
                Loop, % Table.Columns.Count {
                    c := A_Index
                    Cell := Table.Cell(r, c)
                    If (Cell.Selected) {
                        ParaFmt := Cell.Shape.TextFrame2.TextRange.ParagraphFormat
                        ParaFmt.LeftIndent := ParaFmt.LeftIndent + Uni0201_Key설정값
                        ParaFmt.FirstLineIndent := ParaFmt.FirstLineIndent - Uni0201_Key설정값
                        텍스트앞값 := ParaFmt.LeftIndent
                        내어쓰기값 := ParaFmt.FirstLineIndent
                    }
                }
            }
        } 
        Else { ; 일반 도형
            ParaFmt := ppt.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat
            ParaFmt.LeftIndent := ParaFmt.LeftIndent + Uni0201_Key설정값
            ParaFmt.FirstLineIndent := ParaFmt.FirstLineIndent - Uni0201_Key설정값
            텍스트앞값 := ParaFmt.LeftIndent
            내어쓰기값 := ParaFmt.FirstLineIndent
        }
    }
    Else {
        MsgBox, 262208, Message, %msgboxuni_0032%
        Return
    }

    GoSub, 결과툴팁표시
} Catch e {
}
Return





; -----------------------------------------------------------
; 공통 서브루틴
; -----------------------------------------------------------
결과툴팁표시:

; 1 pt = 0.0352778 cm (계산 편의상 0.03528 사용)
변환상수 := 0.03528

; 포인트 값을 cm로 계산
텍스트앞_cm := 텍스트앞값 * 변환상수
내어쓰기_cm := 내어쓰기값 * 변환상수

; 소수점 둘째 자리까지 포맷팅 (SetFormat 또는 Round 함수 사용)
텍스트앞_cm := Round(텍스트앞_cm, 2)
내어쓰기_cm := Round(내어쓰기_cm, 2)

; 툴팁 출력
MouseGetPos, xx1, yy1
ToolTip, ★ Paragraph Indentation : Before text : %텍스트앞_cm% cm _by : %내어쓰기_cm% cm, %xx1%, %yy1%
SetTimer, 정보창툴팁없애기, 800

Return









;--- [2025-10-22 지침 반영: v1 문법] ---

; Ctrl + Alt + Win + ] : 투명도 10% 증가 (점점 안 보이게)
;#!]:: 
Action_Uni0313:
    ChangeTransparency(0.1)
return

; Ctrl + Alt + Win + [ : 투명도 10% 감소 (점점 잘 보이게/진하게)
;#![:: 
Action_Uni0314:
    ChangeTransparency(-0.1)
return

; ==============================================================================
; [핵심 함수] 선택한 개체(도형 또는 표의 특정 셀) 투명도 조절
; ==============================================================================
ChangeTransparency(val) {
    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        ; 아무것도 선택하지 않았거나 슬라이드 배경 등을 선택한 경우 무시
        if (sel.Type = 0 || sel.Type = 1) 
            return 

        hasNoFill := false
        
        try {
            ; 선택된 개체 덩어리(ShapeRange)를 가져옴
            shps := sel.ShapeRange
            Loop % shps.Count {
                shp := shps.Item(A_Index)
                
                ; 1. 선택한 개체가 표(Table)인 경우
                if (shp.HasTable) {
                    tableModified := false
                    
                    ; [핵심 로직] 표 안의 개별 셀들을 하나씩 스캔
                    Loop % shp.Table.Rows.Count {
                        r := A_Index
                        Loop % shp.Table.Columns.Count {
                            c := A_Index
                            cell := shp.Table.Cell(r, c)
                            
                            ; 현재 스캔 중인 셀이 사용자에 의해 '선택된 상태'라면
                            if (cell.Selected) {
                                tableModified := true
                                
                                ; 채우기 없음 확인
                                if (cell.Shape.Fill.Visible = 0) {
                                    hasNoFill := true
                                    continue
                                }
                                
                                ; 투명도 계산 및 적용
                                curr := cell.Shape.Fill.Transparency
                                newTrans := Round(curr + val, 1)
                                if (newTrans > 1.0)
                                    newTrans := 1.0
                                if (newTrans < 0.0)
                                    newTrans := 0.0
                                cell.Shape.Fill.Transparency := newTrans
                            }
                        }
                    }
                    
                    ; 만약 개별 셀을 드래그하지 않고, 표 '전체 외곽선'을 클릭한 경우라면
                    if (!tableModified) {
                        Loop % shp.Table.Rows.Count {
                            r := A_Index
                            Loop % shp.Table.Columns.Count {
                                c := A_Index
                                cell := shp.Table.Cell(r, c)
                                if (cell.Shape.Fill.Visible = 0) {
                                    hasNoFill := true
                                    continue
                                }
                                curr := cell.Shape.Fill.Transparency
                                newTrans := Round(curr + val, 1)
                                if (newTrans > 1.0)
                                    newTrans := 1.0
                                if (newTrans < 0.0)
                                    newTrans := 0.0
                                cell.Shape.Fill.Transparency := newTrans
                            }
                        }
                    }
                }
                
                ; 2. 일반 도형인 경우
                else {
                    if (shp.Fill.Visible = 0) {
                        hasNoFill := true
                        continue
                    }
                    curr := shp.Fill.Transparency
                    newTrans := Round(curr + val, 1)
                    if (newTrans > 1.0)
                        newTrans := 1.0
                    if (newTrans < 0.0)
                        newTrans := 0.0
                    shp.Fill.Transparency := newTrans
                }
            }
        } catch {
            MsgBox, 262208, Message, %msgboxuni_0032%
            return
        }
        
        ; 채우기 없음 상태인 항목이 하나라도 있었다면 알림창 띄우기
        if (hasNoFill) {
            MsgBox, 262208, Message, %msgboxuni_0031%
        }
        
    } catch {
        ; 파워포인트가 실행 중이 아니면 조용히 넘어감
    }
}






;--- [2025-10-22 지침 반영: v1 문법] ---

; Ctrl + Alt + ] : 텍스트 투명도 10% 증가 (점점 안 보이게)
;^!]:: 
Action_Uni0315:

    ChangeTextTransparency(0.1)
return

; Ctrl + Alt + [ : 텍스트 투명도 10% 감소 (점점 진하게)
;^![:: 
Action_Uni0316:

    ChangeTextTransparency(-0.1)
return

; ==============================================================================
; [핵심 함수] 상황별 텍스트 투명도 조절 (커서/도형/표 자동 분별)
; ==============================================================================
ChangeTextTransparency(val) {
    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        type := sel.Type

        ; Type 0: 선택 없음 / 1: 슬라이드 배경 선택 (무시)
        if (type = 0 || type = 1)
            return

        ; =======================================================
        ; [상태 1] 텍스트 커서가 깜빡이거나 텍스트를 드래그한 상태 (1순위)
        ; =======================================================
        if (type = 3) { ; ppSelectionText
            try {
                tr := sel.TextRange2
                UpdateTextTrans(tr, val)
            }
        }
        ; =======================================================
        ; [상태 2 & 3] 도형 테두리나 표(Table)를 선택한 상태
        ; =======================================================
        else if (type = 2) { ; ppSelectionShapes
            shps := sel.ShapeRange
            Loop % shps.Count {
                shp := shps.Item(A_Index)
                
                ; [상태 3] 표(Table)인 경우
                if (shp.HasTable) {
                    tableModified := false
                    
                    ; 표 안의 개별 셀 스캔
                    Loop % shp.Table.Rows.Count {
                        r := A_Index
                        Loop % shp.Table.Columns.Count {
                            c := A_Index
                            cell := shp.Table.Cell(r, c)
                            
                            ; 마우스로 드래그하여 선택된 셀이라면
                            if (cell.Selected) {
                                tableModified := true
                                if (cell.Shape.HasTextFrame) {
                                    tr := cell.Shape.TextFrame2.TextRange
                                    UpdateTextTrans(tr, val)
                                }
                            }
                        }
                    }
                    
                    ; 특정 셀이 아닌 표 테두리 전체를 클릭한 경우
                    if (!tableModified) {
                        Loop % shp.Table.Rows.Count {
                            r := A_Index
                            Loop % shp.Table.Columns.Count {
                                c := A_Index
                                cell := shp.Table.Cell(r, c)
                                if (cell.Shape.HasTextFrame) {
                                    tr := cell.Shape.TextFrame2.TextRange
                                    UpdateTextTrans(tr, val)
                                }
                            }
                        }
                    }
                }
                ; [상태 2] 일반 도형인 경우
                else {
                    if (shp.HasTextFrame) {
                        tr := shp.TextFrame2.TextRange
                        UpdateTextTrans(tr, val)
                    }
                }
            }
        }
    } catch {
        MsgBox, 262208, Message, %msgboxuni_0033%
    }
}

; ==============================================================================
; [보조 함수] 실제 투명도 값을 계산하고 텍스트에 적용
; ==============================================================================
UpdateTextTrans(tr, val) {
    try {
        ; 글자가 아예 없는 빈 텍스트 박스면 스킵
        if (tr.Text = "")
            return
            
        curr := tr.Font.Fill.Transparency
        
        ; 여러 색상이나 투명도가 섞여 있어 값을 읽을 수 없는 경우(msoMixed) 0으로 초기화
        if (curr < 0 || curr > 1.0)
            curr := 0.0
            
        newTrans := Round(curr + val, 1)
        
        if (newTrans > 1.0)
            newTrans := 1.0
        if (newTrans < 0.0)
            newTrans := 0.0
            
        tr.Font.Fill.Transparency := newTrans
    } catch {
        ; 투명도를 지원하지 않는 특수 폰트/개체의 에러 무시
    }
}











;--- [2025-10-22 지침 반영: v1 문법] ---

; Win + Alt + ] : 윤곽선 투명도 10% 증가 (점점 안 보이게)
;^#!]:: 
Action_Uni0317:

    ChangeLineTransparency(0.1)
return

; Win + Alt + [ : 윤곽선 투명도 10% 감소 (점점 진하게)
;^#![:: 
Action_Uni0318:

    ChangeLineTransparency(-0.1)
return

; ==============================================================================
; [핵심 함수] 선택한 개체의 윤곽선(테두리) 투명도 조절 (표 차단 기능 포함)
; ==============================================================================
ChangeLineTransparency(val) {
    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        type := sel.Type

        ; 선택한 항목이 없거나 배경을 선택한 경우 무시
        if (type = 0 || type = 1)
            return

        hasNoLine := false
        isTableSelected := false

        try {
            ; 선택된 개체 덩어리(ShapeRange)를 가져옴
            shps := sel.ShapeRange
            Loop % shps.Count {
                shp := shps.Item(A_Index)
                
                ; ★ [핵심] 선택한 개체가 표(Table)인지 검사
                if (shp.HasTable) {
                    isTableSelected := true
                    continue ; 표인 경우 투명도 조절을 건너뜀
                }
                
                ; 1. 윤곽선 없음(msoFalse = 0) 상태인지 확인
                if (shp.Line.Visible = 0) {
                    hasNoLine := true
                    continue
                }
                
                ; 2. 현재 투명도 가져오기
                curr := shp.Line.Transparency
                
                ; 섞인 값(msoMixed) 오류 방지
                if (curr < 0 || curr > 1.0)
                    curr := 0.0
                    
                newTrans := Round(curr + val, 1)
                
                ; 3. 한계치(0% ~ 100%) 고정
                if (newTrans > 1.0)
                    newTrans := 1.0
                if (newTrans < 0.0)
                    newTrans := 0.0
                    
                ; 새 투명도 적용
                shp.Line.Transparency := newTrans
            }
        } catch {
            MsgBox, 262208, Message, %msgboxuni_0034%
            return
        }
        
        ; ==============================================================================
        ; [알림창 처리 구역] 조건에 맞는 알림창을 최상단(262192/262208)으로 띄움
        ; ==============================================================================
        if (isTableSelected) {
            ; 사장님이 요청하신 테이블 차단 메시지
            MsgBox, 262208, Message, %msgboxuni_0035%
        } else if (hasNoLine) {
            ; 이전 기능들과의 통일성을 위해 '선 없음' 안내도 추가
            MsgBox, 262208, Message, %msgboxuni_0036%
        }
        
    } catch {
        ; 파워포인트 미실행 시 에러 무시
    }
}





















				;영어 대문자로하기
;$!u::
Action_Uni0371:
gosub, 핫키올림확인

send, {Alt down}{h}{Alt up}
sleep, 10
send, {7}
sleep, 10
send, {u}
sleep, 30



작업끝표시(50, "red")


gosub, 키보드올리기

return




				;영어 단어앞 대문자로하기
;$!y::
Action_Uni0372:

gosub, 핫키올림확인

send, {Alt down}{h}{Alt up}
sleep, 10
send, {7}
sleep, 10
send, {c}
sleep, 30



작업끝표시(50, "red")


gosub, 키보드올리기

return






				;영어 소문자로하기
;$!t::
Action_Uni0373:

gosub, 핫키올림확인

send, {Alt down}{h}{Alt up}
sleep, 10
send, {7}
sleep, 10
send, {l}
sleep, 30


작업끝표시(50, "red")


gosub, 키보드올리기

return







				;검정바탕 흰색글씨
;$^/::
Action_Uni0301:

gosub, 핫키올림확인

send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{f}
sleep, 10
send, {right}
sleep, 10
send, {enter}
sleep, 30

send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{o}
sleep, 10
send, {n}
sleep, 30

send, {Alt down}{h}{Alt up}
sleep, 10
send, {t}{c}
sleep, 10
send, {n}
sleep, 30


send, {Alt down}{h}{Alt up}
sleep, 10
send, {f}{c}
sleep, 10
send, {down} ;밝은회색으로
sleep, 10
send, {enter}
sleep, 30



send, {AppsKey}{o}
sleep, 30

send, {esc}
sleep, 10







작업끝표시(50, "red")


gosub, 키보드올리기

return




				;투명바탕 검정글씨
;$^'::
Action_Uni0302:

gosub, 핫키올림확인

send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{f}
sleep, 10
send, {n}
sleep, 30

send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{o}
sleep, 10
send, {n}
sleep, 30


send, {Alt down}{h}{Alt up}
sleep, 10
send, {f}{c}
sleep, 10
send, {down 3} ;어두운회색으로
sleep, 10
send, {right}
sleep, 10
send, {enter}
sleep, 30


send, {Alt down}{h}{Alt up}
sleep, 10
send, {t}{c}
sleep, 10
send, {n}
sleep, 30




send, {AppsKey}{o}
sleep, 30


send, {esc}
sleep, 10






작업끝표시(50, "red")

gosub, 키보드올리기
return











				;채우기없음
;$!/::
Action_Uni0311:

gosub, 핫키올림확인

send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{f}
sleep, 10
send, {n}

send, {AppsKey}{o}
sleep, 30

send, {esc}
sleep, 10


작업끝표시(50, "red")

gosub, 키보드올리기

return





				;라인없음
;$!.::
Action_Uni0312:

gosub, 핫키올림확인

send, {Alt down}{h}{Alt up}
sleep, 10
send, {s}{o}
sleep, 10
send, {n}

send, {AppsKey}{o}
sleep, 30

send, {esc}
sleep, 10


작업끝표시(50, "red")

gosub, 키보드올리기

return









; 12/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위












;사용된 폰트명으로 객체 찾아보기

; ==============================================================================
; PPT 폰트 위치 찾기 (AutoHotkey v1.1) - 최종 통합본 (체크박스 옵션 추가)
; ==============================================================================
; [기능 요약]
; 1. Next/Prev 양방향 탐색 지원
; 2. Flag Tracking 및 Reverse Index 방식 혼용
; 3. '빈 텍스트 상자도 검색' 옵션 추가 (GUI 체크박스)
; ==============================================================================


; [사용자 요청 복원] 툴팁 위치 및 출력 로직



폰트위치찾기실행:
ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경

Gui, 대칭:Destroy
Gui, VBA입력: Destroy

    PPT_FontFinder_Show()
return



; ==============================================================================
; [UI] 초기화 및 GUI 생성
; ==============================================================================
PPT_FontFinder_Show() {
    ; [수정된 부분] 외부에서 읽어온 INI 텍스트 변수들을 모두 global로 선언해 줍니다.
    global PPTGuiFontName, 빈상자도검색, FindFontList, FindFontBtngo, FindFontgotoend, FindFontwinclose, FindFontDes1




    ; 1. PPT 연결 및 초기 설정
    DocFontList := "|" 
    try 
    {

        pptApp := ComObjActive("PowerPoint.Application")
        activePres := pptApp.ActivePresentation
        activeWindow := pptApp.ActiveWindow
        
        ; [기능] 실행 시 현재 슬라이드를 '마지막 슬라이드'로 이동 (기존 로직 유지)
        totalSlides := activePres.Slides.Count
        if (totalSlides > 0) {
            activeWindow.View.GotoSlide(totalSlides)
        }


        ; 폰트 리스트 수집
        Loop, % activePres.Fonts.Count 
        {
            fName := activePres.Fonts.Item(A_Index).Name
            if (fName != "")
                DocFontList .= fName . "|"
        }



    } 
    catch 
    {
        MsgBox, 262208, Message, %msgboxuni_0015%
        return
    }


    ; 2. GUI 생성
    Gui, PPT_Finder:Destroy
    ;Gui, PPT_Finder:New, +AlwaysOnTop +ToolWindow, Font Search Slide
    Gui, PPT_Finder:New, +AlwaysOnTop, Font Search Slide
    Gui, PPT_Finder:Font, s9, 맑은 고딕
    
    ; (1) 라벨
    Gui, PPT_Finder:Add, Text, x20 y15 w300 Center, %FindFontList%
    
    ; (2) 콤보박스
    Gui, PPT_Finder:Add, ComboBox, x20 y40 w300 vPPTGuiFontName, %DocFontList%

    ; (3) 버튼 그룹
    Gui, PPT_Finder:Add, Button, x20 y85 w145 h85 gPPT_BtnPrev default, %FindFontBtngo%

    Gui, PPT_Finder:Font, s8, 맑은 고딕
    Gui, PPT_Finder:Add, Button, x176 y85 w145 h40 gPPT_BtnNext, %FindFontgotoend%
    Gui, PPT_Finder:Font, s9, 맑은 고딕
    Gui, PPT_Finder:Add, Button, x176 y130 w145 h40 gPPT_FinderGuiClose, %FindFontwinclose%
    
    ; (4) [추가] 빈 상자 검색 옵션 (기본값: Checked)
    Gui, PPT_Finder:Add, Checkbox, x20 y180 w300 v빈상자도검색, %FindFontDes1%

    GuiControl, Choose, PPTGuiFontName, 2
    

    Gui, PPT_Finder:Show, w340 h215, Font Search Slide


RestoreCursors()       ; 커서 복구 완료

작업끝표시(50, "red")
    return

}





PPT_BtnPrev:
    Gui, PPT_Finder:Submit, NoHide
    PPT_FindFont("Prev", PPTGuiFontName)
return

PPT_BtnNext:
    Gui, PPT_Finder:Submit, NoHide
;    PPT_FindFont("Next", PPTGuiFontName)

try {
        ; PPT 애플리케이션 객체 연결
        tmpPPT := ComObjActive("PowerPoint.Application")
        ; 전체 슬라이드 개수 확인 (마지막 페이지 번호)
        lastSlideIdx := tmpPPT.ActivePresentation.Slides.Count
        ; 마지막 페이지로 이동
        tmpPPT.ActiveWindow.View.GotoSlide(lastSlideIdx)
        ; 객체 해제 (필수는 아니지만 권장)
        tmpPPT := ""
    } catch {
        ; 오류 발생 시 무시하고 진행
    }

return

; ==============================================================================
; [핵심 로직] Flag Tracking + Index Reverse + Ghost Font Support + Option Check
; ==============================================================================
PPT_FindFont(Direction, TargetFont) {
    global PPT_Search_PassedTarget ; 전역 변수

    TargetFont := Trim(TargetFont)
    if (TargetFont = "") {
        MsgBox, 262208, Message, %msgboxuni_0037%
        return
    }

    try {
        pptApp := ComObjActive("PowerPoint.Application")
    } catch {
        MsgBox, 262208, Message, %msgboxuni_0015%
        return
    }

    try {
        activeWindow := pptApp.ActiveWindow
        activePres := pptApp.ActivePresentation
        
        currSlideIndex := 1
        try {
            currSlideIndex := activeWindow.View.Slide.SlideIndex
        }
        
        ; 현재 선택된 객체 ID
        currSelID := ""
        try {
            if (activeWindow.Selection.Type = 2 || activeWindow.Selection.Type = 3) {
                currSelID := activeWindow.Selection.ShapeRange.Item(1).Id
            }
        }
        
        totalSlides := activePres.Slides.Count
        
        ; --- [다음 찾기: Flag Tracking] ---
        if (Direction = "Next") {
            Loop, %totalSlides% {
                checkSlideIdx := currSlideIndex + A_Index - 1
                if (checkSlideIdx > totalSlides) 
                    Break

                slide := activePres.Slides.Item(checkSlideIdx)
                isCurrentSlide := (checkSlideIdx = currSlideIndex)
                
                ; 상태 설정
                if (isCurrentSlide && currSelID != "") {
                    PPT_Search_PassedTarget := false
                } else {
                    PPT_Search_PassedTarget := true
                }
                
                if (PPT_SearchInSlide_Next(slide, TargetFont, activeWindow, currSelID)) {
                    return 
                }
            }
            MsgBox, 262208, Message, %msgboxuni_0038%`n%TargetFont% 
        } 
        
        ; --- [이전 찾기: Index Reverse] ---
        else {
            Loop, %totalSlides% {
                checkSlideIdx := currSlideIndex - A_Index + 1
                if (checkSlideIdx < 1) 
                    Break

                slide := activePres.Slides.Item(checkSlideIdx)
                shouldSkipSelection := (checkSlideIdx = currSlideIndex)
                
                if (PPT_SearchInSlide_Prev(slide, TargetFont, activeWindow, shouldSkipSelection)) {
                    return
                }
            }
            MsgBox, 262208, Message, %msgboxuni_0038%`n%TargetFont%
        }

    } catch e {
;Gui, PPT_Finder: Destroy
        MsgBox, 262208, Message, %msgboxuni_0029%`n%e%
    }
}

; ==============================================================================
; [NEXT 전용] 플래그 추적 검색 함수 (옵션 반영)
; ==============================================================================
PPT_SearchInSlide_Next(slide, targetFont, activeWindow, currSelID) {
    try {
        shapes := slide.Shapes
        count := shapes.Count
    } catch {
        return false
    }

    Loop, %count% {
        try {
            thisShape := shapes.Item(A_Index)
            
            foundObj := PPT_RecursiveCheck_Next(thisShape, targetFont, currSelID)
            
            if (IsObject(foundObj)) {
                if (activeWindow.View.Slide.SlideIndex != slide.SlideIndex) {
                    activeWindow.View.GotoSlide(slide.SlideIndex)
                }
                try {
                    foundObj.Select()
                } catch {
                    try { 
                        thisShape.Select() 
                    }
                }
                return true
            }
        }
    }
    return false
}

; [NEXT 전용] 재귀 함수
PPT_RecursiveCheck_Next(shp, targetName, currSelID) {
    global PPT_Search_PassedTarget, 빈상자도검색 ; [수정] GUI 변수 가져옴
    
    try {
        ; 1. 기준점 확인
        if (!PPT_Search_PassedTarget && currSelID != "" && shp.Id = currSelID) {
            PPT_Search_PassedTarget := true
            ; 여기서 Return False 하면 안됨. 표/그룹의 경우 내부로 들어가야 함.
        }

        ; 2. 표(Table) 탐색
        if (shp.Type = 19) { 
            tbl := shp.Table
            Loop, % tbl.Rows.Count {
                r := A_Index
                Loop, % tbl.Columns.Count {
                    c := A_Index
                    try {
                        cellShape := tbl.Cell(r, c).Shape
                        res := PPT_RecursiveCheck_Next(cellShape, targetName, currSelID)
                        if (IsObject(res)) {
                            return res
                        }
                    }
                }
            }
            return false
        }

        ; 3. 그룹(Group) 탐색
        if (shp.Type = 6) { 
            Loop, % shp.GroupItems.Count {
                res := PPT_RecursiveCheck_Next(shp.GroupItems.Item(A_Index), targetName, currSelID)
                if (IsObject(res)) {
                    return res
                }
            }
            return false
        }
        
        ; 4. 텍스트 확인 (옵션 적용)
        if (PPT_Search_PassedTarget) {
            ; 깃발 올린 직후의 '자기 자신' 재확인 방지
            if (currSelID != "" && shp.Id = currSelID) {
                return false
            }

            ; [요청 수정 로직 반영]
            ShouldCheck := false
            
            if (빈상자도검색 = 1) {
                ; 체크됨: 텍스트 없어도(프레임만 있어도) 검사
                if (shp.HasTextFrame) {
                    ShouldCheck := true
                }
            } else {
                ; 체크해제: 텍스트가 있어야만 검사
                if (shp.HasTextFrame && shp.TextFrame.HasText) {
                    ShouldCheck := true
                }
            }

            if (ShouldCheck) {
                try {
                    if (PPT_IsFontMatch(shp.TextFrame.TextRange.Font, targetName)) {
                        return shp
                    }
                }
            }
        }
    }
    return false
}

; ==============================================================================
; [PREV 전용] 기존 인덱스 방식 함수 (옵션 반영)
; ==============================================================================
PPT_SearchInSlide_Prev(slide, targetFont, activeWindow, shouldSkipSelection) {
    try {
        shapes := slide.Shapes
        count := shapes.Count
    } catch {
        return false
    }

    startIndex := count + 1
    
    if (shouldSkipSelection) {
        selIndex := PPT_GetSelectionIndex(activeWindow, shapes)
        if (selIndex > 0) {
            startIndex := selIndex
        }
    }

    Loop, %count% {
        idx := count - A_Index + 1
        
        if (shouldSkipSelection && idx >= startIndex) {
            Continue
        }

        try {
            thisShape := shapes.Item(idx)
            foundObj := PPT_CheckFontRecursive_Prev(thisShape, targetFont)
            
            if (IsObject(foundObj)) {
                if (activeWindow.View.Slide.SlideIndex != slide.SlideIndex) {
                    activeWindow.View.GotoSlide(slide.SlideIndex)
                }
                try {
                    foundObj.Select()
                } catch {
                    try { 
                        thisShape.Select() 
                    }
                }
                return true
            }
        }
    }
    return false
}

PPT_CheckFontRecursive_Prev(shp, targetName) {
    global 빈상자도검색 ; [수정] GUI 변수 가져옴

    try {
        ; 1. 표(Table)
        if (shp.Type = 19) { 
            tbl := shp.Table
            Loop, % tbl.Rows.Count {
                r := A_Index
                Loop, % tbl.Columns.Count {
                    c := A_Index
                    try {
                        cellShape := tbl.Cell(r, c).Shape
                        
                        ; [표 셀 내부 옵션 적용]
                        ShouldCheck := false
                        if (빈상자도검색 = 1) {
                            if (cellShape.HasTextFrame)
                                ShouldCheck := true
                        } else {
                            if (cellShape.HasTextFrame && cellShape.TextFrame.HasText)
                                ShouldCheck := true
                        }

                        if (ShouldCheck) {
                             if (PPT_IsFontMatch(cellShape.TextFrame.TextRange.Font, targetName)) {
                                 return cellShape
                             }
                        }
                    }
                }
            }
        }
        
        ; 2. 그룹(Group)
        if (shp.Type = 6) { 
            Loop, % shp.GroupItems.Count {
                foundItem := PPT_CheckFontRecursive_Prev(shp.GroupItems.Item(A_Index), targetName)
                if (IsObject(foundItem)) {
                    return foundItem
                }
            }
        }
        
        ; 3. 일반 텍스트 (옵션 적용)
        ShouldCheck := false
        if (빈상자도검색 = 1) {
            if (shp.HasTextFrame)
                ShouldCheck := true
        } else {
            if (shp.HasTextFrame && shp.TextFrame.HasText)
                ShouldCheck := true
        }

        if (ShouldCheck) {
            try {
                if (PPT_IsFontMatch(shp.TextFrame.TextRange.Font, targetName)) {
                    return shp
                }
            }
        }
    }
    return false
}

; --- 공통 헬퍼 ---
PPT_GetSelectionIndex(activeWindow, shapes) {
    try {
        sel := activeWindow.Selection
        if (sel.Type = 2 || sel.Type = 3) { 
            if (sel.ShapeRange.Count > 0) {
                targetShape := sel.ShapeRange.Item(1)
                try {
                    if (targetShape.ParentGroup) {
                        targetShape := targetShape.ParentGroup
                    }
                }
                targetID := targetShape.Id
                Loop, % shapes.Count {
                    if (shapes.Item(A_Index).Id = targetID) {
                        return A_Index
                    }
                }
            }
        }
    }
    return 0
}

PPT_IsFontMatch(fontObj, needle) {
    try {
        ; 텍스트가 없어도 Font 객체는 존재하므로 이름 비교 가능
        if (InStr(fontObj.Name, needle) 
         || InStr(fontObj.NameFarEast, needle) 
         || InStr(fontObj.NameAscii, needle) 
         || InStr(fontObj.NameComplexScript, needle)) {
            return true
        }
    }
    return false
}




return
;사용된 폰트명으로 슬라이드 찾아가기 끝









; 14/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위










; 긴작업이 끝나면 깜빡이기 =======================================


작업끝표시(깜빡속도, 깜빡배경칼라) {
;sleep, 1
Gui, 깜빡: +AlwaysOnTop -Caption
Gui, 깜빡:Color, %깜빡배경칼라%
Gui, 깜빡:Show, W%A_ScreenWidth% H%A_ScreenHeight%
Sleep, %깜빡속도%
Gui,깜빡: Hide
;sleep, 1
Gui,깜빡: Destroy
;sleep, 1
}




정보창툴팁없애기:
    SetTimer, 정보창툴팁없애기, Off ; 타이머 끄기
;sleep, 1
    ToolTip ; 파라미터 없이 호출하면 툴팁이 즉시 사라짐
;sleep, 1
return



핫키올림확인:

loop
{
    ; Ctrl, Shift, Alt, Win 키들과 더불어 '휠 클릭(MButton)'도 누르지 않은 상태여야 break
    If (!GetKeyState("Control", "P") && !GetKeyState("Shift", "P") && !GetKeyState("Alt", "P") && !GetKeyState("Win", "P") && !GetKeyState("LWin", "P") && !GetKeyState("RWin", "P") && !GetKeyState("MButton", "P"))
    {
        break
    }
    Sleep, 1
}


return








키보드올리기:

Send, {RCtrl up}
Send, {Ctrl up}
Send, {RShift up}
Send, {Shift up}
Send, {RAlt up}
Send, {Alt up}

return






; ==============================================================================
; 슬라이드 마스터 가이드라인 관리 (복사 / 추가 / 삭제 분리형)
; ==============================================================================

; [1] 안내선 복사하기 (Ctrl + win + alt + C)
;$^#!c::
Action_Uni0351:

    ppt := ComObjActive("PowerPoint.Application")

    Try {
        currMaster := ppt.ActiveWindow.View.Slide
        
        GuideData := "PPT_GUIDE_DATA|"
        GuideCount := currMaster.Guides.Count
        
        If (GuideCount = 0) {
            mousegetpos, xx1, yy1
            ToolTip, ★ Nothing Guides lines (0), %xx1%, %yy1%
            SetTimer, 정보창툴팁없애기, 1000
            return
        }

        Loop, %GuideCount% {
            g := currMaster.Guides.Item(A_Index)
            GuideData .= g.Orientation . ":" . g.Position . "|"
        }
        
        Clipboard := GuideData
        
        mousegetpos, xx1, yy1
        ToolTip, ★ Copy! %GuideCount% Guides, %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000
    }
    Catch e {




        currMaster := ppt.ActivePresentation
        
        GuideData := "PPT_GUIDE_DATA|"
        GuideCount := currMaster.Guides.Count
        
        If (GuideCount = 0) {
            mousegetpos, xx1, yy1
            ToolTip, ★ Nothing Guides lines (0), %xx1%, %yy1%
            SetTimer, 정보창툴팁없애기, 1000
            return
        }

        Loop, %GuideCount% {
            g := currMaster.Guides.Item(A_Index)
            GuideData .= g.Orientation . ":" . g.Position . "|"
        }
        
        Clipboard := GuideData
        
        mousegetpos, xx1, yy1
        ToolTip, ★ Copy! %GuideCount% Guides, %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000





        MsgBox, 262208, Message, %msgboxuni_0039%`n%e%
    }
return


; [2] 안내선 붙여넣기 - 추가 모드 (Ctrl + win + alt + V)
; ★ 기존 안내선을 지우지 않고 그 위에 추가합니다.
;$^#!v::
Action_Uni0352:

    IfInString, Clipboard, PPT_GUIDE_DATA|
    {
        ; 데이터 있음
    }
    else
    {
        MsgBox, 262208, Message, %msgboxuni_0040%
        return
    }

    ppt := ComObjActive("PowerPoint.Application")

    Try {
        currMaster := ppt.ActiveWindow.View.Slide
        
        ; ★★★ [변경점] 기존 안내선 삭제 루프를 제거했습니다. ★★★
        
        RawData := Clipboard
        StringReplace, RawData, RawData, PPT_GUIDE_DATA|,, All
        
        Loop, Parse, RawData, |
        {
            if (A_LoopField = "")
                continue
                
            StringSplit, GuideInfo, A_LoopField, :
            
            Orientation := GuideInfo1
            Position := GuideInfo2
            
            currMaster.Guides.Add(Orientation, Position)
        }

        mousegetpos, xx1, yy1
        ToolTip, ★ Paste Guides!, %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000
    }
    Catch e {



        currMaster := ppt.ActivePresentation
        
        ; ★★★ [변경점] 기존 안내선 삭제 루프를 제거했습니다. ★★★
        
        RawData := Clipboard
        StringReplace, RawData, RawData, PPT_GUIDE_DATA|,, All
        
        Loop, Parse, RawData, |
        {
            if (A_LoopField = "")
                continue
                
            StringSplit, GuideInfo, A_LoopField, :
            
            Orientation := GuideInfo1
            Position := GuideInfo2
            
            currMaster.Guides.Add(Orientation, Position)
        }

        mousegetpos, xx1, yy1
        ToolTip, ★ Paste Guides!, %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000


        MsgBox, 262208, Message, %msgboxuni_0041%`n%e%
    }
return


; [3] 안내선 전체 삭제 (Ctrl + win + alt + 0)
; ★ 현재 페이지의 안내선만 깔끔하게 지웁니다.
;$^#!0::
Action_Uni0353:
$^#!Numpad0::
    ppt := ComObjActive("PowerPoint.Application")

    Try {
        currMaster := ppt.ActiveWindow.View.Slide
        
        GuideCount := currMaster.Guides.Count
        
        If (GuideCount = 0) {
            mousegetpos, xx1, yy1
            ToolTip, ★ Nothing guides to delete, %xx1%, %yy1%
            SetTimer, 정보창툴팁없애기, 1000
            return
        }

        ; 기존 안내선 모두 삭제 (인덱스 오류 방지를 위해 항상 1번 아이템을 반복 삭제)
        Loop, %GuideCount% {
            currMaster.Guides.Item(1).Delete()
        }

        mousegetpos, xx1, yy1
        ToolTip, ★ Delete Guides, %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000
    }
    Catch e {

            pres := ppt.ActivePresentation
            GuideCount := pres.Guides.Count
            guideTarget := pres.Guides


        if (GuideCount = 0) {
            MouseGetPos, xx1, yy1
            ToolTip, ★ Nothing guides to delete, %xx1%, %yy1%
            SetTimer, 정보창툴팁없애기, 1000
            return
        }

        Loop, %GuideCount% {
            guideTarget.Item(1).Delete()
        }



        mousegetpos, xx1, yy1
        ToolTip, ★ Delete Guides, %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000


       MsgBox, 262208, Message, %msgboxuni_0042%`n%e%
    }
return












; ==============================================================================
; 객체 고스트 모드 (마스터로 보내기 & 숨기기) 및 복구
; ==============================================================================

; [1] 고스트 모드 실행 (Ctrl + 2)
; 선택 객체를 마스터(레이아웃)에 50% 투명하게 붙여넣고, 원본은 숨김
;$^2::
Action_Uni0341:

    ppt := ComObjActive("PowerPoint.Application")

    Try {
        ; 1. 선택 확인
        selType := ppt.ActiveWindow.Selection.Type
        If (selType = 0 || selType = 1) {
            mousegetpos, xx1, yy1
            ToolTip, ★ Select Object, %xx1%, %yy1%
            SetTimer, 정보창툴팁없애기, 1000
            return
        }
        
        ; 2. 현재 슬라이드 및 마스터 레이아웃 가져오기
        ; (일반 화면에서만 작동하도록 유도)
        If (ppt.ActiveWindow.ViewType != 1 && ppt.ActiveWindow.ViewType != 9) {
             ; 1=Slide view, 9=Normal view
             MsgBox, 262208, Message, %msgboxuni_0043%
             return
        }

        currSlide := ppt.ActiveWindow.View.Slide
        targetLayout := currSlide.CustomLayout ; 현재 슬라이드의 배경(마스터)

        ; 3. 선택 객체 복사
        ppt.ActiveWindow.Selection.ShapeRange.Copy()
        
        ; 4. 마스터 레이아웃에 붙여넣기
        ; Paste는 ShapeRange를 반환함
        pastedShapes := targetLayout.Shapes.Paste()
        
        ; 5. 붙여넣은 객체 속성 변경 (이름 태그 & 투명도 50%)
        Loop, % pastedShapes.Count {
            shp := pastedShapes.Item(A_Index)
            
            ; (1) 나중에 지우기 위해 이름표 붙이기
            shp.Name := "PPT_GHOST_TEMP" 
            
            ; (2) 투명도 50% 적용 (0.5)
            Try shp.Fill.Transparency := 0.5
            Try shp.Line.Transparency := 0.5
            Try shp.PictureFormat.Transparency := 0.5
            Try shp.TextFrame.TextRange.Font.Fill.Transparency := 0.5
            
            ; 그룹 내부 처리
            If (shp.Type = 6) { ; msoGroup
                Loop, % shp.GroupItems.Count {
                    subShp := shp.GroupItems(A_Index)
                    Try subShp.Fill.Transparency := 0.5
                    Try subShp.Line.Transparency := 0.5
                    Try subShp.PictureFormat.Transparency := 0.5
                    Try subShp.TextFrame.TextRange.Font.Fill.Transparency := 0.5
                }
            }
        }
        
        ; 6. 원본 객체 숨기기 (선택된 상태 유지 중이므로 바로 적용)
        hideCount := ppt.ActiveWindow.Selection.ShapeRange.Count
        ppt.ActiveWindow.Selection.ShapeRange.Visible := 0 ; 숨김

        ; 7. 선택 해제 (깔끔하게)
        ppt.ActiveWindow.Selection.Unselect

        mousegetpos, xx1, yy1
        ToolTip, ★ Hide (Transparency 50`%) Lock, %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000
    }
    Catch e {
        MsgBox, 262208, Message, %msgboxuni_0044%`n%e%
    }
return








; [2] 복구 & 청소 (Ctrl + Alt + 2)
; 마스터에 붙였던 임시 객체 삭제 + 현재 슬라이드 모두 표시
;$^!2::
Action_Uni0342:
    Try {
    ppt := ComObjActive("PowerPoint.Application")
        }
    Try {
        ; 1. 현재 슬라이드 및 마스터 레이아웃 확인
        If (ppt.ActiveWindow.ViewType != 1 && ppt.ActiveWindow.ViewType != 9) {
             MsgBox, 262208, Message, %msgboxuni_0043%
             return
        }
        
        currSlide := ppt.ActiveWindow.View.Slide
        targetLayout := currSlide.CustomLayout

        ; 2. 마스터 레이아웃 청소 (PPT_GHOST_TEMP 삭제)
        ; [수정] 인덱스 계산 방식 -> "이름으로 찾아서 없을 때까지 삭제" (가장 안전함)
        delCount := 0
        Loop {
            Try {
                ; "PPT_GHOST_TEMP"라는 이름의 도형을 찾아 삭제 시도
                targetLayout.Shapes.Item("PPT_GHOST_TEMP").Delete()
                delCount++
            }
            Catch {
                ; 더 이상 해당 이름의 도형이 없으면 에러가 나므로 루프 종료
                Break
            }
        }

        ; 3. 현재 슬라이드 모든 객체 표시
        ; (위에서 에러가 나도 Catch로 빠지지 않고 Break로 나오므로 이 부분이 반드시 실행됨)
        showCount := currSlide.Shapes.Count
        Loop, %showCount% {
            currSlide.Shapes.Item(A_Index).Visible := -1 ; 보이기 (msoTrue)
        }

        mousegetpos, xx1, yy1
        ToolTip, ★ Show All (Unlock) , %xx1%, %yy1%
        SetTimer, 정보창툴팁없애기, 1000
    }
    Catch e {
        MsgBox, 262208, Message, %msgboxuni_0044%`n%e%
    }
return












#if
; ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
; 파워포인트가 실행되어있을때만 실행하도록하는 최상위 if 마무리 스크립트 끝=======================












; ==============================================================================
; [함수] PPT 텍스트 입력 상태 감지 (오타 방지 핵심)
; ==============================================================================
IsPptTextEditing() {
    try {
        ppt := ComObjActive("PowerPoint.Application")
        ; Selection.Type 3 = ppSelectionText (텍스트 상자 내 커서 활성 상태)
        if (ppt.ActiveWindow.Selection.Type = 3) {
            return 1
        }
    } catch {
        return 0
    }
    return 0
}

; ==============================================================================
; 1. 스페이스바를 누르면 커서 변경 (PPT 활성 & 텍스트 편집 모드가 아닐 때만)
; ==============================================================================
#If WinActive("ahk_exe POWERPNT.EXE") && !IsPptTextEditing()

$Space::
    ; [최적화 핵심] 루프 없이 '기본 화살표(32654)'만 콕 집어 변경 (속도 향상)
    ; ChangeSingleCursor(32654) 
    ; 아래쪽에서 클릭하면 변하도록 수정

    DidScroll := false     ; 스크롤 여부 변수 초기화
    
    KeyWait, Space         ; 스페이스바를 뗄 때까지 대기
    
    RestoreCursors()       ; 커서 복구
    
    ; 드래그(스크롤)를 안 하고 그냥 뗐다면 -> 로직 수행
    if (DidScroll = false) {

        ; 도형 입력 모드 (F2 상황)
        if (도형입력f2눌림 = 1) {
            SendInput, {Alt down}{h}{Alt up}
           ; Sleep, 1
            SendInput, {s}
           ; Sleep, 1
            SendInput, {h}
           ; Sleep, 1
            도형입력f2눌림 := 0
            return
        }

        ; 도형 수정 모드 (F4 상황)
        if (도형수정f4눌림 = 1) {
            SendInput, {Alt down}{j}{d}{Alt up}
           ; Sleep, 1
            SendInput, {e}
           ; Sleep, 1
            SendInput, {n}
           ; Sleep, 1
            도형수정f4눌림 := 0
            return
        }

        ; 아무 모드도 아니면 공백 입력
        SendInput, {Space}
    }
return

#If ; 블록 닫기



; ==============================================================================
; 2. [핵심 수정] PPT 활성 AND 스페이스바 눌림 AND 텍스트 편집 아님 -> 클릭 스크롤
; ==============================================================================
#If WinActive("ahk_exe POWERPNT.EXE") && GetKeyState("Space", "P") && !IsPptTextEditing()

F4::
도형입력f2눌림=1
return

F2::
도형수정f4눌림=1
return

*LButton::
    ChangeSingleCursor(32654) 

    DidScroll := true ; "나 스크롤 했다"고 표시

    ; 초기 마우스 위치 저장
    MouseGetPos, OldX, OldY
    
    ; 이동 거리를 저장할 변수 초기화
    AccumX := 0
    AccumY := 0
    
    ; ★ 속도 조절
    SpeedFactor := 23.5

    ; 드래그 루프 시작
    Loop
    {
        ; 왼쪽 버튼을 떼면 루프 종료
        If !GetKeyState("LButton", "P")
            Break

        ; 현재 마우스 위치 확인
        MouseGetPos, NewX, NewY
        
        DeltaX := OldX - NewX
        DeltaY := OldY - NewY
        
        ; 거리를 누적 변수에 더함
        AccumX += DeltaX
        AccumY += DeltaY
        
        ; 이전 위치 갱신
        OldX := NewX
        OldY := NewY
        
        ; --- 가로 스크롤 처리 ---
        While (AccumX >= SpeedFactor) {
            SendInput, {WheelRight}
            AccumX -= SpeedFactor
        }
        While (AccumX <= -SpeedFactor) {
            SendInput, {WheelLeft}
            AccumX += SpeedFactor
        }
        
        ; --- 세로 스크롤 처리 ---
        While (AccumY >= SpeedFactor) {
            SendInput, {WheelDown}
            AccumY -= SpeedFactor
        }
        While (AccumY <= -SpeedFactor) {
            SendInput, {WheelUp}
            AccumY += SpeedFactor
        }
        
        Sleep, 1 ; 부드러움을 위해 딜레이
    }
return

#If ; 조건문 끝





; ==============================================================================
; [함수] 단일 커서 변경 (루프 제거 버전)
; ==============================================================================
ChangeSingleCursor(CursorID = "", cx = 0, cy = 0)
{
    VarSetCapacity(AndMask, 128, 0xFF), VarSetCapacity(XorMask, 128, 0)
    SystemCursors := "32512,32513,32514,32515,32516,32642,32643,32644,32645,32646,32648,32649,32650,32651"
    If (CursorID = "") {
        CursorHandle := DllCall("CreateCursor", "Uint", 0, "Int", 0, "Int", 0, "Int", 32, "Int", 32, "Uint", &AndMask, "Uint", &XorMask)
    } Else {
        CursorHandle := DllCall("LoadImage", "Uint", 0, "Uint", CursorID, "Uint", 2, "Int", cx, "Int", cy, "Uint", 0x8000)
    }
    Loop, Parse, SystemCursors, `,
    {
        DllCall("SetSystemCursor", "Uint", DllCall("CopyIcon", "Uint", CursorHandle), "Int", A_LoopField)
    }
}

; ==============================================================================
; [함수] 커서 복구
; ==============================================================================
RestoreCursors()
{
    ; SPI_SETCURSORS (0x0057)를 호출하여 시스템 설정을 리로드합니다.
    DllCall("SystemParametersInfo", "UInt", 0x0057, "UInt", 0, "UInt", 0, "UInt", 0)
}






























;관리자모드 시작
if (UserGrade >= 4)
{



; 일러스트 활성화 되었을때만 아래의 단축키를 적용하는 스크립트 시작 ========================================
#IfWinActive ahk_exe Illustrator.exe




$^!+7::
send, ^!{7}
return

$^F9::
RunAiAction("스트록따로", "유니세프단축키")
return

$^F8::
RunAiAction("스트록같이", "유니세프단축키")
return

/*
$F7::
RunAiAction("앞으로한칸", "유니세프단축키")
return
*/


$+F7::
RunAiAction("제일앞으로", "유니세프단축키")
return

/*
$F8::
RunAiAction("뒤로한칸", "유니세프단축키")
return
*/


$+F8::
RunAiAction("제일뒤로", "유니세프단축키")
return

$+F10::
RunAiAction("가로간격", "유니세프단축키")
return

$+F12::
RunAiAction("세로간격", "유니세프단축키")
return


$^F12::
RunAiAction("수직중앙", "유니세프단축키")
return


$^F10::
RunAiAction("수평중앙", "유니세프단축키")
return





$F9::
RunAiAction("상단정렬", "유니세프단축키")
return

$F11::
RunAiAction("좌측정렬", "유니세프단축키")
return

$F10::
RunAiAction("하단정렬", "유니세프단축키")
return

$F12::
RunAiAction("우측정렬", "유니세프단축키")
return




~MButton & F9::
RunAiAction("텍스트상단정렬", "유니세프단축키")
return

~MButton & F11::
RunAiAction("텍스트좌측정렬", "유니세프단축키")
return


~MButton & F10::
;★★★★★★★★★★★★★★★★★★★★ 마우스버튼 체크기능이 나중으로가야 인식이 잘됨
    If GetKeyState("Ctrl", "P")  ; 컨트롤
    {
RunAiAction("텍스트수평중앙", "유니세프단축키")
return
    }
RunAiAction("텍스트하단정렬", "유니세프단축키")
return


~MButton & F12::
;★★★★★★★★★★★★★★★★★★★★ 마우스버튼 체크기능이 나중으로가야 인식이 잘됨
    If GetKeyState("Ctrl", "P")  ; 컨트롤
    {
RunAiAction("텍스트수직중앙", "유니세프단축키")
return
    }
RunAiAction("텍스트우측정렬", "유니세프단축키")
return



;끼워넣기
$!w::
RunAiAction("상단정렬", "유니세프단축키")
RunAiAction("좌측정렬", "유니세프단축키")
RunAiAction("클리핑마스크", "유니세프단축키")
return








$!+up::
;글자크게
send, ^+.
sleep, 1
return


$!+down::
;글자작게
send, ^+,
sleep, 1
return



$^!left::
;자간좁게
send, !{left}
sleep, 1
return


$^!right::
;자간넓게
send, !{right}
sleep, 1
return



$^!up::
;행간좁게
send, !{up}
sleep, 1
return



$^!down::
;행간넓게
send, !{down}
sleep, 1
return







; -----------------------------------------------------------
; 일러스트레이터 액션 실행 함수 (수정됨)
; -----------------------------------------------------------
RunAiAction(ActionName, SetName) {
    try {
        illu := ComObjActive("Illustrator.Application")
        
        ; DoScript("액션이름", "세트이름")
        illu.DoScript(ActionName, SetName)
    }
    catch e {
        MsgBox, 262208, Message, %msgboxuni_0045%'%SetName% / %ActionName%'
    }
}










;면적계산

$F4::
gosub, 한글체크
sleep, 10

send, !{f}
sleep, 10
send, {r}
sleep, 10
send, {enter}

sleep, 50


; 바탕화면 경로 가져오기
desktopPath := A_Desktop

; 파일 경로 설정
filePath := desktopPath . "\면적계산.txt"


    Sleep, 300 



; 파일이 존재하는 경우 내용 복사 및 삭제 실행
if FileExist(filePath) {
    ; 파일을 읽기 모드로 열기
    FileRead, fileContent, %filePath%
    
    ; 파일 내용이 클립보드로 복사됨
    Clipboard := fileContent


ClipWait, 2
if ErrorLevel
{
    FileDelete, %filePath%
    MsgBox, 262208, Message, %msgboxuni_0022%
    return
}


    ; 알림 표시
    ;MsgBox, 262208, Message, %msgboxuni_0046%
    
    ; 파일 삭제
    FileDelete, %filePath%
    FileDelete, %filePath%


}


return




#If
; 일러스트 활성화 되었을때만 아래의 단축키를 적용하는 스크립트 시작 ========================================


}
;관리자모드 끝










; 한글이 활성화 되었을때만 아래의 단축키를 적용하는 스크립트 시작 ========================================
#IfWinActive ahk_exe Hwp.exe
;ahk_class HwndWrapper[Hwp.exe;;9e0d3140-6311-4cbe-bccf-684dd76a8e86]


$^!Down::

send, +!z

return



$^!Up::

send, +!a

return


$^!Right::

send, +!w

return



$^!Left::

send, +!n

return


$+!Up::

send, !+e

return



$+!Down::

send, !+r

return




$F11::
send, ^+L
return

$F12::
send, ^+r
return

$^F10::
send, ^+C
return




#if
;한글 끝










한글체크:

;현재 일러스트 스크립트에서 활용하고있음
; CapsLock 키보드끄기
SetCapsLockState, off

; 한글로변환 1
; 영어로변환 0
if (IME_CHECK("a") = 1) ;만약 한글이면 시프트스페이스 누르기
{
send, {Shift down}
sleep, 2
send, {space down}
sleep, 2
send, {space up}
sleep, 2
send, {Shift up}
sleep, 200
;{VK15} ;한영키누르기
}
else
{
;;
}
return





;;한영체크 두번째 필요한것
IME_CHECK(WinTitle) 
{
WinGet,hWnd,ID,%WinTitle% 
Return Send_ImeControl(ImmGetDefaultIMEWnd(hWnd),0x005,"") 
}

Send_ImeControl(DefaultIMEWnd, wParam, lParam) 
{
DetectSave := A_DetectHiddenWindows 
DetectHiddenWindows,ON 
SendMessage 0x283, wParam,lParam,,ahk_id %DefaultIMEWnd% 
if (DetectSave <> A_DetectHiddenWindows) 
DetectHiddenWindows,%DetectSave% 
return ErrorLevel 
} 

ImmGetDefaultIMEWnd(hWnd) 
{ 
return DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hWnd, Uint) 
}
return













; 리로드로 화면보호 체크 또는 스크립트 재실행================
$^!+del::
    reload
return







; 스크립트 개발용 메모장 열기 ================
$^+!END::

    Run, "C:\Windows\notepad.exe" %A_ScriptFullPath%
    sleep, 1000
    WinActivate, ahk_exe Notepad.exe
    sleep, 20
return
















$^!+home::







    try {
        ppt := ComObjActive("PowerPoint.Application")
        sel := ppt.ActiveWindow.Selection
        
        윈도우뷰타입 := ppt.ActiveWindow.ViewType
        객체확인타입 := sel.Type
        
        if (객체확인타입 = 0) {
            MsgBox, 262208, Message, %msgboxuni_0016%
            return
        }

        if (객체확인타입 = 2 || 객체확인타입 = 3) {
            sr := sel.ShapeRange
            객체확인 := sr.Type
            선택갯수 := sr.Count
            shape := sr.Item(1)
            
            ; --- 기본 속성 (괄호 제거 및 줄바꿈 적용) ---
            try
                Id := shape.Id 
            catch 
                Id := "N/A" 
                
            try
                Name := shape.Name 
            catch 
                Name := "N/A" 
                
            try
                Visible := shape.Visible 
            catch 
                Visible := "N/A" 
                
            try
                HasTable := shape.HasTable 
            catch 
                HasTable := "N/A" 
                
            try
                HasTextFrame := shape.HasTextFrame 
            catch 
                HasTextFrame := "N/A" 
                
            try
                Connector := shape.Connector 
            catch 
                Connector := "N/A" 
                
            try
                AlternativeText := shape.AlternativeText 
            catch 
                AlternativeText := "N/A" 
            
            ; --- 도형/라인 속성 ---
            try {
                currentStyle := shape.Line.DashStyle
                도형선두께 := shape.Line.Weight
            } catch {
                currentStyle := "N/A"
                도형선두께 := "N/A"
            }
            
            ; --- 텍스트 폰트 라인 속성 ---
            if (HasTextFrame = -1) {
                try { 
                    fontLine := shape.TextFrame2.TextRange.Font.Line
                    선두께 := fontLine.Weight
                    선색상 := fontLine.ForeColor.RGB
                    투명도 := fontLine.Transparency
                    활성도 := fontLine.Visible
                } catch {
                    선두께 := "N/A"
                    선색상 := "N/A"
                    투명도 := "N/A"
                    활성도 := "N/A"
                }
            } else {
                선두께 := "N/A"
                선색상 := "N/A"
                투명도 := "N/A"
                활성도 := "N/A"
            }

            ; --- 테이블 셀 테두리 원시 값 추출 ---
            테이블테두리정보 := ""
            if (HasTable = -1) {
                try {
                    tbl := shape.Table
                    foundCell := false
                    
                    Loop, % tbl.Rows.Count {
                        r := A_Index
                        Loop, % tbl.Columns.Count {
                            c := A_Index
                            if (tbl.Cell(r, c).Selected) {
                                if (!foundCell) {
                                    foundCell := true
                                    cell := tbl.Cell(r, c)
                                    테이블테두리정보 .= "; 기준 셀: [" r "행, " c "열]`n"
                                    
                                    ; 1:Top, 2:Left, 3:Bottom, 4:Right
                                    borderIds := [1, 2, 3, 4]
                                    
                                    for idx, bId in borderIds {
                                        try {
                                            b := cell.Borders.Item(bId)
                                            
                                            try 
                                                bVis := b.Visible 
                                            catch 
                                                bVis := "N/A" 
                                                
                                            try 
                                                bWei := b.Weight 
                                            catch 
                                                bWei := "N/A" 
                                                
                                            try 
                                                bRGB := b.ForeColor.RGB 
                                            catch 
                                                bRGB := "N/A" 
                                            
                                            테이블테두리정보 .= "cell.Borders.Item(" bId ").Visible = " bVis "`n"
                                            테이블테두리정보 .= "cell.Borders.Item(" bId ").Weight = " bWei "`n"
                                            테이블테두리정보 .= "cell.Borders.Item(" bId ").ForeColor.RGB = " bRGB "`n`n"
                                            
                                        } catch {
                                            테이블테두리정보 .= "cell.Borders.Item(" bId ") = N/A`n`n"
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (!foundCell) {
                        테이블테두리정보 := "Selected Cell = N/A`n"
                    }
                } catch {
                    테이블테두리정보 := "Table Properties = N/A`n"
                }
            } else {
                테이블테두리정보 := "HasTable != -1`n"
            }
            
        } else {
            객체확인 := "N/A"
            선택갯수 := "N/A"
            Id := "N/A"
            Name := "N/A"
            Visible := "N/A"
            HasTable := "N/A"
            HasTextFrame := "N/A"
            Connector := "N/A"
            AlternativeText := "N/A"
            currentStyle := "N/A"
            도형선두께 := "N/A"
            선두께 := "N/A"
            선색상 := "N/A"
            투명도 := "N/A"
            활성도 := "N/A"
            테이블테두리정보 := "ShapeRange = N/A`n"
        }

        ; 3. MsgBox 원시 형태(Raw) 출력
        MsgBox, % "shape.Line.DashStyle = " currentStyle "`n"
        . "shape.Line.Weight = " 도형선두께 "`n`n"
        . "ppt.ActiveWindow.ViewType = " 윈도우뷰타입 "`n"
        . "sel.Type = " 객체확인타입 "`n"
        . "sr.Type = " 객체확인 "`n"
        . "sr.Count = " 선택갯수 "`n`n"
        . "shape.Id = " Id "`n"
        . "shape.Name = " Name "`n"
        . "shape.HasTable = " HasTable "`n"
        . "shape.HasTextFrame = " HasTextFrame "`n"
        . "shape.Visible = " Visible "`n"
        . "shape.Connector = " Connector "`n"
        . "shape.AlternativeText = " AlternativeText "`n`n"
        . "shape.TextFrame2.TextRange.Font.Line.Weight = " 선두께 "`n"
        . "shape.TextFrame2.TextRange.Font.Line.ForeColor.RGB = " 선색상 "`n"
        . "shape.TextFrame2.TextRange.Font.Line.Transparency = " 투명도 "`n"
        . "shape.TextFrame2.TextRange.Font.Line.Visible = " 활성도 "`n`n"
        . 테이블테두리정보

    } catch e {
        MsgBox, 262208, Message, Exception Thrown
    }
return
return




































; ------------------------------------------------------------------
; Esc 키 또는 X 버튼으로 GUI 닫기
; (GUI 이름이 'VBA입력'이므로 'VBA입력GuiEscape:' 라벨 필요)
; ------------------------------------------------------------------





PPT_FinderGuiClose:
PPT_FinderGuiEscape:
Gui, PPT_Finder:Destroy
return



처음GuiEscape:
처음GuiClose:
Gui, 처음:Destroy
return


대칭GuiEscape:
대칭GuiClose:
기존회전값 := ""
기존기준점1 := ""
기존기준점2 := ""
기존기준점3 := ""
기존기준점4 := ""
Gui, 대칭:Destroy
return


VBA입력GuiEscape:
VBA입력GuiClose:
Gui, VBA입력:Destroy
return


FontSetGuiEscape:
FontSetGuiClose:
Gui, FontSet:Destroy
return








$^!+f4::
pause
return






$^home::
Action_Uni0441:

MouseGetPos, MouseX, MouseY
PixelGetColor, color, %MouseX%, %MouseY%

Blue:="0x" SubStr(color,3,2) ;substr is to get the piece
Blue:=Blue+0 ;add 0 is to convert it to the current number format
Green:="0x" SubStr(color,5,2)
Green:=Green+0
Red:="0x" SubStr(color,7,2)
Red:=Red+0

msgbox, x : %MouseX%,  y : %MouseY%, color : %color%`n`nRed : %Red%, Green : %Green%, Blue : %Blue%
clipboard = %color%

return






#If WinActive("Google Gemini - Chrome")

;텍스트를 선택하고나서 사이트로 이동
$^PrintScreen::
send, {AppsKey}
sleep, 30
send, {down 3}
sleep, 10
send, {enter}
sleep, 100

return


#if








; 16/10 ◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆버전 2 작업중 위










;================================행간조절함수 시작====
;================================행간조절함수 시작====


; ===================================================================================
; [기능 1용] 재귀적 객체 처리 함수 (ProcessShape_Set)
; (참고: 이 함수는 현재 ^!up, ^!down 핫키에서는 호출되지 않습니다)
; ===================================================================================
ProcessShape_Set(oShape, msoTable, msoGroup, msoTrue)
{
    If (oShape.Type = msoGroup) ; 유형: 그룹 [4]
    {
        For oSubShape in oShape.GroupItems
        {
            ProcessShape_Set(oSubShape, msoTable, msoGroup, msoTrue)
        }
    }
    Else If (oShape.Type = msoTable) ; 유형: 표 [4]
    {
        oTable := oShape.Table
        Loop, % oTable.Rows.Count
        {
            r := A_Index
            Loop, % oTable.Columns.Count
            {
                c := A_Index
                Try
                {
                    oCellShape := oTable.Cell(r, c).Shape
                    ApplyLineSpacing_Set(oCellShape, msoTrue)
                }
                Catch
                {
                    Continue
                }
            }
        }
    }
    Else ; 유형: 일반 도형
    {
        ApplyLineSpacing_Set(oShape, msoTrue)
    }
}

; ===================================================================================
; [기능 1용] 서식 적용 헬퍼 함수 (ApplyLineSpacing_Set)
; ===================================================================================
ApplyLineSpacing_Set(oTargetShape, msoTrue)
{
    If (oTargetShape.HasTextFrame && oTargetShape.TextFrame.HasText)
    {
        Try
        {
            oParaFormat := oTargetShape.TextFrame.TextRange.ParagraphFormat
            oParaFormat.LineRuleWithin := msoTrue ; '배수'로 설정 [6, 7, 8]
            oParaFormat.SpaceWithin := 1.12     ; '1.01'로 고정 (사용자 설정값) [9, 7, 8, 10]
        }
        Catch
        {
        }
    }
}

; ===================================================================================
; [기능 2, 3용] 재귀적 객체 처리 함수 (ProcessShape_Adjust)
; ===================================================================================
ProcessShape_Adjust(oShape, msoTable, msoGroup, msoTrue, adjustment)
{
    If (oShape.Type = msoGroup) ; 유형: 그룹 [4]
    {
        For oSubShape in oShape.GroupItems
        {
            ProcessShape_Adjust(oSubShape, msoTable, msoGroup, msoTrue, adjustment)
        }
    }
    Else If (oShape.Type = msoTable) ; 유형: 표 [4]
    {
        oTable := oShape.Table
        Loop, % oTable.Rows.Count
        {
            r := A_Index
            Loop, % oTable.Columns.Count
            {
                c := A_Index
                Try
                {
                    oCell := oTable.Cell(r, c)
                    
                    ; 해당 셀이 '선택된' 상태인지 확인 [11]
                    If (oCell.Selected)
                    {
                        oCellShape := oCell.Shape
                        ApplyLineSpacing_Adjust(oCellShape, msoTrue, adjustment)
                    }
                }
                Catch
                {
                    Continue
                }
            }
        }
    }
    Else ; 유형: 일반 도형
    {
        ApplyLineSpacing_Adjust(oShape, msoTrue, adjustment)
    }
}

; ===================================================================================
; [기능 2, 3용] 서식 적용 헬퍼 함수 (ApplyLineSpacing_Adjust)
; [수정] "N/A" 문제를 해결하기 위해 로직 변경
; ===================================================================================
ApplyLineSpacing_Adjust(oTargetShape, msoTrue, adjustment)
{
    ; [추가] AdjustSpacing_SharedLogic에서 선언된 전역 변수를 사용
    global g_CurrentSpacing 
    
    If (oTargetShape.HasTextFrame && oTargetShape.TextFrame.HasText)
    {
        Try
        {
            oParaFormat := oTargetShape.TextFrame.TextRange.ParagraphFormat
            
            ; --- [수정된 핵심 로직] ---
            
            ; 1. 현재 '배수' 설정인지 확인 [6]
            If (oParaFormat.LineRuleWithin = msoTrue)
            {
                ; [Case 1: '배수'가 맞음] -> 기존 값에서 조정
                currentSpacing := oParaFormat.SpaceWithin
                newSpacing := currentSpacing + adjustment
            }
            Else
            {
                ; [Case 2: '배수'가 아님 (예: 단일, 고정)]
                ; -> '배수'로 강제 변경하고, 기본값 1.0 (단일)을 기준으로 조정 시작
                oParaFormat.LineRuleWithin := msoTrue
                currentSpacing := 1.0 
                newSpacing := currentSpacing + adjustment
            }
            
            ; 3. (안전 장치) 값이 너무 작아지지 않도록 최소값 0.3로 제한 (사용자 설정값)
            If (newSpacing < 0.3)
                newSpacing := 0.3
            
            ; 4. [수정] 툴팁에 표시할 값을 전역 변수에 저장
            ; 여러 객체 선택 시, 마지막으로 적용된 값이 표시됨
            g_CurrentSpacing := Round(newSpacing, 2)
            
            ; 5. 새 값 적용
            oParaFormat.SpaceWithin := newSpacing
        }
        Catch
        {
            ; COM 오류 발생 시
             g_CurrentSpacing := "Error"
        }
    }
}

;================================행간조절함수 끝====
;================================행간조절함수 끝====



















;색상계열 순서바꾸기
BGRtoRGB(color) {
    if (StrLen(color) != 8) ; 0x 포함 8자리가 아니면 그대로 반환 (오류 방지)
        return color
    Hex := SubStr(color, 3)
    return "0x" . SubStr(Hex, 5, 2) . SubStr(Hex, 3, 2) . SubStr(Hex, 1, 2)
}






;색상계열 순서바꾸기 (테이블라인쪽)
HexToBGR(hexCode) {
    ; 1. 빈 값이 들어오면 기본값 0(검은색) 반환 (에러 방지)
    if (hexCode = "")
        return 0

    ; 2. # 기호 제거 (v1 하위 호환성 및 안정성을 위해 StringReplace 사용)
    StringReplace, hexCode, hexCode, #, , All
    
    ; 3. 색상 코드가 정상적인 6자리인지 확인 (아니면 0 반환)
    if (StrLen(hexCode) != 6)
        return 0
    
    ; 4. R, G, B 두 자리씩 분리
    R := SubStr(hexCode, 1, 2)
    G := SubStr(hexCode, 3, 2)
    B := SubStr(hexCode, 5, 2)
    
    ; 5. B, G, R 순서로 조합하여 "0xC07000" 형태 생성
    hexString := "0x" . B . G . R
    
    ; 6. 숫자형 강제 변환 (파워포인트 COM 객체 호환성 극대화)
    ; v1 문법에 맞게 소수점 없는 정수형으로 확실하게 못을 박아줍니다.
    SetFormat, IntegerFast, d
    return hexString + 0
}











Key-Setting:
;Gui, VBA입력:Destroy

^!+F9::


; 2. 파일 전체 내용을 읽어옵니다.
FileRead, IniContent, %IniPath%

; 3. 각 섹션에서 '폰트 세트 이름'을 추출합니다. (Key= 형식 찾기)
본문_세트명리스트 := GetKeysFromSection(IniContent, "Body_Font_Set")
타이틀_세트명리스트 := GetKeysFromSection(IniContent, "Headings_Font_Set")

; 4. 이전에 저장된 선택값이 있다면 불러옵니다
IniRead, SavedBodySet, %IniPath%, GUI_Settings, LastBodySet, %A_Space%
IniRead, SavedTitleSet, %IniPath%, GUI_Settings, LastTitleSet, %A_Space%

; 드롭다운 리스트에서 미리 선택될 항목 인덱스 찾기
본문세트_인덱스 := 1
타이틀세트_인덱스 := 1

Loop, Parse, 본문_세트명리스트, |
{
    if (A_LoopField = SavedBodySet)
        본문세트_인덱스 := A_Index
}
Loop, Parse, 타이틀_세트명리스트, |
{
    if (A_LoopField = SavedTitleSet)
        타이틀세트_인덱스 := A_Index
}










노트북단독_포토룸로그인확인칼라박스 := "c" . BGRtoRGB(노트북단독_포토룸로그인확인칼라)
노트북단독_RemoveBG로그인확인칼라박스 := "c" . BGRtoRGB(노트북단독_RemoveBG로그인확인칼라)
집컴서브모니터_포토룸로그인확인칼라박스 := "c" . BGRtoRGB(집컴서브모니터_포토룸로그인확인칼라)
집컴서브모니터_RemoveBG로그인확인칼라박스 := "c" . BGRtoRGB(집컴서브모니터_RemoveBG로그인확인칼라)
모든사용자_포토룸로그인확인칼라박스 := "c" . BGRtoRGB(모든사용자_포토룸로그인확인칼라)
모든사용자_RemoveBG로그인확인칼라박스 := "c" . BGRtoRGB(모든사용자_RemoveBG로그인확인칼라)








Gui, FontSet:Destroy
; ==============================================================================
; 메인 GUI 구성
; ==============================================================================
;Gui, FontSet:New, +AlwaysOnTop +ToolWindow, %F9wintitlePreferences%
Gui, FontSet:New, +AlwaysOnTop, %F9wintitlePreferences%
Gui, FontSet:Font, s9, 맑은 고딕


Gui, FontSet:Add, Text, w350 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; [다국어 드롭다운 리스트]
Gui, FontSet:Add, Text, W170 Section, %F9Deslanguage%
Gui, FontSet:Add, DropDownList, x+4 yp-4 w150 vSelectedLang gChangeLang, %LangList%
Gui, FontSet:Add, Text, xs w350 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


; ▶ 본문 폰트 세트
Gui, FontSet:Add, Text, Section, %F9Bodyfontset%
Gui, FontSet:Add, DropDownList, xs w200 v본문세트 Choose%본문세트_인덱스% r10, %본문_세트명리스트%
Gui, FontSet:Add, Button, x+5 yp-1 w120 h25 gEditBodyList, Edit Font List

; ▶ 타이틀 폰트 세트
Gui, FontSet:Add, Text, xs y+10, %F9titlefontset%
Gui, FontSet:Add, DropDownList, xs w200 v타이틀세트 Choose%타이틀세트_인덱스% r10, %타이틀_세트명리스트%
Gui, FontSet:Add, Button, x+5 yp-1 w120 h25 gEditTitleList, Edit Font List

Gui, FontSet:Add, Text, xs w350 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -





    ; ▶ 테이블라인설정 (색상은 16진수 색상값)
    Gui, FontSet:Add, Text, xs, %F9TableBoarder% (Use Hex Color Codes)
    ; 기본 색상


Gui, FontSet:Add, Text, Section %기본테이블라인색상박스% w12 h11 Center, ■
    Gui, FontSet:Add, Text, x+5 w93, %F9Defaultcolor%

    Gui, FontSet:Add, Edit, x+1 yp-6 w60 v기본테이블라인색상, %기본테이블라인색상%

    Gui, FontSet:Add, Text, x+11 yp+6 w93, %F9DefaultWeight%

    Gui, FontSet:Add, Edit, x+3 yp-6 w35 v기본테이블라인두께, %기본테이블라인두께%

    Gui, FontSet:Add, Text, x+1 yp+6 w15, pt

Gui, FontSet:Add, Text, xs w350 h5 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Gui, FontSet:Add, Text, Section %포인트테이블라인색상1박스% xs w11 h12 Center, ■
    Gui, FontSet:Add, Text, x+5 w93, %F9accentcolor1%

    Gui, FontSet:Add, Edit, x+1 yp-6 w60 v포인트테이블라인색상1, %포인트테이블라인색상1%

    Gui, FontSet:Add, Text, x+11 yp+6 w93, %F9accentweight1%

    Gui, FontSet:Add, Edit, x+3 yp-6 w35 v포인트라인두께1, %포인트라인두께1%

    Gui, FontSet:Add, Text, x+1 yp+6 w15, pt

Gui, FontSet:Add, Text, xs w350 h5 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Gui, FontSet:Add, Text, Section %포인트테이블라인색상2박스% xs w11 h12 Center, ■
    Gui, FontSet:Add, Text, x+5 w93, %F9accentcolor2%

    Gui, FontSet:Add, Edit, x+1 yp-6 w60 v포인트테이블라인색상2, %포인트테이블라인색상2%

    Gui, FontSet:Add, Text, x+11 yp+6 w93, %F9accentweight2%

    Gui, FontSet:Add, Edit, x+3 yp-6 w35 v포인트라인두께2, %포인트라인두께2%

    Gui, FontSet:Add, Text, x+1 yp+6 w15, pt






   Gui, FontSet:Font, s9, 맑은 고딕

    Gui, FontSet:Add, Text, xs w350 h15 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

    ; ▶ 추가 리본메뉴 설정
    Gui, FontSet:Add, Text, xs, %F9DesAddRibbon% (Uni-Key)
    Gui, FontSet:Add, Text, Section xs w80, %F9Addshape%
    Gui, FontSet:Add, Edit, x+5 yp-4 w30 v단축키도형병합영어, %단축키도형병합영어%
    Gui, FontSet:Add, Edit, x+5 w30 v단축키도형병합숫자, %단축키도형병합숫자%

    Gui, FontSet:Add, Text, Section xs w80, %F9Addcrop%
    Gui, FontSet:Add, Edit, x+5 yp-4 w30 v자르기영어, %자르기영어%
    Gui, FontSet:Add, Edit, x+5 w30 v자르기숫자, %자르기숫자%



    Gui, FontSet:Add, Button, x168 y359 w168 h49 g리본메뉴추가, %F9BtnAddRibbon%



Gui, FontSet:Add, Text, xs w350 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; ▶ 시작프로그램 등록
Gui, FontSet:Add, Text, Section, %F9DesWindowStart%
; -- 대체텍스트 & 체크박스(기본 체크) --
Gui, FontSet:Add, Button, w207 h30 Section gAddStartup, %F9BtnwindowStart%
Gui, FontSet:Add, Button, x+10 w107 h30 gRemoveStartup, %F9BtnStartDisable%



Gui, FontSet:Add, Text, xs w350 h10 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

; ▶ 대체텍스트
Gui, FontSet:Add, Text, Section, %F9DesAltText% (Mouseover image in PDF)
; -- 대체텍스트 & 체크박스(기본 체크) --
Gui, FontSet:Add, Edit, w157 v대체텍스트내용1 Section, %대체텍스트내용1%
Gui, FontSet:Add, Edit, x+10 w157 v대체텍스트내용2, %대체텍스트내용2%



;관리자모드 시작
if (UserGrade >=4)
{

    Gui, FontSet:Add, Text, xs w350 h20 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Gui, FontSet:Font, s8, 맑은 고딕

if (노트북단독=1)
{
Gui, FontSet:Add, Text, Section xs %노트북단독_포토룸로그인확인칼라박스% w12 h12 Center, ■
    Gui, FontSet:Add, Edit, x+5 yp-4 w65 v노트북단독_포토룸로그인확인칼라, %노트북단독_포토룸로그인확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v노트북단독_포토룸다운로드확인칼라, %노트북단독_포토룸다운로드확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v노트북단독_포토룸사람확인칼라, %노트북단독_포토룸사람확인칼라%
    Gui, FontSet:Add, Text, x+5 yp+5, 노트북_포토룸

Gui, FontSet:Add, Text, Section xs %노트북단독_RemoveBG로그인확인칼라박스% w12 h12 Center, ■
    Gui, FontSet:Add, Edit, x+5 yp-4 w65 v노트북단독_RemoveBG로그인확인칼라, %노트북단독_RemoveBG로그인확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v노트북단독_RemoveBG다운로드확인칼라, %노트북단독_RemoveBG다운로드확인칼라%
    Gui, FontSet:Add, Text, x+5 yp+5, 노트북_removebg
}


if (집컴서브모니터=1)
{
Gui, FontSet:Add, Text, Section xs %집컴서브모니터_포토룸로그인확인칼라박스% w12 h12 Center, ■
    Gui, FontSet:Add, Edit, x+5 yp-4 w65 v집컴서브모니터_포토룸로그인확인칼라, %집컴서브모니터_포토룸로그인확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v집컴서브모니터_포토룸다운로드확인칼라, %집컴서브모니터_포토룸다운로드확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v집컴서브모니터_포토룸사람확인칼라, %집컴서브모니터_포토룸사람확인칼라%
    Gui, FontSet:Add, Text, x+5 yp+5, 집컴터_포토룸

Gui, FontSet:Add, Text, Section xs %집컴서브모니터_RemoveBG로그인확인칼라박스% w12 h12 Center, ■
    Gui, FontSet:Add, Edit, x+5 yp-4 w65 v집컴서브모니터_RemoveBG로그인확인칼라, %집컴서브모니터_RemoveBG로그인확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v집컴서브모니터_RemoveBG다운로드확인칼라, %집컴서브모니터_RemoveBG다운로드확인칼라%
    Gui, FontSet:Add, Text, x+5 yp+5, 집컴터_removebg
}


if (모든사용자=1)
{
Gui, FontSet:Add, Text, Section xs %모든사용자_포토룸로그인확인칼라박스% w12 h12 Center, ■
    Gui, FontSet:Add, Edit, x+5 yp-4 w65 v모든사용자_포토룸로그인확인칼라, %모든사용자_포토룸로그인확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v모든사용자_포토룸다운로드확인칼라, %모든사용자_포토룸다운로드확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v모든사용자_포토룸사람확인칼라, %모든사용자_포토룸사람확인칼라%
    Gui, FontSet:Add, Text, x+5 yp+5, 모든사용자_포토룸

Gui, FontSet:Add, Text, Section xs %모든사용자_RemoveBG로그인확인칼라박스% w12 h12 Center, ■
    Gui, FontSet:Add, Edit, x+5 yp-4 w65 v모든사용자_RemoveBG로그인확인칼라, %모든사용자_RemoveBG로그인확인칼라%
    Gui, FontSet:Add, Edit, x+5 w65 v모든사용자_RemoveBG다운로드확인칼라, %모든사용자_RemoveBG다운로드확인칼라%
    Gui, FontSet:Add, Text, x+5 yp+5, 모든사용자_removebg
}

}
;관리자모드 끝


    Gui, FontSet:Add, Text, xs w350 h30 c0xC0C0C0, - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


   Gui, FontSet:Font, s10, 맑은 고딕

; 저장 버튼
Gui, FontSet:Add, Button, xs w328 h40 g폰트세트저장 default, %F9BtnSavePreferences%




Gui, FontSet:Show, w350, %F9wintitlePreferences%
return










;==============================================================================
; [서브 루틴] 시작프로그램 등록 (단축 아이콘 생성)
;==============================================================================
AddStartup:
Gui, FontSet:Destroy
    ; 1. 현재 파일명에서 확장자(.exe 또는 .ahk)를 제외한 순수 이름만 추출
    SplitPath, A_ScriptName, , , , FileNameNoExt
    
    ; 2. 윈도우 시작프로그램 폴더 경로에 만들 단축 아이콘 이름 지정
    ; A_Startup 변수는 로그인한 사용자의 시작프로그램 폴더 경로를 자동으로 찾습니다.
    ShortcutPath := A_Startup . "\" . FileNameNoExt . ".lnk"
    
    ; 3. 현재 파일의 전체 경로(A_ScriptFullPath)를 타겟으로 단축 아이콘 생성
    FileCreateShortcut, %A_ScriptFullPath%, %ShortcutPath%
    
    if (ErrorLevel) {
        MsgBox, 262208, Message, %msgboxuni_0055%
    } else {
        MsgBox, 262208, Message, %msgboxuni_0056%
    }
return

;==============================================================================
; [서브 루틴] 시작프로그램 해제 (단축 아이콘 삭제)
;==============================================================================
RemoveStartup:
Gui, FontSet:Destroy
    SplitPath, A_ScriptName, , , , FileNameNoExt
    ShortcutPath := A_Startup . "\" . FileNameNoExt . ".lnk"
    
    ; 해당 경로에 단축 아이콘이 존재하는지 확인 후 삭제
    if FileExist(ShortcutPath) {
        FileDelete, %ShortcutPath%
        MsgBox, 262208, Message, %msgboxuni_0057%
    } else {
        MsgBox, 262208, Message, %msgboxuni_0058%
    }
return





; ==============================================================================
; 동작 로직 (리스트 수정) - 헤더 제외 버전
; ==============================================================================

; [리스트 변경] - 본문 버튼 클릭 시
EditBodyList:
    TargetSectionName := "Body_Font_Set"
    Gosub, OpenEditorWindow
return

; [리스트 변경] - 타이틀 버튼 클릭 시
EditTitleList:
    TargetSectionName := "Headings_Font_Set"
    Gosub, OpenEditorWindow
return

; 공통 에디터 창 열기
OpenEditorWindow:
    ; 1. 원본 파일 읽기
    FileRead, FullContent, %IniPath%
    
    ; 2. 정규식을 사용하여 섹션 헤더([Section])를 제외한 내부 내용만 추출
    ; \K 옵션: 앞의 헤더 패턴은 찾되, 결과값에는 포함하지 않음
    Needle_Load := "ims`a)^\s*\[\Q" . TargetSectionName . "\E\]\R?\K.*?(?=(?:^\s*\[)|\z)"
    
    if RegExMatch(FullContent, Needle_Load, OnlyContent) {
        ; 헤더를 제외한 알맹이만 UI용 변수에 할당 (앞뒤 불필요 공백 제거)
        DisplayContent := Trim(OnlyContent, " `t`r`n")
    } else {
        ; 섹션이 아예 없거나 비어있는 경우
        DisplayContent := "" 
    }

    ; 수정 GUI 구성
    Gui, Editor:Destroy
    ; Gui, Editor:New, +OwnerFontSet  +ToolWindow, Edit ★ %TargetSectionName%
    Gui, Editor:New, +OwnerFontSet +AlwaysOnTop , Edit ★ %TargetSectionName%
    Gui, Editor:Font, s9 cGray, 맑은 고딕
    
    ; 좌측 가이드 문구 설정
    폰트작성설명=
    (

%F9DesEditFont1%

font name=                        ￣￣￣￣￣
(
font name1
font name2
font name3
font name4
font name5
font name6
font name7
font name8
font name9
`)                                       __________

General Sans=                     ￣￣￣￣￣
(
General Sans Light
General Sans Light
General Sans Light
General Sans Medium
General Sans Medium
General Sans Medium
General Sans Semibold
General Sans Semibold
General Sans Semibold
`)                                       __________

NotoSans=                         ￣￣￣￣￣
(
NotoSans-Thin
NotoSans-ExtraLight
NotoSans-Light
NotoSans-Regular
NotoSans-Medium
NotoSans-SemiBold
NotoSans-Bold
NotoSans-ExtraBold
NotoSans-Black      
`)                                       __________

)


    폰트작성설명2=
    (
%F9DesEditFont2%
%F9DesEditFont3%
    )


    ; UI 요소 배치
    Gui, Editor:Add, Text, y24, %폰트작성설명%
    Gui, Editor:Add, Text, x200 y30 w420 cBlue, %폰트작성설명2%
    
    ; [중요] 에디트 박스에 헤더가 제거된 %DisplayContent%를 직접 삽입
Gui, Editor:Font, s9 cBlack, 맑은 고딕
    Gui, Editor:Add, Edit, w400 h590 vEditedContent, %DisplayContent%
    Gui, Editor:Add, Button, w400 h40 gSaveAndReload default, Save Updated Font List
    
    Gui, Editor:Show, h710
return









; [수정된 리스트 저장] 버튼 클릭 시
SaveAndReload:
    Gui, Editor:Submit, NoHide
    
    ; ==============================================================================
    ; [사전 검증 1] '=' 앞의 폰트 이름에 금지된 특수문자가 있는지 검사
    ; ==============================================================================
    ; 금지 문자: ( ) , ; ` %
    ; 각 줄을 돌면서 '=' 앞부분만 떼어내어 검사합니다.
    Loop, Parse, EditedContent, `n, `r
    {
        SplitPos := InStr(A_LoopField, "=")
        if (SplitPos) 
        {
            FontNamePrefix := Trim(SubStr(A_LoopField, 1, SplitPos - 1))
            
            ; 금지 문자가 하나라도 포함되어 있는지 정규식으로 확인
            if RegExMatch(FontNamePrefix, "[\(\),\;`\%]") 
            {
                MsgBox, 262208, Message, %msgboxuni_0059%`n  (  )  `,  `;  ``  `% `n`n%FontNamePrefix%
                return ; 저장을 중단하고 원래 에디터 화면으로 되돌아감
            }
        }
    }

    ; ==============================================================================
    ; [사전 검증 2] '(' 바로 다음에는 반드시 폰트 이름이 9줄인지 & 빈 줄이 없는지 검사
    ; ==============================================================================
    ; "( 부터 ) 까지의 덩어리"를 모두 찾아냅니다.
    Pos := 1
    While RegExMatch(EditedContent, "\(\s*(.*?)\s*\)", MatchBlock, Pos)
    {
        ; 찾아낸 덩어리의 내용(괄호 안쪽 텍스트)
        BlockContent := Trim(MatchBlock1, " `t`r`n")
        
        LineCount := 0
        HasEmptyLine := false ; ★ 빈 줄 감지용 스위치
        
        Loop, Parse, BlockContent, `n, `r
        {
            if (Trim(A_LoopField) = "") 
            {
                ; 텍스트가 없는 빈 줄(엔터)이 감지됨
                HasEmptyLine := true 
            }
            else 
            {
                LineCount++
            }
        }
        
        ; ★ 1. 중간에 빈 줄(엔터)이 있는 경우 즉시 차단
        if (HasEmptyLine = true)
        {
            MsgBox, 262208, Message, %msgboxuni_0060%`n`n%BlockContent%
            return ; 저장을 중단
        }
        
        ; ★ 2. 정확히 9줄이 아니면 에러 발생
        if (LineCount != 9)
        {
            MsgBox, 262208, Message, %msgboxuni_0061%`n`n(%LineCount%)`n%BlockContent%
            return ; 저장을 중단
        }
        
        ; 다음 덩어리를 찾기 위해 검색 위치 이동
        Pos += StrLen(MatchBlock)
    }

    ; --- (여기서부터는 기존 저장 로직과 동일) ---

    Gui, Editor:Destroy
    Gui, FontSet:Destroy

    ; 3. [핵심] 파일 저장 전, 숨겼던 섹션 헤더를 다시 결합
    EditedContent_WithHeader := "[" . TargetSectionName . "]`r`n" . RTrim(EditedContent, " `t`r`n") . "`r`n`r`n"

    ; 원본 파일 다시 읽기 (교체 대상 확인용)
    FileRead, OldFullContent, %IniPath%
    
    ; 교체할 영역(헤더 포함 전체)을 찾는 정규식
    Needle_Save := "ims`a)^\s*\[\Q" . TargetSectionName . "\E\].*?(?=(?:^\s*\[)|\z)"
    
    if RegExMatch(OldFullContent, Needle_Save)
    {
        NewFullContent := RegExReplace(OldFullContent, Needle_Save, EditedContent_WithHeader)
    }
    else
    {
        NewFullContent := RTrim(OldFullContent, " `t`r`n") . "`r`n`r`n" . EditedContent_WithHeader
    }
    
    ; 파일 끝부분 불필요한 공백 최종 정리
    NewFullContent := RTrim(NewFullContent, " `t`r`n")
    
    ; 파일 저장 (UTF-16 LE BOM 규격 유지)
    FileDelete, %IniPath%
    FileAppend, %NewFullContent%, %IniPath%, UTF-16
    FileSetAttrib, +H, %IniPath%
    
    Reload
return






; ==============================================================================
; 동작 로직 (설정값 적용 및 저장)
; ==============================================================================
폰트세트저장:

    Gui, FontSet:Submit, NoHide

    FileRead, CurrentIniContent, %IniPath%

    ; A. 본문 폰트 처리 (1 ~ 9)
    BodyFontLines := GetMultiLineValue(CurrentIniContent, "Body_Font_Set", 본문세트)
    Loop, Parse, BodyFontLines, `n, `r
    {
        CurrentLine := Trim(A_LoopField)
        if (CurrentLine = "")
            continue
        if (A_Index > 9)
            break
        IniWrite, %CurrentLine%, %IniPath%, 본문폰트 기본설정, 한글폰트%A_Index%
        IniWrite, %CurrentLine%, %IniPath%, 본문폰트 기본설정, 영문폰트%A_Index%
    }

    ; B. 타이틀 폰트 처리 (11 ~ 99)
    TitleFontLines := GetMultiLineValue(CurrentIniContent, "Headings_Font_Set", 타이틀세트)
    Loop, Parse, TitleFontLines, `n, `r
    {
        CurrentLine := Trim(A_LoopField)
        if (CurrentLine = "")
            continue
        if (A_Index > 9)
            break
        Idx := A_Index * 11
        IniWrite, %CurrentLine%, %IniPath%, 타이틀본트 기본설정, 한글폰트%Idx%
        IniWrite, %CurrentLine%, %IniPath%, 타이틀본트 기본설정, 영문폰트%Idx%
    }
    
    ; 마지막 선택값 저장
    IniWrite, %본문세트%, %IniPath%, GUI_Settings, LastBodySet
    IniWrite, %타이틀세트%, %IniPath%, GUI_Settings, LastTitleSet




    ; ==============================================================================
    ; [단축키] 저장
    ; ==============================================================================
    IniWrite, %단축키도형병합영어%, %IniPath%, 단축키, 단축키도형병합영어
    IniWrite, %단축키도형병합숫자%, %IniPath%, 단축키, 단축키도형병합숫자
    IniWrite, %자르기영어%, %IniPath%, 단축키, 자르기영어
    IniWrite, %자르기숫자%, %IniPath%, 단축키, 자르기숫자


    ; ==============================================================================
    ; [디자인] 저장
    ; ==============================================================================
    IniWrite, %기본테이블라인색상%, %IniPath%, 디자인, 기본테이블라인색상
    IniWrite, %포인트테이블라인색상1%, %IniPath%, 디자인, 포인트테이블라인색상1
    IniWrite, %포인트테이블라인색상2%, %IniPath%, 디자인, 포인트테이블라인색상2

    IniWrite, %기본테이블라인두께%, %IniPath%, 디자인, 기본테이블라인두께
    IniWrite, %포인트라인두께1%, %IniPath%, 디자인, 포인트라인두께1
    IniWrite, %포인트라인두께2%, %IniPath%, 디자인, 포인트라인두께2



    ; ==============================================================================
    ; [이미지수정] 저장 (빈 값이면 빈 값으로 저장됨)
    ; ==============================================================================
    ; 노트북 단독
    IniWrite, %노트북단독_포토룸로그인확인칼라%, %IniPath%, 이미지수정, 노트북단독_포토룸로그인확인칼라
    IniWrite, %노트북단독_포토룸다운로드확인칼라%, %IniPath%, 이미지수정, 노트북단독_포토룸다운로드확인칼라
    IniWrite, %노트북단독_포토룸사람확인칼라%, %IniPath%, 이미지수정, 노트북단독_포토룸사람확인칼라
    IniWrite, %노트북단독_RemoveBG로그인확인칼라%, %IniPath%, 이미지수정, 노트북단독_RemoveBG로그인확인칼라
    IniWrite, %노트북단독_RemoveBG다운로드확인칼라%, %IniPath%, 이미지수정, 노트북단독_RemoveBG다운로드확인칼라

    ; 집컴 서브모니터
    IniWrite, %집컴서브모니터_포토룸로그인확인칼라%, %IniPath%, 이미지수정, 집컴서브모니터_포토룸로그인확인칼라
    IniWrite, %집컴서브모니터_포토룸다운로드확인칼라%, %IniPath%, 이미지수정, 집컴서브모니터_포토룸다운로드확인칼라
    IniWrite, %집컴서브모니터_포토룸사람확인칼라%, %IniPath%, 이미지수정, 집컴서브모니터_포토룸사람확인칼라
    IniWrite, %집컴서브모니터_RemoveBG로그인확인칼라%, %IniPath%, 이미지수정, 집컴서브모니터_RemoveBG로그인확인칼라
    IniWrite, %집컴서브모니터_RemoveBG다운로드확인칼라%, %IniPath%, 이미지수정, 집컴서브모니터_RemoveBG다운로드확인칼라

    ; 모든 사용자
    IniWrite, %모든사용자_포토룸로그인확인칼라%, %IniPath%, 이미지수정, 모든사용자_포토룸로그인확인칼라
    IniWrite, %모든사용자_포토룸다운로드확인칼라%, %IniPath%, 이미지수정, 모든사용자_포토룸다운로드확인칼라
    IniWrite, %모든사용자_포토룸사람확인칼라%, %IniPath%, 이미지수정, 모든사용자_포토룸사람확인칼라
    IniWrite, %모든사용자_RemoveBG로그인확인칼라%, %IniPath%, 이미지수정, 모든사용자_RemoveBG로그인확인칼라
    IniWrite, %모든사용자_RemoveBG다운로드확인칼라%, %IniPath%, 이미지수정, 모든사용자_RemoveBG다운로드확인칼라

    IniWrite, %대체텍스트내용1%, %IniFile%, 기본설정, 대체텍스트내용1
    IniWrite, %대체텍스트내용2%, %IniFile%, 기본설정, 대체텍스트내용2


Gui, Editor:Destroy
Gui, FontSet:Destroy
reload
; 설정값이 INI 파일에 저장되었습니다.`n`n[본문] 1~9`n[타이틀] 11~99
return




GuiClose:
GuiEscape:
ExitApp
return

EditorGuiClose:
EditorGuiEscape:
Gui, Editor:Destroy
return




; ==============================================================================
; 파싱 함수들
; ==============================================================================
GetKeysFromSection(Content, SectionName) {
    Keys := ""
    InTargetSection := false
    Loop, Parse, Content, `n, `r
    {
        Line := Trim(A_LoopField)
        if (RegExMatch(Line, "^\[(.*)\]$", match)) {
            if (match1 = SectionName)
                InTargetSection := true
            else
                InTargetSection := false
            continue
        }
        if (InTargetSection && RegExMatch(Line, "^(.*?)=$", match)) {
            if (Keys != "")
                Keys .= "|"
            Keys .= match1
        }
    }
    return Keys
}

GetMultiLineValue(Content, SectionName, KeyName) {
    ResultText := ""
    State := 0 
    Loop, Parse, Content, `n, `r
    {
        Line := Trim(A_LoopField)
        if (State = 0) {
            if (Line = "[" . SectionName . "]")
                State := 1
        }
        else if (State = 1) {
            if (RegExMatch(Line, "^\[.*\]$"))
                return "" 
            if (Line = KeyName . "=")
                State := 2
        }
        else if (State = 2) {
            if (Line = "(")
                State := 3
        }
        else if (State = 3) {
            if (Line = ")")
                break 
            if (ResultText != "")
                ResultText .= "`n"
            ResultText .= Line
        }
    }
    return ResultText
}










로그아웃:
토큰초기화=
    IniWrite, %토큰초기화%, %IniPath%, 인증정보, 토큰
RestoreCursors()       ; 커서 복구
exitapp
return



; ===================================================================================
; ★★★★★★★★★★ 특수문자 입력 [단축키] Ctrl + 넘버패드 별표* 또는 Ctrl + 8 (*)
; ===================================================================================

특수문자입력창:

$^NumpadMult::
$^8::
    Gui_Draw()
return



; ===================================================================================
; [GUI 생성 함수]
; ===================================================================================
Gui_Draw() {
    ; 1. 기존 창 제거

    Gui, MyCharGui:Destroy
    
    ; 2. GUI 기본 설정
    ;Gui, MyCharGui:New, +AlwaysOnTop +ToolWindow, Symbols (★)
    Gui, MyCharGui:New, +AlwaysOnTop, Symbols (★)
    Gui, MyCharGui:Color, White
    Gui, MyCharGui:Font, s10, 맑은 고딕 
    
    ; =========================================================
    ; [데이터 로드] INI 파일에서 불러오기
    ; =========================================================
    IniFile := "UNI-Value.ini"
    IniSection := "특수문자"
    IniKey := "Data"

    ; (1) INI 읽기
    IniRead, RawData, %IniFile%, %IniSection%, %IniKey%, %A_Space%

    ; (2) 데이터가 아예 없을 경우를 대비한 기본값
    if (RawData = "" || RawData = "ERROR") {
        RawData = 
        (LTrim Join||
·,∙,◦,▪︎,▫︎,⁎,▸,▹,☉,※,✳,✱,
➔,➜,➠,➥,➧,➲,➾,↣,↠,⭆,↪,⇥
《,》,【,】,〔,〕,﹛,﹜,❬,❭,❮,❯,❰,❱
√,✓,✔,✖︎,✗,✘,☑,☒,🆇,☆,★,◆
Ⅰ,Ⅱ,Ⅲ,Ⅳ,Ⅴ,Ⅵ,Ⅶ,Ⅷ,Ⅸ,Ⅹ,Ⅺ
ⅰ,ⅱ,ⅲ,ⅳ,ⅴ,ⅵ,ⅶ,ⅷ,ⅸ,ⅹ
⓪,①,②,③,④,⑤,⑥,⑦,⑧,⑨,⑩
⓿,❶,❷,❸,❹,❺,❻,❼,❽,❾,❿
⬅,⬆,⬇,⮕,⬈,⬉,⬊,⬋,⬌,⬍
⭠,⭡,⭢,⭣,⭤,⭥,⮂,⮃,↤,↥,↦,↧
⇕,⇖,⇗,⇘,⇙,⇚,⇛,⇠,⇡,⇢,⇣
⇦,⇧,⇨,⇩,⇪,⏎
㎧,°,℃,㎜,㎝,㎞,㎟,㎠,㎡,㎢
㎣,㎤,㎥,㎦,㎖,㎘,_,₩,$,¥
        )
    }







    ; (3) 복구 ('||' -> 줄바꿈)
    StringReplace, FinalData, RawData, ||, `n, All

    ; =========================================================
    ; [화면 배치]
    ; =========================================================
    Gui, MyCharGui:Add, Text, x10 y10 w0 h0 Section, 
    Create_Buttons(FinalData)

; ▼▼▼▼▼ [버튼 디자인 수정: 검정 배경 + 흰색 글씨] ▼▼▼▼▼
    
    ; 1. 검정 배경 만들기 (Progress 컨트롤 활용)
    ;    Disabled: 클릭 안 되게(뒤에 깔림), Background000000: 검정색
    Gui, MyCharGui:Add, Progress, x10 y+20 w130 h30 Disabled Background000000
    
    ; 2. 흰색 글씨 및 클릭 기능 입히기 (Text 컨트롤)
    ;    xp yp: 바로 위 컨트롤(검정 배경)과 같은 위치에 겹침
    ;    BackgroundTrans: 배경 투명, cWhite: 흰색 글씨, 0x200: 텍스트 세로 가운데 정렬
    Gui, MyCharGui:Font, s10 bold cWhite ; 폰트 설정 (굵게, 흰색)
    Gui, MyCharGui:Add, Text, xp yp w130 h30 BackgroundTrans 0x200 center gOpenCharEditor, Edit Symbols (★)
    Gui, MyCharGui:Font ; 폰트 설정 초기화 (이후 컨트롤에 영향 안 주게)

    ; 3. 오른쪽에 '닫기' 버튼 (기존 유지)
    ;    xp+140: 겹쳐진 텍스트(xp)에서 너비(130)+간격(10) 만큼 이동
    ;    yp: 같은 높이
    Gui, MyCharGui:Add, Button, xp+160 yp w130 h30 gMyCharGuiGuiClose, Close (Esc)
    
    ; ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    Gui, MyCharGui:Show, AutoSize Center
    return
}

; ===================================================================================
; [버튼 생성 로직]
; ===================================================================================
Create_Buttons(StringData) {
    if (StringData = "")
        return

    Loop, Parse, StringData, `n, `r 
    {
        CurrentLine := A_LoopField
        if (Trim(CurrentLine) = "")
            continue
            
        Loop, Parse, CurrentLine, % ",", %A_Space%%A_Tab% 
        {
            Symbol := A_LoopField
            if (Symbol = "")
                continue
                
            DisplayText := (StrLen(Symbol) = 0) ? "Err" : Symbol

            if (A_Index = 1)
                Gui, MyCharGui:Add, Button, Section xs y+9 w30 h30 gInsertChar, %DisplayText%
            else
                Gui, MyCharGui:Add, Button, ys w30 h30 gInsertChar, %DisplayText%
        }
    }
}

; ===================================================================================
; [동작] 입력 및 복사
; ===================================================================================
InsertChar:
    Gui, MyCharGui:Submit, NoHide
    GuiControlGet, ClickText,, %A_GuiControl%
    
    if (ClickText = "Err")
        return

    Clipboard := ClickText
    Gui, MyCharGui:Destroy
    Gui, 대칭:Destroy
    sleep, 10
    SendInput, %ClickText%    
return

MyCharGuiGuiClose:
MyCharGuiGuiEscape:
    Gui, MyCharGui:Destroy
return


; ==============================================================================
; ▼▼▼ [편집기 실행 라벨] (요청하신 코드를 여기에 연결했습니다) ▼▼▼
; ==============================================================================
OpenCharEditor:
    ; 기존 입력창은 닫지 않고 위에 띄웁니다 (필요시 Gui, MyCharGui:Destroy 추가 가능)

   ; Gui, CharEditor:New, +AlwaysOnTop +ToolWindow, Customize Symbol (★)
    Gui, CharEditor:New, +AlwaysOnTop, Customize Symbol (★)
    Gui, CharEditor:Color, White
    Gui, CharEditor:Font, s10, 맑은 고딕

    ; --- 1. 상단 도움말 ---
    Gui, CharEditor:Add, Text, x15 y10 w400 cBlue, → ★, ◆, ★ %F0Dessymbols1%
    Gui, CharEditor:Add, Text, x15 y+5 w400 cBlue, ⭣ %F0Dessymbols2%

    Gui, CharEditor:Add, Text, x15 y+10 w400 cblack, %F0Dessymbols3%





    Gui, CharEditor:Font, s9, 맑은 고딕

    ; --- 2. 데이터 로드 ---
    IniFile := "UNI-Value.ini"
    IniSection := "특수문자"
    IniKey := "Data"

    IniRead, SavedData, %IniFile%, %IniSection%, %IniKey%, %A_Space%
    StringReplace, DisplayData, SavedData, ||, `n, All

    ; --- 3. 에디트 박스 ---
    Gui, CharEditor:Font, s11, 맑은 고딕
    Gui, CharEditor:Add, Edit, x15 y+5 w450 h330 vMyEditBox +Multi +WantTab +VScroll, %DisplayData%

    ; --- 4. 저장 버튼 ---
    Gui, CharEditor:Font, s10, 맑은 고딕
    Gui, CharEditor:Add, Button, x15 y+15 w450 h40 gSaveIniData default, %F0Btnsymbolssave%

    Gui, CharEditor:Show, w480 h490, Customize Symbol (★)
return



CharEditorGuiClose:
CharEditorGuiEscape:
    Gui, CharEditor:Destroy
return


; ==============================================================================
; [저장 로직]
; ==============================================================================
SaveIniData:
    Gui, CharEditor:Submit, NoHide
    
    ; [수정된 로직] 줄바꿈을 단순 치환하기 전에, 각 줄의 끝 콤마(,)를 확인하고 제거합니다.
    SaveStr := ""
    Loop, Parse, MyEditBox, `n, `r 
    {
        ; RTrim 함수를 사용하여 현재 줄의 오른쪽 끝에 있는 '콤마(,)'와 '공백/탭'을 모두 잘라냄
        cleanLine := RTrim(A_LoopField, " ,`t")
        
        ; 첫 번째 줄이면 바로 넣고, 두 번째 줄부터는 앞에 '||' 구분자를 붙여서 연결
        if (A_Index = 1)
            SaveStr := cleanLine
        else
            SaveStr .= "||" . cleanLine
    }
    
    ; INI 파일 기록
    IniWrite, %SaveStr%, %IniFile%, %IniSection%, %IniKey%

    Gui, CharEditor:Destroy
    Gui, MyCharGui:Destroy

    MsgBox, 262208, Message, %msgboxuni_0062%
    
    ; (선택사항) 저장 후 입력창 바로 갱신하고 싶으면 아래 주석 해제
    ; Gui_Draw() 
return







리본메뉴추가:

    Gui, FontSet:Submit, NoHide
    Gui, FontSet:Destroy


; 1. 타겟 파일 경로 설정
targetFile := "C:\Users\" . A_UserName . "\AppData\Local\Microsoft\Office\PowerPoint.officeUI"

; 2. 삽입할 핵심 XML 코드 (Tabs 부분)
userXML = 
(
<mso:tabs>

<mso:tab idQ="mso:TabSlideMasterHome">

<mso:group idQ="x1:ACROBAT_SHARE" visible="false"/>

<mso:group idQ="mso:GroupOfficeExtensionsAddinFlyout" visible="false"/>

<mso:group id="mso_c1.5F32896" label="UNI-Key" autoScale="true">

<mso:gallery idQ="mso:CombineShapesGallery" showInRibbon="false" visible="true"/>
<mso:control idQ="mso:PictureCropTools" visible="true"/>


</mso:group>

</mso:tab>




<mso:tab idQ="mso:TabHome">

<mso:group idQ="x1:ACROBAT_SHARE" visible="false"/>

<mso:group idQ="mso:GroupOfficeExtensionsAddinFlyout" visible="false"/>

<mso:group id="mso_c1.74A44B5" label="UNI-Key" imageMso="AppointmentColorDialog" autoScale="true">

<mso:gallery idQ="mso:CombineShapesGallery" showInRibbon="false" visible="true"/>
<mso:control idQ="mso:PictureCropTools" visible="true"/>


</mso:group>

</mso:tab>


<mso:tab idQ="mso:HelpTab">

<mso:group idQ="mso:GroupHelpAndSupport" visible="false"/></mso:tab>

</mso:tabs>
)


; ==============================================================================
; 시나리오 A: 파일이 없는 경우 (새로 생성)
; ==============================================================================
If !FileExist(targetFile)
{
    ; 기본 뼈대(Header/Footer)를 포함하여 전체 내용을 작성
    ; QAT 태그는 비어있어도 상관없으며, 사용자의 Tab 설정을 Ribbon 안에 넣음
    newFileContent = 
    (
<mso:customUI xmlns:x1="PDFMaker.OfficeAddin" xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui">
    <mso:ribbon>
        <mso:qat/>
%userXML%
    </mso:ribbon>
</mso:customUI>
    )

    ; 파일 생성 및 쓰기 (UTF-8)
    FileAppend, %newFileContent%, %targetFile%, UTF-8
    
    If (ErrorLevel = 0)
        MsgBox, 262208, Message, %msgboxuni_0047%
    Else
        MsgBox, 262208, Message, %msgboxuni_0048%
    return
}

; ==============================================================================
; 시나리오 B: 파일이 이미 있는 경우 (내용 추가)
; ==============================================================================
FileRead, fileContent, %targetFile%

; 중복 방지 체크
If InStr(fileContent, "label=""UNI-Key""")
{
    MsgBox, 262208, Message, %msgboxuni_0049%
    return
}

; 닫는 태그(</mso:ribbon>)를 찾아서 그 앞에 삽입
searchString := "</mso:ribbon>"
replaceString := "`n" . userXML . "`n" . searchString

StringReplace, newContent, fileContent, %searchString%, %replaceString%

If (ErrorLevel != 0)
{
    MsgBox, 262208, Message, %msgboxuni_0050%
    return
}

; 기존 파일 백업 및 덮어쓰기
FileCopy, %targetFile%, %targetFile%.bak, 1
FileDelete, %targetFile%
FileAppend, %newContent%, %targetFile%, UTF-8

MsgBox, 262208, Message, %msgboxuni_0051%

return










$^#!f12::
run, ★0WindowSpy.ahk
return













/*
BlockKeyboardInputs(state = "On")
{
   static keys
   keys=Space,Enter,Tab,Esc,BackSpace,Del,Ins,Home,End,PgDn,PgUp,Up,Down,Left,Right,CtrlBreak,ScrollLock,PrintScreen,CapsLock
,Pause,AppsKey,LWin,LWin,NumLock,Numpad0,Numpad1,Numpad2,Numpad3,Numpad4,Numpad5,Numpad6,Numpad7,Numpad8,Numpad9,NumpadDot
,NumpadDiv,NumpadMult,NumpadAdd,NumpadSub,NumpadEnter,NumpadIns,NumpadEnd,NumpadDown,NumpadPgDn,NumpadLeft,NumpadClear
,NumpadRight,NumpadHome,NumpadUp,NumpadPgUp,NumpadDel,Media_Next,Media_Play_Pause,Media_Prev,Media_Stop,Volume_Down,Volume_Up
,Volume_Mute,Browser_Back,Browser_Favorites,Browser_Home,Browser_Refresh,Browser_Search,Browser_Stop,Launch_App1,Launch_App2
,Launch_Mail,Launch_Media,F1,F2,F3,F5,F6,F7,F8,F9,F10,F11,F12,F13,F14,F15,F16,F17,F18,F19,F20,F21,F22
,1,2,3,4,5,6,7,8,9,0,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z
,²,&,e,",',(,-,e,_,c,a,),=,$,￡,u,*,~,#,{,[,|,``,\,^,@,],},;,:,!,?,.,/,§,<,>,vkBC
,Ctrl,Alt,Shift,LControl,RControl,LAlt,RAlt,LShift,RShift
   Loop,Parse,keys, `,
      Hotkey, *%A_LoopField%, KeyboardDummyLabel, %state% UseErrorLevel
   Return
; 단축키에는 라벨이 필요하므로 아무 작업도 수행하지 않는 라벨
KeyboardDummyLabel:
Return
}


;스크립트 사용방법
BlockKeyboardInputs("On")  ; 입력 차단합니다.
BlockKeyboardInputs("Off")  ; 입력 차단을 해제합니다.
,LControl,RControl,LAlt,RAlt,LShift,RShift
*/




RemoveToolTip:
    ToolTip
return







; ==============================================================================
; [이벤트] 언어 변경 시 INI 업데이트 후 Reload 
; ==============================================================================

    
ChangeLang:
    Gui, FontSet:Submit, NoHide
    GuiControlGet, SelectedLang, FontSet:
    RegExMatch(SelectedLang, "\((\d+)\)", Match)
    
    if (Match1 != "")
    {
        IniWrite, %Match1%, %IniFile%, 시스템, LangID
        Reload
    }
return







;--- [2025-10-22 지침 반영: v1 문법] ---

핫키설정변경:

ChangeSingleCursor(32514) ; 커서를 '바쁨'으로 변경

    ; 중복 실행 방지를 위해 기존 GUI 창 파괴
    Gui, Main:Destroy
    Gui, Container:Destroy  ; 컨테이너 파괴
    Gui, Child:Destroy

    창가로크기 := 790
    창가로크기3 := 창가로크기 + 155
    버튼너비 := 창가로크기3 - 40 
    저장버튼너비 := 버튼너비 - 50
    리셋버튼x := 저장버튼너비 + 20
    
    ; 하단 버튼이 y=756에 위치하므로, 스크롤 영역의 높이를 700으로 맞춤
    ViewHeight := 700 

    ;Gui, Main: New, +AlwaysOnTop +ToolWindow +Resize +0x300000 +HwndhMain
    Gui, Main: New, +AlwaysOnTop +Resize +0x300000 +HwndhMain
    Gui, Main: Font, s9, 맑은 고딕

    ; --- 고정되는 상단 헤더 (Main GUI 소속) ---
    Gui, Main: Add, Text, x35 y20 w200 h20, | %HotkeyEditDes1%
    Gui, Main: Add, Text, x+265 y20 w180 h20, | %HotkeyEditDes2%
    Gui, Main: Add, Text, x+39 y20 w100 h20, | %HotkeyEditDes3%
    Gui, Main: Add, Text, x+53 y20 w100 h20,   + / -
    Gui, Main: Add, Text, x20 y50 w%버튼너비% h1 0x10

    Gui, Main: Font, s10 c215F9A, 맑은 고딕
    Gui, Main: Add, Button, x20 y756 w%저장버튼너비% h45 gBtnSave, %HotkeyBtnSave%
    
    ; --- 실행키 리셋 버튼 ---
    Gui, Main: Add, Progress, x%리셋버튼x% y756 w50 h45 BackgroundBlack Disabled
    Gui, Main: Font, s10 cWhite bold, 맑은 고딕
    Gui, Main: Add, Text, xp yp wp hp Center 0x200 BackgroundTrans g실행키리셋, Reset


    ; ==============================================================================
    ; ★ [핵심 최적화] INI 파일을 한 번만 읽어서 RAM(배열)에 통째로 저장 (디스크 I/O 최소화)
    ; ==============================================================================
    IniRead, RawHotkeys, %IniFile%, 단축키
    MemHotkeys := {}
    DynamicHotkeyList := []
    TempList := {}
    Loop, Parse, RawHotkeys, `n, `r
    {
        SplitPos := InStr(A_LoopField, "=")
        if (SplitPos) {
            key := Trim(SubStr(A_LoopField, 1, SplitPos-1))
            val := Trim(SubStr(A_LoopField, SplitPos+1))
            MemHotkeys[key] := val
            
            StringSplit, IdParts, key, _
            CurrentID := IdParts1
            if (CurrentID != "" && InStr(CurrentID, "Uni") = 1 && !TempList.HasKey(CurrentID)) {
                TempList[CurrentID] := true
                DynamicHotkeyList.Push(CurrentID)
            }
        }
    }

    IniRead, RawFixedKeys, %IniFile%, 고정키
    MemFixedKeys := {}
    FixedHotkeyList := []
    TempFixedList := {}
    Loop, Parse, RawFixedKeys, `n, `r
    {
        SplitPos := InStr(A_LoopField, "=")
        if (SplitPos) {
            key := Trim(SubStr(A_LoopField, 1, SplitPos-1))
            val := Trim(SubStr(A_LoopField, SplitPos+1))
            MemFixedKeys[key] := val
            
            StringSplit, IdParts, key, _
            CurrentID := IdParts1
            if (CurrentID != "" && InStr(CurrentID, "Uni") = 1 && !TempFixedList.HasKey(CurrentID)) {
                TempFixedList[CurrentID] := true
                FixedHotkeyList.Push(CurrentID)
            }
        }
    }

    IniRead, RawLang, %IniFile%, %LangSec%
    MemLang := {}
    Loop, Parse, RawLang, `n, `r
    {
        SplitPos := InStr(A_LoopField, "=")
        if (SplitPos)
            MemLang[Trim(SubStr(A_LoopField, 1, SplitPos-1))] := Trim(SubStr(A_LoopField, SplitPos+1))
    }
    ; ==============================================================================


    ;Gui, Container: New, +AlwaysOnTop +ToolWindow +ParentMain -Caption +HwndhContainer
    Gui, Container: New, +AlwaysOnTop +ParentMain -Caption +HwndhContainer
    ;Gui, Child: New, +AlwaysOnTop +ToolWindow +Parent%hContainer% -Caption +HwndhChild
    Gui, Child: New, +AlwaysOnTop +Parent%hContainer% -Caption +HwndhChild

    Gui, Child: Font, s9, 맑은 고딕

    CurrentY := 20

    ; ==============================================================================
    ; 1. [단축키] 섹션 - 동적 핫키 (편집 가능)
    ; ==============================================================================
    For Index, thisID in DynamicHotkeyList {
        ; 메모리(RAM)에서 즉시 값을 가져옴
        KeyDesc := MemLang[thisID "_Key설명"]
        
        if (KeyDesc != "") {
            StringReplace, KeyDesc, KeyDesc, \n, `n, All
            thisName := "[" . thisID . "] `n`n" . KeyDesc
        } else {
            thisName := "[" . thisID . "] 설명 지정 안 됨"
        }

        StringReplace, DummyVar, thisName, `n, `n, UseErrorLevel
        LineCount := ErrorLevel
        
        AddHeight := LineCount * 17
        TextH := 25 + AddHeight
        RowStep := 45 + AddHeight
        
        ; IniRead 대신 초고속 메모리 변수 호출
        v1 := MemHotkeys[thisID "_Key1"]
        v2 := MemHotkeys[thisID "_Key2"]
        v3 := MemHotkeys[thisID "_Key3"]
        v4 := MemHotkeys[thisID "_Key4"]
        v설정값 := MemHotkeys[thisID "_Key설정값"]
        
        c1 := GetChoice(ModList, v1)
        c2 := GetChoice(ModList, v2)
        c3 := GetChoice(ModList, v3)
        c4 := GetChoice(KeyList, v4)

        DropY := CurrentY - 4

        Gui, Child: Font, s9 cBlack norm, 맑은 고딕
        Gui, Child: Add, Text, x35 y%CurrentY% yp+14 w420 h%TextH%, %thisName%

        Gui, Child: Font, s10 cBlack norm, 맑은 고딕
        Gui, Child: Add, DropDownList, x500 y%DropY% yp-3 w50 vGui_%thisID%_1 Choose%c1%, %ModList%
        Gui, Child: Add, DropDownList, x+10 y%DropY% yp0 w50 vGui_%thisID%_2 Choose%c2%, %ModList%
        Gui, Child: Add, DropDownList, x+10 y%DropY% yp0 w50 vGui_%thisID%_3 Choose%c3%, %ModList%
        Gui, Child: Add, DropDownList, x+50 y%DropY% yp0 w90 vGui_%thisID%_4 Choose%c4%, %KeyList%
        
        if (MemHotkeys.HasKey(thisID "_Key설정값")) {
            HasKey설정값_%thisID% := true
            Gui, Child: Add, Edit, x+60 y%DropY% yp0 w40 h25 vGui_%thisID%_설정값 Center, %v설정값%
        }
        
        LineY := CurrentY + RowStep - 10
        Gui, Child: Add, Text, x20 y%LineY% w%버튼너비% h1 0x10
        CurrentY += RowStep
    }

    ; ==============================================================================
    ; 2. [고정키] 섹션 - 고정 핫키 (텍스트로만 표시)
    ; ==============================================================================
    if (FixedHotkeyList.Length() > 0) {
        CurrentY += 20 
        Gui, Child: Font, s12 c215F9A bold, 맑은 고딕
        Gui, Child: Add, Text, x35 y%CurrentY% w200 h25, | %HotkeyEditFixed%
        CurrentY += 30
        
        Gui, Child: Add, Progress, x20 y%CurrentY% w%버튼너비% h2 c215F9A Background215F9A -Smooth, 100
        CurrentY += 15
        
        For Index, thisID in FixedHotkeyList {
            KeyDesc := MemLang[thisID "_Key설명"]
            
            if (KeyDesc != "") {
                StringReplace, KeyDesc, KeyDesc, \n, `n, All
                thisName := "[" . thisID . "] `n`n" . KeyDesc
            } else {
                thisName := "[" . thisID . "] 설명 지정 안 됨"
            }
            
            StringReplace, DummyVar, thisName, `n, `n, UseErrorLevel
            LineCount := ErrorLevel
            
            AddHeight := LineCount * 17
            TextH := 25 + AddHeight
            RowStep := 45 + AddHeight
            
            ; IniRead 대신 초고속 메모리 변수 호출
            v1 := MemFixedKeys[thisID "_Key1"]
            v2 := MemFixedKeys[thisID "_Key2"]
            v3 := MemFixedKeys[thisID "_Key3"]
            v4 := MemFixedKeys[thisID "_Key4"]
            v설정값 := MemFixedKeys[thisID "_Key설정값"]

            DropY := CurrentY - 4

            Gui, Child: Font, s9 cBlack norm, 맑은 고딕
            Gui, Child: Add, Text, x35 y%CurrentY% yp+14 w420 h%TextH%, %thisName%
            
            Gui, Child: Font, s10 cBlack norm, 맑은 고딕
            Gui, Child: Add, Text, x500 y%DropY% yp0 w50 cRed, %v1%
            Gui, Child: Add, Text, x+10 y%DropY% yp0 w50 cRed, %v2%
            Gui, Child: Add, Text, x+10 y%DropY% yp0 w50 cRed, %v3%
            Gui, Child: Add, Text, x+50 y%DropY% yp0 w90 cRed, %v4%
            
            if (v설정값 != "") {
                Gui, Child: Add, Text, x+60 y%DropY% yp0 w40 cGray Center, %v설정값%
            }
            
            LineY := CurrentY + RowStep - 10
            Gui, Child: Add, Text, x20 y%LineY% w%버튼너비% h1 0x10
            CurrentY += RowStep
        }
    }

    Gui, Main: Show, w960 h800, %HotkeyEditDestitle%
    Gui, Container: Show, x0 y51 w960 h%ViewHeight%

    Global ChildHeight := CurrentY + 10 
    Gui, Child: Show, x0 y0 w960 h%ChildHeight%

    Global PageSize := ViewHeight, ScrollPos := 0
    OnMessage(0x115, "OnScroll")
    OnMessage(0x20A, "OnWheel")
    UpdateScrollbar()

RestoreCursors()       ; 커서 복구 완료

return









; ==============================================================================
; [마우스 휠 오작동 방지] 드롭다운 리스트 위에서 스크롤 휠 차단
; ==============================================================================
#If WinActive("ahk_class AutoHotkeyGUI") && IsHoveringDropDownList()
WheelUp::return
WheelDown::return
#If

IsHoveringDropDownList() {
    ; 현재 마우스가 위치한 곳의 컨트롤(버튼, 에디트, 드롭다운 등) 이름을 가져옵니다.
    MouseGetPos, , , , ctrlName
    
    ; 오토핫키에서 DropDownList의 윈도우 시스템 내부 클래스 이름은 "ComboBox" 입니다.
    ; 마우스가 올려진 컨트롤 이름이 "ComboBox"로 시작한다면 true를 반환하여 휠을 차단합니다.
    if (InStr(ctrlName, "ComboBox") = 1)
        return true
        
    return false
}






; ==============================================================================
; [3단계: 서브루틴 및 함수 영역] 로직 처리
; ==============================================================================

InitHotkeys:
    ; ---------------------------------------------------------
    ; ★ [핵심 최적화] INI 파일을 한 번만 읽어서 RAM(배열)에 통째로 저장
    ; ---------------------------------------------------------
    IniRead, RawHotkeys, %IniFile%, 단축키
    MemHotkeys := {}
    DynamicHotkeyList_Init := []
    TempList_Init := {}
    Loop, Parse, RawHotkeys, `n, `r
    {
        SplitPos := InStr(A_LoopField, "=")
        if (SplitPos) {
            key := Trim(SubStr(A_LoopField, 1, SplitPos-1))
            val := Trim(SubStr(A_LoopField, SplitPos+1))
            MemHotkeys[key] := val
            
            StringSplit, IdParts, key, _
            CurrentID := IdParts1
            if (CurrentID != "" && InStr(CurrentID, "Uni") = 1 && !TempList_Init.HasKey(CurrentID)) {
                TempList_Init[CurrentID] := true
                DynamicHotkeyList_Init.Push(CurrentID)
            }
        }
    }

    IniRead, RawFixedKeys, %IniFile%, 고정키
    MemFixedKeys := {}
    FixedHotkeyList_Init := []
    TempFixedList_Init := {}
    Loop, Parse, RawFixedKeys, `n, `r
    {
        SplitPos := InStr(A_LoopField, "=")
        if (SplitPos) {
            key := Trim(SubStr(A_LoopField, 1, SplitPos-1))
            val := Trim(SubStr(A_LoopField, SplitPos+1))
            MemFixedKeys[key] := val
            
            StringSplit, IdParts, key, _
            CurrentID := IdParts1
            if (CurrentID != "" && InStr(CurrentID, "Uni") = 1 && !TempFixedList_Init.HasKey(CurrentID)) {
                TempFixedList_Init[CurrentID] := true
                FixedHotkeyList_Init.Push(CurrentID)
            }
        }
    }
    ; ---------------------------------------------------------






; ==============================================================================
    ; ★ [핵심 추가] 지금부터 등록되는 핫키는 파워포인트가 활성화될 때만 작동하도록 선언
    ; ==============================================================================
    Hotkey, IfWinActive, ahk_class PPTFrameClass



    ; 1. [단축키] 동적 등록
    For Index, thisID in DynamicHotkeyList_Init {
        k1 := Trim(MemHotkeys[thisID "_Key1"])
        k2 := Trim(MemHotkeys[thisID "_Key2"])
        k3 := Trim(MemHotkeys[thisID "_Key3"])
        k4 := Trim(MemHotkeys[thisID "_Key4"])
        v설정값 := MemHotkeys[thisID "_Key설정값"]

        Saved_%thisID%_1 := k1, Saved_%thisID%_2 := k2
        Saved_%thisID%_3 := k3, Saved_%thisID%_4 := k4
        Saved_%thisID%_설정값 := v설정값

        if (k4 = "") 
            continue

        mods := ""
        if (k1="Ctrl" || k2="Ctrl" || k3="Ctrl")
            mods .= "^"
        if (k1="Alt" || k2="Alt" || k3="Alt")
            mods .= "!"
        if (k1="Shift" || k2="Shift" || k3="Shift")
            mods .= "+"
        if (k1="Win" || k2="Win" || k3="Win")
            mods .= "#"
        
        fullKey := mods . k4
        targetLabel := "Action_" . thisID
        
        if IsLabel(targetLabel) {
            Hotkey, $%fullKey%, %targetLabel%, On
        }
    }
    
    ; 2. [고정키] 동적 등록
    For Index, thisID in FixedHotkeyList_Init {
        k1 := Trim(MemFixedKeys[thisID "_Key1"])
        k2 := Trim(MemFixedKeys[thisID "_Key2"])
        k3 := Trim(MemFixedKeys[thisID "_Key3"])
        k4 := Trim(MemFixedKeys[thisID "_Key4"])
        v설정값 := MemFixedKeys[thisID "_Key설정값"]

        Saved_%thisID%_1 := k1, Saved_%thisID%_2 := k2
        Saved_%thisID%_3 := k3, Saved_%thisID%_4 := k4
        Saved_%thisID%_설정값 := v설정값

        if (k4 = "") 
            continue
            
        if (k4 = "⭢")
            k4 := "Right"
        else if (k4 = "⭠")
            k4 := "Left"
        else if (k4 = "⭡")
            k4 := "Up"
        else if (k4 = "⭣")
            k4 := "Down"

        mods := ""
        if (k1="Ctrl" || k2="Ctrl" || k3="Ctrl")
            mods .= "^"
        if (k1="Alt" || k2="Alt" || k3="Alt")
            mods .= "!"
        if (k1="Shift" || k2="Shift" || k3="Shift")
            mods .= "+"
        if (k1="Win" || k2="Win" || k3="Win")
            mods .= "#"
        
        fullKey := mods . k4
        targetLabel := "Action_" . thisID
        
        if IsLabel(targetLabel) {
            try {
                Hotkey, $%fullKey%, %targetLabel%, On
            } catch {
            }
        }
    }



; ==============================================================================
    ; ★ [안전장치 추가] 혹시 나중에 파워포인트 외에 작동해야 할 전역 단축키를
    ; 추가할 수도 있으니, 핫키 조건을 다시 '모든 창(조건 없음)'으로 초기화해 둡니다.
    ; ==============================================================================
    Hotkey, IfWinActive

return









GetChoice(list, val) {
    if (val = "" || val = " ") return 1
    Loop, Parse, list, |
    {
        if (A_LoopField = val)
            return A_Index
    }
    return 1
}

UpdateScroll(NewPos) {
    Global ChildHeight, PageSize, ScrollPos
    MaxPos := ChildHeight - PageSize
    
    if (MaxPos < 0) 
        MaxPos := 0
        
    NewPos := (NewPos < 0) ? 0 : (NewPos > MaxPos) ? MaxPos : NewPos
    
    if (ScrollPos != NewPos) {
        ScrollPos := NewPos
        TargetY := 0 - ScrollPos 
        Gui, Child: Show, x0 y%TargetY%
        UpdateScrollbar()
    }
}

UpdateScrollbar() {
    Global hMain, ChildHeight, PageSize, ScrollPos
    VarSetCapacity(si, 28, 0), NumPut(28, si, 0, "UInt"), NumPut(0x17, si, 4, "UInt")
    NumPut(0, si, 8, "Int"), NumPut(ChildHeight, si, 12, "Int")
    NumPut(PageSize, si, 16, "UInt"), NumPut(ScrollPos, si, 20, "Int")
    DllCall("SetScrollInfo", "Ptr", hMain, "Int", 1, "Ptr", &si, "Int", 1)
}

OnWheel(W, L, M, H) {
    Global ScrollPos
    Dir := (W >> 16 > 32768) ? 40 : -40
    UpdateScroll(ScrollPos + Dir)
}

OnScroll(W, L, M, H) {
    Global ScrollPos
    Action := W & 0xFFFF
    
    if (Action = 0) {
        UpdateScroll(ScrollPos - 10)
    } else if (Action = 1) {
        UpdateScroll(ScrollPos + 10)
    } else if (Action = 2) {
        UpdateScroll(ScrollPos - 50)
    } else if (Action = 3) {
        UpdateScroll(ScrollPos + 50)
    } else if (Action = 5) {
        VarSetCapacity(si, 28, 0), NumPut(28, si, 0, "UInt"), NumPut(0x10, si, 4, "UInt")
        DllCall("GetScrollInfo", "Ptr", H, "Int", 1, "Ptr", &si)
        UpdateScroll(NumGet(si, 24, "Int"))
    }
}




;--- [2025-10-22 지침 반영: v1 문법] ---

BtnSave:
    Gui, 대칭:Destroy
    Gui, VBA입력:Destroy

    Gui, Child: Submit, NoHide 
    
    ; ==============================================================================
    ; [신규 추가] 시스템 보호키(중복키추가방지) 목록 미리 불러오기
    ; ==============================================================================
    ReservedKeys := {}
    Loop, 100 ; 넉넉하게 최대 100개의 보호키 셋업을 확인합니다 (중간에 번호가 비어도 안전하게 건너뜀)
    {
        prefix := Format("{:03}", A_Index)
        IniRead, rK4, %IniFile%, 중복키추가방지, %prefix%실행키d, %A_Space%
        rK4 := Trim(rK4)
        
        ; 실행키가 없으면 이번 번호는 패스
        if (rK4 = "")
            continue

        IniRead, rK1, %IniFile%, 중복키추가방지, %prefix%보조키a, %A_Space%
        IniRead, rK2, %IniFile%, 중복키추가방지, %prefix%보조키b, %A_Space%
        IniRead, rK3, %IniFile%, 중복키추가방지, %prefix%보조키c, %A_Space%

        rK1 := Trim(rK1), rK2 := Trim(rK2), rK3 := Trim(rK3)

        ; 시스템 보호키도 기호(^, !, +, #)로 변환
        rMods := ""
        if (rK1="Ctrl" || rK2="Ctrl" || rK3="Ctrl")
            rMods .= "^"
        if (rK1="Alt" || rK2="Alt" || rK3="Alt")
            rMods .= "!"
        if (rK1="Shift" || rK2="Shift" || rK3="Shift")
            rMods .= "+"
        if (rK1="Win" || rK2="Win" || rK3="Win")
            rMods .= "#"

        ; ★ [추가된 부분] 충돌 시 보여줄 친절한 안내 문구 조합 (예: "Ctrl + Alt + 1")
        DisplayKey := ""
        if (rK1 != "")
            DisplayKey .= rK1 . " + "
        if (rK2 != "")
            DisplayKey .= rK2 . " + "
        if (rK3 != "")
            DisplayKey .= rK3 . " + "
        DisplayKey .= rK4

        ; 변환된 시스템 조합을 키(Key)로, 텍스트 조합을 값(Value)으로 저장
        ReservedKeys[rMods . rK4] := DisplayKey
    }
    ; ==============================================================================

    UsedKeys := {}
    DuplicateItems := "", SelfDuplicateItems := "", EmptyKeyItems := "", SystemDuplicateItems := ""
    
    For Index, thisID in DynamicHotkeyList {
        k1 := Trim(Gui_%thisID%_1), k2 := Trim(Gui_%thisID%_2)
        k3 := Trim(Gui_%thisID%_3), k4 := Trim(Gui_%thisID%_4)
        
        isModSelected := (k1 != "" || k2 != "" || k3 != "")
        isKeyEmpty := (k4 == "")
        
        if (isModSelected && isKeyEmpty) {
            EmptyKeyItems .= "▶ [" . thisID . "]`n"
            continue 
        }
        if (isKeyEmpty)
            continue

        hasSelfDup := false
        if (k1 != "" && (k1 == k2 || k1 == k3))
            hasSelfDup := true
        if (!hasSelfDup && k2 != "" && k2 == k3)
            hasSelfDup := true
        
        if (hasSelfDup) {
            SelfDuplicateItems .= "▶ [" . thisID . "] (보조키 중복 선택)`n"
            continue
        }

        mods := ""
        if (k1="Ctrl" || k2="Ctrl" || k3="Ctrl")
            mods .= "^"
        if (k1="Alt" || k2="Alt" || k3="Alt")
            mods .= "!"
        if (k1="Shift" || k2="Shift" || k3="Shift")
            mods .= "+"
        if (k1="Win" || k2="Win" || k3="Win")
            mods .= "#"
            
        fullKeyString := mods . k4
        
        ; ★ [신규 수정] 시스템 보호키 충돌 시, 어떤 키와 겹치는지 보여줌
        if (ReservedKeys.HasKey(fullKeyString)) {
            SystemDuplicateItems .= "▶ [" . thisID . "] ↔ 시스템 예약키 (" . ReservedKeys[fullKeyString] . ")`n"
            continue 
        }
        
        if (UsedKeys.HasKey(fullKeyString)) {
            DuplicateItems .= "▶ [" . UsedKeys[fullKeyString] . "] ↔ [" . thisID . "]`n"
        } else {
            UsedKeys[fullKeyString] := thisID
        }
    }
    
    if (EmptyKeyItems != "" || SelfDuplicateItems != "" || DuplicateItems != "" || SystemDuplicateItems != "") {
        ErrorMsg := "단축키 설정에 오류가 있습니다. 다음 항목을 확인해주세요:`n`n"
        
        if (SystemDuplicateItems != "")
            ErrorMsg .= "[시스템 단축키 충돌]`n" . SystemDuplicateItems . "`n"
        if (EmptyKeyItems != "")
            ErrorMsg .= "[실행키 누락]`n" . EmptyKeyItems . "`n"
        if (SelfDuplicateItems != "")
            ErrorMsg .= "[보조키 단일 중복]`n" . SelfDuplicateItems . "`n"
        if (DuplicateItems != "")
            ErrorMsg .= "[단축키 상호 충돌]`n" . DuplicateItems
            
        MsgBox, 262208, Message, %ErrorMsg%
        return
    }

    For Index, thisID in DynamicHotkeyList {
        Loop, 4 {
            val := Gui_%thisID%_%A_Index%
            IniWrite, %val%, %IniFile%, 단축키, %thisID%_Key%A_Index%
        }
        
        if (HasKey설정값_%thisID%) {
            val설정값 := Gui_%thisID%_설정값
            IniWrite, %val설정값%, %IniFile%, 단축키, %thisID%_Key설정값
        }
    }
    
    MsgBox, 262208, Message, %msgboxuni_0063%
    Reload
return




실행키리셋:
    ; 4100 옵션: 4096(항상 위) + 4(예/아니오 버튼)
    MsgBox, 4100, Message, %msgboxuni_0064%
    
    ; 아니오(No)를 눌렀을 때의 동작
    IfMsgBox, No
        return
        
    ; 예(Yes)를 눌렀을 때 실행될 코드
    MsgBox, 262208, Message, %msgboxuni_0065%
return




GuiMainClose:
MainGuiClose:
MainGuiEscape:
    Gui, Main:Destroy
    Gui, Container:Destroy
    Gui, Child:Destroy
return







;==============================================================================
; [메뉴 클릭 시 실행되는 라벨]
;==============================================================================

OpenContact:
; 사장님의 카카오톡 오픈프로필이나 문의 웹사이트 주소를 넣으세요.
; Run, https://www.shutterstock.com/ko/g/signist
MsgBox, 262208, Message, %msgboxuni_0066%
return

OpenManual:
; 깃허브의 README나 노션 설명서 주소를 넣으세요.
; Run, https://www.shutterstock.com/ko/g/signist
MsgBox, 262208, Message, %msgboxuni_0067%
return

ExitMenu:
RestoreCursors()       ; 커서 복구
ExitApp

