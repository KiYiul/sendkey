' 2025-05-21
' kiyiul@asianaidt.com
' 주기적으로 sendkey 사용


' define
Dim logFile, fso, stream, userProfile, currentDate, WshShell

' 변수 설정
Set objShell = CreateObject("WScript.Shell")
userProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
currentDate = Year(Now) & "-" & Right("00" & Month(Now), 2) & "-" & Right("00" & Day(Now), 2)
logFile = userProfile & "\WorkTime_" & currentDate & ".log"

' 객체 생성
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set stream = CreateObject("ADODB.Stream")

' Stream 설정: 인코딩
stream.Charset = "utf-8"
stream.Open

' 파일이 존재하지 않으면 새로 생성, 있으면 끝으로 추가
If Not fso.FileExists(logFile) Then
    ' 파일이 없으면 새로 생성
    stream.WriteText "Started: " & Now & vbCrLf
    stream.SaveToFile logFile, 2 ' 2는 '추가' 모드
Else
    ' 파일이 있으면 기존 내용 뒤에 추가
    stream.Position = stream.Size ' 파일 끝으로 이동
    stream.WriteText "Started: " & Now & vbCrLf
    stream.SaveToFile logFile, 2 ' 2는 '추가' 모드
End If

' 무한 loop 시작
Do
	' 시간 확인
	currentTime = Now
	currentHour = Hour(currentTime)  ' 현재 시간의 시간 부분 (0-23)
	currentMinute = Minute(currentTime)  ' 현재 분 (0-59)

	' 시간에 맞추어서 키 전송, 8시 ~ 17시에만 전송
	If (currentHour >= 8 And currentHour < 12 And currentMinute <= 30) Or (currentHour >= 13 And currentHour < 17) Then
		' 키 전송 (루프 시작)
		WshShell.SendKeys "{SCROLLLOCK}"
		WshShell.SendKeys "{SCROLLLOCK}"

		WScript.Sleep 600000   ' 10분
		'WScript.Sleep 600000   ' 15분
		
		' 로그 기록
		stream.Position = stream.Size  ' 파일 끝으로 이동
		stream.WriteText "Looping: " & Now & vbCrLf
		stream.SaveToFile logFile, 2 ' 2는 '쓰기를 추가' 모드
	End If

    ' 종료 시간 확인 (17시 이후에는 SendKeys를 하지 않도록, 루프를 계속)
    currentTime = Now
    currentHour = Hour(currentTime)
    If currentHour >= 17 Then
        ' 17시 이후에는 키 전송을 하지 않고, 루프를 계속 실행하여 기다리기만 함
        WScript.Sleep 60000  ' 1분 대기 후 루프 계속
        Continue Do  ' 루프의 현재 반복을 중단하고, 다음 반복으로 넘어감
    End If

    ' 10초마다 루프를 돌며 시간 체크
    WScript.Sleep 10000  ' 10초 대기
Loop

' 파일 닫기
stream.Close

' 객체 정리
Set stream = Nothing
Set fso = Nothing
Set WshShell = Nothing