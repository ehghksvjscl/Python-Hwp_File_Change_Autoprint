# hwp 파일 field명을 이용한 자동화

## 1. 사용 패키지
- os
- wincom32.client (hwp 파일 제어)
- win32event (윈도우 print 제어 [기본 값 : 기본 프린터]
- win32process (윈도우 프로세서 제어)
- datetime -> timedelta, datetime (시간 계산 및 가져오기)
- win32com.shell.shell -> ShellExecuteEx (윈도우 cmd 제어)
- win32com.shell -> shellcon (윈도우 cmd 제어)
- pytimekr -> pytimekr (공유일 가져오기)

## 2. 실행 준비
  1. C드라이브에 auto_print라는 폴더에 clone 한다.
      ex) C:\auto_pirnt\3day.bat ~~~
  2. 보안모듈(Automation).zip을 압축을 풀고 레지스트에 등록한다.
  3. bat파일을 실행 시킨다.
  
 ## 3. 실행
 3days.bat or 30days.bat file exec
