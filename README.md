혼데춘WEB매니저

## config.ini 

~~~.ini
[Settings]
keyfilePath = # ssh 키 파일 경로
keyPassphrase = # ssh 키 파일 passphrase
username = # 로그인 할 계정명
host = # 호스트
port = # ssh 포트
ytapikey = # 유튜브 API 키
appPath = /home/ubuntu/hondaechun # 서버 DB 파일 경로
startCommand = sudo docker run -d -p 80:80 -v /home/ubuntu/hondaechun:/app hondaechun
# 서비스 시작 명령어
stopCommand = sudo docker stop $(sudo docker ps -aq)
# 서비스 중지 명령어
allowedExtensions = .zip,.rar,.7z,.mp3,.mp4,.m4a,.flac,.alac,.wav,.ogg,.opus
# 카카오톡 대화 파일에서 추출할 파일 확장자 (쉼표로 구분, 공백 없음)
~~~
