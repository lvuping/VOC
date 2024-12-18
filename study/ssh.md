ssh -Y lime@192.168.50.39
mssh -Y lime@168.126.186.47

cat ~/.ssh/config 
위의 경로가서 아래와 같이 입력한다. 
그 후 비밀번호 입력 

Host Mac_Mini_internal
    HostName 192.168.50.39  # Mac Mini internal
    User lime

Host Mac_Mini_external
    HostName 168.126.186.47  # Mac Mini external
    User lime