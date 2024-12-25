# installer-excel-xlam
{template} for excel xlam installer


- src 폴더의 파일을 APPDATA\Microsoft\AddIns에 복사합니다.
- xlam파일은 레지스트리에 등록합니다. (별도 addin 추가과정을 대신합니다.)
- 엑셀을 실행 중이라면 재실행이 필요하여 이를 확인합니다.

<br />

# How to Use
- 설치할 파일을 하위폴더`src`에 넣습니다.
- exe 파일을 실행해서 절차에 따라 진행합니다.

<br />

# How to build .EXE file
- **윈도우 powershell**을 실행합니다.
- PS2EXE 모듈 설치를 합니다.   
  `Install-Module -Name ps2exe -Scope CurrentUser -Force`
     
- 아래 명령어로 exe파일을 생성합니다.    
  ```Invoke-ps2exe .\installer-excel-xlam-exe.ps1 .\installer-excel-xlam.exe -noConsole -RequireAdmin```

<br />
<br />


  ![image](https://github.com/user-attachments/assets/52dc639b-0aec-4dcd-9989-6addd70ec309)
