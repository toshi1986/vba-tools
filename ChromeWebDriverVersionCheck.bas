Attribute VB_Name = "ChromeWebDriverVersionCheck"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, _
                          ByVal lpReturnedString As String, _
                          ByVal nSize As Long, _
                          ByVal lpFileName As String) As Long
                          
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpString As Any, _
                          ByVal lpFileName As String) As Long
                          
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                        (ByVal pCaller As Long, _
                         ByVal szURL As String, _
                         ByVal szFileName As String, _
                         ByVal dwReserved As Long, _
                         ByVal lpfnCB As Long) As Long

Const ConfigFile As String = "config.ini"
Const SecName As String = "getChromeWebDriver" '---�Z�N�V������
Const KeyName As String = "localchromedversion" '---�L�[��
Const Default As String = "1" '---�f�t�H���g�l
Dim RtnCD As Long
Dim RtnStr As String

Dim intFF As Long
Dim strURL As String

Dim LocalChromeVersion As Long
Dim ChromeWebDriverVersion As Long

Dim objFSO As Scripting.FileSystemObject  '--- �uMicrosoft Scripting Runtime�v�̎Q�Ɛݒ��L���ɂ��Ă���
Dim objFolder As Scripting.Folder
Dim objFile As Scripting.File
Dim objRE As RegExp  '--- �uMicrosoft VBScript Regular Expressions 5.5�v�̎Q�Ɛݒ��L���ɂ��Ă���
Dim xmlhttp As MSXML2.XMLHTTP60 '--- �uMicrosoft XML v6.0�v�̎Q�Ɛݒ��L���ɂ��Ă���
Dim htmlDoc As HTMLDocument '--- �uMicrosoft HTML Object Library�v�̎Q�Ɛݒ��L���ɂ��Ă���
Dim element As HTMLLIElement

Const ChromePath As String = "C:\Program Files (x86)\Google\Chrome\Application" '---Chrome.exe���ۑ�����Ă���p�X
Const WebDriverPath As String = "C:\Program Files\SeleniumBasic" '---WebDriver���ۑ�����Ă���p�X
Const CheckFolder As String = "drivercheck"
Const zipFile As String = "howa1.zip"
Public Function ChromeWebDriverVersionCheck()
'====================================================================================================
'Python�̏ꍇ�̓R�`�����Q�Ɓ@���@https://github.com/sbfm/getChromeWebDriver
'====================================================================================================
    LocalChromeVersion = GetLocalChromeVersion '---Chrome.exe�̃o�[�W�������擾����
    ChromeWebDriverVersion = GetChromeWebDriverVersion '---ChromeDriver�̃o�[�W�������擾����
    If LocalChromeVersion > ChromeWebDriverVersion Then '---Chrome.exe��ChromeDriver���V������΍X�V����
        Call UpdateChromeWebDriver(LocalChromeVersion)
        RtnCD = WritePrivateProfileString(SecName, KeyName, CStr(LocalChromeVersion), WebDriverPath & "\" & ConfigFile) '---ini�̃o�[�W�������㏑������
    End If
End Function
Function GetLocalChromeVersion() As Long
'====================================================================================================
'Chrome.exe�Ɠ����t�H���_�Ƀo�[�W�������̃t�H���_������̂ł�������o�[�W�������擾����
'====================================================================================================
    GetLocalChromeVersion = 1 '---�����l
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objRE = CreateObject("VBScript.RegExp")
    
    With objRE '---���K�\���̏���
        .Pattern = "([0-9]+(?=\.))"
        .Global = True
    End With
    For Each objFolder In objFSO.GetFolder(ChromePath).SubFolders
        If objRE.Test(objFolder.Name) Then '---���K�\���Ńt�H���_�����}�b�`�����ăo�[�W�������擾����
            If GetLocalChromeVersion < Val(objFolder.Name) Then GetLocalChromeVersion = Val(objFolder.Name)
        End If
    Next objFolder
    
    Set objRE = Nothing
    Set objFSO = Nothing
End Function
Function GetChromeWebDriverVersion() As Long
'====================================================================================================
'WebDriver�Ɠ����t�H���_��config.ini��ۑ����ăo�[�W�������Ǘ����Ď擾����
'====================================================================================================
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(WebDriverPath & "\" & ConfigFile) Then
        GetChromeWebDriverVersion = ReadIni(WebDriverPath & "\" & ConfigFile, SecName, KeyName, Default)
    Else '---config.ini�������Ȃ���
        GetChromeWebDriverVersion = Default '---�����l
        intFF = FreeFile
        Open WebDriverPath & "\" & ConfigFile For Output As #intFF '---getChromeWebDriver.py�ɍ��킹��ini����������
            Print #intFF, "[getChromeWebDriver]"
            Print #intFF, "localchromedversion = 1"
            Print #intFF, "chromepath = " & ChromePath
            Print #intFF, "tempdirectory = "
            Print #intFF, "useproxy = false"
            Print #intFF, "httpproxy = "
            Print #intFF, "httpsproxy = "
        Close
    End If
    
    If objFSO.FolderExists(WebDriverPath & "\" & CheckFolder) = False Then '---zip�ۑ��p�t�H���_������������쐬���Ă���
        objFSO.CreateFolder WebDriverPath & "\" & CheckFolder
    End If
    Set objFSO = Nothing
End Function
Function UpdateChromeWebDriver(TargetChromeWebDriverVersion As Long)
'====================================================================================================
'https://chromedriver.chromium.org/downloads�̃\�[�X�𕪉�����Chrome�̃o�[�W�����ɍ��킹��WebDriver���_�E�����[�h���ĕۑ�����
'====================================================================================================
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    Set htmlDoc = CreateObject("HTMLFile")
    'Set htmlDoc = New HTMLDocument
    
    With xmlhttp
        .Open "GET", "https://chromedriver.chromium.org/downloads" '---Chrome�h���C�o�[�����J����Ă���T�C�g����\�[�X���擾
        .Send
        Do While .readyState < 4
            DoEvents
        Loop
        If .responseText <> "" Then
            htmlDoc.Write .responseText
            For Each element In htmlDoc.getElementsByTagName("li")
                If element.innerHTML Like "*If you are using Chrome version " & TargetChromeWebDriverVersion & "*" Then '---Chrome�o�[�W�����ʂ̐����h���C�o�[���W
                    strURL = element.getElementsByTagName("a")(0).getAttribute("href") '---�A���J�[�����𒊏o
                    strURL = Split(strURL, "=")(1) '---GET�����ɕ���
                    strURL = Replace(strURL, "/", "") '---�Ō�̃X���b�V������������
                    strURL = "https://chromedriver.storage.googleapis.com/" & strURL & "/chromedriver_win32.zip" '---�K���URL�Ƀo�[�W�����𖄂ߍ���
                    
                    URLDownloadToFile 0, strURL, WebDriverPath & "\" & CheckFolder & "\" & zipFile, 0, 0 '---�t�@�C�����_�E�����[�h����WinAPI
                    Kill WebDriverPath & "\" & "chromedriver.exe" '---���̃h���C�o�[���폜����
                    With CreateObject("Shell.Application") '---zip������̃t�H���_�Ɍ����ĉ𓀂���
                        .Namespace(WebDriverPath).CopyHere .Namespace(WebDriverPath & "\" & CheckFolder & "\" & zipFile).Items
                    End With
                    Exit For
                End If
            Next
        End If
    End With
    
    Set htmlDoc = Nothing
    Set xmlhttp = Nothing
End Function
Public Function ReadIni(ByVal FName As String, ByVal SName As String, ByVal KName As String, ByVal Default As String) As String
'====================================================================================================
'WinAPI��ini�t�@�C����������擾����
'====================================================================================================
    RtnStr = Space$(256)
    RtnCD = GetPrivateProfileString(SName, KName, Default, RtnStr, 255, FName)
        
    ' �߂�l�ݒ�
    If RtnCD > 0 Then
        If InStr(RtnStr, Chr$(0)) > 0 Then
            ReadIni = Left$(RtnStr, InStr(RtnStr, Chr$(0)) - 1)
        Else
            ReadIni = ""
        End If
    Else
        ReadIni = Default
    End If
End Function



