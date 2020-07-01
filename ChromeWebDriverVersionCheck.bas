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
Const SecName As String = "getChromeWebDriver" '---セクション名
Const KeyName As String = "localchromedversion" '---キー名
Const Default As String = "1" '---デフォルト値
Dim RtnCD As Long
Dim RtnStr As String

Dim intFF As Long
Dim strURL As String

Dim LocalChromeVersion As Long
Dim ChromeWebDriverVersion As Long

Dim objFSO As Scripting.FileSystemObject  '--- 「Microsoft Scripting Runtime」の参照設定を有効にしておく
Dim objFolder As Scripting.Folder
Dim objFile As Scripting.File
Dim objRE As RegExp  '--- 「Microsoft VBScript Regular Expressions 5.5」の参照設定を有効にしておく
Dim xmlhttp As MSXML2.XMLHTTP60 '--- 「Microsoft XML v6.0」の参照設定を有効にしておく
Dim htmlDoc As HTMLDocument '--- 「Microsoft HTML Object Library」の参照設定を有効にしておく
Dim element As HTMLLIElement

Const ChromePath As String = "C:\Program Files (x86)\Google\Chrome\Application" '---Chrome.exeが保存されているパス
Const WebDriverPath As String = "C:\Program Files\SeleniumBasic" '---WebDriverが保存されているパス
Const CheckFolder As String = "drivercheck"
Const zipFile As String = "howa1.zip"
Public Function ChromeWebDriverVersionCheck()
'====================================================================================================
'Pythonの場合はコチラを参照　→　https://github.com/sbfm/getChromeWebDriver
'====================================================================================================
    LocalChromeVersion = GetLocalChromeVersion '---Chrome.exeのバージョンを取得する
    ChromeWebDriverVersion = GetChromeWebDriverVersion '---ChromeDriverのバージョンを取得する
    If LocalChromeVersion > ChromeWebDriverVersion Then '---Chrome.exeがChromeDriverより新しければ更新する
        Call UpdateChromeWebDriver(LocalChromeVersion)
        RtnCD = WritePrivateProfileString(SecName, KeyName, CStr(LocalChromeVersion), WebDriverPath & "\" & ConfigFile) '---iniのバージョンを上書きする
    End If
End Function
Function GetLocalChromeVersion() As Long
'====================================================================================================
'Chrome.exeと同じフォルダにバージョン名のフォルダがあるのでそこからバージョンを取得する
'====================================================================================================
    GetLocalChromeVersion = 1 '---初期値
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objRE = CreateObject("VBScript.RegExp")
    
    With objRE '---正規表現の準備
        .Pattern = "([0-9]+(?=\.))"
        .Global = True
    End With
    For Each objFolder In objFSO.GetFolder(ChromePath).SubFolders
        If objRE.Test(objFolder.Name) Then '---正規表現でフォルダ名をマッチさせてバージョンを取得する
            If GetLocalChromeVersion < Val(objFolder.Name) Then GetLocalChromeVersion = Val(objFolder.Name)
        End If
    Next objFolder
    
    Set objRE = Nothing
    Set objFSO = Nothing
End Function
Function GetChromeWebDriverVersion() As Long
'====================================================================================================
'WebDriverと同じフォルダにconfig.iniを保存してバージョンを管理して取得する
'====================================================================================================
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(WebDriverPath & "\" & ConfigFile) Then
        GetChromeWebDriverVersion = ReadIni(WebDriverPath & "\" & ConfigFile, SecName, KeyName, Default)
    Else '---config.iniが無いなら作る
        GetChromeWebDriverVersion = Default '---初期値
        intFF = FreeFile
        Open WebDriverPath & "\" & ConfigFile For Output As #intFF '---getChromeWebDriver.pyに合わせたiniを準備する
            Print #intFF, "[getChromeWebDriver]"
            Print #intFF, "localchromedversion = 1"
            Print #intFF, "chromepath = " & ChromePath
            Print #intFF, "tempdirectory = "
            Print #intFF, "useproxy = false"
            Print #intFF, "httpproxy = "
            Print #intFF, "httpsproxy = "
        Close
    End If
    
    If objFSO.FolderExists(WebDriverPath & "\" & CheckFolder) = False Then '---zip保存用フォルダが無かったら作成しておく
        objFSO.CreateFolder WebDriverPath & "\" & CheckFolder
    End If
    Set objFSO = Nothing
End Function
Function UpdateChromeWebDriver(TargetChromeWebDriverVersion As Long)
'====================================================================================================
'https://chromedriver.chromium.org/downloadsのソースを分解してChromeのバージョンに合わせたWebDriverをダウンロードして保存する
'====================================================================================================
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    Set htmlDoc = CreateObject("HTMLFile")
    'Set htmlDoc = New HTMLDocument
    
    With xmlhttp
        .Open "GET", "https://chromedriver.chromium.org/downloads" '---Chromeドライバーが公開されているサイトからソースを取得
        .Send
        Do While .readyState < 4
            DoEvents
        Loop
        If .responseText <> "" Then
            htmlDoc.Write .responseText
            For Each element In htmlDoc.getElementsByTagName("li")
                If element.innerHTML Like "*If you are using Chrome version " & TargetChromeWebDriverVersion & "*" Then '---Chromeバージョン別の推奨ドライバー収集
                    strURL = element.getElementsByTagName("a")(0).getAttribute("href") '---アンカー部分を抽出
                    strURL = Split(strURL, "=")(1) '---GET部分に分解
                    strURL = Replace(strURL, "/", "") '---最後のスラッシュを消したい
                    strURL = "https://chromedriver.storage.googleapis.com/" & strURL & "/chromedriver_win32.zip" '---規定のURLにバージョンを埋め込む
                    
                    URLDownloadToFile 0, strURL, WebDriverPath & "\" & CheckFolder & "\" & zipFile, 0, 0 '---ファイルをダウンロードするWinAPI
                    Kill WebDriverPath & "\" & "chromedriver.exe" '---今のドライバーを削除する
                    With CreateObject("Shell.Application") '---zipを既定のフォルダに向けて解凍する
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
'WinAPIでiniファイルから情報を取得する
'====================================================================================================
    RtnStr = Space$(256)
    RtnCD = GetPrivateProfileString(SName, KName, Default, RtnStr, 255, FName)
        
    ' 戻り値設定
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



