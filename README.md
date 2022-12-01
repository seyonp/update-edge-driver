# Update Edge WebDriver

## VBS file
Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'C:\Users\MSDemo\Desktop\Edge WebDriver\edge-webdriver-update - Copy.xlsm'!Module1.UpdateEdgeWebDriver"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing

## Macro code
'Attribute VB_Name = "WebDriverManager4SeleniumBasic"

'0) Download exe for Selenium Basic library installation: https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
'1) Install Selenium Basic library (by default, it will install to: C:\Users\USERNAME\AppData\Local\SeleniumBasic\edgedriver.exe)
'2) Download (any ver) Edge WebDriver from: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
'3) Replace the edgedriver.exe with the currently existing exe
'4) Place this script inside a new module of a new Macro-enabled Workbook (.xlsm)
'5) Tools > References > tick Selenium Type Library
'6) Save the .xlsm 
'7) Save the .vbs file that will run the above macro
'8) Set up scheduled task in Task Scheduler
'9) What the scheduled task will do: Download the matching version Edge Driver if not already matching > Open Edge browser to complete Driver download > Closer browser automatically

Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Enum BrowserName
    Chrome
    Edge
End Enum


'// Win32 API for file download
#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
'#Else
'Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
'    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If


Private Property Get fso() 'As FileSystemObject
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property


'// Default path of downloaded WebDriver zip
Public Property Get ZipPath(Browser As BrowserName) As String
    Dim path_download As String
    path_download = CreateObject("Shell.Application").Namespace("shell:Downloads").self.Path
    Select Case Browser
    Case BrowserName.Chrome
        ZipPath = path_download & "\chromedriver_win32.zip"
    Case BrowserName.Edge
        Select Case Is64BitOS
            Case True: ZipPath = path_download & "\edgedriver_win64.zip"
            Case Else: ZipPath = path_download & "\edgedriver_win32.zip"
        End Select
    End Select
End Property


'// WebDriver executable file location
Public Property Get WebDriverPath(Browser As BrowserName) As String
    Dim SeleniumBasicParentPath As String
    SeleniumBasicParentPath = Environ("LOCALAPPDATA") & "\SeleniumBasic\"
    
    If Not fso.FolderExists(SeleniumBasicParentPath) Then
        SeleniumBasicParentPath = Environ("ProgramFiles") & "\SeleniumBasic\"
    End If
    
    If Not fso.FolderExists(SeleniumBasicParentPath) Then
        SeleniumBasicParentPath = Environ("ProgramFiles(x86)") & "\SeleniumBasic\"
    End If
    
    Select Case Browser
        Case BrowserName.Chrome: WebDriverPath = SeleniumBasicParentPath & "chromedriver.exe"
        Case BrowserName.Edge:   WebDriverPath = SeleniumBasicParentPath & "edgedriver.exe"
    End Select
End Property



'// Read browser version from registry
'// Output example "94.0.992.31"
Public Property Get BrowserVersion(Browser As BrowserName)
    Dim reg_version As String
    Select Case Browser
        Case BrowserName.Chrome: reg_version = "HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\BLBeacon\version"
        Case BrowserName.Edge:   reg_version = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Edge\BLBeacon\version"
    End Select
    
    On Error GoTo Catch
    BrowserVersion = CreateObject("WScript.Shell").RegRead(reg_version)
    Exit Property
    
Catch:
    Err.Raise 4000, , "Failed to get version information. No browser installed."
End Property
'// Output example "94"
Public Property Get BrowserVersionToMajor(Browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(Browser), ".")
    BrowserVersionToMajor = vers(0)
End Property
'// Output example "94.0"
Public Property Get BrowserVersionToMinor(Browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(Browser), ".")
    BrowserVersionToMinor = Join(Array(vers(0), vers(1)), ".")
End Property
'// Output example "94.0.992"
Public Property Get BrowserVersionToBuild(Browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(Browser), ".")
    BrowserVersionToBuild = Join(Array(vers(0), vers(1), vers(2)), ".")
End Property


'// Determine if OS is 64Bit
Public Property Get Is64BitOS() As Boolean
    Dim arch As String
    'Return value one of: "AMD64","IA64","x86"
    arch = CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE")
    'If you are running 32bitOffice on 64bitOS, check that the original OS bit number is saved in PROCESSOR_ARCHITEW6432.
    If InStr(arch, "64") = 0 Then arch = CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITEW6432")
    Is64BitOS = InStr(arch, "64")
End Property




'// If you omit the third argument, it will be downloaded to the download folder
'// Example of use: DownloadWebDriver Edge, "94.0.992.31"
'//
'// If you use the BrowserVersion property as the second argument, you can download the WebDriver compatible with the current browser
'// Example of use: DownloadWebDriver Edge, BrowserVersion(Edge)
'//
'// If you specify the path in the third argument, you can save it in any location with any name.
'// Example of use: DownloadWebDriver Edge, "94.0.992.31", "C:\Users\yamato\Desktop\edgedriver_94.zip"
Public Function DownloadWebDriver(Browser As BrowserName, ver_webdriver As String, Optional path_save_to As String) As String
    Dim url As String
    Select Case Browser
    Case BrowserName.Chrome
        url = Replace("https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip", "{version}", ver_webdriver)
    Case BrowserName.Edge
        Select Case Is64BitOS
            Case True: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip", "{version}", ver_webdriver)
            Case Else: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win32.zip", "{version}", ver_webdriver)
        End Select
    End Select
    
    If path_save_to = "" Then path_save_to = ZipPath(Browser)   'Default is: "C:Users\USERNAME\Downloads\~~~.zip"
    
    DeleteUrlCacheEntry url
    Dim ret As Long
    ret = URLDownloadToFile(0, url, path_save_to, 0, 0)
    If ret <> 0 Then Err.Raise 4001, , "Download failed : " & url
    
    DownloadWebDriver = path_save_to
End Function

'// Extracting the contents of zip and saving the executable file to the specified location
'// Create a temp folder so as not to overwrite the original executable file, then move the executable file to the target path
'// Normally, when extracting a zip, you specify the destination folder of the zip, but this function specifies the path of the WebDriver executable file, so be careful! (Extracts only exe)
'// Example of use:
'// Extract "C:\Users\yamato\Downloads\chromedriver_win32.zip", "C:\Users\yamato\Downloads\chromedriver_94.exe"
Sub Extract(path_zip As String, path_save_to As String)
    Debug.Print "Extracting zip"
    
    Dim folder_temp
    folder_temp = fso.BuildPath(fso.GetParentFolderName(path_save_to), fso.GetTempName)
    fso.CreateFolder folder_temp
    Debug.Print "    temp folder: " & folder_temp
    
    'Since it was detected as malware when deployed using PowerShell,
    'Extract zip using MS-deprecated Shell.Application
    On Error GoTo Catch
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    
    'Copy the files in the zip file to the specified folder
    'If you don't evaluate the string with a () and pass it to Namespace, an error will occur
    sh.Namespace((folder_temp)).CopyHere sh.Namespace((path_zip)).Items
    
    Dim path_exe As String
    path_exe = fso.BuildPath(folder_temp, Dir(folder_temp & "\*.exe"))
    
    If fso.FileExists(path_save_to) Then fso.DeleteFile path_save_to
    fso.CopyFile path_exe, path_save_to, True
    
    fso.DeleteFolder folder_temp
    Debug.Print "    Extracted WebDriver to: " & path_save_to
    Exit Sub
    
Catch:
    fso.DeleteFolder folder_temp
    Err.Raise 4002, , "Zip extraction failed. Cause: " & Err.Description
    Exit Sub
End Sub


' // Basically, you should download WebDriver of exactly the same version as the browser version.
Public Function RequestWebDriverVersion(ver_chrome)
    Dim http 'As XMLHTTP60
    Dim url As String
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & ver_chrome
    http.Open "GET", url, False
    http.send
    
    If http.statusText <> "OK" Then
        Err.Raise 4003, "Failed to connect to server"
        Exit Function
    End If

    RequestWebDriverVersion = http.responseText
End Function


'// Automatically download the WebDriver that matches the browser version, extract the zip, and place the WebDriver exe in a specific folder
'// By default, WebDriver is downloaded to C:\Users\USERNAME\Downloads, and then placed in C:\Users\USERNAME\AppData\SeleniumBasic\edgedriver.exe
'// If you specify the second argument, you can place the Driver in any folder/filename
'// Even if the folder in the middle of the specified path does not exist, it will be automatically created
'// Example of use:
'// InstallWebDriver Edge, "C:\Users\USERNAME\Desktop\a\b\c\edgedriver_123.exe"
Public Sub InstallWebDriver(Browser As BrowserName, Optional path_driver As String)
    Debug.Print "Installing WebDriver...."
    
    Dim ver_browser   As String
    Dim ver_webdriver As String
    ver_browser = BrowserVersion(Browser)
    Select Case Browser
        Case BrowserName.Chrome: ver_webdriver = RequestWebDriverVersion(BrowserVersionToBuild(Browser))
        Case BrowserName.Edge:   ver_webdriver = ver_browser
    End Select
    
    Debug.Print "   Browser          : Ver. " & ver_browser
    Debug.Print "   Matching WebDriver : Ver. " & ver_webdriver
    
    Dim path_zip As String
    path_zip = DownloadWebDriver(Browser, ver_webdriver)
    
    Do Until fso.FileExists(ZipPath(Browser))
        DoEvents
    Loop
    Debug.Print "   Download completed: " & path_zip
    
    If path_driver = "" Then path_driver = WebDriverPath(Browser)
    
    If Not fso.FolderExists(fso.GetParentFolderName(path_driver)) Then
        Debug.Print "   Creating a destination folder for WebDriver..."
        CreateFolderEx fso.GetParentFolderName(path_driver)
    End If
    
    Extract path_zip, path_driver

End Sub



' // A function that checks the existence of all folders included in the path and creates folders
' // Example of use:
' // CreateFolderEx "C:\a\b\c\d\e\"
Public Sub CreateFolderEx(path_folder As String)
    '// Recurse until the parent folder can no longer be traced back
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    '// Create a folder that does not exist in the middle
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub



'// Replacing SeleniumBasic's Driver.Start with this eliminates the need for extra operations when upgrading or distributing to new PCs
Public Sub SafeOpen(Driver As Selenium.WebDriver, Browser As BrowserName, Optional CustomDriverPath As String)
    
    If Not IsOnline Then Err.Raise 4005, , "You are offline. Please connect to the internet": Exit Sub
    
    Dim DriverPath As String
    DriverPath = IIf(CustomDriverPath <> "", CustomDriverPath, WebDriverPath(Browser))
    
    '// Driver update process:
    If Not IsLatestDriver(Browser, DriverPath) Then
        Dim TmpDriver As String
        If fso.FileExists(DriverPath) Then TmpDriver = BuckupTempDriver(DriverPath)
        
        Call InstallWebDriver(Browser, DriverPath)
    End If
    
    On Error GoTo Catch
    Select Case Browser
        Case BrowserName.Chrome: Driver.Start "chrome"
        Case BrowserName.Edge: Driver.Start "edge"
    End Select
    
    If TmpDriver <> "" Then Call DeleteTempDriver(TmpDriver)
    Exit Sub
    
Catch:
    If TmpDriver <> "" Then Call RollbackDriver(TmpDriver, DriverPath)
    Err.Raise Err.Number, , Err.Description
    
End Sub



'// Determine if the PC is online
'// Since the request destination is google, it is rare that the page cannot be opened due to a failure
Public Function IsOnline() As Boolean
    Dim http
    Dim url As String
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://www.google.co.jp/"
    http.Open "GET", url, False
    http.send
    
    Select Case http.statusText
        Case "OK": IsOnline = True
        Case Else: IsOnline = False
    End Select
End Function


'// Check driver version
Function DriverVersion(DriverPath As String) As String
    If Not fso.FileExists(DriverPath) Then DriverVersion = "": Exit Function
    
    Dim ret As String
    ret = CreateObject("WScript.Shell").Exec(DriverPath & " -version").StdOut.ReadLine
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "\d+\.\d+\.\d+(\.\d+|)"
    DriverVersion = reg.Execute(ret)(0).Value
End Function

'// Check if the latest driver is installed
Function IsLatestDriver(Browser As BrowserName, DriverPath As String) As Boolean
    Select Case Browser
    
    Case BrowserName.Edge
        IsLatestDriver = BrowserVersion(Edge) = DriverVersion(DriverPath)
    
    Case BrowserName.Chrome
        IsLatestDriver = RequestWebDriverVersion(BrowserVersionToBuild(Chrome)) = DriverVersion(DriverPath)
    
    End Select
End Function

'// Save WebDriver to temp folder
Function BuckupTempDriver(DriverPath As String) As String
    Dim TmpFolder As String
    TmpFolder = fso.BuildPath(fso.GetParentFolderName(DriverPath), fso.GetTempName)
    fso.CreateFolder TmpFolder
    
    Dim TmpDriver As String
    TmpDriver = fso.BuildPath(TmpFolder, "\webdriver.exe")
    fso.MoveFile DriverPath, TmpDriver
    
    BuckupTempDriver = TmpDriver
End Function

'// Copy the WebDriver from the temp folder to the WebDriver location
Sub RollbackDriver(TmpDriverPath As String, DriverPath As String)
    fso.CopyFile TmpDriverPath, DriverPath, True
    fso.DeleteFolder fso.GetParentFolderName(TmpDriverPath)
End Sub

'// Delete the temp WebDriver
Sub DeleteTempDriver(TmpDriverPath As String)
    fso.DeleteFolder fso.GetParentFolderName(TmpDriverPath)
End Sub


Public Sub UpdateEdgeWebDriver()
    Dim Driver As New Selenium.EdgeDriver
    SafeOpen Driver, Edge, "C:\EUC_EdgeWebDriver\msedgedriver.exe"
    Driver.Get "https://www.google.co.jp/?q=selenium"
    Sleep 3000
    Driver.Close
End Sub
