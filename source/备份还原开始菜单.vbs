' 判断系统版本号
SysVersion=GetSystemVersion()
if SysVersion(0)<10 then
    msgbox "本脚本只支持windows10以上系统",16
    wscript.quit
end if

if SysVersion(2)<15063 then
    msgbox "你的系统版本过低，脚本暂不支持1703(15063)以前的系统",16
    wscript.quit
end if


set FSO=createobject("Scripting.FileSystemObject")
PathOfCurrenScript = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
if not FSO.FileExists(PathOfCurrenScript &"\GUI.html") and FSO.FileExists(PathOfCurrenScript &"\GUI.hta") then
    msgbox "缺少 .\source\GUI.html"
    wscript.quit
end if
if not FSO.FileExists(PathOfCurrenScript &"\Compress.ps1") then
    msgbox "缺少 .\source\Compress.ps1"
    wscript.quit
end if
if not FSO.FileExists(PathOfCurrenScript &"\RBST.ico1") and FSO.FileExists(PathOfCurrenScript &"\GUI.hta") then
    msgbox "缺少 .\source\RBST.ico1"
    wscript.quit
end if

WindowStyle_Debug = 0
if not Fso.FileExists(PathOfCurrenScript&"\config.json") then
'═════════════════════ 设置 模块开始 ═════════════════════
Path=""
TriggerDay=0
DeleteBeforeDay=TriggerDay*2
DeleteToReCycleBin = true
BeforeRestore=true
SeletcRestoreFolderByCustom=false

Set outFile = FSO.CreateTextFile(PathOfCurrenScript&"\config.json", True)
JsonDataFront = "{" & vbcrlf
JsonDataBack = vbcrlf& "}"
JsonDataMiddle = _
    """Path"":"""& Path &""","& vbcrlf &_
    """BackUpTaskBar"":false,"& vbcrlf &_
    """TriggerDay"":"& TriggerDay &","& vbcrlf &_
    """DeleteBeforeDay"":"& DeleteBeforeDay &","& vbcrlf &_
    """DeleteToReCycleBin"":"& LCase(DeleteToReCycleBin) &","& vbcrlf &_
    """LastRestorePath"":"""","& vbcrlf &_
    """BeforeRestore"":"& LCase(BeforeRestore) &","& vbcrlf &_
    """RestoreStartMode"":1,"& vbcrlf &_
    """RestoreTaskBarMode"":0,"& vbcrlf &_
    """SeletcRestoreFolderByCustom"":"& LCase(SeletcRestoreFolderByCustom) &","& vbcrlf &_
    """AutoClose"":false"
JsonData = JsonDataFront & JsonDataMiddle & JsonDataBack

outFile.Write JsonData
outFile.Close
Set outFile = Nothing

else
    set oFile = FSO.opentextfile(PathOfCurrenScript&"\config.json",1,true)
    contents = oFile.readall
    set ConfigJson = ParseJson(contents)
    Path = ConfigJson.Path
    BackUpTaskBar = ConfigJson.BackUpTaskBar
    TriggerDay = ConfigJson.TriggerDay
    DeleteBeforeDay = ConfigJson.DeleteBeforeDay
    DeleteToReCycleBin = ConfigJson.DeleteToReCycleBin
    BeforeRestore = ConfigJson.BeforeRestore
    SeletcRestoreFolderByCustom = ConfigJson.SeletcRestoreFolderByCustom
end if



Function ParseJson(strJson) '解析Json
    Set html = CreateObject("htmlfile")
    Set window = html.parentWindow
    window.execScript "var json = " & strJson, "JScript"
    Set ParseJson = window.json
End Function


'声明
'脚本：备份还原开始菜单
'版本：beta 0.6.2
'说明：本脚本可以备份和还原开始菜单布局


'作者：YUPHIZ
'版权：此脚本版权归YUPHIZ所有，备份还原方法借鉴BackupSML和winaero tweaker，并优化改进
        '凡用此脚本从事法律不允许的事情的，均与本作者无关
        '此脚本遵循 gpl3.0 and later 协议




if path="" then
    set FSO=createobject("Scripting.FileSystemObject")
    PathOfCurrenScript = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
    Path=FSO.GetParentFolderName(PathOfCurrenScript)
elseif instr(1,path,":",1) <=0 then
    set FSO=createobject("Scripting.FileSystemObject")
    PathOfCurrenScript = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
    Path=FSO.GetParentFolderName(PathOfCurrenScript) &"\"&path
end if



' 建立环境文件夹
title="开始菜单备份"
set FSO=CreateObject("scripting.filesystemobject")
if not FSO.FolderExists(Path) then
    Ask=CreateObject("Wscript.Shell").Popup( _
        "未找到备份路径"& vbcrlf & path & vbcrlf &_
        "是否创建？",0,"请选择",1+16+512 _
    )
    if Ask=1 then
        call CreateFolder(Path)
        if not FSO.FolderExists(Path) then
             msgbox "建立文件夹失败，请手动操作",1+16
             Wscript.quit
        end if
    else
        Wscript.quit
    end if
end if

select case WScript.Arguments.count
    case 5 ' 【还原2.2】需要的管理员
        if WScript.Arguments(1)="--Restore" then
            RestoreTaskBarMode = WScript.Arguments(2)
            RestoreFolder = WScript.Arguments(4)
            if RestoreTaskBarMode = "" then RestoreTaskBarMode=null
            username = WScript.Arguments(3)
            if RestoreFolder = "" then RestoreFolder = null
            call RestoreStartLayout(Path,"--WithStartLayout--AllUser",RestoreTaskBarMode,UserName,RestoreFolder) '管理员运行
            Wscript.quit
        end if 
    case 1 ' 任务计划和HtaGui备份相关
        if WScript.Arguments(0) = "--HiddenBackup" then
            call BackUpStartLayout(Path,"--Auto",null,"notips")
            if DeleteBeforeDay > 0 then
                call DeleteFilesTree(Path&"\开始菜单备份")
            end if
            Wscript.quit
        elseif WScript.Arguments(0) = "--HiddenBackup--WithTaskBar" then
            call BackUpStartLayout(Path,"--Auto","--WithTaskBar","notips")
            if DeleteBeforeDay > 0 then
                call DeleteFilesTree(Path&"\开始菜单备份")
            end if
            Wscript.quit
        elseif WScript.Arguments(0) = "--HTAGUI" then
            set Shell = CreateObject("WScript.Shell")
            if Shell.Appactivate("备份还原开始菜单GUI v0.6.2 By --@YUPHIZ") then
                Wscript.quit
            End If
            CreateObject("Wscript.Shell").run "mshta """&PathOfCurrenScript&"\gui.hta"""
            Wscript.quit
        end if
    case 2 ' HtaGui还原相关
        ArgumentMode = WScript.Arguments(0)
        JsonContent = WScript.Arguments(1)
        if ArgumentMode = "--JsonBackup" or ArgumentMode = "--jsonReStore" or ArgumentMode = "--restartexplorer" or ArgumentMode = "--removeSchTask" or ArgumentMode = "--ReSetStartLayout" or ArgumentMode = "--JsonBackup--Save" or  ArgumentMode = "--jsonReStore--Save" then
            call GuiLauncher(ArgumentMode,JsonContent)
        end if
        wscript.quit
    case else
        call SatrtLayout()
end select

function GuiLauncher(Mode,JsonContent)
    if JsonContent <> "" and instr(1,Mode,"json",1) > 0 then
        set ConfigJsonFromGui = ParseJson(JsonContent)
        if ConfigJsonFromGui is nothing then wscript.quit
        Path = ConfigJsonFromGui.Path
        BackUpTaskBar = ConfigJsonFromGui.BackUpTaskBar
        BeforeRestore = ConfigJsonFromGui.BeforeRestore
        SeletcRestoreFolderByCustom = ConfigJsonFromGui.SeletcRestoreFolderByCustom
        TriggerDay = ConfigJsonFromGui.TriggerDay
        DeleteBeforeDay = ConfigJsonFromGui.DeleteBeforeDay
        DeleteToReCycleBin = ConfigJsonFromGui.DeleteToReCycleBin
        RestoreStartMode = ConfigJsonFromGui.RestoreStartMode
        RestoreTaskBarMode = ConfigJsonFromGui.RestoreTaskBarMode
        LastRestorePath = ConfigJsonFromGui.LastRestorePath
        AutoClose = ConfigJsonFromGui.AutoClose
        IsNeedUpdateSchTask = ConfigJsonFromGui.IsNeedUpdateSchTask

        if path = "" then
            path = FSO.GetParentFolderName(PathOfCurrenScript)
        end if
        if BackUpTaskBar then
            WithTaskBar = "--WithTaskBar"
        else
            WithTaskBar = null
        end if
        
        FolderName = null
        tips = null
        if Mode = "--JsonBackup" then
            call BackUpStartLayout(Path,FolderName,WithTaskBar,tips)
        end if
        if Mode = "--jsonReStore" then
            RestoreFolder = LastRestorePath
            if RestoreStartMode = "3" then
                UserName = CreateObject("WScript.Network").UserName
                call RunAs(RestoreTaskBarMode,UserName,RestoreFolder)
            else 
                UserName = null
                call RestoreStartLayout(Path,RestoreStartMode,RestoreTaskBarMode,UserName,RestoreFolder)
            end if
        end if
        select case IsNeedUpdateSchTask
            case "Enable-First"
                call enableTasksch(WithTaskBar,IsNeedUpdateSchTask)
            case "Enable"
                call enableTasksch(WithTaskBar,IsNeedUpdateSchTask)
            case "Disable"
                call removeTasksch(null)
        end select

    elseif instr(1,Mode,"json",1) <=0 then
        select case mode
            case "--restartexplorer" 'Y'
                call RestartExplorer()
            case "--removeSchTask"
                call removeTasksch("istips")
            case "--ReSetStartLayout"
                call ResetStartLayout()
        end select
    end if
end function


' 主界面函数
function SatrtLayout()
    Ask=inputbox( _
        "※ 此为旧版ui，请使用【双击运行】启动新GUI"&vbcrlf&vbcrlf& _
        "   ↑↑↑↑↑↑↑ 旧版ui可能有功能缺失"&vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
        "※ 【备份选项】"&vbcrlf&vbcrlf& _
        "    1    备 份 开始菜单"&vbcrlf&vbcrlf& _
        "    1.1  备 份 开始菜单和任务栏 "&vbcrlf&vbcrlf&vbcrlf& _
        "※ 【还原选项】"&vbcrlf&vbcrlf& _
        "    2     还 原 开始菜单、且不覆盖快捷方式"&vbcrlf&vbcrlf& _
        "    2.1   还 原 开始菜单+ 还原个人用户快捷方式"&vbcrlf&vbcrlf& _
        "    2.2   还 原 开始菜单+ 还原所有用户快捷方式"&vbcrlf&vbcrlf&vbcrlf& _
        "    2.3   还 原 任务栏，且不覆盖快捷方式"&vbcrlf&vbcrlf& _
        "    2.4   还 原 任务栏、且覆盖快捷方式"&vbcrlf&vbcrlf&vbcrlf& _
        "    2.5   选项 2   + 选项 2.3"&vbcrlf&vbcrlf& _
        "    2.6   选项 2.1 + 选项 2.3"&vbcrlf&vbcrlf& _
        "    2.7   选项 2.2 + 选项 2.3"&vbcrlf&vbcrlf& _
        "    2.8   选项 2   + 选项 2.4"&vbcrlf&vbcrlf& _
        "    2.9   选项 2.1 + 选项 2.4"&vbcrlf&vbcrlf& _
        "    2.10  选项 2.2 + 选项 2.4"&vbcrlf&vbcrlf&vbcrlf& _
        "    0.0   重置开始菜单"&vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
        "※ 【定时备份选项】"&vbcrlf&vbcrlf&_
        "    3   开 启（刷新）定时自动备份 不备份任务栏"&vbcrlf&vbcrlf&_
        "    3.1 开 启（刷新）定时自动备份 备份任务栏"&vbcrlf&vbcrlf&_
        "    4   禁 用 定时自动备份"&vbcrlf&vbcrlf&_
        "    5   卸 载 定时自动备份 "&vbcrlf&vbcrlf&_
        "    0  管 理 备份 "&vbcrlf&vbcrlf,_
        title&" 小工具",_
        "输入对应的序号(1或2或3) 此为旧版ui，请使用【双击运行】启动新GUI")
    select case True
        case Ask=""
            Wscript.quit
            msgbox "空值"
        case Ask="1"
            call BackUpStartLayout(Path,null,null,null)
        case Ask="1.1"
            call BackUpStartLayout(Path,null,"--WithTaskBar",null)
             
        case Ask="2"
            ' 还原开始菜单、且只还原已有的快捷方式
            call RestoreStartLayout(Path,"--WithStartLayout",null,null,null)
        case Ask="2.1"
            ' 还原开始菜单，且还原个人用户快捷方式
            call RestoreStartLayout(Path,"--WithStartLayout--User",null,null,null)
        case Ask="2.2"
            ' 还原开始菜单、且还原所有用户快捷方式
            RestoreTaskBarMode = null
            UserName = CreateObject("WScript.Network").UserName
            call RunAs(RestoreTaskBarMode,UserName,null)

        case Ask="2.3"
            ' 还原任务栏，不覆盖快捷方式
            call RestoreStartLayout(Path,null,"--WithTaskBar",null,null,null)
        case Ask="2.4"
            ' 还原任务栏，且覆盖快捷方式
            call RestoreStartLayout(Path,null,"--WithTaskBar--WithLnk",null,null)

        case Ask="2.5"
            ' 选项2 + 选项2.3
            call RestoreStartLayout(Path,"--WithStartLayout","--WithTaskBar",null,null)
        case Ask="2.6"
            ' 选项2.1 + 选项2.3
            call RestoreStartLayout(Path,"--WithStartLayout--User","--WithTaskBar",null,null)
        case Ask="2.7"
            ' 选项2.2 + 选项2.3
            RestoreTaskBarMode = "--WithTaskBar"
            UserName = CreateObject("WScript.Network").UserName
            call RunAs(RestoreTaskBarMode,UserName,null)
        case Ask="2.8"
            ' 选项2 + 选项2.4
            call RestoreStartLayout(Path,"--WithStartLayout","--WithTaskBar--WithLnk",null,null)
        case Ask="2.9"
            ' 选项2.1 + 选项2.4
            call RestoreStartLayout(Path,"--WithStartLayout--User","--WithTaskBar--WithLnk",null,null)
        case Ask="2.10"
            ' 选项2.2 + 选项2.4
            RestoreTaskBarMode = "--WithTaskBar--WithLnk"
            UserName = CreateObject("WScript.Network").UserName
            call RunAs(RestoreTaskBarMode,UserName,null)

        case Ask="0.0"
            call ResetStartLayout()
        case Ask="3"
            call enableTasksch(null,"--WithRunBackup")
        case Ask="3.1"
            call enableTasksch("--WithTaskBar","--WithRunBackup")
        case Ask="4"
            call disabledTasksch(null)
        case Ask="5"
            call removeTasksch("istips")
        case Ask="0"
            CreateObject("Wscript.shell").run """"&Path&"\开始菜单备份"""
        case else
            call SatrtLayout()
            Wscript.quit
    end select
end function



' 备份开始菜单或任务栏
function BackUpStartLayout(Path,Folder,WithTaskBar,tips)
    set Shell=CreateObject("Wscript.shell")
    Set Env=Shell.Environment("Process")
    set FSO=CreateObject("scripting.filesystemobject")
    UserName=CreateObject("WScript.Network").UserName
    UserDomain=CreateObject("WScript.Network").UserDomain
    LocalAppData = Env.Item("LocalAppData")
    FileLayout = LocalAppData &"\Microsoft\Windows\Shell\DefaultLayouts.xml"
    StartBin = LocalAppData &"\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start.bin"
    Start2Bin = LocalAppData &"\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"
    OemDefaultStartJson = LocalAppData &"\Microsoft\Windows\Shell\startlayout.json"

    RegPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
    RootFolder=Path&"\开始菜单备份\"
    if not FSO.FolderExists(RootFolder) then
        Call CreateFolder(RootFolder)
    end if
    
    if IsNull(Folder) Then
        if IsNull(WithTaskBar) then
            Folder = RootFolder&UserDomain&"_"&UserName&"_"&year(now)&"_"&Month(now)&"_"&day(now)&"_"&int(timer)&"_Os"&SysVersion(2)
        elseif WithTaskBar="--WithTaskBar" then
            Folder = RootFolder&UserDomain&"_"&UserName&"_"&year(now)&"_"&Month(now)&"_"&day(now)&"_"&int(timer)&"_WithTaskBar"&"_Os"&SysVersion(2)
        end if
    elseif Folder="--Auto" then
        if IsNull(WithTaskBar) then
            Folder = RootFolder&UserDomain&"_"&UserName&"_"&year(now)&"_"&Month(now)&"_"&day(now)&"_"&int(timer)&"_Os"&SysVersion(2)&"_AutoBakvp"
        elseif WithTaskBar="--WithTaskBar" then
            Folder = RootFolder&UserDomain&"_"&UserName&"_"&year(now)&"_"&Month(now)&"_"&day(now)&"_"&int(timer)&"_WithTaskBar"&"_Os"&SysVersion(2)&"_AutoBakvp"
        end if
    
    end if

' ———— 异步处理的mshta提示
    if IsNull(tips) then ProcessID = Start_InvokeTip("备份")

    if not FSO.FolderExists(Folder) then
        Call CreateFolder(Folder)
    end if

    RegFile=Folder&"\开始菜单备份.reg"

    Shell.run "cmd /c echo y|REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount """&RegFile&"""" ,WindowStyle_Debug,true
    if not FSO.FileExists(RegFile) then
        Shell.popup "不能备份开始菜单注册表，可能脚本不支持此系统版本",0,"错误",16
    end if
    
    
    if  FSO.FileExists(FileLayout) then
        FSO.CopyFile FileLayout, Folder&"\DefaultLayouts.xml"
    end if

    if  FSO.FileExists(StartBin) then
        FSO.CopyFile StartBin, Folder&"\start.bin"
    end if
    if  FSO.FileExists(Start2Bin) then
        FSO.CopyFile Start2Bin, Folder&"\start2.bin"
    end if
    if  FSO.FileExists(OemDefaultStartJson) then
        FSO.CopyFile OemDefaultStartJson, Folder&"\startlayout.json" 'H'
    end if
     


    if IsNull(WithTaskBar) then
        OtherRegFile = Folder&"\其他开始菜单设置.reg"
        Shell.run "cmd /c echo y|REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced """&OtherRegFile&"""" ,WindowStyle_Debug,true
        
        OtherStartRegFile = Folder&"\其他开始菜单设置.reg"
        contents = ReadFile(OtherRegFile)
        Newcontents = FilterReg(contents,"start","--include")
        TotalLine = split(Newcontents,vbcrlf)
        TotalLine = ubound(TotalLine)
        if TotalLine >= 3 then
            call WriteFile(OtherStartRegFile,Newcontents)
        end if

        if IsNull(tips) then 
            WindowStyle = 1
        else 
            WindowStyle = 0
        end if
        
        Shell.run "powershell -noprofile -executionpolicy Bypass -file """&PathOfCurrenScript &"\Compress.ps1"" """&Folder&"""" ,WindowStyle_Debug,true
        



    elseif WithTaskBar="--WithTaskBar" then
        OtherRegFile = Folder&"\其他任务栏设置.reg"
        Shell.run "cmd /c echo y|REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced """&OtherRegFile&"""" ,WindowStyle_Debug,true

        OtherStartRegFile = Folder&"\其他开始菜单设置.reg"
        contents = ReadFile(OtherRegFile)
        Newcontents = FilterReg(contents,"start","--include")
        TotalLine = split(Newcontents,vbcrlf)
        TotalLine = ubound(TotalLine)
        if TotalLine >= 3 then
            call WriteFile(OtherStartRegFile,Newcontents)
        end if

        OtherStartRegFile = Folder&"\其他任务栏设置.reg"
        contents = ReadFile(OtherRegFile)
        Newcontents = FilterReg(contents,"start","--exclude")
        TotalLine = split(Newcontents,vbcrlf)
        TotalLine = ubound(TotalLine)
        if TotalLine >= 3 then
            call WriteFile(OtherStartRegFile,Newcontents)
        end if

        RegTaskBarFile=Folder&"\任务栏备份.reg"
        Shell.run "cmd /c echo y|REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband """&RegTaskBarFile&"""" ,WindowStyle_Debug,true
        if not FSO.FileExists(RegTaskBarFile) then
            Shell.popup "不能备份任务栏注册表，可能脚本不支持此系统版本",0,"错误",16
        end if

        RegToolBarFile=Folder&"\工具栏备份.reg"
        Shell.run "cmd /c echo y|REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Desktop """&RegToolBarFile&"""" ,WindowStyle_Debug,true
        if not FSO.FileExists(RegToolBarFile) then 
            Shell.popup "不能备份工具栏栏注册表，可能脚本不支持此系统版本",0,"错误",16
        end if

        if IsNull(tips) then 
            WindowStyle = 1
        else 
            WindowStyle = 0
        end if
        Shell.run "powershell -noprofile -executionpolicy Bypass -file """&PathOfCurrenScript &"\Compress.ps1"" """&Folder &""" -WithTaskBar",WindowStyle_Debug,true
    end If
    
' ———— 关闭异步处理的mshta提示
    if IsNull(tips) then call Stop_InvokeTip(ProcessID)

    if IsNull(WithTaskBar) Then
        if FSO.FileExists(RegFile) then
            Info = "备份 开始菜单 操作完成"
            PopupTimeout = 1
        else
            Info = "备份 开始菜单 失败"
            infocode = 16
            PopupTimeout = 3
        end if
    else
        if FSO.FileExists(RegFile) and FSO.FileExists(RegTaskBarFile) then
            Info = "备份 开始菜单 和 任务栏 操作完成"
            PopupTimeout = 1
        else 
            Info = "备份 开始菜单 和 任务栏 失败"
            infocode = 16
            PopupTimeout = 3
        end if
    end if
if IsNull(tips) then Shell.popup info,PopupTimeout,"备份结果提示",infocode
end function




' 还原开始菜单或任务栏
function RestoreStartLayout(Path,RestoreStartMode,RestoreTaskBarMode,UserName,RestoreFolder)
    set Shell=CreateObject("Wscript.shell")
    Set Env=Shell.Environment("Process")

    UserDomain=CreateObject("WScript.Network").UserDomain
    startFolderAllUser = Shell.SpecialFolders("AllUsersStartMenu")
    
    if IsNull(UserName) then 
        UserName = CreateObject("WScript.Network").UserName
        LocalAppData = Env.Item("LocalAppData")
        startFolderUser = Shell.SpecialFolders("programs")
    else
        SystemDrive = Env.Item("SystemDrive")
        LocalAppData = SystemDrive&"\Users\"&UserName&"\AppData\Local"

        startFolderUser = SystemDrive&"\Users\"&UserName&"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    end if

    
    set FSO=CreateObject("scripting.filesystemobject")
    FileLayout = LocalAppData &"\Microsoft\Windows\Shell\DefaultLayouts.xml"
    StartBin = LocalAppData &"\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start.bin"
    Start2Bin = LocalAppData &"\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"

    if isnull(RestoreFolder) then 
        if SeletcRestoreFolderByCustom=false then
            RootFolderReStore = Path&"\开始菜单备份"
        elseif SeletcRestoreFolderByCustom=true then
            RootFolderReStore = null
        end if
        RestoreFolder = SelectFolder(RootFolderReStore,title,RestoreStartMode,RestoreTaskBarMode,UserName)
    else
        RootFolderReStore = null
    end if

    do until FSO.FileExists( RestoreFolder&"\开始菜单备份.reg")
        title = "所选文件夹不包含还原所需文件" &vbcrlf &vbcrlf &_ 
                "请重新选择要还原的文件夹目录，选择根目录是刷新"
        RestoreFolder = SelectFolder(RootFolderReStore,title,RestoreStartMode,RestoreTaskBarMode,UserName)
    loop
    

' ———— 异步处理的mshta提示
    ProcessID = Start_InvokeTip("还原")
    
    ' 还原前备份
    if not ((IsNull(RestoreStartMode) or RestoreStartMode="0") and (IsNull(RestoreTaskBarMode) or RestoreTaskBarMode="0")) Then
        if BeforeRestore=true then
            BeforeRestoreFolder=Path&"\开始菜单备份\————————"&UserDomain&"_"&UserName&"_还原前的备份"
            call BackUpStartLayout(Path,BeforeRestoreFolder,"--WithTaskBar","notips")
        end if
    end if

    
' ———— 还原解压开始菜单快捷方式
    if RestoreStartMode="--WithStartLayout--User" or RestoreStartMode="2" then
        Unzip RestoreFolder& "\开始菜单快捷方式User.zip", startFolderUser
    elseif RestoreStartMode="--WithStartLayout--AllUser" or RestoreStartMode="3" then
        Unzip RestoreFolder& "\开始菜单快捷方式User.zip", startFolderUser
        Unzip RestoreFolder& "\开始菜单快捷方式Alluser.zip", startFolderAllUser
    end if

' ———— 还原开始菜单
    if (not IsNull(RestoreStartMode)) and RestoreStartMode<>"0" then
        if FSO.FileExists( RestoreFolder&"\DefaultLayouts.xml") then
            FSO.CopyFile RestoreFolder&"\DefaultLayouts.xml",FileLayout
        end if

        ' 适配windows11
        if  FSO.FileExists( RestoreFolder&"\start.bin") then
            FSO.CopyFile RestoreFolder&"\start.bin",StartBin
        end if
        if  FSO.FileExists( RestoreFolder&"\start2.bin") then 'U'
            FSO.CopyFile RestoreFolder&"\start2.bin",Start2Bin
        end if
        
        RegFile=RestoreFolder&"\开始菜单备份.reg"
        Shell.run "REG IMPORT """&RegFile&"""" ,WindowStyle_Debug,true

        StartRegFile = RestoreFolder& "\其他开始菜单设置.reg"
        if FSO.FileExists(StartRegFile) then
            Shell.run "REG IMPORT """&StartRegFile&"""" ,WindowStyle_Debug,true
        end if

        call KillStartMenuProcess()
    end if


' ———— 还原解压任务栏快捷方式
    if RestoreTaskBarMode="--WithTaskBar--WithLnk" or RestoreTaskBarMode="2" then
        AppData = Env.Item("AppData")
        TaskBarFolder=AppData&"\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        if FSO.FileExists( RestoreFolder& "\任务栏快捷方式.zip") then
            Unzip RestoreFolder& "\任务栏快捷方式.zip", TaskBarFolder
        end if
    end if

' ———— 还原任务栏
    if (not isnull(RestoreTaskBarMode)) and RestoreTaskBarMode<> "0" then
        RegTaskBarFile = RestoreFolder&"\任务栏备份.reg"
        if FSO.FileExists(RegTaskBarFile) then
            Shell.run "REG IMPORT """&RegTaskBarFile&"""" ,WindowStyle_Debug,true
            IsNeedRestartExplorer = true
        end if

        RegToolBarFile = RestoreFolder&"\工具栏备份.reg"
        if FSO.FileExists(RegToolBarFile) then
            Shell.run "REG IMPORT """&RegToolBarFile&"""" ,WindowStyle_Debug,true
            IsNeedRestartExplorer = true
        end if

        TaskRegFile = RestoreFolder& "\其他任务栏设置.reg"
        if FSO.FileExists(TaskRegFile) then
            TaskBarReg = ReadFile(TaskRegFile)
            Shell.run "REG IMPORT """&TaskRegFile&"""" ,WindowStyle_Debug,true
            IsNeedRestartExplorer = true
        end if

        if IsNeedRestartExplorer = true then RestartExplorer()
    end if

' ———— 关闭异步处理的mshta提示
    call Stop_InvokeTip(ProcessID)

' ———— 提示操作完成
    if (not IsNull(RestoreStartMode) and RestoreStartMode<>"0")  and (IsNull(RestoreTaskBarMode) or RestoreTaskBarMode="0") then
        Info = "还原 开始菜单 操作完成"
    elseif (IsNull(RestoreStartMode) or RestoreStartMode="0") and (not IsNull(RestoreTaskBarMode) and RestoreTaskBarMode<>"0") Then
        Info = "还原 任务栏和工具栏 操作完成"
    elseif (not IsNull(RestoreStartMode) and RestoreStartMode<>"0") and(not IsNull(RestoreTaskBarMode) and RestoreTaskBarMode<>"0") Then
        Info = "还原 开始菜单、任务栏和工具栏 操作完成"
    elseif (IsNull(RestoreStartMode) or RestoreStartMode="0") and (IsNull(RestoreTaskBarMode) or RestoreTaskBarMode="0") Then
        Info = "您没有选择要还原的操作"
    end if
Shell.popup Info, 1
end function



' 选择还原的文件夹
Function SelectFolder(default,title,RestoreStartMode,RestoreTaskBarMode,UserName)
    If IsNull(default) Then
        default = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    End If

    If IsNull(title) Then
        title = "选择要还原的文件夹目录，选择根目录是刷新"
    End If

    Set Folder = CreateObject("Shell.Application").BrowseForFolder(0, title, 0, default)
    If Folder Is Nothing Then
        Wscript.quit
    elseif Folder.Self.Path=default then
' ————  重新载入文件夹
        call RestoreStartLayout(Path,RestoreStartMode,RestoreTaskBarMode,UserName,null)
        Wscript.quit
    else
        SelectFolder = Folder.Self.Path
    End If
End Function



' 重置开始菜单或任务栏
Function ResetStartLayout()
    set FSO=CreateObject("scripting.filesystemobject")
    
    UserName=CreateObject("WScript.Network").UserName
    UserDomain=CreateObject("WScript.Network").UserDomain

    ' 重置前备份
    if BeforeRestore=true then
        BeforeRestoreFolder=Path&"\开始菜单备份\————————"&UserDomain&"_"&UserName&"_还原前的备份"
        call BackUpStartLayout(Path,BeforeRestoreFolder,"--WithTaskBar","notips")
    end if


    set Shell = CreateObject("Wscript.shell")
    Set Env = Shell.Environment("Process")
    LocalAppData = Env.Item("LocalAppData")
    FileLayout = LocalAppData &"\Microsoft\Windows\Shell\DefaultLayouts.xml"
    StartBin = LocalAppData &"\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start.bin"
    Start2Bin = LocalAppData &"\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"

    SysVersion = GetSystemVersion()
    if FSO.FileExists(FileLayout) and SysVersion(2)<17744 then
        FSO.deleteFile FileLayout
    end if

    Shell.run "reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount /f",WindowStyle_Debug,true

    if FSO.FileExists(StartBin) then
        FSO.deleteFile StartBin
    end if
    if FSO.FileExists(Start2Bin) then
        FSO.deleteFile Start2Bin
    end if
    
    On Error Resume Next
    Shell.regdelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_Layout"
    on error goto 0
    call KillStartMenuProcess()
    Shell.popup "重置菜单操作完成",1
end function



' 创建文件夹
Function CreateFolder(path)
    set FSO=CreateObject("scripting.filesystemobject")
    Set getDrivers=FSO.Drives
    
    Set getDrive=FSO.GetDrive(FSO.GetDriveName(path))
    If Not FSO.FolderExists(fso.GetParentFolderName(path)) Then
      CreateFolder FSO.GetParentFolderName(path)
    End If
    FSO.CreateFolder(path)
End Function


' 压缩函数，不支持unicode路径
Function Zip(ZipSourcePath,ZipFile)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(ZipFile, True)
        f.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
        f.Close
    Set Shell = CreateObject("Shell.Application")
    Set Source = Shell.NameSpace(ZipSourcePath)
    
    Set Target = Shell.NameSpace(ZipFile)
    intOptions = 4+256+1024
    Target.CopyHere ZipSourcePath, intOptions
    Do
        WScript.Sleep 1000
    Loop Until Target.Items.Count > 0
End Function


' 解压函数，支持unicode路径
Sub UnZip(ZipFile,TargetPath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If NOT fso.FolderExists(TargetPath) Then
        fso.CreateFolder(TargetPath)
    End If
    Set Shell = CreateObject("Shell.Application")
    Set Source = Shell.NameSpace(ZipFile)
    Set FolderItem = Source.Items()
    Set Target = Shell.NameSpace(TargetPath)
    intOptions =  4+16+256+1024
    Target.CopyHere FolderItem, intOptions
End Sub


' 关闭异步处理的操作提示的函数
function Stop_InvokeTip(processid)
set WR=getobject("winmgmts:\\.\root\cimv2") 
set ps=WR.execquery("select * from win32_process where processid = "&processid)
for each Oneof in ps
    ' msgbox  Oneof.name
    if Oneof.name = "mshta.exe" then 
        CreateObject("WScript.Shell").run "taskkill /im " &processid& " /f",0
        exit for
    end if
next
end function

' 异步处理的操作提示的函数
function Start_InvokeTip(StringTips)
        Set Shell = CreateObject("WScript.Shell")
        mshta = "mshta ""vbscript:createobject(""wscript.shell"").popup(""vbs后台正在"&StringTips&"……"",0,""vbs后台正在"&StringTips&"……"") & window.close"""
        Set oExec = Shell.exec(mshta)
Start_InvokeTip = oexec.processid
end function



' 管理员运行
function RunAs(RestoreTaskBarMode,UserName,RestoreFolder)
    set Shell=CreateObject("Shell.Application") 
    Shell.ShellExecute "wscript.exe" _ 
    , """" & WScript.ScriptFullName & """ RunAsAdministrator --Restore """&RestoreTaskBarMode&""" """&UserName&""" """&RestoreFolder&"""", , "runas", 1
    WScript.Quit
end function


' 启用任务计划
sub enableTasksch(WithTaskBar,WithRunBackup)  ' 启用本脚本任务计划
    set FSO=createobject("Scripting.FileSystemObject")
    PathOfCurrenScript = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
    UserName=CreateObject("WScript.Network").UserName
    UserDomain=CreateObject("WScript.Network").UserDomain
    
    if WithRunBackup = "--WithRunBackup" then
        call BackUpStartLayout(Path,"--Auto",WithTaskBar,null)
    elseif WithRunBackup = "Enable-First" then
        call BackUpStartLayout(Path,"--Auto",WithTaskBar,"notips")
    end if

    if isNUll(WithTaskBar) then
        Argument = "--HiddenBackup"
    elseif WithTaskBar = "--WithTaskBar" then 'Z'
        Argument = "--HiddenBackup--WithTaskBar"
    end if

    

    Set ShellTask=createobject("Schedule.Service")
    call ShellTask.connect()
    Set rootFolder=ShellTask.getfolder("\")
    
    Set taskDefinition=ShellTask.NewTask(0)
    
    Set Settings = taskDefinition.Settings
    Settings.StartWhenAvailable = True
    Settings.DisallowStartIfOnBatteries = false
    Settings.ExecutionTimeLimit= "PT5M"
    Set triggers = taskDefinition.Triggers

    TriggerTypeDaily= 2
    Set trigger = triggers.Create(TriggerTypeDaily)
    monthnow=month(now)
    if len(monthnow)=1 then monthnow="0"&month(now)
    daynow=day(now)
    if len(daynow)=1 then daynow="0"&day(now)
    startTime = year(now)&"-"&monthnow&"-"&daynow&"T02:00:00"
    trigger.StartBoundary = startTime
    
    trigger.DaysInterval = TriggerDay
    trigger.Id = "StartMenuBvkupDailyTrigger"
    trigger.Enabled = True

    Set Action = taskDefinition.Actions.Create(0)
    Action.Path = "wscript"
    Action.Arguments= _
        """"&PathOfCurrenScript&"\备份还原开始菜单.vbs"" "&Argument&""
    
    CreateOrUpdate=6
    TaskPath="YuphizScript\"&UserName&"\"&title
    TaskName="自动备份"
    Call rootFolder.RegisterTaskDefinition( _
        TaskPath&"\"&TaskName, taskDefinition, CreateOrUpdate, _
        Empty , Empty,3)
end sub


' 禁用任务计划
sub disabledTasksch(tips)
    set Shell = CreateObject("Wscript.Shell") 
    UserName = CreateObject("WScript.Network").UserName
    TaskPath = "YuphizScript\"&UserName&"\"&title
    ThisTask = TaskPath&"\自动备份"
    if isnull(tips) then 
        WindowStyle = 1
    elseif tips="notips" then
        WindowStyle = 0
    end if 
    Shell.run( _
        "cmd /c "&_
        "@echo off &"&_
        "for %i in ("&ThisTask&") do (SCHTASKS /change /disable /tn %i)" _
    ),WindowStyle,true
    if isnull(tips) then Shell.popup "成功停用【"&title&"】",1
end sub


' 移除任务计划
sub removeTasksch(tips)
    UserName=CreateObject("WScript.Network").UserName
    set Shell=CreateObject("Wscript.Shell") 

    if not isnull(tips) then
        askdelete=Shell.popup( _
        "防误操作，真的要删除【"&title&"】吗？", _
        0, _
        "防误操作，请再确认",_
        1+48+256+4096 _
        )
        if askdelete=2 then wscript.quit
    end if

    set ShellTask=createobject("Schedule.Service")
    call ShellTask.connect()
    On Error Resume Next
    set rootFolder=ShellTask.getfolder("\YuphizScript\"&UserName)
    call rootFolder.DeleteTask(title&"\自动备份",0)
    rootFolder.deleteFolder title,0
    ' if err.number<> 0 then
    '     Shell.Popup "不存在自动备份的任务计划"
    '     exit sub
    ' end if
    On Error goto 0
    '  wscript.sleep 700
    ' Shell.popup _
    '   "成功卸载【"&title&"】"&vbcrlf&_
    '   "已删除全部的任务计划",_
    ' 3
end sub



' 关闭开始菜单进程
sub KillStartMenuProcess()
Set Shell=CreateObject("WScript.Shell")
set WinRC=getobject("winmgmts:\\.\root\cimv2")
set GetProcess=WinRC.execquery("select * from win32_process where name='StartMenuExperienceHost.exe'")
     if GetProcess.count>=1 then 
         Shell.run "taskkill /im StartMenuExperienceHost.exe /f",WindowStyle_Debug
     else
         Shell.run "taskkill /im ShellExperienceHost.exe /f",WindowStyle_Debug
     end if
end sub



' 获取系统版本号
function GetSystemVersion()
    set WinRC=getobject("winmgmts:\\.\root\cimv2")
    set GetSystemInfos=WinRC.execquery("select * from Win32_OperatingSystem")
    for each GetSystemInfo in GetSystemInfos
        SystemVersion=split(GetSystemInfo.Version,".")
    next
GetSystemVersion=SystemVersion
end function



'遍历文件夹并删除符合条件的文件夹
function DeleteFilesTree(PathSource)
set Shell=CreateObject("Wscript.Shell")
set FSO=CreateObject("Scripting.FileSystemObject")
    UserName=CreateObject("WScript.Network").UserName
    UserDomain=CreateObject("WScript.Network").UserDomain
    FolderNameKey=UserDomain&"_"&UserName

If Not FSO.FolderExists(PathSource) Then Wscript.quit

set Folders=FSO.GetFolder(PathSource)
set SubFolders=Folders.SubFolders

for each Oneof in SubFolders
    daycomparison=DateDiff("d",Oneof.DateLastModified,now)
    if daycomparison>=DeleteBeforeDay and InStr(1,Oneof.Name,FolderNameKey,1) > 0  and InStr(1,Oneof.Name,"_AutoBakvp",1) > 0  then
        message= _
            "文件夹：" & Folders &vbcrlf&_
            "文件：" & Oneof.Name &vbcrlf&_
            "文件最后修改日期：" & Oneof.DateLastModified &vbcrlf&_
            "文件相差日期：" & daycomparison
        msgBox message
        'wscript.quit
            if DeleteToReCycleBin = false then
                FSO.DeleteFolder Oneof,True
            elseif DeleteToReCycleBin = true then
                call DeleteFileToRecycle(Oneof,39)
             end if
    end if
next

end function



' 删除文件到回收站
function DeleteFileToRecycle(PathFile,IsConfirm)
    Set objReg=GetObject( _
"winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    objReg.GetBinaryValue &H80000001, _
        "Software\Microsoft\Windows\CurrentVersion\Explorer", _
        "ShellState", _
        ValueStateArray
    ValueBackupState=ValueStateArray 'P'
    
    ValueStateArray(4)=IsConfirm '删除确认，39为静默删除，35为删除前确认
    objReg.SetBinaryValue &H80000001, _
        "Software\Microsoft\Windows\CurrentVersion\Explorer", _
        "ShellState", _
        ValueStateArray
CreateObject("Shell.Application").NameSpace(0).ParseName(PathFile).InvokeVerb("delete")
    objReg.SetBinaryValue &H80000001, _
        "Software\Microsoft\Windows\CurrentVersion\Explorer", _
        "ShellState", _
        ValueBackupState
end function

' 读取unicode字符文件
Function ReadFile(FilePath)
    Dim String
    Set Stream = CreateObject("ADODB.stream")
    Stream.Type = 2
    Stream.mode = 3
    Stream.charset = "UTF-16LE"
    Stream.Open
    Stream.loadfromfile FilePath
    String = Stream.readtext
    Stream.Close
    Set Stream = Nothing
    ReadFile = String
End Function

' 写入unicode字符文件
Sub WriteFile(FilePath,Msg)
    Set fso = WScript.CreateObject("Scripting.Filesystemobject")
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = 2
    Stream.Open
    Stream.Charset = "UTF-16LE"
    Stream.Position = Stream.Size
    Stream.WriteText Msg
    Dim FolderArray,Folder
    FilePath = Replace(FilePath,"/","\")
    FolderArray = Split(FilePath, "\")
    If UBound(FolderArray) <> 0 Then
        For i = 0 To UBound(FolderArray)-1
            Folder = Folder & FolderArray(i)
            If fso.folderExists(Folder) = flase Then
                fso.createfolder (Folder)
            End If
            Folder = Folder & "\"
        Next
    End If
    Stream.SaveToFile FilePath, 2
    Stream.Close
    set Stream = nothing 'I'
End Sub

' 过滤想要的注册表
function FilterReg(contents,string,mode)
    contentsArray = split(contents,vbcrlf)
    Newcontents = "Windows Registry Editor Version 5.00"
    TotalLine = ubound(contentsArray)
    for i=0 to TotalLine
        if i < 3 and i > 0 then
            Newcontents = Newcontents &vbcrlf & contentsArray(i)
        elseif i >=3 then
            KeyInstr = Instr(1, contentsArray(i), string, 1)
            SectionInstr = Instr(1, contentsArray(i), "[", 1)

            if mode = "--include" then
                if KeyInstr > 0 or SectionInstr > 0 then Newcontents = Newcontents &vbcrlf & contentsArray(i)
            elseif mode = "--exclude" then
                if (contentsArray(i) <> "") and (KeyInstr <= 0 or SectionInstr > 0) then Newcontents = Newcontents &vbcrlf & contentsArray(i)
            end if
        end if
    next

    ' 移出空Section
    NewcontentsArray = split(Newcontents,vbcrlf)
    TotalLine = ubound(NewcontentsArray)
    Newcontents = "Windows Registry Editor Version 5.00"
    for i=0 to TotalLine
        if i < 3 and i > 0 then
            Newcontents = Newcontents &vbcrlf & NewcontentsArray(i)
        elseif i >=3 then
            NextLineCount = i+1
            if NextLineCount > TotalLine then NextLineCount = TotalLine
            Lastkey = Instr(1, NewcontentsArray(i), "[", 1)
            Section = Instr(1, NewcontentsArray(NextLineCount), "[", 1)

            if Lastkey <= 0 or isNull(Lastkey) then
                if Section <= 0 or isNull(Section) then
                    Newcontents = Newcontents &vbcrlf & NewcontentsArray(i)
                elseif Section > 0 then
                    Newcontents = Newcontents &vbcrlf & NewcontentsArray(i) &vbcrlf
                end if
            elseif Lastkey > 0 and (Section <= 0 or isNull(Section)) then
                Newcontents = Newcontents &vbcrlf & NewcontentsArray(i)
            end if
        end if
    next
FilterReg = Newcontents
end function

'重启资源管理器并重新打开上次的目录
sub RestartExplorer()    '重启资源管理器并重新打开上次的目录
    On Error Resume Next

    Dim ArrayPathFonders(), oAppShell, WindowOfoAppShell, Shell
    Set Dictionary = CreateObject("Scripting.Dictionary")
    Set oAppShell = CreateObject("Shell.Application")
    Set WindowOfoAppShell=oAppShell.Windows()
    Set Shell = CreateObject("WScript.Shell")

    Dictionary.Add "::{679F85CB-0220-4080-B29B-5540CC05AAB6}",null
    Dictionary.Add "::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",null
    Dictionary.Add "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}",null
    Dictionary.Add Shell.SpecialFolders("Desktop"),null
    Dictionary.Add Shell.SpecialFolders("AllUsersDesktop"),null
    
    n=-1
    For Each Oneof in WindowOfoAppShell
        if Instr(1, Oneof.FullName, "\explorer.exe", 1) > 0 Then
            If Not Dictionary.Exists(Oneof.Document.Folder.Self.Path) Then
                n = n + 1
                ReDim Preserve ArrayPathFonders(n)
                ArrayPathFonders(n) = Oneof.Document.Folder.Self.Path
                Dictionary.Add Oneof.Document.Folder.Self.Path ,null
            end if
        end if
    Next

    Shell.Run "Tskill explorer",WindowStyle_Debug,True
    if err.number = -2147024894 then 
        err.number = 0
        Shell.Run "taskkill /im explorer.exe /f",WindowStyle_Debug,True
        ' Shell.Run "explorer",0,True
        Shell.Run "cmd /c start explorer",WindowStyle_Debug,True
        if err.number = -2147024894 then
            msgbox "重启资源管理器失败，请手动重启"
            exit sub
        end if
    end if

For Each Oneof in ArrayPathFonders
    Shell.Run """"&Oneof&""""
Next
end sub
