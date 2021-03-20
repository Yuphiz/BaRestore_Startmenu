'═════════════════════ 设置 模块开始 ═════════════════════
Path=""
'备份文件放在哪里，空值就和脚本在同一个文件夹

BeforeRestore=1
'还原前是否备份

SeletcRestoreFolderByCustom=0
'自由选择要还原的备份文件夹，0是只加载【开始菜单备份】里面的文件夹，1是可以自由选择文件夹，不过要自己辨别文件夹里面的文件是否正确，建议默认0

TriggerDay=30
'多少天自动备份一次，开启功能3才生效，修改完需要再次启用功能3才会刷新

DeleteBeforeDay=100000000000
'警告：此脚本是alpah版，建议先不开启这个自动清理功能，等测试bug
'自动清理多少天前的备份，0为不自动清理，不自动清理请注意文件夹大小，快捷方式多的话，一个备份有2m左右大小，开启功能3才生效
'注意，手动备份和自动备份的名称不一样，自动清理不会清理手动备份

deleteTo=1
'清理的备份文件删除到哪里，1是放到回收站，0是直接删除，开启自动清理才生效
'————————————— 设置 模块结束 ——————————————




'声明
'脚本：备份还原开始菜单
'版本：Alpha 0.2
'说明：本脚本可以备份和还原开始菜单布局


'作者：YUPHIZ
'版权：此脚本版权归YUPHIZ所有，备份还原方法借鉴BackupSML和winaero tweaker，并优化改进
          '凡用此脚本从事法律不允许的事情的，均与本作者无关
          '此脚本遵循 CC BY-NC-SA 4.0协议(署名-禁止商业-相同方式共享)



' 判断系统版本号
SysVersion=GetSystemVersion()
if SysVersion(0)<>10 then 
      msgbox "本脚本只支持windows10系统",16
      wscript.quit
end if

if SysVersion(2)<15063 then
      msgbox "你的win10版本过低，脚本暂不支持1703以前的系统",16
      wscript.quit
end if


if path="" then
      set FSO=createobject("Scripting.FileSystemObject")
      PathOfCurrenScript = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
      Path=PathOfCurrenScript
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



' 【还原2.4】需要的管理员
select case WScript.Arguments.count
      case 2
            if WScript.Arguments(1)="--Restore" then 
                  call RestoreStartLayout(Path,"--AllUser") '管理员运行
                  Wscript.quit
            end if 
      case 1
            if WScript.Arguments(0) = "--HiddenBackup" then
                  call BackUpStartLayout(Path,"--Auto","notips")
                  if DeleteBeforeDay > 0 then
                        call DeleteFilesTree(Path&"\开始菜单备份")
                  end if
                  Wscript.quit
            end if
end select


' 主界面
call SatrtLayout()
function SatrtLayout()
      Ask=inputbox( _
            "※ 【备份】"&vbcrlf&vbcrlf& _
            "    1  备 份 开始菜单"&vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
            "※ 【还原】"&vbcrlf&vbcrlf& _
            "    2  只 还 原 已安装快捷方式 "&vbcrlf&vbcrlf& _
            "        2.2   还 原 个人用户快捷方式 "&vbcrlf&vbcrlf& _
            "        2.4   还 原 所有快捷方式 "&vbcrlf&vbcrlf& _
            "        2.9   重 置 开始菜单"&vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
            "※ 【定时备份】"&vbcrlf&vbcrlf&_
            "    3  开 启 定时自动备份 "&vbcrlf&vbcrlf&_
            "    4  禁 用 定时自动备份 "&vbcrlf&vbcrlf&_
            "    5  卸 载 定时自动备份 "&vbcrlf&vbcrlf,_
            title&" 小工具",_
            "请输入对应的序号(1或2或3)")
      select case True
            case Ask=""
                  Wscript.quit
                  msgbox "空值"
           case Ask="1"
                 call BackUpStartLayout(Path,null,null)
            case Ask="2"
                  call RestoreStartLayout(Path,null)
            case Ask="2.2"
                  call RestoreStartLayout(Path,"--User")
            case Ask="2.4"
                  call RunAs()
            case Ask="2.9"
                  call ResetStartLayout()
            case Ask="3"
                  call enableTasksch()
            case Ask="4"
                  call disabledTasksch()
            case Ask="5"
                  call removeTasksch()
            case else
                  call SatrtLayout()
                  Wscript.quit
      end select
end function


' 备份开始菜单
function BackUpStartLayout(Path,Folder,tips)
      set Shell=CreateObject("Wscript.shell")
      Set Env=Shell.Environment("Process")
      set FSO=CreateObject("scripting.filesystemobject")
      UserName=CreateObject("WScript.Network").UserName
      UserDomain=CreateObject("WScript.Network").UserDomain
      SystemDrive=Env.Item("SystemDrive")
      
      FileLayout = SystemDrive&"\users\"&UserName&"\AppData\Local\Microsoft\Windows\Shell\DefaultLayouts.xml"

      RegPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
      RootFolder=Path&"\开始菜单备份\"
      if not FSO.FolderExists(RootFolder) then
            Call CreateFolder(RootFolder) 'I'
      end if
       'Y'
      if IsNull(Folder) Then
            Folder = RootFolder&UserDomain&"_"&UserName&"_"&year(now)&"_"&Month(now)&"_"&day(now)&"_"&int(timer)
      elseif Folder="--Auto" then
            Folder = RootFolder&UserDomain&"_"&UserName&"_"&year(now)&"_"&Month(now)&"_"&day(now)&"_"&int(timer)&"_AutoBakvp"
    
      end if
      
      
      startFolderUser=Shell.SpecialFolders("programs")
      startFolderAllUser=Shell.SpecialFolders("AllUsersStartMenu")
      
      if not FSO.FolderExists(Folder) then
            Call CreateFolder(Folder)
      end if

      RegFile=Folder&"\开始菜单备份.reg"
      startmenulayout=Folder&"\startmenulayout.xml"
      
      
      Shell.run "cmd /c echo y|REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount """&RegFile&"""" ,0,true
      if not FSO.FileExists(RegFile) then
            Shell.popup "不能备份注册表，可能脚本不支持此系统版本",0,"错误",16
      end if
      
      Shell.run "powershell Export-startlayout -path """""""& startmenulayout &" """""" " ,0
      
      if  FSO.FileExists(FileLayout) then
      FSO.CopyFile FileLayout, Folder&"\DefaultLayouts.xml" 'P'
      end if
      
      
      Shell.run "powershell  Compress-Archive '" & startFolderUser & "\*' ' "&Folder& "\开始菜单快捷方式User.zip' -force",0
      Shell.run "powershell  Compress-Archive '" & startFolderAllUser & "\*' ' "&Folder& "\开始菜单快捷方式Alluser.zip' -force",0

      if IsNull(tips) then Shell.popup "操作完成",1
end function




' 还原开始菜单
function RestoreStartLayout(Path,Restorelnk)
      set Shell=CreateObject("Wscript.shell")
      Set Env=Shell.Environment("Process")
      
      UserName=CreateObject("WScript.Network").UserName
      UserDomain=CreateObject("WScript.Network").UserDomain
      set FSO=CreateObject("scripting.filesystemobject")
      SystemDrive=Env.Item("SystemDrive")
      FileLayout = SystemDrive&"\users\"&UserName&"\AppData\Local\Microsoft\Windows\Shell\DefaultLayouts.xml"
      if SeletcRestoreFolderByCustom=0 then
            RestoreFolder=SelectFolder(Path&"\开始菜单备份",Null,null)
      elseif SeletcRestoreFolderByCustom=1 then
            RestoreFolder=SelectFolder(Null,Null,null)
      end if
      startFolderUser=Shell.SpecialFolders("programs")
      startFolderAllUser=Shell.SpecialFolders("AllUsersStartMenu")


      ' 还原前备份
      if BeforeRestore=1 then
            BeforeRestoreFolder=Path&"\开始菜单备份\————————"&UserDomain&"_"&UserName&"_还原前的备份"
            call BackUpStartLayout(Path,BeforeRestoreFolder,"notips")
      end if

      if Restorelnk="--User" then   
            Shell.run "powershell $Host.UI.RawUI.WindowTitle = '正在还原……这不需要很久，请耐心等候'; Expand-Archive '"&RestoreFolder&"\开始菜单快捷方式User.zip' '"&startFolderUser&"'   -ErrorAction 'SilentlyContinue' ",1,true
      elseif Restorelnk="--AllUser" then   
            Shell.run "powershell $Host.UI.RawUI.WindowTitle = '正在还原……这不需要很久，请耐心等候';Expand-Archive '"&RestoreFolder&"\开始菜单快捷方式User.zip' '"&startFolderUser&"'   -ErrorAction 'SilentlyContinue'; Expand-Archive '"&RestoreFolder&"\开始菜单快捷方式Alluser.zip' '"&startFolderAllUser&"'   -ErrorAction 'SilentlyContinue' ",1,true
      end if 'Z'

      if  FSO.FileExists( RestoreFolder&"\DefaultLayouts.xml") then
            FSO.CopyFile RestoreFolder&"\DefaultLayouts.xml",FileLayout
      end if
      
      RegFile=RestoreFolder&"\开始菜单备份.reg"
      
      Shell.run "REG IMPORT """&RegFile&"""" ,0,true
      call KillStartMenuProcess()
Shell.popup "操作完成",1
end function


' 选择还原的文件夹
Function SelectFolder(default,title,Restorelnk)
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
        call RestoreStartLayout(Path,Restorelnk)
        Wscript.quit
    else
        SelectFolder = Folder.Self.Path
    End If
End Function



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



' 重置开始菜单
Function ResetStartLayout()
      set FSO=CreateObject("scripting.filesystemobject")
      
      UserName=CreateObject("WScript.Network").UserName
      UserDomain=CreateObject("WScript.Network").UserDomain

      ' 重置前备份
      if BeforeRestore=1 then
            BeforeRestoreFolder=Path&"\开始菜单备份\————————"&UserDomain&"_"&UserName&"_还原前的备份"
            call BackUpStartLayout(Path,BeforeRestoreFolder,"notips")
      end if

      set Shell = CreateObject("Wscript.shell")
      Set Env = Shell.Environment("Process")
      SystemDrive = Env.Item("SystemDrive")
      FileLayout = SystemDrive&"\users\"&UserName&"\AppData\Local\Microsoft\Windows\Shell\DefaultLayouts.xml"
      SysVersion = GetSystemVersion()
      if  FSO.FileExists(FileLayout) and SysVersion(2)<17744 then
            FSO.deleteFile FileLayout
      end if

      set Shell=CreateObject("Wscript.shell")
      
      Shell.run "reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount /f",0,true
      call KillStartMenuProcess()
      Shell.popup "操作完成",1
end function


' 管理员运行
function RunAs()
    set Shell=CreateObject("Shell.Application") 
    Shell.ShellExecute "wscript.exe" _ 
    , """" & WScript.ScriptFullName & """ RunAsAdministrator --Restore", , "runas", 1 
    WScript.Quit 
end function


' 启用任务计划
sub enableTasksch()  '启用本脚本任务计划
      set FSO=createobject("Scripting.FileSystemObject")
      PathOfCurrenScript = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
      UserName=CreateObject("WScript.Network").UserName
      UserDomain=CreateObject("WScript.Network").UserDomain
      
      call BackUpStartLayout(Path,"--Auto",null)
      
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
      daynow=day(now) 'U'
      if len(daynow)=1 then daynow="0"&day(now)
      startTime = year(now)&"-"&monthnow&"-"&daynow&"T02:00:00"
      trigger.StartBoundary = startTime
      
      trigger.DaysInterval = TriggerDay
      trigger.Id = "DailyTriggerId"
      trigger.Enabled = True
      Set Action = taskDefinition.Actions.Create(0)
      Action.Path = "wscript"
      Action.Arguments= _
            """"&PathOfCurrenScript&"\备份还原开始菜单.vbs"" --HiddenBackup"
      CreateOrUpdate=6
      TaskPath="YuphizScript\"&UserName&"\"&title
      TaskName="自动备份"
      Call rootFolder.RegisterTaskDefinition( _
            TaskPath&"\"&TaskName, taskDefinition, CreateOrUpdate, _
            Empty , Empty,3)
end sub


' 禁用任务计划
sub disabledTasksch()
    set Shell=CreateObject("Wscript.Shell") 
      UserName=CreateObject("WScript.Network").UserName
TaskPath="YuphizScript\"&UserName&"\"&title
ThisTask=TaskPath&"\自动备份"
            Shell.run( _
            "cmd /c "&_
            "@echo off &"&_
            "for %i in ("&ThisTask&") do (SCHTASKS /change /disable /tn %i)" _
      ),1,true
    Shell.popup _
        "成功停用【"&title&"】",1
end sub


' 移除任务计划
sub removeTasksch()
UserName=CreateObject("WScript.Network").UserName
    set Shell=CreateObject("Wscript.Shell") 


    askdelete=Shell.popup( _
        "防误操作，真的要删除【"&title&"】吗？", _
        0, _
        "防误操作，请再确认",_
        1+48+256+4096 _
    )
    if askdelete=2 then wscript.quit

      set ShellTask=createobject("Schedule.Service")
      call ShellTask.connect()
      set rootFolder=ShellTask.getfolder("\YuphizScript\"&UserName)
      set taskDefinition=ShellTask.NewTask(0)
       
            call rootFolder.DeleteTask(title&"\自动备份",0)
      On Error goto 0
      rootFolder.deleteFolder title,0
     wscript.sleep 700
    Shell.popup _
        "成功卸载【"&title&"】"&vbcrlf&_
        "已删除全部的任务计划",_
    3
end sub



' 关闭开始菜单进程
sub KillStartMenuProcess()
Set Shell=CreateObject("WScript.Shell")
set WinRC=getobject("winmgmts:\\.\root\cimv2")
set GetProcess=WinRC.execquery("select * from win32_process where name='StartMenuExperienceHost.exe'")
     if GetProcess.count>=1 then 
             Shell.run "taskkill /im StartMenuExperienceHost.exe /f",0
     else
             Shell.run "taskkill /im ShellExperienceHost.exe /f",0
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
      if daycomparison>=DeleteBeforeDay and InStr(1,Oneof.Name,FolderNameKey,1)  and InStr(1,Oneof.Name,"_AutoBakvp",1)  then
            message= _
                "文件夹：" & Folders &vbcrlf&_
                "文件：" & Oneof.Name &vbcrlf&_
                "文件最后修改日期：" & Oneof.DateLastModified &vbcrlf&_
                "文件相差日期：" & daycomparison
                  if DeleteTo=0 then
                        FSO.DeleteFolder Oneof,True
                  elseif DeleteTo=1 then
                        call DeleteFileToRecycle(Oneof,39)
                   end if
      end if
next 'H'
end function



' 删除文件到回收站 函数开始
function DeleteFileToRecycle(PathFile,IsConfirm)
      Set objReg=GetObject( _
"winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
      objReg.GetBinaryValue &H80000001, _
            "Software\Microsoft\Windows\CurrentVersion\Explorer", _
            "ShellState", _
            ValueStateArray
      ValueBackupState=ValueStateArray
      
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
