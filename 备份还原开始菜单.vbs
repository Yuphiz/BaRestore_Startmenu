'�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T ���� ģ�鿪ʼ �T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T
Path=""
'�����ļ����������ֵ�ͺͽű���ͬһ���ļ���

BeforeRestore=1
'��ԭǰ�Ƿ񱸷�

SeletcRestoreFolderByCustom=0
'����ѡ��Ҫ��ԭ�ı����ļ��У�0��ֻ���ء���ʼ�˵����ݡ�������ļ��У�1�ǿ�������ѡ���ļ��У�����Ҫ�Լ�����ļ���������ļ��Ƿ���ȷ������Ĭ��0

TriggerDay=30
'�������Զ�����һ�Σ���������3����Ч���޸�����Ҫ�ٴ����ù���3�Ż�ˢ��

DeleteBeforeDay=100000000000
'���棺�˽ű���alpah�棬�����Ȳ���������Զ������ܣ��Ȳ���bug
'�Զ����������ǰ�ı��ݣ�0Ϊ���Զ��������Զ�������ע���ļ��д�С����ݷ�ʽ��Ļ���һ��������2m���Ҵ�С����������3����Ч
'ע�⣬�ֶ����ݺ��Զ����ݵ����Ʋ�һ�����Զ������������ֶ�����

deleteTo=1
'����ı����ļ�ɾ�������1�Ƿŵ�����վ��0��ֱ��ɾ���������Զ��������Ч
'�������������������������� ���� ģ����� ����������������������������




'����
'�ű������ݻ�ԭ��ʼ�˵�
'�汾��Alpha 0.2
'˵�������ű����Ա��ݺͻ�ԭ��ʼ�˵�����


'���ߣ�YUPHIZ
'��Ȩ���˽ű���Ȩ��YUPHIZ���У����ݻ�ԭ�������BackupSML��winaero tweaker�����Ż��Ľ�
          '���ô˽ű����·��ɲ����������ģ����뱾�����޹�
          '�˽ű���ѭ CC BY-NC-SA 4.0Э��(����-��ֹ��ҵ-��ͬ��ʽ����)



' �ж�ϵͳ�汾��
SysVersion=GetSystemVersion()
if SysVersion(0)<>10 then 
      msgbox "���ű�ֻ֧��windows10ϵͳ",16
      wscript.quit
end if

if SysVersion(2)<15063 then
      msgbox "���win10�汾���ͣ��ű��ݲ�֧��1703��ǰ��ϵͳ",16
      wscript.quit
end if


if path="" then
      set FSO=createobject("Scripting.FileSystemObject")
      PathOfCurrenScript = FSO.GetFile(Wscript.ScriptFullName).ParentFolder.Path
      Path=PathOfCurrenScript
end if



' ���������ļ���
title="��ʼ�˵�����"
set FSO=CreateObject("scripting.filesystemobject")
if not FSO.FolderExists(Path) then
      Ask=CreateObject("Wscript.Shell").Popup( _
            "δ�ҵ�����·��"& vbcrlf & path & vbcrlf &_
            "�Ƿ񴴽���",0,"��ѡ��",1+16+512 _
      )
      if Ask=1 then
            call CreateFolder(Path)
            if not FSO.FolderExists(Path) then
                   msgbox "�����ļ���ʧ�ܣ����ֶ�����",1+16
                   Wscript.quit
            end if
    else
            Wscript.quit
    end if
end if



' ����ԭ2.4����Ҫ�Ĺ���Ա
select case WScript.Arguments.count
      case 2
            if WScript.Arguments(1)="--Restore" then 
                  call RestoreStartLayout(Path,"--AllUser") '����Ա����
                  Wscript.quit
            end if 
      case 1
            if WScript.Arguments(0) = "--HiddenBackup" then
                  call BackUpStartLayout(Path,"--Auto","notips")
                  if DeleteBeforeDay > 0 then
                        call DeleteFilesTree(Path&"\��ʼ�˵�����")
                  end if
                  Wscript.quit
            end if
end select


' ������
call SatrtLayout()
function SatrtLayout()
      Ask=inputbox( _
            "�� �����ݡ�"&vbcrlf&vbcrlf& _
            "    1  �� �� ��ʼ�˵�"&vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
            "�� ����ԭ��"&vbcrlf&vbcrlf& _
            "    2  ֻ �� ԭ �Ѱ�װ��ݷ�ʽ "&vbcrlf&vbcrlf& _
            "        2.2   �� ԭ �����û���ݷ�ʽ "&vbcrlf&vbcrlf& _
            "        2.4   �� ԭ ���п�ݷ�ʽ "&vbcrlf&vbcrlf& _
            "        2.9   �� �� ��ʼ�˵�"&vbcrlf&vbcrlf&vbcrlf&vbcrlf& _
            "�� ����ʱ���ݡ�"&vbcrlf&vbcrlf&_
            "    3  �� �� ��ʱ�Զ����� "&vbcrlf&vbcrlf&_
            "    4  �� �� ��ʱ�Զ����� "&vbcrlf&vbcrlf&_
            "    5  ж �� ��ʱ�Զ����� "&vbcrlf&vbcrlf,_
            title&" С����",_
            "�������Ӧ�����(1��2��3)")
      select case True
            case Ask=""
                  Wscript.quit
                  msgbox "��ֵ"
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


' ���ݿ�ʼ�˵�
function BackUpStartLayout(Path,Folder,tips)
      set Shell=CreateObject("Wscript.shell")
      Set Env=Shell.Environment("Process")
      set FSO=CreateObject("scripting.filesystemobject")
      UserName=CreateObject("WScript.Network").UserName
      UserDomain=CreateObject("WScript.Network").UserDomain
      SystemDrive=Env.Item("SystemDrive")
      
      FileLayout = SystemDrive&"\users\"&UserName&"\AppData\Local\Microsoft\Windows\Shell\DefaultLayouts.xml"

      RegPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
      RootFolder=Path&"\��ʼ�˵�����\"
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

      RegFile=Folder&"\��ʼ�˵�����.reg"
      startmenulayout=Folder&"\startmenulayout.xml"
      
      
      Shell.run "cmd /c echo y|REG EXPORT HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount """&RegFile&"""" ,0,true
      if not FSO.FileExists(RegFile) then
            Shell.popup "���ܱ���ע������ܽű���֧�ִ�ϵͳ�汾",0,"����",16
      end if
      
      Shell.run "powershell Export-startlayout -path """""""& startmenulayout &" """""" " ,0
      
      if  FSO.FileExists(FileLayout) then
      FSO.CopyFile FileLayout, Folder&"\DefaultLayouts.xml" 'P'
      end if
      
      
      Shell.run "powershell  Compress-Archive '" & startFolderUser & "\*' ' "&Folder& "\��ʼ�˵���ݷ�ʽUser.zip' -force",0
      Shell.run "powershell  Compress-Archive '" & startFolderAllUser & "\*' ' "&Folder& "\��ʼ�˵���ݷ�ʽAlluser.zip' -force",0

      if IsNull(tips) then Shell.popup "�������",1
end function




' ��ԭ��ʼ�˵�
function RestoreStartLayout(Path,Restorelnk)
      set Shell=CreateObject("Wscript.shell")
      Set Env=Shell.Environment("Process")
      
      UserName=CreateObject("WScript.Network").UserName
      UserDomain=CreateObject("WScript.Network").UserDomain
      set FSO=CreateObject("scripting.filesystemobject")
      SystemDrive=Env.Item("SystemDrive")
      FileLayout = SystemDrive&"\users\"&UserName&"\AppData\Local\Microsoft\Windows\Shell\DefaultLayouts.xml"
      if SeletcRestoreFolderByCustom=0 then
            RestoreFolder=SelectFolder(Path&"\��ʼ�˵�����",Null,null)
      elseif SeletcRestoreFolderByCustom=1 then
            RestoreFolder=SelectFolder(Null,Null,null)
      end if
      startFolderUser=Shell.SpecialFolders("programs")
      startFolderAllUser=Shell.SpecialFolders("AllUsersStartMenu")


      ' ��ԭǰ����
      if BeforeRestore=1 then
            BeforeRestoreFolder=Path&"\��ʼ�˵�����\����������������"&UserDomain&"_"&UserName&"_��ԭǰ�ı���"
            call BackUpStartLayout(Path,BeforeRestoreFolder,"notips")
      end if

      if Restorelnk="--User" then   
            Shell.run "powershell $Host.UI.RawUI.WindowTitle = '���ڻ�ԭ�����ⲻ��Ҫ�ܾã������ĵȺ�'; Expand-Archive '"&RestoreFolder&"\��ʼ�˵���ݷ�ʽUser.zip' '"&startFolderUser&"'   -ErrorAction 'SilentlyContinue' ",1,true
      elseif Restorelnk="--AllUser" then   
            Shell.run "powershell $Host.UI.RawUI.WindowTitle = '���ڻ�ԭ�����ⲻ��Ҫ�ܾã������ĵȺ�';Expand-Archive '"&RestoreFolder&"\��ʼ�˵���ݷ�ʽUser.zip' '"&startFolderUser&"'   -ErrorAction 'SilentlyContinue'; Expand-Archive '"&RestoreFolder&"\��ʼ�˵���ݷ�ʽAlluser.zip' '"&startFolderAllUser&"'   -ErrorAction 'SilentlyContinue' ",1,true
      end if 'Z'

      if  FSO.FileExists( RestoreFolder&"\DefaultLayouts.xml") then
            FSO.CopyFile RestoreFolder&"\DefaultLayouts.xml",FileLayout
      end if
      
      RegFile=RestoreFolder&"\��ʼ�˵�����.reg"
      
      Shell.run "REG IMPORT """&RegFile&"""" ,0,true
      call KillStartMenuProcess()
Shell.popup "�������",1
end function


' ѡ��ԭ���ļ���
Function SelectFolder(default,title,Restorelnk)
    If IsNull(default) Then
        default = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    End If
    If IsNull(title) Then
        title = "ѡ��Ҫ��ԭ���ļ���Ŀ¼��ѡ���Ŀ¼��ˢ��"
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



' �����ļ���
Function CreateFolder(path)
    set FSO=CreateObject("scripting.filesystemobject")
    Set getDrivers=FSO.Drives
    
    Set getDrive=FSO.GetDrive(FSO.GetDriveName(path))
    If Not FSO.FolderExists(fso.GetParentFolderName(path)) Then
        CreateFolder FSO.GetParentFolderName(path)
    End If
    FSO.CreateFolder(path)
End Function



' ���ÿ�ʼ�˵�
Function ResetStartLayout()
      set FSO=CreateObject("scripting.filesystemobject")
      
      UserName=CreateObject("WScript.Network").UserName
      UserDomain=CreateObject("WScript.Network").UserDomain

      ' ����ǰ����
      if BeforeRestore=1 then
            BeforeRestoreFolder=Path&"\��ʼ�˵�����\����������������"&UserDomain&"_"&UserName&"_��ԭǰ�ı���"
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
      Shell.popup "�������",1
end function


' ����Ա����
function RunAs()
    set Shell=CreateObject("Shell.Application") 
    Shell.ShellExecute "wscript.exe" _ 
    , """" & WScript.ScriptFullName & """ RunAsAdministrator --Restore", , "runas", 1 
    WScript.Quit 
end function


' ��������ƻ�
sub enableTasksch()  '���ñ��ű�����ƻ�
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
            """"&PathOfCurrenScript&"\���ݻ�ԭ��ʼ�˵�.vbs"" --HiddenBackup"
      CreateOrUpdate=6
      TaskPath="YuphizScript\"&UserName&"\"&title
      TaskName="�Զ�����"
      Call rootFolder.RegisterTaskDefinition( _
            TaskPath&"\"&TaskName, taskDefinition, CreateOrUpdate, _
            Empty , Empty,3)
end sub


' ��������ƻ�
sub disabledTasksch()
    set Shell=CreateObject("Wscript.Shell") 
      UserName=CreateObject("WScript.Network").UserName
TaskPath="YuphizScript\"&UserName&"\"&title
ThisTask=TaskPath&"\�Զ�����"
            Shell.run( _
            "cmd /c "&_
            "@echo off &"&_
            "for %i in ("&ThisTask&") do (SCHTASKS /change /disable /tn %i)" _
      ),1,true
    Shell.popup _
        "�ɹ�ͣ�á�"&title&"��",1
end sub


' �Ƴ�����ƻ�
sub removeTasksch()
UserName=CreateObject("WScript.Network").UserName
    set Shell=CreateObject("Wscript.Shell") 


    askdelete=Shell.popup( _
        "������������Ҫɾ����"&title&"����", _
        0, _
        "�������������ȷ��",_
        1+48+256+4096 _
    )
    if askdelete=2 then wscript.quit

      set ShellTask=createobject("Schedule.Service")
      call ShellTask.connect()
      set rootFolder=ShellTask.getfolder("\YuphizScript\"&UserName)
      set taskDefinition=ShellTask.NewTask(0)
       
            call rootFolder.DeleteTask(title&"\�Զ�����",0)
      On Error goto 0
      rootFolder.deleteFolder title,0
     wscript.sleep 700
    Shell.popup _
        "�ɹ�ж�ء�"&title&"��"&vbcrlf&_
        "��ɾ��ȫ��������ƻ�",_
    3
end sub



' �رտ�ʼ�˵�����
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



' ��ȡϵͳ�汾��
function GetSystemVersion()
      set WinRC=getobject("winmgmts:\\.\root\cimv2")
      set GetSystemInfos=WinRC.execquery("select * from Win32_OperatingSystem")
      for each GetSystemInfo in GetSystemInfos
            SystemVersion=split(GetSystemInfo.Version,".")
      next
GetSystemVersion=SystemVersion
end function



'�����ļ��в�ɾ�������������ļ���
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
                "�ļ��У�" & Folders &vbcrlf&_
                "�ļ���" & Oneof.Name &vbcrlf&_
                "�ļ�����޸����ڣ�" & Oneof.DateLastModified &vbcrlf&_
                "�ļ�������ڣ�" & daycomparison
                  if DeleteTo=0 then
                        FSO.DeleteFolder Oneof,True
                  elseif DeleteTo=1 then
                        call DeleteFileToRecycle(Oneof,39)
                   end if
      end if
next 'H'
end function



' ɾ���ļ�������վ ������ʼ
function DeleteFileToRecycle(PathFile,IsConfirm)
      Set objReg=GetObject( _
"winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
      objReg.GetBinaryValue &H80000001, _
            "Software\Microsoft\Windows\CurrentVersion\Explorer", _
            "ShellState", _
            ValueStateArray
      ValueBackupState=ValueStateArray
      
      ValueStateArray(4)=IsConfirm 'ɾ��ȷ�ϣ�39Ϊ��Ĭɾ����35Ϊɾ��ǰȷ��
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
