param (
    $BackvpFolder,
    [switch]$WithTaskBar
)
$Host.UI.RawUI.WindowTitle = '正在备份……这不需要很久，请耐心等候'

Export-startlayout -path "$BackvpFolder\startmenulayout.xml"
<#
'声明
'脚本：备份还原开始菜单
'版本：beta 0.6.2
'说明：本脚本可以备份和还原开始菜单布局


'作者：YUPHIZ
'版权：此脚本版权归YUPHIZ所有
    '凡用此脚本从事法律不允许的事情的，均与本作者无关
    '此脚本遵循 gpl3.0 and later协议 #>
function Compress-Files($SourcePath,$ZipFilenPath){
    if (test-path $ZipFilenPath){remove-item $ZipFilenPath -Recurse -force}
    Add-Type -Assembly System.IO.Compression
    Add-Type -Assembly System.IO.Compression.FileSystem #H#
    $Archive = [System.IO.Compression.ZipFile]::open($ZipFilenPath,[System.IO.Compression.ZipArchiveMode]::update)
    $FrontStringCount = ($SourcePath.length)+1 #Y#
    $AllDir = (Get-ChildItem $SourcePath -Recurse -force  -ErrorAction Ignore).fullName
    $Total = $AllDir.count
    $count=0
    $zipTime = [System.Diagnostics.Stopwatch]::StartNew() #P#
    foreach ($OneOfFile in $Alldir){
        $count++
        $attribute = (Get-ItemProperty $OneOfFile).attributes
        $entry = $OneOfFile.Substring($FrontStringCount,($OneOfFile.length)-$FrontStringCount)
        if ($attribute -match [io.fileattributes]::Directory){
            if (!(get-ChildItem $OneOfFile -force -ErrorAction Ignore)) {continue}
            $entry = "$entry\"
            [void]$Archive.CreateEntry($entry)
        }else{
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive,$OneOfFile,$entry)
        }
        $Archive.getentry($entry).ExternalAttributes = $attribute #Z#
        if ($UnzipTime.Elapsed.TotalMilliseconds -ge 500) { 
            [console]::write("已完成{0} / {1}",$count,$Total)
            $zipTime.Reset(); $zipTime.Start()
        }
    }
    $Archive.Dispose()
} #I#

$StartUserFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
$StartAllUserFolder = "$env:ProgramData\Microsoft\Windows\Start Menu"

Compress-Files $StartUserFolder "$BackvpFolder\开始菜单快捷方式User.zip" #U#
Compress-Files $StartAllUserFolder "$BackvpFolder\开始菜单快捷方式Alluser.zip"
if ($WithTaskBar){
    $TaskBarFolder = "$env:AppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    Compress-Files $TaskBarFolder "$BackvpFolder\任务栏快捷方式.zip"
}

write-host 备份结束