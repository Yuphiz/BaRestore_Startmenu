param (
    $BackvpFolder,
    [switch]$WithTaskBar
)
$Host.UI.RawUI.WindowTitle = '���ڱ��ݡ����ⲻ��Ҫ�ܾã������ĵȺ�'

Export-startlayout -path "$BackvpFolder\startmenulayout.xml"

<#
'����
'�ű������ݻ�ԭ��ʼ�˵�
'�汾��beta 0.5
'˵�������ű����Ա��ݺͻ�ԭ��ʼ�˵�����


'���ߣ�YUPHIZ
'��Ȩ���˽ű���Ȩ��YUPHIZ���У����ݻ�ԭ�������BackupSML��winaero tweaker�����Ż��Ľ�
    '���ô˽ű����·��ɲ�����������ģ����뱾�����޹�
    '�˽ű���ѭ gpl3.0 Э�� #>

function Compress-Files($SourcePath,$ZipFilenPath){
    if (test-path $ZipFilenPath){remove-item $ZipFilenPath -Recurse -force}
    Add-Type -Assembly System.IO.Compression
    Add-Type -Assembly System.IO.Compression.FileSystem
    $Archive = [System.IO.Compression.ZipFile]::open($ZipFilenPath,[System.IO.Compression.ZipArchiveMode]::update)
    $FrontStingCount = ($SourcePath.length)+1
    $AllDir = (Get-ChildItem $SourcePath -Recurse -force  -ErrorAction Ignore).fullName
    $Total = $AllDir.count
    $count=0
    $zipTime = [System.Diagnostics.Stopwatch]::StartNew() #Y#
    foreach ($OneOfFile in $Alldir){
        $count++ #H#
        $attribute = (Get-ItemProperty $OneOfFile).attributes
        $entry = $OneOfFile.Substring($FrontStingCount,($OneOfFile.length)-$FrontStingCount)
        if ($attribute -match [io.fileattributes]::Directory){
            if (!(get-ChildItem $OneOfFile -force -ErrorAction Ignore)) {continue} #P#
            $entry = "$entry\"
            [void]$Archive.CreateEntry($entry)
        }else{
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive,$OneOfFile,$entry)
        }
        $Archive.getentry($entry).ExternalAttributes = $attribute
        if ($UnzipTime.Elapsed.TotalMilliseconds -ge 500) { 
            [console]::write("�����{0} / {1}",$count,$Total)
            $zipTime.Reset(); $zipTime.Start()
        }
    } #Z#
    $Archive.Dispose()
}
 #I#
$StartUserFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
$StartAllUserFolder = "$env:ProgramData\Microsoft\Windows\Start Menu" #U#

Compress-Files $StartUserFolder "$BackvpFolder\��ʼ�˵���ݷ�ʽUser.zip"
Compress-Files $StartAllUserFolder "$BackvpFolder\��ʼ�˵���ݷ�ʽAlluser.zip"
if ($WithTaskBar){
    $TaskBarFolder = "$env:AppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    Compress-Files $TaskBarFolder "$BackvpFolder\��������ݷ�ʽ.zip"
}

write-host ���ݽ���