param(
    [Parameter(Mandatory = $true)]
    [string]$ParentFolder,

    [Parameter(Mandatory = $true)]
    [string]$LessonTitle,

    [datetime]$StartTime = (Get-Date)
)

function ConvertFrom-UnicodeCodes {
    param([int[]]$Codes)
    return -join ($Codes | ForEach-Object { [char]$_ })
}

$resolvedParent = (Resolve-Path -LiteralPath $ParentFolder).Path
if (-not (Test-Path -LiteralPath $resolvedParent -PathType Container)) {
    throw 'ParentFolder must be an existing directory.'
}

$safeTitle = ($LessonTitle -replace '[\\/:*?"<>|]', '-' -replace '\s+', ' ').Trim(' ', '.', '-')
if ([string]::IsNullOrWhiteSpace($safeTitle)) {
    throw 'LessonTitle must contain a meaningful title after sanitization.'
}

$stamp = $StartTime.ToString('yyyyMMdd-HHmmss')
$baseName = "${stamp}_Trae_${safeTitle}"
$taskRoot = Join-Path $resolvedParent $baseName
$suffix = 2
while (Test-Path -LiteralPath $taskRoot) {
    $taskRoot = Join-Path $resolvedParent ("{0}-{1:D2}" -f $baseName, $suffix)
    $suffix++
}

$taskRoot = [System.IO.Path]::GetFullPath($taskRoot)
if (-not $taskRoot.StartsWith($resolvedParent, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Resolved task directory escaped ParentFolder.'
}

New-Item -ItemType Directory -Path $taskRoot | Out-Null
$workFilesName = ConvertFrom-UnicodeCodes @(0x5DE5,0x4F5C,0x6587,0x4EF6)
$networkSourcesName = ConvertFrom-UnicodeCodes @(0x7F51,0x7EDC,0x8865,0x5145,0x8D44,0x6599)
$workspaceInfoName = ConvertFrom-UnicodeCodes @(0x5DE5,0x4F5C,0x533A,0x4FE1,0x606F,0x002E,0x006D,0x0064)
New-Item -ItemType Directory -Path (Join-Path $taskRoot $workFilesName) | Out-Null
New-Item -ItemType Directory -Path (Join-Path $taskRoot $networkSourcesName) | Out-Null

$workspaceHeading = ConvertFrom-UnicodeCodes @(0x0054,0x0072,0x0061,0x0065,0x5907,0x8BFE,0x5DE5,0x4F5C,0x533A)
$lessonLabel = ConvertFrom-UnicodeCodes @(0x8BFE,0x6807,0x9898)
$createdLabel = ConvertFrom-UnicodeCodes @(0x521B,0x5EFA,0x65F6,0x95F4)
$sourceRootLabel = ConvertFrom-UnicodeCodes @(0x539F,0x59CB,0x8D44,0x6599,0x6839,0x76EE,0x5F55)
$workspaceRootLabel = ConvertFrom-UnicodeCodes @(0x5DE5,0x4F5C,0x533A,0x6839,0x76EE,0x5F55)
$stageLabel = ConvertFrom-UnicodeCodes @(0x5F53,0x524D,0x9636,0x6BB5)
$curationStage = ConvertFrom-UnicodeCodes @(0x8D44,0x6599,0x7B5B,0x9009)
$workspaceRule = ConvertFrom-UnicodeCodes @(0x672C,0x5DE5,0x4F5C,0x6D41,0x751F,0x6210,0x7684,0x5168,0x90E8,0x6587,0x4EF6,0x5FC5,0x987B,0x4FDD,0x5B58,0x5728,0x672C,0x76EE,0x5F55,0x5185,0xFF1B,0x539F,0x59CB,0x8D44,0x6599,0x76EE,0x5F55,0x53EA,0x8BFB,0x3002)

$info = @(
    "# $workspaceHeading"
    ''
    "- ${lessonLabel}: $LessonTitle"
    "- ${createdLabel}: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    "- ${sourceRootLabel}: $resolvedParent"
    "- ${workspaceRootLabel}: $taskRoot"
    "- ${stageLabel}: $curationStage"
    ''
    "> $workspaceRule"
) -join [Environment]::NewLine

$infoPath = Join-Path $taskRoot $workspaceInfoName
Set-Content -LiteralPath $infoPath -Value $info -Encoding utf8

[ordered]@{
    lesson_title = $LessonTitle
    source_root = $resolvedParent
    task_root = $taskRoot
    info_file = $infoPath
} | ConvertTo-Json -Depth 3
