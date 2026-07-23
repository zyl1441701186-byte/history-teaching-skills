param(
    [ValidateSet('dark', 'light')]
    [string]$Theme = 'dark',

    [string]$SkillRoot = (Split-Path -Parent $PSScriptRoot)
)

$resolvedSkillRoot = [System.IO.Path]::GetFullPath($SkillRoot)
$templateName = "cloud-school-$Theme.pptx"
$templatePath = Join-Path $resolvedSkillRoot "private-assets\templates\$templateName"
$resolvedTemplatePath = [System.IO.Path]::GetFullPath($templatePath)

if (-not $resolvedTemplatePath.StartsWith($resolvedSkillRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Resolved private template path escaped the Skill root.'
}

if (Test-Path -LiteralPath $resolvedTemplatePath -PathType Leaf) {
    $result = [ordered]@{
        mode = 'private-template'
        theme = $Theme
        path = $resolvedTemplatePath
        exists = $true
    }
}
else {
    $result = [ordered]@{
        mode = 'written-fallback'
        theme = $Theme
        path = $null
        exists = $false
    }
}

$result | ConvertTo-Json -Depth 3
