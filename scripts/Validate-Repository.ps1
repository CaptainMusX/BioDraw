[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$errors = [System.Collections.Generic.List[string]]::new()

function Assert-RepositoryCondition {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) {
        $script:errors.Add($Message)
    }
}

function Get-OccurrenceCount {
    param([string]$Text, [string]$Needle)
    if ([string]::IsNullOrEmpty($Text) -or [string]::IsNullOrEmpty($Needle)) {
        return 0
    }
    return ([regex]::Matches($Text, [regex]::Escape($Needle))).Count
}

$projectPath = Join-Path $repositoryRoot 'BioDraw\BioDraw.csproj'
$projectText = Get-Content -LiteralPath $projectPath -Raw
[xml]$projectXml = $projectText
$namespaceManager = [System.Xml.XmlNamespaceManager]::new($projectXml.NameTable)
$namespaceManager.AddNamespace('msb', 'http://schemas.microsoft.com/developer/msbuild/2003')

$projectIncludes = @(
    $projectXml.SelectNodes('//msb:Compile[@Include] | //msb:EmbeddedResource[@Include] | //msb:None[@Include]', $namespaceManager)
)
foreach ($node in $projectIncludes) {
    $include = [string]$node.Include
    if ([string]::IsNullOrWhiteSpace($include) -or $include.StartsWith('..\')) {
        continue
    }
    $resolved = Join-Path (Split-Path -Parent $projectPath) $include
    Assert-RepositoryCondition (Test-Path -LiteralPath $resolved) "项目引用不存在：$include"
}

$packagePath = Join-Path $repositoryRoot 'BioDraw\packages.config'
[xml]$packageXml = Get-Content -LiteralPath $packagePath -Raw
$webViewPackage = @($packageXml.packages.package | Where-Object { $_.id -eq 'Microsoft.Web.WebView2' })
Assert-RepositoryCondition ($webViewPackage.Count -eq 1) 'packages.config 必须且只能包含一个 Microsoft.Web.WebView2 版本。'
if ($webViewPackage.Count -eq 1) {
    $packageVersion = [string]$webViewPackage[0].version
    Assert-RepositoryCondition ($projectText.Contains("Microsoft.Web.WebView2.$packageVersion")) 'WebView2 的 packages.config 与项目 HintPath 版本不一致。'
}

$applicationVersionNode = $projectXml.SelectSingleNode('//msb:ApplicationVersion', $namespaceManager)
$applicationVersion = if ($null -eq $applicationVersionNode) { '' } else { [string]$applicationVersionNode.InnerText }
$assemblyInfo = Get-Content -LiteralPath (Join-Path $repositoryRoot 'BioDraw\Properties\AssemblyInfo.cs') -Raw
$fileVersionMatch = [regex]::Match($assemblyInfo, 'AssemblyFileVersion\("([0-9.]+)"\)')
Assert-RepositoryCondition $fileVersionMatch.Success 'AssemblyInfo.cs 缺少 AssemblyFileVersion。'
if ($fileVersionMatch.Success) {
    Assert-RepositoryCondition ($fileVersionMatch.Groups[1].Value -eq $applicationVersion) 'ApplicationVersion 与 AssemblyFileVersion 不一致。'
}

$webUiPath = Join-Path $repositoryRoot 'BioDraw\WebUI'
$htmlFiles = @(Get-ChildItem -LiteralPath $webUiPath -Filter '*.html' -File)
Assert-RepositoryCondition ($htmlFiles.Count -eq 7) 'WebUI HTML 资源数量应为 7。'
foreach ($file in $htmlFiles) {
    $html = Get-Content -LiteralPath $file.FullName -Raw
    Assert-RepositoryCondition ((Get-OccurrenceCount $html 'STYLES_PLACEHOLDER') -eq 1) "$($file.Name) 必须包含一次样式占位符。"
    Assert-RepositoryCondition ((Get-OccurrenceCount $html 'BRIDGE_PLACEHOLDER') -eq 1) "$($file.Name) 必须包含一次桥接脚本占位符。"
    Assert-RepositoryCondition ($html.Contains('lang="zh-CN"')) "$($file.Name) 缺少中文语言声明。"
    Assert-RepositoryCondition ($html.Contains('Content-Security-Policy')) "$($file.Name) 缺少内容安全策略。"
    Assert-RepositoryCondition ($html -match '<button[^>]+class="[^"]*tl tl-close[^"]*"[^>]+aria-label="关闭"') "$($file.Name) 的关闭控件不可访问。"
    Assert-RepositoryCondition (-not $html.Contains('.innerHTML')) "$($file.Name) 使用 innerHTML，可能引入本地脚本注入。"
    Assert-RepositoryCondition (-not ($html -match '(?i)(src|href)\s*=\s*["'']https?://')) "$($file.Name) 不应加载远程资源。"

    $scriptIndex = 0
    foreach ($scriptMatch in [regex]::Matches($html, '(?is)<script[^>]*>(.*?)</script>')) {
        $scriptIndex++
        $tempScript = Join-Path ([System.IO.Path]::GetTempPath()) ("biodraw-webui-{0}-{1}.js" -f [guid]::NewGuid().ToString('N'), $scriptIndex)
        try {
            [System.IO.File]::WriteAllText($tempScript, $scriptMatch.Groups[1].Value, [System.Text.UTF8Encoding]::new($false))
            if (Get-Command node -ErrorAction SilentlyContinue) {
                & node --check $tempScript 2>&1 | Out-Null
                Assert-RepositoryCondition ($LASTEXITCODE -eq 0) "$($file.Name) 的第 $scriptIndex 个脚本存在语法错误。"
            }
        }
        finally {
            if (Test-Path -LiteralPath $tempScript) {
                [System.IO.File]::Delete($tempScript)
            }
        }
    }
}

$bridgeText = Get-Content -LiteralPath (Join-Path $webUiPath 'bridge.js') -Raw
Assert-RepositoryCondition ($bridgeText.Contains('postMessage({ action,')) 'WebView 桥接应发送结构化对象。'
Assert-RepositoryCondition (-not $bridgeText.Contains('JSON.stringify({ action,')) 'WebView 桥接不应重复 JSON 编码。'

$webDialogSources = @(
    Get-ChildItem -LiteralPath (Join-Path $repositoryRoot 'BioDraw') -Filter '*WebDialog.cs' -File |
        ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
) -join [Environment]::NewLine
Assert-RepositoryCondition (-not ($webDialogSources -match 'Thread\.Sleep\(')) 'WebView 对话框不得阻塞 UI 线程。'
Assert-RepositoryCondition ($projectText.Contains('<SignManifests>true</SignManifests>')) 'VSTO 项目必须启用 ClickOnce 清单签名。'
Assert-RepositoryCondition ($projectText.Contains('ValidateManifestSigningKey')) '项目必须在缺少本地签名证书时给出明确错误。'

$trackedFiles = @(& git -C $repositoryRoot ls-files)
$sensitiveTracked = @($trackedFiles | Where-Object {
    $_ -match '(?i)(^|/)(packages|nuget\.exe)(/|$)' -or $_ -match '(?i)\.(pfx|pfx\.backup|snk)$'
})
Assert-RepositoryCondition ($sensitiveTracked.Count -eq 0) ("仓库仍跟踪依赖缓存或私钥文件：" + ($sensitiveTracked -join ', '))

if ($errors.Count -gt 0) {
    foreach ($errorMessage in $errors) {
        Write-Error $errorMessage -ErrorAction Continue
    }
    exit 1
}

Write-Host "Repository validation passed: $($htmlFiles.Count) WebUI pages and $($projectIncludes.Count) project resources checked."
