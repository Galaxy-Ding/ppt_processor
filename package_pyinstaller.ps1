[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

$date    = Get-Date -Format "yyyyMMdd"
$exeName = "一键生成发包规范工具_$date"

pip install --upgrade pyinstaller

pyinstaller --noconfirm --onefile --windowed `
    --icon=.\icon.ico `
    --name $exeName `
    --add-data "icon.ico;." `
    --add-data "config;config" `
    --add-data "examples;examples" `
    --add-data "exporters;exporters" `
    --add-data "extractors;extractors" `
    --add-data "ui;ui" `
    --add-data "utils;utils" `
    --add-data "core;core" `
    --add-data "content_models.py;." `
    --add-data "C:\Users\CAA\.conda\envs\py39\Lib\site-packages\zhconv;zhconv" `
    --add-data "ppt_reader.py;." `
    main.py

# 把 config 目录复制到 dist，同级可直接编辑
$dist = Join-Path (Get-Location) "dist"
$cfgSrc = Join-Path (Get-Location) "config"
$cfgDst = Join-Path $dist "config"
if (Test-Path $cfgSrc) {
    New-Item -ItemType Directory -Force -Path $cfgDst | Out-Null
    Copy-Item "$cfgSrc\*" $cfgDst -Recurse -Force
    Write-Host "配置已复制到: $cfgDst（可直接编辑）" -ForegroundColor Yellow
}

Write-Host "打包完成：$dist\$exeName.exe"
Copy-Item -Path "$dist\$exeName.exe" -Destination "D:\download\" -Force