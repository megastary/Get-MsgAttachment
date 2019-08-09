<#PSScriptInfo

.VERSION 1.1

.GUID a787a311-5b25-444b-ae20-b5512ea460ef

.AUTHOR Jakub Šindelář

.COMPANYNAME Mountfield a.s.

.COPYRIGHT Jakub Šindelář

.TAGS msg outlook message mail email attachment extract

.LICENSEURI https://opensource.org/licenses/MIT

.PROJECTURI https://github.com/megastary/Get-MsgAttachment

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
 1.0 - [2019-08-08] - Vytvořen skript umožňující extrahování příloh ze souboru msg.
 1.1 - [2019-08-09] - Upravena logika detekce, odebrána podpora PDF a PNG, zachováno pouze WAV.
#>

<#
.SYNOPSIS
 Vytvořen skript umožňující extrahování příloh ze souboru msg.

.DESCRIPTION
 Vytvořen skript umožňující extrahování příloh ze souboru msg.

.PARAMETER Path
 Cesta k msg souboru.

.PARAMETER OutputPath
 Cesta, kam se uloží extrahované přílohy.

.PARAMETER SpecificExtension
 V případě, že je žádáno extrahování souborů se specifickou koncovkou.

.PARAMETER SevenZipPath
 Cesta k 7-zip exe souboru.

.INPUTS
 Žádné.

.OUTPUTS
 Žádné.

.EXAMPLE
 .\Get-MsgAttachment.ps1 -Path "C:\Data\Test.msg" -OutputPath "C:\Attachments" -SevenZipPath "C:\bin\7za.exe"
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true,
    Position=0,     
    HelpMessage="Zadejte cestu k souboru typu 'msg'.")]
    [ValidateNotNullOrEmpty()]
    [String]$Path,
    [Parameter(Mandatory=$true,
    Position=1,       
    HelpMessage="Zadejte cestu ke složce, do které se uloží přílohy extrahované ze souboru.")]
    [ValidateNotNullOrEmpty()]
    [String]$OutputPath,
    [Parameter(Mandatory=$true,
    Position=2,
    HelpMessage="Zadejte cestu k souboru 7z.exe, který extrahuje zprávu.")]
    [ValidateNotNullOrEmpty()]
    [String]$SevenZipPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

If (!((Test-Path -Path $Path) -and ($Path -like "*.msg"))) {
    Write-Error -Message "Zadejte správnou cestu k .msg souboru!"
}

If (!((Test-Path -Path $SevenZipPath) -and ($SevenZipPath -like "*\7z.exe"))) {
    Write-Output "Stáhněte si nejnovější exe z webu https://www.7-zip.org/download.html"
    Write-Error -Message "Zadejte správnou cestu k 7z.exe souboru!"
}

# Vytvoříme složku s názvem  zprávy, kam se extrahuje obsah e-mailu.
#$MessageName = (Get-ItemProperty -Path $Path).Name.Substring(0,(Get-ItemProperty -Path $Path).Name.Length-4)
$TempPath = Join-Path -Path $OutputPath -ChildPath "Temp"
try {
    New-Item -Path $TempPath -ItemType Directory -Force | Out-Null
} catch {
    Write-Error -Message "Nepodařilo se vytvořit složku $Temp v cestě $OutputPath. Originální zpráva chyby: $_"
}

# Extrahujeme zprávu
try {
    Start-Process -FilePath $SevenZipPath -ArgumentList "x ""$Path"" -y -o""$TempPath""" -Wait -NoNewWindow | Out-Null
} catch {
    Write-Error -Message "Nepodařilo se extrahovat soubor $Path. Originální zpráva chyby: $_"
}

Get-ChildItem -Path $TempPath -Directory -Filter "*attach*" | ForEach-Object {
    if (Get-ChildItem -Path $_.FullName -Filter "__substg1.0_37010102") {
        try {
            $Content = (Get-FileHash -Path $(Join-Path -Path $_.FullName -ChildPath "__substg1.0_370E001F") -ErrorAction Stop).Hash
            if ($Content -match "C0C46C8FDD4F328E583942A8ADEC745017EB87EE26980A7AC57329B48FB813AF") {
                # Found WAV
                Copy-Item -Path $(Join-Path -Path $_.FullName -ChildPath "__substg1.0_37010102") -Destination $(Join-Path -Path $OutputPath -ChildPath "$(Get-Date -Format "yyyy-MM-ddTHH-mm-ss").wav")
            } else {
                Write-Output "Content Unknown"
            }
        } catch {
            Write-Output "Unsupported content"
        }
    }
}

try {
    Remove-Item -Path $TempPath -Force -Recurse
} catch {
    Write-Error -Message "Nepodařilo se odstranit složku $TempPath. Originální zpráva chyby: $_"
}
