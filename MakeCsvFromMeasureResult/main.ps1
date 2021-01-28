using namespace System.Collections.Generic
function Test-Admin {
    $wid = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $prp = New-Object System.Security.Principal.WindowsPrincipal($wid)
    $adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    $prp.IsInRole($adm)
}
Function Get-MeasureFilename() {
    
    #アセンブリのロード
    Add-Type -AssemblyName System.Windows.Forms

    #ダイアログインスタンス生成
    $dialog = New-Object Windows.Forms.OpenFileDialog

    $dialog.Title = "測定ファイルを選択してください。"
    $dialog.Filter = "測定ファイル(*.txt) | *.txt"
    $dialog.InitialDirectory = "C:\"

    #ダイアログ表示
    $result = $dialog.ShowDialog()

    #「開くボタン」押下ならファイル名フルパスをリターン
    If ($result -eq "OK") {
        Write-Output $dialog.FileName 
    }
    Else {
        Break
    }

}
function Split-DataLine {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            Position = 0,
            #    ParameterSetName="LiteralPath",
            #    ValueFromPipelineByPropertyName=$true,
            HelpMessage = "trim target string")]
        [ValidateNotNullOrEmpty()]
        [string]
        $inputstr
    )
    [string]$ss = $inputstr -replace "　", " "
    [string]$ss = $ss -replace "^\s+", ''
    [string]$ss = $ss -replace "\s+$", ''
    [string[]]$ans = $ss -split "[\s,/<=>]+"
    Write-Output $ans
}

function Get-Splitback {
    param (
        # this method can make one line
        [Parameter(Mandatory = $true,
            Position = 0,
            #    ParameterSetName="LiteralPath",
            #    ValueFromPipelineByPropertyName=$true,
            HelpMessage = "splitback target string")]
        [List[string]]
        $targetList
    )
    [string]$ans = ""
    [string]$targetchild = ""
    foreach ($targetchild in $targetList) {
        if ($ans -eq "") {
            $ans = $targetchild
        }
        else {
            $ans += ",$targetchild"
        }
    }
    Write-Output $ans
}

if ([int]$psversiontable.psversion.major -lt 6) {
    Write-Host "PowerShell version need 6 or later" -BackgroundColor Red -ForegroundColor White
}
else {
    Write-Host "PowerShell version is Fit" -BackgroundColor Green -ForegroundColor White
}


# if ((Test-Admin) -eq $false) {
#     Write-Host 'You need Administrator privileges to run this.' -BackgroundColor Red -ForegroundColor White
#     # Abort the script
#     # this will work only if you are actually running a script
#     # if you did not save your script, the ISE editor runs it as a series
#     # of individual commands, so break will not break then.
#     return
# }

Set-Location ..
[string[]]$DataArray = (Get-Content -Path sampledata/14828-sample.txt -Encoding shift-jis)

[string]$mydialog = Get-MeasureFilename
[string[]]$DataArray = (Get-Content -Path $mydialog -Encoding shift-jis)
Write-Host Line count is $DataArray.Count

[string[]]$headerArray = @("点", "平面", "円")
[string]$ActualDataHead = "理論値"
[string]$TypeString = "This is measurement type"
[string]$NumberString = "This is measurement number of every type"

[bool]$isMaching = $false
[string]$CompareBuffer = ""

[List[List[string]]]$HappyList = [List[List[string]]]::new()

foreach ($DataLine in $DataArray) {
    if ($isMaching) {
        $CompareBuffer = Split-DataLine($DataLine)
        if ($CompareBuffer[0][0] -eq $ActualDataHead[0]) {
            [List[string]]$BuffList = [List[string]]::new(5)
            $BuffList.Add($TypeString)
            $BuffList.Add($NumberString)
            $BuffList.Add($CompareBuffer[1])
            $BuffList.Add($CompareBuffer[2])
            $BuffList.Add($CompareBuffer[3])
            $HappyList.Add($BuffList)
            
        }
        $isMaching = $false
    }
    else {
        $isMaching = $false
        foreach ($header in $headerArray) {
            if ($header[0] -eq $DataLine[0]) {
                $TypeString = $header
                [string[]]$CompareBuffer = Split-DataLine($DataLine)
                $NumberString = $CompareBuffer[0].Substring($header.Length)
                $isMaching = $true
            }
        }
    }
}


[List[string]]$outdata = ""
foreach ($outdata in $HappyList) {
    Get-Splitback($outdata) | Out-File -LiteralPath test.csv -Encoding shift-jis -Append
}


Write-Host "Finish"
# [Console]::ReadKey($true) | Out-Null


