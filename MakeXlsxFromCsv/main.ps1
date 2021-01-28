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

    $dialog.Title = "csvファイルを選択してください。"
    $dialog.Filter = "csvファイル(*.csv) | *.csv"
    $dialog.InitialDirectory = "$($env:USERPROFILE)\Desktop"

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

# Set-Location ..
# [string[]]$DataArray = (Get-Content -Path sampledata/14828-sample.txt -Encoding shift-jis)
# $mydialog = "C:\Users\take\Documents\GitHub\ParseMeasureResultTxt\sampledata\sampleParsed.csv"
[string]$mydialog = Get-MeasureFilename
[List[string]]$DataArray = (Get-Content -Path $mydialog -Encoding shift-jis)
Write-Host Line count is $DataArray.Count
$DataArray.Insert(0,"Type,Point,L,W,H")

try {
    $excel = New-Object -ComObject Excel.Application
    # $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.WorkSheets(1)
    $sheet.Name = "clipboard"
    $sheet.Range($sheet.Cells(1, "B"), $sheet.Cells($DataArray.Count, "C")).Merge($true)
    $sheet.Range($sheet.Cells(1, "D"), $sheet.Cells($DataArray.Count, "G")).Merge($true)
    $sheet.Range($sheet.Cells(1, "H"), $sheet.Cells($DataArray.Count, "I")).Merge($true)
    $sheet.Range($sheet.Cells(1, "J"), $sheet.Cells($DataArray.Count, "K")).Merge($true)
    $sheet.Range($sheet.Cells(1, "L"), $sheet.Cells($DataArray.Count, "O")).Merge($true)
    $sheet.Range($sheet.Cells(1, "P"), $sheet.Cells($DataArray.Count, "Q")).Merge($true)
    $sheet.Range($sheet.Cells(1, "R"), $sheet.Cells($DataArray.Count, "S")).Merge($true)
    $sheet.Range($sheet.Cells(1, "T"), $sheet.Cells($DataArray.Count, "W")).Merge($true)
    $sheet.Range($sheet.Cells(1, "X"), $sheet.Cells($DataArray.Count, "Y")).Merge($true)
    $sheet.Range($sheet.Cells(1, "Z"), $sheet.Cells($DataArray.Count, "AA")).Merge($true)
    # $array1 = @(@(1,"a"),@(2,"b"),@(3,"c"),@(4,"d"))
    #(A,B, ,D, , , ,H, ,J, ,L, , ,O, , ,R, ,T, , , ,X, ,Z)
    #(A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z)
    [string[]]$ans = $null
    $array1 = (1, 2, $null, 3, $null, $null, $null, 4, $null, 5, $null, 6, $null, $null, $null, 7, $null, 8, $null, 9, $null, $null, $null, 10, $null, 11)
    for ($i = 0; $i -lt $DataArray.Count; $i++) {
        $ans = $DataArray[$i] -split ","
        $array1 = (
            $ans[0],
            $ans[1], $null,
            $ans[2], $null, $null, $null,
            $null, $null,
            $null, $null,
            $ans[3], $null, $null, $null,
            $null, $null,
            $null, $null,
            $ans[4], $null, $null, $null,
            $null, $null,
            $null
        )
        $sheet.Range($sheet.Cells($i + 1, "A"), $sheet.Cells($i + 1, "AA")).Value(10) = $array1
    }
    

    if (!(Test-Path("output"))) {
        New-Item -Path "output" -ItemType Directory
    }

    $book.SaveAs("$(Get-Location)\output\tt.xlsx")
    $excel.DisplayAlerts = $true
    $excel.ScreenUpdating = $true
    $excel.EnableEvents = $true
    $excel.Quit()
}
finally {
    # $excel = $Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)  | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [GC]::collect()  
}

Write-Host "Finish" -ForegroundColor Green
# [Console]::ReadKey($true) | Out-Null


