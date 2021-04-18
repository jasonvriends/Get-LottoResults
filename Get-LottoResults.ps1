<#
    .SYNOPSIS
    Downloads Lotto results into CSV format.

    .DESCRIPTION
    Get-LottoResults.ps1 script downloads Lotto results from lottodatabase.com into Comma Separated Values (CSV) format.

    .INPUTS
    None. You cannot pipe objects to Get-LottoResults.ps1

    .OUTPUTS
    $scriptpath\Outputs

    .EXAMPLE
    PS> .\Get-LottoResults.ps1

    .LINK
    https://github.com/jasonvriends/Get-LottoResults.git

#>

# Global variables
$quote = '"'
$currentYear = get-date -Format yyyy
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$super7CSV = "$scriptPath\Outputs\Super7.csv"
$lotto649CSV = "$scriptPath\Outputs\Lotto649.csv"
$lottoMaxCSV = "$scriptPath\Outputs\LottoMax.csv"
$atlantic49CSV = "$scriptPath\Outputs\Atlantic49.csv"

# Create Outputs folder
if (!(Test-Path $scriptPath\Outputs))
{
    Write-Host "Creating the Outputs folder."
    New-Item -itemType Directory -Path $scriptPath -Name Outputs
}

# Get-Results function
function Get-Results {

    [CmdletBinding()]
    param (
        [string] $Lottery,      # e.g. lotto-max
        [int]    $TotalNumbers, # e.g. 8
        [int]    $Year          # e.g. 2000
    )

    Write-Host "Lottery: $Lottery"
    Write-Host "TotalNumbers: $TotalNumbers"
    Write-Host "Year: $Year"
    Write-Host ""

    $drawResults = New-Object -TypeName "System.Collections.ArrayList"

    # Build CSV Header
    $header = $quote + "Date" + $quote + ","
    For ($i=1; $i -le $TotalNumbers+1; $i++) {
        if ($i -ne $TotalNumbers+1) {
            $header = $header + $quote + "N$i" + $quote + ","
        } else {
            $header = $header + $quote + "B" + $quote
        }
        
    }

    Write-Host "CSV Header: $header"
    Write-Host ""

    $drawResults += $header

    # Build URL
    $url = "https://www.lottodatabase.com/lotto-database/canadian-lotteries/$Lottery/draw-history/$Year"
    $req = Invoke-WebRequest -UseBasicParsing $url

    Write-Host $url
        
    # Parse HTML

    $html = New-Object -Com "HTMLFile"
    
    try {
        # MS Office is installed
        $html.IHTMLDocument2_write($req.Content)
    }
    catch {
        # MS Office not installed
        $src = [System.Text.Encoding]::Unicode.GetBytes($req.Content)
        $html.write($src)
    }

    # Get draw dates from HTML
    Write-Host "Get draw dates from HTML"
    $drawDates   = $html.getElementsByTagName("div") | Where {$_.getAttributeNode('class').Value -eq 'col s_3_12'}
    $drawDates | ForEach-Object {

        $row = $_.innerhtml
        $row = $row | Get-Date -Format 'yyyy-MM-dd'
        $drawResults += $quote + $row + $quote

    }

    # Get draw numbers from HTML
    Write-Host "Get draw numbers from HTML"
    $result = ""
    $count1 = 1
    $count2 = 0
    $drawNumbers = $html.getElementsByTagName("span") | Where {$_.getAttributeNode('class').Value -eq 'white ball'}
    $drawNumbers | ForEach-Object {

        $row = $_.innerhtml

        $result = $result + $quote + $row + $quote + ","

        if ($count1 -eq $TotalNumbers) {
            $count2 = $count2 + 1
            $drawResults[$count2] = $drawResults[$count2] + "," + $result
            $count1 = 0
            $result = ""
        }
            
        $count1 = $count1 + 1            

    }

    # Get draw bonus from HTML
    Write-Host "Get draw bonus from HTML"
    $result = ""
    $count1 = 1
    $count2 = 0
    $drawBonus   = $html.getElementsByTagName("span") | Where {$_.getAttributeNode('class').Value -eq 'grey ball'}
    $drawBonus | ForEach-Object {

        $row = $_.innerhtml
        $row = $row -replace "<BR><SPAN class=bonus>Bonus</SPAN>",""

        $result = $quote + $row + $quote

        $count2 = $count2 + 1
        $drawResults[$count2] = $drawResults[$count2] + $result
        $count1 = 0
        $result = ""
        $count1 = $count1 + 1  

    }

    Write-Host ""

    return $drawResults

}

####################################################

# Get-Super7 function
function Get-Super7 {
    
    [int] $Start=1994
    [int] $End=2009
    
    For ($i=$Start; $i -le $End; $i++) {

        Write-Host $i

        $temp = Get-Results -Lottery "super-7" -Year $i -TotalNumbers 7

        if ($i -eq $Start) {

            $super7 = $super7 + $temp

        } else {

            $temp = $temp -replace '"Date","N1","N2","N3","N4","N5","N6","N7","B"'
            $super7 = $super7 + $temp

        }

    }

    Set-Content $super7CSV -Value $null
    Foreach ($arr in $super7) {
          $arr -join ',' | Add-Content $super7CSV
    }
    
    $a = Import-Csv -Delimiter "," $super7CSV
    $a | Sort-Object -Property Date | Export-Csv -NoTypeInformation $super7CSV

}

# Get-Lotto649 function
function Get-Lotto649 {

    [int] $Start=1982
    [int] $End=$currentYear
    
    For ($i=$Start; $i -le $End; $i++) {

        Write-Host $i

        $temp = Get-Results -Lottery "lotto-649" -Year $i -TotalNumbers 6

        if ($i -eq $Start) {

            $lotto649 = $lotto649 + $temp

        } else {

            $temp = $temp -replace '"Date","N1","N2","N3","N4","N5","N6","B"'
            $lotto649 = $lotto649 + $temp

        }

    }

    Set-Content $lotto649CSV -Value $null
    Foreach ($arr in $lotto649) {
          $arr -join ',' | Add-Content $lotto649CSV
    }
    
    $a = Import-Csv -Delimiter "," $lotto649CSV
    $a | Sort-Object -Property Date | Export-Csv -NoTypeInformation $lotto649CSV

}

# Get-LottoMax function
function Get-LottoMax {

    [int] $Start=2009
    [int] $End=$currentYear
    
    For ($i=$Start; $i -le $End; $i++) {

        Write-Host $i

        $temp = Get-Results -Lottery "lotto-max" -Year $i -TotalNumbers 7

        if ($i -eq $Start) {

            $lottomax = $lottomax + $temp

        } else {

            $temp = $temp -replace '"Date","N1","N2","N3","N4","N5","N6","N7","B"'
            $lottomax = $lottomax + $temp

        }

    }

    Set-Content $lottoMaxCSV -Value $null
    Foreach ($arr in $lottomax) {
          $arr -join ',' | Add-Content $lottoMaxCSV
    }
    
    $a = Import-Csv -Delimiter "," $lottoMaxCSV
    $a | Sort-Object -Property Date | Export-Csv -NoTypeInformation $lottoMaxCSV

}

# Get-Atlantic49 function
function Get-Atlantic49 {

    [int] $Start=2002
    [int] $End=$currentYear
    
    For ($i=$Start; $i -le $End; $i++) {

        Write-Host $i

        $temp = Get-Results -Lottery "atlantic-49" -Year $i -TotalNumbers 6

        if ($i -eq $Start) {

            $atlantic49 = $atlantic49 + $temp

        } else {

            $temp = $temp -replace '"Date","N1","N2","N3","N4","N5","N6","B"'
            $atlantic49 = $atlantic49 + $temp

        }

    }

    Set-Content $atlantic49CSV -Value $null
    Foreach ($arr in $atlantic49) {
          $arr -join ',' | Add-Content $atlantic49CSV
    }
    
    $a = Import-Csv -Delimiter "," $atlantic49CSV
    $a | Sort-Object -Property Date | Export-Csv -NoTypeInformation $atlantic49CSV

}

# Update numbers

Get-Super7
Get-Lotto649
Get-LottoMax
Get-Atlantic49
