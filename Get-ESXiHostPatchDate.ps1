param(
    [Parameter(Mandatory=$true)][String]$ESXiHost
    )

Connect-VIServer $ESXiHost > $Null

Get-VMHost $ESXiHost | ForEach-Object -Process { 
  [PSCustomObject]@{
    'Host'        = $_.Name
    'Version'     = $_.version
    'Build'       = $_.build
    'LastPatchDate' = [datetime]((Get-ESXCli -VMHost $_).software.vib.list() |
      Select-Object -Property installdate -ExpandProperty installdate |
      Sort-Object -Descending)[0]
  } 
}

Disconnect-VIServer $ESXiHost -Confirm:$false