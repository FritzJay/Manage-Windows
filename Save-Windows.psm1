$JSONDirectory = "$($env:APPDATA)\ManageWindows"
$JSONPath = "$JSONDirectory\SavedWindows.json"

Function Save-Windows {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $SavedWindows = New-Object -TypeName PSObject

  if (-not (Test-Path $JSONDirectory)) {
    New-Item -ItemType Directory -Path $JSONDirectory
  }

  if (Test-Path $JSONPath) {
    $SavedWindows = Get-Content -Path $JSONPath | ConvertFrom-Json
  }

  $SavedWindows | Add-Member -MemberType NoteProperty -Name $Name -Value $(Get-Windows) -Force

  $SavedWindows | ConvertTo-Json | Out-File -FilePath $JSONPath
}

Export-ModuleMember -Function Save-Windows