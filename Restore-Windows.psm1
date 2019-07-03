$JSONDirectory = "$($env:APPDATA)\ManageWindows"
$JSONPath = "$JSONDirectory\SavedWindows.json"

Function Restore-Windows {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Name
  )
  
  Begin {
    Try {
      [void][Window]
    }
    Catch {
      Add-Type @"
              using System;
              using System.Runtime.InteropServices;
              public class Window {
                [DllImport("User32.dll")]
                public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
              }
"@
    }
  }
  Process {
    # Take note of all process ids that match the process names of the one that were saved
    

    if (-not (Test-Path $JSONPath)) {
      Write-Error("There was an error retrieving previously saved window layouts. Have you saved any using Save-Windows?")
      return
    }

    $JSON = Get-Content -Path $JSONPath | ConvertFrom-Json

    $SavedWindowsExist = [bool](Get-Member -InputObject $JSON -Name $Name -MemberType Properties)
    if (-not $SavedWindowsExist) {
      Write-Error("Invalid Name given. Are you sure you used the correct name?")
    }

    $JSON.$Name | ForEach-Object {
      if ($_.ProcessName -eq 'ii') {
        Restore-Explorer-Window -URL $_.URL -Left $_.Left -Top $_.Top -Width $_.Width -Height $_.Height
      }
      else {
        Restore-Application-Window -Arguments $_.Arguments -ProcessName $_.ProcessName -ID $_.id -Path $_.Path -Left $_.Left -Top $_.Top -Height $_.Height -Width $_.Width
      }
    }
  }
}

function Restore-Application-Window {
  [CmdletBinding()]
  param (
    [AllowNull()]
    [Parameter(Mandatory = $true)]
    [string[]]$Arguments,
    [Parameter(Mandatory = $true)]
    [string]$ProcessName,
    [Parameter(Mandatory = $true)]
    [int32]$ID,
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [int32]$Left,
    [Parameter(Mandatory = $true)]
    [int32]$Top,
    [Parameter(Mandatory = $true)]
    [int32]$Width,
    [Parameter(Mandatory = $true)]
    [int32]$Height
  )

  $ExistingProcessIDs = Get-Process -Name $ProcessName | Select-Object -Property id

  $ProcessIsAlreadyRunning = $ExistingProcessIDs -contains $ID
  if (-not $ProcessIsAlreadyRunning) {
    if ($Arguments) {
      Start-Process -FilePath $Path -ArgumentList $Arguments
    }
    else {
      Start-Process -FilePath $Path
    }
  }

  $MainWindowHandle = $null
  while ($null -eq $MainWindowHandle) {
    Start-Sleep -Seconds 0.25
    if ($ProcessIsAlreadyRunning) {
      $MainWindowHandle = (Get-Process -Id $ID).MainWindowHandle
    }
    else {
      $NewProcesses = (Get-Process -Name $ProcessName | Where-Object {
          $ExistingProcessIDs -notcontains $_.id
        })
      $NewProcessesWithMainWindows = $NewProcesses | Where-Object { ($null -ne $_.MainWindowHandle) -and ($_.MainWindowHandle -ne 0) }
      $MainWindowHandle = $NewProcessesWithMainWindows[0].MainWindowHandle
    }
  }

  [Window]::MoveWindow($MainWindowHandle, $Right, $Top, $Width, $Height, $True) | Out-Null
}

function Restore-Explorer-Window {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$URL,
    [Parameter(Mandatory = $true)]
    [int32]$Left,
    [Parameter(Mandatory = $true)]
    [int32]$Top,
    [Parameter(Mandatory = $true)]
    [int32]$Width,
    [Parameter(Mandatory = $true)]
    [int32]$Height
  )
  $LocationURL = if ($URL) {
    $URL
  }
  else {
    $env:USERPROFILE
  }
  $ShellApplication = New-Object -ComObject "Shell.Application"
  $ExistingExplorerWindowsCount = $ShellApplication.Windows().Count
  $ShellApplication.Explore($LocationURL)
  $NewestExplorerWindow = $ShellApplication.Windows()[$ExistingExplorerWindowsCount]
  while ($null -eq $NewestExplorerWindow) {
    Start-Sleep -Seconds 0.25
    $NewestExplorerWindow = $ShellApplication.Windows()[$ExistingExplorerWindowsCount]
  }
  $NewestExplorerWindow.Left = $Left
  $NewestExplorerWindow.Top = $Top
  $NewestExplorerWindow.Width = $Width
  $NewestExplorerWindow.Height = $Height
}

Export-ModuleMember -Function Restore-Windows