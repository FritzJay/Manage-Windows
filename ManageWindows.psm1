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
      # Restore File Explorer Windows
      if ($_.ProcessName -eq 'ii') {
        $LocationURL = if ($_.LocationURL) {
          $_.LocationURL
        }
        else {
          $env:USERPROFILE
        }
        Write-host("File Explorer $LocationURL")
        $ShellApplication = New-Object -ComObject "Shell.Application"
        $ExistingExplorerWindowsCount = $ShellApplication.Windows().Count
        Write-Output($ExistingExplorerWindowsCount)
        $ShellApplication.Explore($LocationURL)
        $NewestExplorerWindow = $ShellApplication.Windows()[$ExistingExplorerWindowsCount]
        while ($null -eq $NewestExplorerWindow) {
          Start-Sleep -Seconds 1
          $NewestExplorerWindow = $ShellApplication.Windows()[$ExistingExplorerWindowsCount]
        }
        Write-Output('NewstExplorerWindow:')
        Write-Output($NewestExplorerWindow)
        $NewestExplorerWindow.Left = $NewestExplorerWindow.Left
        $NewestExplorerWindow.Top = $NewestExplorerWindow.Top
        $NewestExplorerWindow.Width = $NewestExplorerWindow.Right - $_.Left
        $NewestExplorerWindow.Height = $NewestExplorerWindow.Bottom - $_.Top

      }
      # Restore Application Windows
      <#
      else {
        $Process = $null
        if ($_.Arguments) {
          $Process = (Start-Process -FilePath $_.Path -PassThru -ArgumentList $ArgumentList | Select-Object -First 1)[0]
        }
        else {
          $Process = (Start-Process -FilePath $_.Path -PassThru | Select-Object -First 1)[0]
        }
        while (UnableToMoveWindow($Process)) {
          Write-Host("Waiting for $($_.ProcessName)...")
          Start-Sleep -Seconds 1
          $Process.Refresh()
        }
      }
      #> 
    }
  }
}

function UnableToMoveWindow($Process) {
  if ($null -eq $Process.MainWindowHandle) {
    return $true
  }
  return ([Window]::MoveWindow($Process.MainWindowHandle, $x, $y, $($_.Right - $_.Left), $($_.Bottom - $_.Top), $True) -eq $false)
}

Export-ModuleMember -Function Save-Windows, Restore-Windows

