. "$PSScriptRoot/Save-Windows.ps1"

Function Restore-Windows {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  Process {
    if (-not (Test-Path $JSONPath)) {
      Write-Error("There was an error retrieving previously saved window layouts. Have you saved any using Save-Windows?")
      return
    }

    $JSON = Get-Content -Path $JSONPath | ConvertFrom-Json

    $SavedWindowsExist = [bool](Get-Member -InputObject $JSON -Name $Name -MemberType Properties)
    if (-not $SavedWindowsExist) {
      Write-Error("Invalid Name given. Are you sure you used the correct name?")
    }

    $PreviouslyModifiedWindows = New-Object System.Collections.Generic.List[string]

    $JSON.$Name | ForEach-Object {
      if ($_.ProcessName -eq 'ii') {
        $ShellApplication = New-Object -ComObject "Shell.Application"
        $ModifiedWindow = Restore-ExplorerWindow -ShellApplication $ShellApplication -URL $_.LocationURL -Left $_.Left -Top $_.Top -Width $_.Width -Height $_.Height -PreviouslyModifiedWindows $PreviouslyModifiedWindows
        $PreviouslyModifiedWindows.Add($ModifiedWindow)
      }
      else {
        $ModifiedWindow = Restore-ApplicationWindow -Arguments $_.Arguments -ProcessName $_.ProcessName -ID $_.id -Path $_.Path -Left $_.Left -Top $_.Top -Height $_.Height -Width $_.Width -PreviouslyModifiedWindows $PreviouslyModifiedWindows
        $PreviouslyModifiedWindows.Add($ModifiedWindow)
      }
    }
  }
}

function Restore-ApplicationWindow {
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
    [int32]$Height,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[string]]$PreviouslyModifiedWindows
  )

  $ExistingProcesses = @()
  Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | ForEach-Object { $ExistingProcesses += $_ }

  $SameProcesses = @()
  $ExistingProcesses | Where-Object { $_.Id -eq $ID } | ForEach-Object { $SameProcesses += $_ }

  $SimilarProcesses = @()
  $ExistingProcesses | Where-Object {
    ($null -ne $_.MainWindowHandle) -and
    ($_.MainWindowHandle -ne 0) -and
    ($PreviouslyModifiedWindows -notcontains $_.MainWindowHandle)
  } | ForEach-Object { $SimilarProcesses += $_ }

  if (($SameProcesses.Count -lt 1) -and ($SimilarProcesses.Count -lt 1)) {
    try {
      Start-ApplicationProcess -Path $Path -Arguments $Arguments
    }
    catch {
      Write-Warning "Unable to start the process '$Path'. Aborting."
      return
    }
  }
  else {
    Write-Debug "The process for '$Path' is already running."
  }

  $MainWindowHandle = $null
  while ($null -eq $MainWindowHandle) {
    Start-Sleep -Seconds 0.25
    if ($SimilarProcesses.Count -gt 0) {
      $MainWindowHandle = $SimilarProcesses[0].MainWindowHandle
      Write-Debug "Main window from similar process: '$MainWindowHandle'"
    }
    elseif ($SameProcesses.Count -gt 0) {
      $MainWindowHandle = $SameProcesses[0].MainWindowHandle
      Write-Debug "Main window from the same process: '$MainWindowHandle'"
    }
    else {
      $NewProcesses = @()
      Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object {
        $ExistingProcesses -notcontains $_.id
      } | ForEach-Object { $NewProcesses += $_ }

      $NewProcessesWithMainWindows = @()
      $NewProcesses | Where-Object { ($null -ne $_.MainWindowHandle) -and ($_.MainWindowHandle -ne 0) } | ForEach-Object { $NewProcessesWithMainWindows += $_ }
      Write-Debug "New $ProcessName processes with main windows: '$NewProcessesWithMainWindows'"

      $MainWindowHandle = if ($NewProcessesWithMainWindows.Count -gt 0) {
        $NewProcessesWithMainWindows[0].MainWindowHandle
      }
      Write-Debug "Main Window Handle: $MainWindowHandle"
    }
  }

  try {
    [Window]::MoveWindow($MainWindowHandle, $Left, $Top, $Width, $Height, $True) | Out-Null
    Write-Debug "Successfuly moved the window '$MainWindowHandle'"
  }
  catch {
    Write-Warning "There was an error moving the window '$MainWindowHandle'. Aborting."
    return $MainWindowHandle
  }

  return $MainWindowHandle
}

function Start-ApplicationProcess {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [string[]]$Arguments = @()
  )
  
  Write-Debug "Starting a new process for '$Path'."
  if ($Arguments.Count -gt 0) {
    Start-Process -FilePath $Path -ArgumentList $Arguments
  }
  else {
    Start-Process -FilePath $Path
  }
}

function Restore-ExplorerWindow {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    $ShellApplication,
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
    [int32]$Height,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[string]]$PreviouslyModifiedWindows
  )

  $LocationURL = if ($URL) {
    $URL
  }
  else {
    $env:USERPROFILE
  }

  Write-Debug "Restoring explorer window to location URL: '$LocationURL'."

  $ShellApplication.Windows() | ForEach-Object {
    Write-Debug $_.HWND
  }

  $PreviouslyModifiedWindows | ForEach-Object {
    Write-Debug $_
  }

  $ExistingExplorerWindowsThatHaveNotBeenModified = @()
  $ShellApplication.Windows() | Where-Object { $PreviouslyModifiedWindows -notcontains $_.HWND } | ForEach-Object { $ExistingExplorerWindowsThatHaveNotBeenModified += $_ }
  Write-Debug $ExistingExplorerWindowsThatHaveNotBeenModified.ToString()
  Write-Debug "There are $($ExistingExplorerWindowsThatHaveNotBeenModified.Count) explorer windows, that have not been modified, already open."

  $ExplorerWindow = $null
  if ($ExistingExplorerWindowsThatHaveNotBeenModified.Count -gt 0) {
    $ExplorerWindow = $ExistingExplorerWindowsThatHaveNotBeenModified[0]
  }
  else {
    $ShellApplication.Explore($LocationURL)
    
    $ExplorerWindows = @()
    while ($ExplorerWindows.Count -lt 1) {
      Start-Sleep -Seconds 0.25
      $ShellApplication.Windows() | Where-Object {
        ($ExistingExplorerWindowsThatHaveNotBeenModified.HWND -notcontains $_.HWND) -and
        ($PreviouslyModifiedWindows -notcontains $_.HWND) -and
        ($null -ne $_.HWND) -and
        ($_.HWND -ne 0)
      } | ForEach-Object { $ExplorerWindows += $_ }
    }

    $ExplorerWindow = $ExplorerWindows[0]
  }

  $ExplorerWindow.Left = $Left
  $ExplorerWindow.Top = $Top
  $ExplorerWindow.Width = $Width
  $ExplorerWindow.Height = $Height

  return $ExplorerWindow.HWND
}

Export-ModuleMember -Function Restore-Windows
