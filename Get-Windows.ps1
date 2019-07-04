Function Get-Windows {
  Get-Explorer-Windows
  Get-Application-Windows
}

Function Get-Application-Windows {
  $ScreenSize = Get-WmiObject -Class Win32_DesktopMonitor | Select-Object ScreenWidth, ScreenHeight

  $Processes = Get-Process | Where-Object { $_.MainWindowHandle -ne [System.IntPtr]::Zero }

  $Windows = $Processes | ForEach-Object {
    $Rectangle = New-Object RECT
    [Window]::GetWindowRect($_.MainWindowHandle, [ref]$Rectangle)
    return New-Object -TypeName PSObject -Property @{
      ProcessName = $_.ProcessName
      id          = $_.id
      Left        = $Rectangle.Left
      Top         = $Rectangle.Top
      Width       = $Rectangle.Right - $Rectangle.Left
      Height      = $Rectangle.Bottom - $Rectangle.Top
      LocationURL = $null
      Path        = $_.Path
      Arguments   = $_.Arguments
    }
  } | Where-Object { $_ -ne $true }

  $Windows | Where-Object { $_ -ne $true -and $_.Left -ne 0 -and $_.Right -ne $ScreenSize.ScreenWidth -and $_.ProcessName -ne 'ApplicationFrameHost' }
}

Function Get-Explorer-Windows {
  $app = New-Object -COM 'Shell.Application'

  $Windows = $app.Windows() | Where-Object Name -EQ 'File Explorer' | Select-Object -Property 'Left', 'Top', 'Width', 'Height', 'LocationURL' 

  $Windows | ForEach-Object { New-Object -TypeName PSObject -Property @{
      ProcessName = 'ii'
      id          = $null
      Left        = $_.Left
      Top         = $_.Top
      Width       = $_.Width
      Height      = $_.Height
      LocationURL = $_.LocationURL
      Path        = $null
      Arguments   = $null
    } }
}

Export-ModuleMember -Function Get-Windows
