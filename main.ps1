Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsFormsIntegration
Add-Type -LiteralPath $PSScriptRoot\MSTSCLib.dll
Add-Type -LiteralPath $PSScriptRoot\AxMSTSCLib.dll

$JobScripts = @{}
$JobScripts.Window = Get-Content -LiteralPath $PSScriptRoot\Window.job.ps1 -Raw
$JobScripts.VMStatus = Get-Content -LiteralPath $PSScriptRoot\VMStatus.job.ps1 -Raw

function Start-VMRDC {
  [CmdletBinding()]
  param(
    [Parameter(
      Position = 0,
      Mandatory = $true
    )]
    [Microsoft.HyperV.PowerShell.VirtualMachine[]]
    $VMs,

    [ValidateSet(640, 800, 1024)]
    [int]
    $DisplayWidth = 800
  )

  function New-UIPowerShell {
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    #$rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
  
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
  
    $ps
  }

  try {
    $Threads = @{ # Runspaces hosting WPF *must* be STA; hence the function.
      Window   = New-UIPowerShell
      VMStatus = New-UIPowerShell
    }
    
    # Once synchronized, hashtable values are shared between all in-process
    # runspaces where it is used.
    $SynchronizedData = [hashtable]::Synchronized(@{
      VMStatusJobState = "Init"
      VMStatuses = @{}
      Connected = $false
      ConfirmWindowExit = $true
    })

    $VMs |
      ForEach-Object {
        $SynchronizedData.VMStatuses.$($_.Id.ToString()) = [PSCustomObject]@{
          Name = $_.Name
          Id = $_.Id.ToString()
          Status = "Unknown"
          StatusRequested = $false
          WmiVM = $null
          WmiVMSummary = $null
        }
      }
    
    $cmd = [System.Management.Automation.Runspaces.Command]::new(
      $script:JobScripts.Window,
      $true # Script, not command.
    )
    $cmd.Parameters.Add("VMs", $VMs)
    $cmd.Parameters.Add("DisplayWidth", $DisplayWidth)
    $cmd.Parameters.Add("SynchronizedData", $SynchronizedData)
    
    $Threads.Window.Commands.AddCommand($cmd) |
      Out-Null

    $cmd = [System.Management.Automation.Runspaces.Command]::new(
      $script:JobScripts.VMStatus,
      $true # Script, not command.
    )
    $cmd.Parameters.Add("SynchronizedData", $SynchronizedData)
    
    $Threads.VMStatus.Commands.AddCommand($cmd) |
      Out-Null

    $Threads.VMStatus.BeginInvoke() | Out-Null    
    $Threads.Window.BeginInvoke() | Out-Null

    [PSCustomObject]@{
      PSTypeName       = "VMRDCData"
      Threads          = $Threads
      SynchronizedData = $SynchronizedData
    }
  }catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}

function Stop-VMRDC {
  [CmdletBinding(
    PositionalBinding = $false
  )]
  param(
    [Parameter(
      Position = 0,
      Mandatory = $true
    )]
    [PSTypeName("VMRDCData")]
    [PSCustomObject]
    $VMRDCData
  )
  try {
    # This override of Dispatcher.Invoke is very forgiving -- if the action
    # cannot be dispatched within the timeout it simply moves on; it doesn't
    # even throw an exception.
    $VMRDCData.SynchronizedData.Window.Dispatcher.Invoke(
    [System.Windows.Threading.DispatcherPriority]::Normal,
    [timespan]::FromSeconds(1),
    [action]{
      $VMRDCData.SynchronizedData.Window.Close()
    })
    
    do { # This is facile -- mature code would loop ~4-5 times, then throw.
      Start-Sleep -Seconds 1
    } until (
      $VMRDCData.SynchronizedData.VMStatusJobState -eq "Off" -and
      $VMRDCData.Threads.Window.InvocationStateInfo.State -eq "Completed" -and
      $VMRDCData.Threads.VMStatus.InvocationStateInfo.State -eq "Completed"
    )
    
    # Runspaces exist in the global scope; those not explicitly disposed of
    # accumulate. In PowerShell 5, you can observe this using 'Get-Runspace'.
    $VMRDCData.Threads.GetEnumerator() |
      ForEach-Object {
        $_.Value.Runspace.Dispose()
      }
  } catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}