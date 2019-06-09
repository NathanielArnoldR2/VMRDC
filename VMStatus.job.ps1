[CmdletBinding()]
param(
  [Parameter(
    Mandatory = $true
  )]
  [hashtable]
  $SynchronizedData
)
try {
  if ($SynchronizedData.VMStatusJobState -eq "Init") {
    $PSDefaultParameterValues = @{
      "Get-WmiObject:Namespace" = "root\virtualization\v2"
    }

    $SynchronizedData.VMStatusJobState = "Running"
  }

  while ($true) {
    if ($SynchronizedData.VMStatusJobState -eq "Stop") {
      break
    }

    foreach ($vm in $SynchronizedData.VMStatuses.GetEnumerator()) {
      # Speed is important. Status for a vm is only updated on request, *or*
      # when the active feed to a vm is in disconnected state.
      if ($vm.Value.StatusRequested -eq $false -and ($SynchronizedData.Connected -or $vm.Key -ne $SynchronizedData.ConnectionTargets.Current)) {
        continue
      }

      # ...and wmi object paths are cached after first retrieval.
      if ($null -eq $vm.Value.WmiVMSummary) {
        if ($null -eq $vm.Value.WmiVM) {
          $vmObj = Get-WmiObject -Class Msvm_ComputerSystem -Filter "Name = `"$($vm.Key)`""

          $vm.Value.WmiVM = $vmObj.Path
        } else {
          $vmObj = [wmi]$vm.Value.WmiVM
        }

        $sumObj = $vmObj.GetRelated("Msvm_SummaryInformation")

        $vm.Value.WmiVMSummary = $sumObj.Path
      } else {
        $sumObj = [wmi]$vm.Value.WmiVMSummary
      }

      $vm.Value.StatusRequested = $false
      if ($null -eq $sumObj.ThumbnailImage) {
        $vm.Value.Status = "Off"
      } else {
        $vm.Value.Status = "Running"
      }
    }

    # The remainder of the loop is concerned with updating the "NoVideo"
    # interface for a disconnected vm feed, giving the user a reason --
    #  "Off" or "Running" but disconnected -- and exposing the option
    # to reconnect if the latter.
    #
    # The cross-thread dispatches needed to update the UI are *expensive*,
    # hence we cache known information and update only as changes are found.
    
    if ($SynchronizedData.Connected -or $null -eq $SynchronizedData.ConnectionTargets.Current) {
      $knownId = $null
      $knownStatus = $null

      $knownId_Date = $null

      continue
    }

    $currentId = $SynchronizedData.ConnectionTargets.Current

    if ($currentId -ne $knownId) {
      $knownId = $currentId
      $knownStatus = $null

      $knownId_Date = [datetime]::Now

      continue
    }

    if (([datetime]::Now - $knownId_Date).TotalSeconds -lt 1) {
      continue
    }

    $currentStatus = $SynchronizedData.VMStatuses.$currentId.Status

    if ($currentStatus -eq $knownStatus) {
      continue
    }

    if ($currentStatus -eq "Off") {
      $SynchronizedData.nvReason.Dispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Normal,
        [timespan]::FromSeconds(1),
        [action]{
          $SynchronizedData.nvReconnect.IsEnabled = $false
          $SynchronizedData.nvReconnect.Visibility = [System.Windows.Visibility]::Hidden
          $SynchronizedData.nvReason.Visibility = [System.Windows.Visibility]::Visible
          $SynchronizedData.nvReason.Text = "VM is off."
        }
      )
    } elseif ($currentStatus -eq "Running") {
      $SynchronizedData.nvReason.Dispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Normal,
        [timespan]::FromSeconds(1),
        [action]{
          $SynchronizedData.nvReconnect.Visibility = [System.Windows.Visibility]::Visible
          $SynchronizedData.nvReconnect.IsEnabled = $true
          $SynchronizedData.nvReason.Visibility = [System.Windows.Visibility]::Visible
          $SynchronizedData.nvReason.Text = "VM console feed is disconnected."
        }
      )
    } else {
      $SynchronizedData.nvReason.Dispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Normal,
        [timespan]::FromSeconds(1),
        [action]{
          $SynchronizedData.nvReconnect.IsEnabled = $false
          $SynchronizedData.nvReconnect.Visibility = [System.Windows.Visibility]::Hidden
          $SynchronizedData.nvReason.Visibility = [System.Windows.Visibility]::Hidden
          $SynchronizedData.nvReason.Text = "Unknown."
        }
      )
    }

    $knownStatus = $currentStatus
  }

  $SynchronizedData.VMStatusJobState = "Off"
} catch {
  $PSCmdlet.ThrowTerminatingError($_)
}