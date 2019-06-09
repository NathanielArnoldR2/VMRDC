[CmdletBinding()]
param(
  [Parameter(
    Mandatory = $true
  )]
  [Microsoft.HyperV.PowerShell.VirtualMachine[]]
  $VMs,

  [Parameter(
    Mandatory = $true
  )]
  [ValidateSet(640, 800, 1024)]
  [int]
  $DisplayWidth,

  [Parameter(
    Mandatory = $true
  )]
  [hashtable]
  $SynchronizedData
)
try {
  $SynchronizedData.ConnectionTargets = @{
    Current = $null
    Next    = $null
  }

  $VMListItems = @(
    $VMs |
      Select-Object Name,Id
  )

  # The window is equipped to handle display with aspect ratio between 4:3 and
  # 16:9, with horizontal size -ge desired $DisplayWidth and vertical size -ge
  # corresponding $DisplayMinHeight. Display with size not conforming will be
  # "letterboxed" horizontally, vertically, or both.
  $DisplayMinHeight = $DisplayWidth / 16 * 9
  $DisplayMaxHeight = $DisplayWidth / 4 * 3

  $windowXml = [xml]@"
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  xmlns:axrdp="clr-namespace:AxMSTSCLib;assembly=AxMSTSCLib"
  WindowStyle="None"
  ShowInTaskbar="False"
  SizeToContent="WidthAndHeight"
  ResizeMode="NoResize">
    <StackPanel>
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition />
                <ColumnDefinition Width="Auto" />
            </Grid.ColumnDefinitions>
            <ComboBox
              x:Name="VMList"
              HorizontalContentAlignment="Center"
              VerticalContentAlignment="Center"
              Padding="10"/>
            <Button
              x:Name="ExitButton"
              Grid.Column="1"
              Content="Exit"
              Padding="10"/>
        </Grid>
        <Grid
          x:Name="nvTextGrid"
          Visibility="Visible"
          Width="$DisplayWidth"
          Height="$DisplayMaxHeight"
          Background="Black">
            <StackPanel
              HorizontalAlignment="Center"
              VerticalAlignment="Center">
                <TextBlock
                  x:Name="nvText"
                  HorizontalAlignment="Center"
                  Foreground="White"
                  Text="No Video"/>
                <TextBlock
                  x:Name="nvReason"
                  Visibility="Hidden"
                  HorizontalAlignment="Center"
                  Foreground="White"
                  Text="Unknown"/>
                <Button
                  x:Name="nvReconnect"
                  IsEnabled="False"
                  Visibility="Hidden"
                  HorizontalAlignment="Center"
                  Content="Reconnect"/>
            </StackPanel>
        </Grid>
        <Grid
          x:Name="ConnectionHostGrid"
          Visibility="Collapsed"
          Width="$DisplayWidth"
          MinHeight="$DisplayMinHeight"
          MaxHeight="$DisplayMaxHeight"
          Background="LightGray">
            <WindowsFormsHost x:Name="ConnectionHost"
            HorizontalAlignment="Center"
            VerticalAlignment="Center">
                <axrdp:AxMsRdpClient8NotSafeForScripting x:Name="Connection" />
            </WindowsFormsHost>
        </Grid>
    </StackPanel>
</Window>
"@

  $SynchronizedData.Window =
  $Window = [System.Windows.Markup.XamlReader]::Load(
    [System.Xml.XmlNodeReader]::new($windowXml)
  )
  $Window.Add_Closing({
    $SynchronizedData.VMStatusJobState = "Stop"

    # Since this is an ActiveX component it must be disposed of manually before
    # window close; otherwise, an exception will be thrown, or PowerShell will
    # simply terminate with an application error.
    $Connection.Dispose()
  })

  $Window.FindName("VMList") |
    ForEach-Object {
      $_.FontSize = $_.FontSize * 2
      $_.ItemsSource = $VMListItems
      $_.DisplayMemberPath = "Name"
      $_.SelectedValuePath = "Id"
      $_.Add_SelectionChanged({
        param($obj, $evtArgs)

        $Window.Title = $obj.SelectedItem.Name

        # See below for full explanation of patterns required by MsRdpClient's
        # asynchronous, event-driven model.

        $SynchronizedData.ConnectionTargets.Next = $obj.SelectedItem.Id.ToString()

        if ($Connection.Connected) {
          $Connection.Disconnect() # Feed may not be changed while connected.
        } else {
          Invoke-VMRDCConnection
        }
      })
    }

  $Window.FindName("ExitButton") |
    ForEach-Object {
      $_.FontSize = $Window.FindName("VMList").FontSize
      $_.Add_Click({
        if ($SynchronizedData.ConfirmWindowExit) {
          $result = [System.Windows.MessageBox]::Show(
            "Are you sure you want to exit?",
            "Confirm Exit",
            [System.Windows.MessageBoxButton]::OKCancel,
            [System.Windows.MessageBoxImage]::Question,
            [System.Windows.MessageBoxResult]::Cancel,
            [System.Windows.MessageBoxOptions]::None
          )
        }
      
        if (
          $SynchronizedData.ConfirmWindowExit -eq $false -or
          $result -eq "OK"
        ) {
          $Window.Close()
        }
      })
    }

  $nvTextGrid = $Window.FindName("nvTextGrid")

  $Window.FindName("nvText") |
    ForEach-Object {
      $_.FontSize = $_.FontSize * 4
    }

  $SynchronizedData.nvReason = $nvReason = $Window.FindName("nvReason")
  $SynchronizedData.nvReconnect = $nvReconnect = $Window.FindName("nvReconnect")
  $nvReconnect.Add_Click({
    $Connection.Connect()
  })

  $ConnectionHostGrid = $Window.FindName("ConnectionHostGrid")

  $ConnectionHost = $Window.FindName("ConnectionHost")

  $Connection = $Window.FindName("Connection")

  # If the rdp control had been defined and added to the tree in code instead
  # of XAML, it would have been necessary to EndInit() before setting any of
  # these properties.

  # If either of these settings are not modified from default ($false and
  # $true, respectively) the connection will fail without visible error.
  #
  # The C# examples I studied for initial drafts of this code used a specific
  # interface to apply these settings, in a manner similar to how I still set 
  # "DisableCredentialsDelegation" below. Testing established this was not
  # needed. Furthermore, when they were modified using this interface the
  # 'AdvancedSettings' continued to show the incorrect value.
  $Connection.AdvancedSettings.EnableCredSspSupport = $true
  $Connection.AdvancedSettings.NegotiateSecurityLayer = $false
  
  # If this is not set, a Windows security prompt for host-level credentials
  # (not to the vm) will display prior to vm connection.
  [MSTSCLib.IMsRdpExtendedSettings] |
    ForEach-Object {
      $_.GetProperty("Property").SetValue(
        $Connection.GetOcx(),
        $true,
        @("DisableCredentialsDelegation")
      )
    }

  # Provided the authentication-related settings above are in place, three
  # things are still needed:
  #
  #   - The server network address of the VM host -- "localhost", in this case.
  #
  #   - The port at which the VM console is presented (2179, instead of the RDP
  #     default 3389). The C# examples I studied for my initial drafts of this
  #     code set this via the 'AdvancedSettings2' interface, but testing has
  #     established this is not needed.
  #
  #   - The VM Id is sent to the server as the preconnection blob, governing
  #     which VM's video output is seen. To facilitate switching between VM
  #     feeds this is done within Invoke-VMRDCConnection.
  $Connection.Server = "localhost"
  $Connection.AdvancedSettings.RDPPort = 2179

  # The MsRdpClient is designed for an event-driven interaction model, and
  # trying to subvert it by (e.g.) starting a connection and waiting on it
  # will either block the UI thread or crash PowerShell *hard*, as will
  # any attempt to directly query or manipulate the connection from the
  # VMStatus thread.
  #
  # Hence the use of 'Invoke-VMRDCConnection' called from different contexts,
  # and the separate 'Connected' synchronized data to convey connected state.

  function Invoke-VMRDCConnection {
    if ($null -ne $SynchronizedData.ConnectionTargets.Next) {
      $StatusRoot = $SynchronizedData.VMStatuses.$($SynchronizedData.ConnectionTargets.Next)

      # Get an updated report from VMStatus thread.
      $StatusRoot.Status = "Unknown"
      $StatusRoot.StatusRequested = $true

      while ($StatusRoot.Status -eq "Unknown") {}

      # The C# examples I studied for initial drafts used 'AdvancedSettings9'
      # for this setting. Testing has established this is not needed.
      $Connection.AdvancedSettings.PCB = $SynchronizedData.ConnectionTargets.Next
    }

    if ($null -ne $SynchronizedData.ConnectionTargets.Next -and $StatusRoot.Status -eq "Running") {
      # Will also show the connection interface elements and hide the NoVideo
      # interface elements, as governed by the OnConnected event handler.
      $Connection.Connect()
    } else {
      # Hide the connection interface elements, and show the (basic) NoVideo
      # interface elements.
      $ConnectionHostGrid.Visibility = [System.Windows.Visibility]::Collapsed
      $nvReason.Visibility = [System.Windows.Visibility]::Hidden
      $nvReconnect.IsEnabled = $false
      $nvReconnect.Visibility = [System.Windows.Visibility]::Hidden
      $nvTextGrid.Visibility = [System.Windows.Visibility]::Visible
    }

    # 'Next' (if any) becomes current.
    if ($null -ne $SynchronizedData.ConnectionTargets.Next) {
      $SynchronizedData.ConnectionTargets.Current = $SynchronizedData.ConnectionTargets.Next
      $SynchronizedData.ConnectionTargets.Next = $null
    }
  }

  $Connection.Add_OnConnected({
    $SynchronizedData.Connected = $true

    # Show the connection interface elements and hide the NoVideo interface
    # elements.
    $nvReason.Visibility = [System.Windows.Visibility]::Hidden
    $nvReconnect.IsEnabled = $false
    $nvReconnect.Visibility = [System.Windows.Visibility]::Hidden
    $nvTextGrid.Visibility = [System.Windows.Visibility]::Collapsed
    $ConnectionHostGrid.Visibility = [System.Windows.Visibility]::Visible
  })
  $Connection.Add_OnDisconnected({
    $SynchronizedData.Connected = $false

    Invoke-VMRDCConnection
  })
  $Connection.Add_OnRemoteDesktopSizeChange({
    # Change to rdp display dimensions must be made with SmartSizing off to
    # avoid erroneous behavior when desktop dimensions change while aspect
    # ratio does not.
    $Connection.AdvancedSettings.SmartSizing = $false

    # Note that dimension changes are applied to the WPF ConnectionHost,
    # rather than directly to the ActiveX connection.

    if (
      $Connection.DesktopWidth -ge $DisplayWidth -or
      $Connection.DesktopHeight -ge $DisplayMaxHeight
    ) { # Scaled/SmartSized to fit display dimensions.
      if (($Connection.DesktopWidth / $Connection.DesktopHeight) -ge 4/3) {
        # Limited by width; height scaled proportionally. Aspect ratios -gt
        # 16:9 (e.g. 16:10) will be "letterboxed" horizontally.
        $ConnectionHost.Width = $DisplayWidth
        $ConnectionHost.Height = $Connection.DesktopHeight / $Connection.DesktopWidth * $DisplayWidth
      } else {
        # Limited by height; width scaled proportionally. Aspect ratios -lt
        # 4:3 (e.g. 5:4) will be "letterboxed" vertically.
        $ConnectionHost.Width = $Connection.DesktopWidth / $Connection.DesktopHeight * $DisplayMaxHeight
        $ConnectionHost.Height = $DisplayMaxHeight
      }
      $Connection.AdvancedSettings.SmartSizing = $true
    } else { # No scaling required; presented as-is.

      # Display size remains @ $DisplayWidth and at least $DisplayMinHeight;
      # horizontal and vertical letterboxing will be used if resolution is,
      # less than either of these dimensions.
      $ConnectionHost.Width = $Connection.DesktopWidth
      $ConnectionHost.Height = $Connection.DesktopHeight
    }
  })

  $Window.FindName("VMList").SelectedValue = $VMListItems[0].Id
  
  $Window.ShowDialog()

} catch {
  $PSCmdlet.ThrowTerminatingError($_)
}