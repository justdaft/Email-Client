Add-Type -AssemblyName presentationframework
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#Clear screen for errors
cls

$DebugPreference = "continue"
$VerbosePreference = "continue"
$WarningPreference = "continue"

[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Email Client" Height="250.667" Width="540" ResizeMode="NoResize">
    <Grid Margin="-9,-9,-8,-6" Height="238" VerticalAlignment="Top" Background="{DynamicResource {x:Static SystemColors.HighlightBrushKey}}">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="19*"/>
            <ColumnDefinition Width="32*"/>
            <ColumnDefinition Width="52*"/>
            <ColumnDefinition Width="18*"/>
            <ColumnDefinition Width="11*"/>
            <ColumnDefinition Width="4*"/>
            <ColumnDefinition Width="20*"/>
            <ColumnDefinition Width="257*"/>
            <ColumnDefinition Width="27*"/>
            <ColumnDefinition Width="111*"/>
        </Grid.ColumnDefinitions>
        <Label Content="Subject" HorizontalAlignment="Left" Height="25" Margin="0,10,0,0" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="7" Grid.Column="1"/>
        <TextBox x:Name="txtSubject" HorizontalAlignment="Left" Height="25" TextWrapping="Wrap" VerticalAlignment="Top" Width="300" Margin="0,35,0,0" Grid.ColumnSpan="7" Grid.Column="1"/>
        <Label Content="Message" HorizontalAlignment="Left" Margin="0,60,0,0" VerticalAlignment="Top" Height="25" Width="200" Grid.ColumnSpan="7" Grid.Column="1"/>
        <TextBox x:Name="txtMessage" HorizontalAlignment="Left" Height="135" TextWrapping="Wrap" VerticalAlignment="Top" Width="300" Margin="0,85,0,0" Grid.ColumnSpan="7" Grid.Column="1"/>
        <Label Content="SMTP Server" Grid.Column="7" HorizontalAlignment="Left" Height="25" Margin="175,10,0,0" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="3"/>
        <Label Content="From" Grid.Column="7" HorizontalAlignment="Left" Margin="175,60,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.184,0.808" Height="25" Width="200" Grid.ColumnSpan="3"/>
        <TextBox x:Name="txtFrom" Grid.Column="7" HorizontalAlignment="Left" Height="25" Margin="175,85,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" RenderTransformOrigin="0.008,-1.087" Grid.ColumnSpan="3"/>
        <Label Content="To" Grid.Column="7" HorizontalAlignment="Left" Margin="175,110,0,0" VerticalAlignment="Top" Width="200" Height="25" Grid.ColumnSpan="3"/>
        <TextBox x:Name="txtTo" HorizontalAlignment="Left" Height="25" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" Grid.Column="7" Margin="175,135,0,0" Grid.ColumnSpan="3"/>
        <Button x:Name="cmdSend" Content="Send" Grid.Column="7" HorizontalAlignment="Left" Margin="175,165,0,0" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="3" Height="25"/>
        <ComboBox x:Name="txtSMTP" Grid.ColumnSpan="3" Grid.Column="7" HorizontalAlignment="Left" Margin="175,35,0,0" VerticalAlignment="Top" Width="200" Height="25">
            <ComboBoxItem Content="smtp.gmail.com" HorizontalAlignment="Left" Width="198"/>
            <ComboBoxItem Content="relay.nhs.uk" HorizontalAlignment="Left" Width="198"/>
            <ComboBoxItem Content="localhost" HorizontalAlignment="Left" Width="198"/>
        </ComboBox>
		<Button x:Name="cmdClear" Content="Clear" Grid.Column="7" HorizontalAlignment="Left" Margin="175,195,0,0" VerticalAlignment="Top" Width="90" RenderTransformOrigin="0.436,0.288" Grid.ColumnSpan="2" Height="25"/>
        <Button x:Name="cmdInfo" Content="Info" Grid.Column="9" HorizontalAlignment="Left" Height="25" Margin="0,195,0,0" VerticalAlignment="Top" Width="90" RenderTransformOrigin="0.381,1.167"/>
    </Grid>
</Window>
"@

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $GLOBAL:Window=[Windows.Markup.XamlReader]::Load( $reader )
	
	$cmdclear = $Window.FindName("cmdClear")
	$cmdinfo = $Window.FindName("cmdInfo")
	
	
		$subject = $Window.findname("txtSubject")   
        $message = $Window.findname("txtMessage")   
        $smtp = $Window.findname("txtSMTP")   
        $from = $Window.findname("txtFrom")   
        $to = $Window.FindName("txtTo")
        $cmdsend = $Window.FindName("cmdSend")
		$bar = $Window.FindName("Bar")
	
	#Send button 
	$cmdsend.Add_Click({
     
#            $subject = $Window.findname("txtSubject")   
#            $message = $Window.findname("txtMessage")   
#            $smtp = $Window.findname("txtSMTP")   
#            $from = $Window.findname("txtFrom")   
#            $to = $Window.FindName("txtTo")
#            $cmdsend = $Window.FindName("cmdSend")
#			$bar = $Window.FindName("Bar")
		
      #If message doesnt work first time, try 3 times.
      for ( [int]$attempt = 1; $attempt -le 3; $attempt++ )
        { 
            [bool]$success = $false;

            try
                {
                    Send-MailMessage -to $to.Text -from $from.Text -SmtpServer $SMTP.SelectionBoxItem -Body $message.Text -Subject "$(get-date) $($subject.text)" -ErrorAction Stop 
                    $success = $true;
                    Write-Debug "Send Email Attempt $attempt succeeded.";
                }
            catch [System.Exception]
                {
                    Write-Debug "$("-" * 80)`r`nSend Email Attempt $attempt failed:`r`n"
                    Write-Debug $Error[0]
                    Write-Debug $("-" * 80); 
                    Write-Debug "`r`n"
                }

                 #If the message succeeded
             	 if($success)
				{ 
			  		[System.Windows.Forms.MessageBox]::Show("The email was sent successfully." , "Success!")
				 
				#Exit the loop
				break
				}	  
        }
	})
	
	#Info Button
	$cmdinfo.Add_Click({
		[System.Windows.Forms.MessageBox]::Show("This email cll.com" , "Info")
	})

	#Clear Button
	$cmdclear.Add_Click({ 
		

		
		$message.Text = ""
		$subject.Text = ""
		$smtp.Text = ""
		$to.Text = ""
		$from.Text = ""
		$GLOBAL:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $GLOBAL:Window.UpdateLayout() }, $null, $null)
	})
	
	#region
	
		$Window.Add_Loaded({
		
	    $Global:updatelayout = [Windows.Input.InputEventHandler]{ $Global:progressBar.UpdateLayout() }
		
	    $GLOBAL:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $GLOBAL:Window.UpdateLayout() }, $null, $null)

	    ##Configure a timer to refresh window##

	    #Create Timer object

	    Write-Verbose "Creating timer object"

	    $Global:timer = new-object System.Windows.Threading.DispatcherTimer

	    #Fire off every 5 seconds

	    Write-Verbose "Adding 5 second interval to timer object"

	    $timer.Interval = [TimeSpan]"0:0:5.00"

	    #Add event per tick

	    Write-Verbose "Adding Tick Event to timer object"

	    $timer.Add_Tick({

	        [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }

	        #Write-Verbose "Updating Window"

	    })

	    #Start timer

	    Write-Verbose "Starting Timer"

	    $timer.Start()

	    If (-NOT $timer.IsEnabled) {

	        $Window.Close()

	    }

	})  
	#endregion 

	$Window.ShowDialog() | out-null