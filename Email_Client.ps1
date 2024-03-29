﻿	#Version 1.1.1 27/09/2012
	#Changelog:
	#+water marks

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
    Title="Email Client" Height="240" Width="530" ResizeMode="NoResize">
    <Window.Resources>
        <!--ROUND TEXTBOX -->
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="KeyboardNavigation.TabNavigation" Value="None" />
            <Setter Property="AllowDrop" Value="true" />
            <Setter Property="Background" Value="Transparent"></Setter>
            <Setter Property="HorizontalContentAlignment" Value="Stretch" />
            <Setter Property="VerticalContentAlignment" Value="Stretch" />
            <Setter Property="FontFamily" Value="Segoe UI" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="Padding" Value="8,5,3,3" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Grid>
                            <Border x:Name="BorderBase" Background="White"   BorderThickness="3"
                            BorderBrush="SkyBlue" CornerRadius="10" />
                            <Label x:Name="TextPrompt" Content="{TemplateBinding Tag}" Visibility="Collapsed" Focusable="False"  Foreground="Silver"></Label>
                            <ScrollViewer Margin="0" x:Name="PART_ContentHost" Foreground="{DynamicResource OutsideFontColor}" />
                        </Grid>
                        <ControlTemplate.Triggers>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsFocused" Value="False"></Condition>
                                    <Condition Property="Text" Value=""></Condition>
                                </MultiTrigger.Conditions>
                                <MultiTrigger.Setters>
                                    <Setter Property="Visibility" TargetName="TextPrompt" Value="Visible"></Setter>
                                </MultiTrigger.Setters>
                            </MultiTrigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderThickness" TargetName="BorderBase" Value="2.4,2.4,1,1"></Setter>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Foreground" Value="DimGray" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="-9,-9,-8,-6" Height="238" VerticalAlignment="Top" Background="{DynamicResource {x:Static SystemColors.HighlightBrushKey}}">

        <Label Content="Subject" HorizontalAlignment="Left" Height="25" Margin="13,5,0,0" VerticalAlignment="Top" Width="200"/>
        <TextBox x:Name="txtSubject" HorizontalAlignment="Left" Tag="  Email Subject" Height="25" TextWrapping="Wrap" VerticalAlignment="Top" Width="300" Margin="13,30,0,0"/>
        <Label Content="Message" HorizontalAlignment="Left" Margin="13,55,0,0" VerticalAlignment="Top" Height="25" Width="200"/>
        <TextBox x:Name="txtMessage" HorizontalAlignment="Left" Tag=" Email Body " Height="135" TextWrapping="Wrap" VerticalAlignment="Top" Width="300" Margin="13,80,0,0" SelectionBrush="White" Foreground="White">
            <TextBox.BorderBrush>
                <LinearGradientBrush EndPoint="0,20" MappingMode="Absolute" StartPoint="0,0">
                    <GradientStop Color="#FFABADB3" Offset="0.05"/>
                    <GradientStop Color="#FFE2E3EA" Offset="0.07"/>
                    <GradientStop Color="White" Offset="1"/>
                </LinearGradientBrush>
            </TextBox.BorderBrush>
        </TextBox>
        <Label Content="SMTP Server" HorizontalAlignment="Left" Height="25" Margin="328,5,0,0" VerticalAlignment="Top" Width="200"/>
        <Label Content="From" HorizontalAlignment="Left" Margin="328,55,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.184,0.808" Height="25" Width="200"/>
        <TextBox x:Name="txtFrom" Tag=" Sender's Email" HorizontalAlignment="Left" Height="25" Margin="328,80,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" RenderTransformOrigin="0.008,-1.087"/>
        <Label Content="To" HorizontalAlignment="Left" Margin="328,105,0,0" VerticalAlignment="Top" Width="200" Height="25"/>
        <TextBox x:Name="txtTo" HorizontalAlignment="Left" Height="25" Tag=" Recipient Email" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" Margin="328,130,0,0"/>
        <Button x:Name="cmdSend" Content="Send" HorizontalAlignment="Left" Margin="328,160,0,0" VerticalAlignment="Top" Width="200" Height="25"/>
        <ComboBox x:Name="txtSMTP" HorizontalAlignment="Left" Margin="328,30,0,0" VerticalAlignment="Top" Width="200" Height="25">
            <ComboBoxItem Content="smtp.gmail.com" HorizontalAlignment="Left" Width="198"/>
            <ComboBoxItem Content="relay.nhs.uk" HorizontalAlignment="Left" Width="198"/>
            <ComboBoxItem Content="localhost" HorizontalAlignment="Left" Width="198"/>
            <ComboBoxItem Content="smtp.live.com" HorizontalAlignment="Left" Width="198"/>
            <ComboBoxItem Content="smtp.mail.yahoo.com" HorizontalAlignment="Left" Width="198"/>
        </ComboBox>
        <Button x:Name="cmdClear" Content="Clear" HorizontalAlignment="Left" Margin="328,190,0,0" VerticalAlignment="Top" Width="90" RenderTransformOrigin="0.436,0.288" Height="25"/>
        <Button x:Name="cmdInfo" Content="Info" HorizontalAlignment="Left" Height="25" Margin="438,190,0,0" VerticalAlignment="Top" Width="90" RenderTransformOrigin="0.381,1.167"/>
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

	#Send button 
	$cmdsend.Add_Click({
     
#            $subject = $Window.findname("txtSubject")   
#            $message = $Window.findname("txtMessage")   
#            $smtp = $Window.findname("txtSMTP")   
#            $from = $Window.findname("txtFrom")   
#            $to = $Window.FindName("txtTo")
#            $cmdsend = $Window.FindName("cmdSend")
#			 $bar = $Window.FindName("Bar")
		
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
		[System.Windows.Forms.MessageBox]::Show("This program was created by Logan Westbury. If you have any questions feel free to contact me at logan.westbury@gmail.com" , "Info")
	})

	#Clear Button
	$cmdclear.Add_Click({ 
		
		$message.Text = ""
		$subject.Text = ""
		$smtp.Text = ""
		$to.Text = ""
		$from.Text = ""
		#$GLOBAL:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $GLOBAL:Window.UpdateLayout() }, $null, $null)
	})
	
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
	
	$Window.ShowDialog() | out-null