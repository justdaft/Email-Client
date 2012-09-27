Add-Type -AssemblyName presentationframework

#Clear screen for errors
cls

$DebugPreference = "continue"
$VerbosePreference = "continue"
$WarningPreference = "continue"

[xml]$XAML = @"
<Window
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MainWindow" Height="318" Width="540">
    <Grid Margin="-9,-9,-8,-5" Background="{DynamicResource {x:Static SystemColors.ControlDarkBrushKey}}" Height="302" VerticalAlignment="Top">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="121*"/>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition Width="15*"/>
            <ColumnDefinition Width="16*"/>
            <ColumnDefinition Width="260*"/>
            <ColumnDefinition Width="27*"/>
            <ColumnDefinition Width="110*"/>
        </Grid.ColumnDefinitions>
        <Label Content="Subject" HorizontalAlignment="Left" Height="25" Margin="10,10,0,0" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="5"/>
        <TextBox x:Name="txtSubject" HorizontalAlignment="Left" Height="25" TextWrapping="Wrap" VerticalAlignment="Top" Width="300" Margin="10,35,0,0" Grid.ColumnSpan="5"/>
        <Label Content="Message" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top" Height="25" Width="200" Grid.ColumnSpan="5"/>
        <TextBox x:Name="txtMessage" HorizontalAlignment="Left" Height="200" TextWrapping="Wrap" VerticalAlignment="Top" Width="300" Margin="10,85,0,0" Grid.ColumnSpan="5"/>
        <Label Content="SMTP Server" Grid.Column="4" HorizontalAlignment="Left" Height="25" Margin="175,10,0,0" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="3"/>
        <Label Content="From" Grid.Column="4" HorizontalAlignment="Left" Margin="175,60,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.184,0.808" Height="25" Width="200" Grid.ColumnSpan="3"/>
        <TextBox x:Name="txtFrom" Grid.Column="4" HorizontalAlignment="Left" Height="25" Margin="175,85,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" RenderTransformOrigin="0.008,-1.087" Grid.ColumnSpan="3"/>
        <TextBox x:Name="txtSMTP" Grid.Column="4" HorizontalAlignment="Left" Height="25" Margin="175,34,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="3"/>
        <Label Content="To" Grid.Column="4" HorizontalAlignment="Left" Margin="175,110,0,0" VerticalAlignment="Top" Width="200" Height="25" Grid.ColumnSpan="3"/>
        <TextBox x:Name="txtTo" HorizontalAlignment="Left" Height="25" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" Grid.Column="4" Margin="175,135,0,0" Grid.ColumnSpan="3"/>
        <Button x:Name="cmdSend" Content="Send" Grid.Column="4" HorizontalAlignment="Left" Margin="175,165,0,0" VerticalAlignment="Top" Width="200" Grid.ColumnSpan="3" Height="22"/>
    </Grid>
</Window>
"@

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $Window=[Windows.Markup.XamlReader]::Load( $reader )

    $cmdsend = $Window.FindName("cmdSend")

        $cmdsend.Add_Click({
       
            Write-Debug "BUTTON CLICKED"
            $subject = $Window.findname("txtSubject")   
            $message = $Window.findname("txtMessage")   
            $smtp = $Window.findname("txtSMTP")   
            $from = $Window.findname("txtFrom")   
            $to = $Window.FindName("txtTo")
            $cmdsend = $Window.FindName("cmdSend")

        #Email
        Send-MailMessage -To $to.text -From $from.text -Body $message.text -Subject $subject.text -SmtpServer $smtp.text
   
    })
   
$Window.ShowDialog() | out-null