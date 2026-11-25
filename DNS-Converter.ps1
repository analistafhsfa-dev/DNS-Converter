$ErrorActionPreference = "Stop"
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DNS Converter"
        Width="820" Height="560" MinWidth="700" MinHeight="500"
        WindowStartupLocation="CenterScreen"
        Background="#F3F6FB" Foreground="#1F2937"

        FontFamily="Segoe UI" ResizeMode="NoResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="#2563EB"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="18,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#111827"/>
            <Setter Property="Foreground" Value="#F9FAFB"/>
            <Setter Property="CaretBrush" Value="#111827"/>
            <Setter Property="BorderBrush" Value="#374151"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="13,12"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>
    <Grid Margin="28">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,26">
            <Border Width="154" Height="54" CornerRadius="15" Background="#DBEAFE" HorizontalAlignment="Left" Margin="0,0,0,16">
                <TextBlock Text="DNS Converter" Foreground="#1D4ED8" FontSize="17" FontWeight="Bold"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
			
			<!-- Janela principal do DNS Converter 
            <TextBlock Text="DNS Converter" FontSize="30" FontWeight="Bold" Foreground="#111827"/> -->
            <TextBlock Text="Consultare endereços da Internet." FontSize="14"
                       Foreground="#6B7280" Margin="0,6,0,0"/>
        </StackPanel>

        <TextBlock Grid.Row="1" Text="DOMÍNIO OU LINK" FontSize="11" FontWeight="Bold"
                   Foreground="#6B7280" Margin="0,0,0,9"/>

        <Grid Grid.Row="2" Margin="0,0,0,18">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="TxtDomain" Grid.Column="0" Height="48"
                     ToolTip="Ex.: exemplo.com, https://exemplo.com/pagina"
                     Margin="0,0,10,0"/>
            <Button x:Name="BtnLookup" Grid.Column="1" Content="Buscar IP  ↵" Height="48"
                    Padding="20,0" IsDefault="True"/>
        </Grid>

        <Border Grid.Row="3" Background="#FFFFFF" BorderBrush="#1F2937" BorderThickness="1"
                CornerRadius="12" Padding="20" Margin="0,0,0,18">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="0,0,0,16">
                    <TextBlock Text="RESULTADO DA CONSULTA" FontSize="11" FontWeight="Bold"
                               Foreground="#6B7280" VerticalAlignment="Center"/>
                    <TextBlock x:Name="LblStatus" Text="Aguardando consulta" FontSize="12"
                               Foreground="#6B7280" DockPanel.Dock="Right" HorizontalAlignment="Right"/>
                </DockPanel>
                <StackPanel Grid.Row="1" Margin="0,0,0,14">
                    <TextBlock Text="Endereço IP" FontSize="13" Foreground="#6B7280"/>
                    <TextBlock x:Name="LblPrimaryIP" Text="—" FontSize="27" FontWeight="SemiBold"
                               Foreground="#1D4ED8" Margin="0,4,0,0" TextWrapping="Wrap"/>
                </StackPanel>
                <Grid Grid.Row="2">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Text="Todos os IPs encontrados" FontSize="12" FontWeight="SemiBold"
                               Foreground="#374151" Margin="0,0,0,8"/>
                    <ListBox x:Name="LstIPs" Grid.Row="1" Background="#F9FAFB" Foreground="#374151"
                             BorderBrush="#D1D5DB" BorderThickness="1" Padding="7"
                             FontFamily="Consolas" FontSize="13"/>
                </Grid>
            </Grid>
        </Border>

        <DockPanel Grid.Row="4">
            <TextBlock Text="DNS • IPv4 / IPv6 • Desenvolvido por Guilherme - Analista Infra e Segurança                                "
                       FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
            <Button x:Name="BtnClear" Content="Limpar" Background="#E5E7EB"
                    Padding="14,7" FontSize="12" DockPanel.Dock="Right"/>
        </DockPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

$TxtDomain   = $Window.FindName("TxtDomain")
$BtnLookup   = $Window.FindName("BtnLookup")
$BtnClear    = $Window.FindName("BtnClear")
$LblStatus   = $Window.FindName("LblStatus")
$LblPrimaryIP = $Window.FindName("LblPrimaryIP")
$LstIPs      = $Window.FindName("LstIPs")

function Resolve-DomainFromInput {
    param([Parameter(Mandatory=$true)][string]$InputText)

    $value = $InputText.Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Digite um domínio ou link para consultar."
    }

    # Aceita links com http://, https:// ou sem protocolo.
    if ($value -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
        $value = "https://$value"
    }

    $uri = $null
    if (-not [Uri]::TryCreate($value, [UriKind]::Absolute, [ref]$uri)) {
        throw "O link informado não é válido."
    }

    if ($uri.Scheme -notin @("http", "https")) {
        throw "Informe um endereço HTTP/HTTPS ou somente o domínio."
    }

    $hostName = $uri.DnsSafeHost.TrimEnd(".")
    if ([string]::IsNullOrWhiteSpace($hostName)) {
        throw "Não foi possível identificar o domínio no link."
    }

    # Se foi informado um IP diretamente, não é necessário consultar DNS.
    $parsedIP = $null
    if ([System.Net.IPAddress]::TryParse($hostName, [ref]$parsedIP)) {
        return [pscustomobject]@{
            Host = $hostName
            IPs  = @($parsedIP.IPAddressToString)
        }
    }

    $addresses = [System.Net.Dns]::GetHostAddresses($hostName) |
        ForEach-Object { $_.IPAddressToString } |
        Sort-Object -Unique

    if (-not $addresses -or $addresses.Count -eq 0) {
        throw "Nenhum endereço IP foi encontrado para $hostName."
    }

    return [pscustomobject]@{
        Host = $hostName
        IPs  = @($addresses)
    }
}

function Invoke-DNSLookup {
    $BtnLookup.IsEnabled = $false
    $BtnLookup.Content = "Consultando..."
    $LblStatus.Text = "Consultando DNS"
    $LblStatus.Foreground = [Windows.Media.Brushes]::LightSkyBlue
    $LblPrimaryIP.Text = "—"
    $LstIPs.Items.Clear()

    try {
        $result = Resolve-DomainFromInput -InputText $TxtDomain.Text
        $LblPrimaryIP.Text = $result.IPs[0]
        foreach ($ip in $result.IPs) {
            [void]$LstIPs.Items.Add($ip)
        }
        $LblStatus.Text = "Concluído • $($result.Host)"
        $LblStatus.Foreground = [Windows.Media.Brushes]::LightGreen
    }
    catch {
        $LblStatus.Text = "Falha na consulta"
        $LblStatus.Foreground = [Windows.Media.Brushes]::Salmon
        $LblPrimaryIP.Text = "Não encontrado"
        [void]$LstIPs.Items.Add($_.Exception.Message)
    }
    finally {
        $BtnLookup.IsEnabled = $true
        $BtnLookup.Content = "Buscar IP  ↵"
    }
}

$BtnLookup.Add_Click({ Invoke-DNSLookup })
$TxtDomain.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
        Invoke-DNSLookup
        $_.Handled = $true
    }
})
$BtnClear.Add_Click({
    $TxtDomain.Clear()
    $LstIPs.Items.Clear()
    $LblPrimaryIP.Text = "—"
    $LblStatus.Text = "Aguardando consulta"
    $LblStatus.Foreground = [Windows.Media.Brushes]::Gray
    $TxtDomain.Focus()
})

$Window.Add_ContentRendered({ $TxtDomain.Focus() })
[void]$Window.ShowDialog()
