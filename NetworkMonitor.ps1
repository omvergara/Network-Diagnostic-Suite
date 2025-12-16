<#
    ===========================================================================
    SUITE DE DIAGNÓSTICO DE RED - V1.5 (GITHUB RELEASE)
    Autor: Mauricio Vergara
    
    CARACTERÍSTICAS:
    - Monitor en tiempo real con gráficas de latencia.
    - Sistema "Double-Check" para evitar falsos positivos en Wi-Fi.
    - Carga interactiva de Impresoras (evita bloqueos por IPs inexistentes).
    - Exportación automática de logs a CSV.
    - Interfaz gráfica completa sin dependencias externas.
    ===========================================================================
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# ==================================================
# 1. CARGA DE CONFIGURACIÓN (JSON)
# ==================================================
$RutaBase = [System.AppDomain]::CurrentDomain.BaseDirectory
$RutaConfig = Join-Path $RutaBase "config_monitor.json"
$Hostname = $env:COMPUTERNAME
$Usuario = "$env:USERDOMAIN\$env:USERNAME"

# Configuración por defecto (SANITIZADA PARA GITHUB)
$ConfigDefault = @{
    General = @{ 
        NombreTecnico = $env:USERNAME 
        IntervaloSeg = 2          
        TimeoutPing = 1000        
        UmbralLentitud = 150 
        AlertaSonora = $true 
        LogPath = "$env:USERPROFILE\Desktop\ReporteRed_$(Get-Date -Format 'yyyyMMdd').csv" 
    }
    Email = @{ 
        Activar = $false 
        ServidorSMTP = "smtp.office365.com" 
        Puerto = 587 
        Usuario = "usuario@empresa.com" 
        Password = "" 
        Desde = "alerta@empresa.com" 
        Para = "soporte@empresa.com" 
    }
    Servicios = @( 
        @{ Nombre="Gateway"; IP="AUTO"; Color="Blue" }, 
        @{ Nombre="Google"; IP="8.8.8.8"; Color="Green" }, 
        @{ Nombre="Cloudflare"; IP="1.1.1.1"; Color="OrangeRed" }, 
        @{ Nombre="Office365"; IP="outlook.office365.com"; Color="Purple" } 
    )
    Impresoras = @( 
        @{ Nombre="Ejemplo HP"; IP="192.168.1.50"; Color="Teal" }
    )
}

if (-not (Test-Path $RutaConfig)) { 
    $ConfigDefault | ConvertTo-Json -Depth 4 | Out-File $RutaConfig -Encoding UTF8
    $DatosConfig = $ConfigDefault 
} else { 
    try { $DatosConfig = Get-Content $RutaConfig -Raw | ConvertFrom-Json } catch { $DatosConfig = $ConfigDefault } 
}

$Config = @{ 
    IntervaloSegundos = [Math]::Max(2, [int]$DatosConfig.General.IntervaloSeg); 
    MaxPuntosGrafica = 80; 
    TimeoutPing = [Math]::Max(500, [int]$DatosConfig.General.TimeoutPing); 
    UmbralLentitud = $DatosConfig.General.UmbralLentitud; 
    RutaLog = $DatosConfig.General.LogPath; 
    AlertaSonora = $DatosConfig.General.AlertaSonora; 
    NombreTecnico = $DatosConfig.General.NombreTecnico 
}
$CorreoConfig = $DatosConfig.Email
function Get-Color($c){ try{ return [System.Drawing.Color]::FromName($c) }catch{ return [System.Drawing.Color]::Black } }

$GatewayIP = "127.0.0.1"
try { $GwReal = (Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty NextHop -First 1); if($GwReal){ $GatewayIP = $GwReal } } catch {}

$DestinosServicios = [Ordered]@{}; foreach($i in $DatosConfig.Servicios){ $ip = if($i.IP -eq "AUTO"){$GatewayIP}else{$i.IP}; $DestinosServicios[$i.Nombre] = @{IP=$ip; Color=(Get-Color $i.Color)} }

if (-not (Test-Path $Config.RutaLog)) { "Fecha;Hora;Destino;IP;Latencia_ms;Estado;Info_Extra" | Out-File $Config.RutaLog -Encoding UTF8 }

$script:TiempoGlobal = 0; $script:SeriesDict = @{}; $script:Destinos = [Ordered]@{}; $script:EstadoRed = @{}; $script:ModoTimer = "Infinito"; $script:SegundosRestantes = 0
$script:ContadorFallosConsecutivos = 0; $script:UltimoCorreoEnviado = (Get-Date).AddMinutes(-20)
try{ $IPLocal = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike "169.*" -and $_.IPAddress -notlike "127.*"}).IPAddress[0] }catch{ $IPLocal="Sin Red" }
try{ $MacAddr = (Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | Select-Object -First 1).MacAddress }catch{ $MacAddr="N/A" }

# ==================================================
# 2. INTERFAZ GRÁFICA
# ==================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Suite de Red Corporativa - $Hostname"
$form.Size = New-Object System.Drawing.Size(1420, 900)
$form.StartPosition = "CenterScreen"; $form.BackColor = [System.Drawing.Color]::WhiteSmoke

$tabControl = New-Object System.Windows.Forms.TabControl; $tabControl.Dock = "Fill"; $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tabInicio  = New-Object System.Windows.Forms.TabPage; $tabInicio.Text  = "  INICIO / AYUDA  "
$tabMonitor = New-Object System.Windows.Forms.TabPage; $tabMonitor.Text = "  MONITOR EN VIVO  "
$tabDiag    = New-Object System.Windows.Forms.TabPage; $tabDiag.Text    = "  ANÁLISIS DE RED  "
$tabControl.Controls.AddRange(@($tabInicio, $tabMonitor, $tabDiag))
$form.Controls.Add($tabControl)

# --- PESTAÑA 1: INICIO (TEXTO ORIGINAL RESTAURADO) ---
$tabInicio.BackColor = [System.Drawing.Color]::WhiteSmoke
$panelHelpContainer = New-Object System.Windows.Forms.Panel; $panelHelpContainer.Dock = "Fill"; $panelHelpContainer.Padding = New-Object System.Windows.Forms.Padding(40); $tabInicio.Controls.Add($panelHelpContainer)
$txtHelp = New-Object System.Windows.Forms.RichTextBox; $txtHelp.Dock = "Fill"; $txtHelp.BorderStyle = "FixedSingle"; $txtHelp.BackColor = [System.Drawing.Color]::White; $txtHelp.ReadOnly = $true; $panelHelpContainer.Controls.Add($txtHelp)

function Add-StyledText($text, $size, $bold, $color, $newline=$true) {
    $txtHelp.SelectionStart = $txtHelp.TextLength
    $style = if($bold){ [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $txtHelp.SelectionFont = New-Object System.Drawing.Font("Segoe UI", $size, $style)
    $txtHelp.SelectionColor = $color
    $txtHelp.SelectionIndent = 20; $txtHelp.SelectionRightIndent = 20
    $txtHelp.AppendText($text); if($newline){ $txtHelp.AppendText("`n") }
}

Add-StyledText "`n" 5 $false [System.Drawing.Color]::White
Add-StyledText "SUITE DE DIAGNÓSTICO DE RED" 24 $true [System.Drawing.Color]::DarkBlue
Add-StyledText "Herramienta Corporativa - Desarrollado por: $($Config.NombreTecnico)" 12 $true [System.Drawing.Color]::DimGray
Add-StyledText "`nEsta aplicación es una herramienta portátil diseñada para diagnosticar el estado de la red en tiempo real, identificar lentitud, caídas de servicio y analizar la calidad de la señal Wi-Fi sin requerir instalación.`n" 11 $false [System.Drawing.Color]::Black

Add-StyledText "MÓDULOS DISPONIBLES" 14 $true [System.Drawing.Color]::SteelBlue
Add-StyledText "1. MONITOR EN VIVO (Pestaña 2):" 11 $true [System.Drawing.Color]::Black
Add-StyledText "   Es un 'Electrocardiograma' de su red. Muestra gráficas en vivo hacia servidores críticos." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "   - Si la línea es VERDE: La conexión es estable." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "   - Si la línea es NARANJA: La red está lenta." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "   - Si la línea es ROJA: Se ha caído el servicio." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "   * Nota: Se genera automáticamente un reporte en Excel (.csv) en su Escritorio." 11 $false [System.Drawing.Color]::DimGray

Add-StyledText "`n2. ANÁLISIS DE RED (Pestaña 3):" 11 $true [System.Drawing.Color]::Black
Add-StyledText "   Herramientas para técnicos:" 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "   - Escáner Wi-Fi: Muestra la potencia de señal (%), canal y banda." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "   - Traza: Permite ver la ruta de conexión para saber dónde se corta internet." 11 $false [System.Drawing.Color]::DimGray

Add-StyledText "`nPERSONALIZACIÓN PARA SU EMPRESA" 14 $true [System.Drawing.Color]::SteelBlue
Add-StyledText "Esta herramienta es configurable. Al ejecutarla por primera vez, se creó un archivo llamado:" 11 $false [System.Drawing.Color]::Black
Add-StyledText "config_monitor.json" 11 $true [System.Drawing.Color]::DarkRed
Add-StyledText "en la misma carpeta donde está este programa." 11 $false [System.Drawing.Color]::Black
Add-StyledText "`nInstrucciones para adaptar:" 11 $true [System.Drawing.Color]::Black
Add-StyledText "1. Cierre esta aplicación." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "2. Abra el archivo 'config_monitor.json' con el Bloc de Notas." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "3. En la sección 'Email', configure sus credenciales para recibir alertas automáticas." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "4. Guarde el archivo y vuelva a abrir esta aplicación." 11 $false [System.Drawing.Color]::DimGray
Add-StyledText "`nConsejo: Para un diagnóstico real, cierre Spotify, YouTube o Descargas antes de iniciar." 11 $true [System.Drawing.Color]::DarkOrange

# --- PESTAÑA 2: MONITOR ---
$panelFooter = New-Object System.Windows.Forms.Panel; $panelFooter.Dock = "Bottom"; $panelFooter.Height = 30; $panelFooter.BackColor = [System.Drawing.Color]::LightGray
$lblStatus = New-Object System.Windows.Forms.Label; $lblStatus.Text = "Equipo: $Hostname | Usuario: $Usuario | IP: $IPLocal | MAC: $MacAddr"; $lblStatus.AutoSize=$true; $lblStatus.Location=New-Object System.Drawing.Point(10,5); $panelFooter.Controls.Add($lblStatus); $tabMonitor.Controls.Add($panelFooter)

$panelChecks = New-Object System.Windows.Forms.Panel; $panelChecks.Dock = "Bottom"; $panelChecks.Height = 50; $panelChecks.BackColor = [System.Drawing.Color]::White; $panelChecks.BorderStyle = "FixedSingle"
$lblFilter = New-Object System.Windows.Forms.Label; $lblFilter.Text = "Filtros:"; $lblFilter.Location = New-Object System.Drawing.Point(10, 15); $lblFilter.AutoSize=$true; $lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $flowPanel.Location = New-Object System.Drawing.Point(70, 5); $flowPanel.Size = New-Object System.Drawing.Size(1200, 40); $flowPanel.FlowDirection = "LeftToRight"; $flowPanel.WrapContents = $false; $flowPanel.Anchor="Left, Right"
$panelChecks.Controls.Add($lblFilter); $panelChecks.Controls.Add($flowPanel); $tabMonitor.Controls.Add($panelChecks)

$panelHeader = New-Object System.Windows.Forms.Panel; $panelHeader.Dock = "Top"; $panelHeader.Height = 85; $panelHeader.BackColor = [System.Drawing.Color]::WhiteSmoke
$grpTime = New-Object System.Windows.Forms.GroupBox; $grpTime.Text = "Temporizador"; $grpTime.Location = New-Object System.Drawing.Point(10, 5); $grpTime.Size = New-Object System.Drawing.Size(350, 70)
$btnInf = New-Object System.Windows.Forms.RadioButton; $btnInf.Text="Infinito"; $btnInf.Checked=$true; $btnInf.Location=New-Object System.Drawing.Point(15,25); $btnInf.AutoSize=$true
$btn5m  = New-Object System.Windows.Forms.RadioButton; $btn5m.Text="5 Min"; $btn5m.Location=New-Object System.Drawing.Point(90,25); $btn5m.AutoSize=$true
$btn10m = New-Object System.Windows.Forms.RadioButton; $btn10m.Text="10 Min"; $btn10m.Location=New-Object System.Drawing.Point(160,25); $btn10m.AutoSize=$true
$lblTimer = New-Object System.Windows.Forms.Label; $lblTimer.Text = "MODO: INFINITO"; $lblTimer.Location = New-Object System.Drawing.Point(230, 25); $lblTimer.AutoSize=$true; $lblTimer.ForeColor = [System.Drawing.Color]::DarkRed; $lblTimer.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpTime.Controls.AddRange(@($btnInf, $btn5m, $btn10m, $lblTimer))

$grpScen = New-Object System.Windows.Forms.GroupBox; $grpScen.Text = "Seleccionar Objetivo"; $grpScen.Location = New-Object System.Drawing.Point(370, 5); $grpScen.Size = New-Object System.Drawing.Size(800, 70)
$btnServ = New-Object System.Windows.Forms.Button; $btnServ.Text="SERVICIOS WEB"; $btnServ.Location=New-Object System.Drawing.Point(15,25); $btnServ.Size=New-Object System.Drawing.Size(110,30); $btnServ.BackColor=[System.Drawing.Color]::SteelBlue; $btnServ.ForeColor="White"
$btnImp = New-Object System.Windows.Forms.Button; $btnImp.Text="IMPRESORAS"; $btnImp.Location=New-Object System.Drawing.Point(130,25); $btnImp.Size=New-Object System.Drawing.Size(110,30); $btnImp.BackColor=[System.Drawing.Color]::SteelBlue; $btnImp.ForeColor="White"

$lblManual = New-Object System.Windows.Forms.Label; $lblManual.Text = "Manual (Ej: 1.1.1.1, google.com):"; $lblManual.Location = New-Object System.Drawing.Point(250, 30); $lblManual.AutoSize=$true
$txtCustom = New-Object System.Windows.Forms.TextBox; $txtCustom.Location=New-Object System.Drawing.Point(460,28); $txtCustom.Size=New-Object System.Drawing.Size(220,25); $txtCustom.PlaceholderText="Escriba IP o Dominio..."
$btnManual = New-Object System.Windows.Forms.Button; $btnManual.Text="TEST"; $btnManual.Location=New-Object System.Drawing.Point(690,25); $btnManual.Size=New-Object System.Drawing.Size(80,30); $btnManual.BackColor=[System.Drawing.Color]::DimGray; $btnManual.ForeColor="White"

$grpScen.Controls.AddRange(@($btnServ, $btnImp, $lblManual, $txtCustom, $btnManual))
$btnToggle = New-Object System.Windows.Forms.Button; $btnToggle.Text = "PAUSAR TEST"; $btnToggle.Location = New-Object System.Drawing.Point(1200, 10); $btnToggle.Size=New-Object System.Drawing.Size(150,60); $btnToggle.Anchor="Right"; $btnToggle.BackColor = [System.Drawing.Color]::DarkGreen; $btnToggle.ForeColor="White"; $btnToggle.Font=New-Object System.Drawing.Font("Arial",10,"Bold")
$panelHeader.Controls.AddRange(@($grpTime, $grpScen, $btnToggle))
$tabMonitor.Controls.Add($panelHeader)

$panelCentral = New-Object System.Windows.Forms.Panel; $panelCentral.Dock = "Fill"; $panelCentral.Padding = New-Object System.Windows.Forms.Padding(10); $tabMonitor.Controls.Add($panelCentral); $panelCentral.BringToFront()
$list = New-Object System.Windows.Forms.ListView; $list.View = 'Details'; $list.FullRowSelect = $true; $list.GridLines = $true; $list.Dock = "Left"; $list.Width = 450; $list.Font = New-Object System.Drawing.Font("Consolas", 9)
$list.Columns.Add("Hora", 70); $list.Columns.Add("Destino", 90); $list.Columns.Add("IP", 90); $list.Columns.Add("Ms", 40); $list.Columns.Add("Prom", 50); $list.Columns.Add("Estado", 80)
$panelCentral.Controls.Add($list)
$splitter = New-Object System.Windows.Forms.Splitter; $splitter.Dock = "Left"; $splitter.Width = 10; $splitter.BackColor = [System.Drawing.Color]::WhiteSmoke; $panelCentral.Controls.Add($splitter)
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart; $chart.Dock = "Fill"; $chart.BackColor = [System.Drawing.Color]::White; $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea; $area.AxisX.Title="Tiempo"; $area.AxisY.Title="Ms"; $area.AxisX.MajorGrid.LineColor="LightGray"; $area.AxisY.MajorGrid.LineColor="LightGray"; $chart.ChartAreas.Add($area); $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend; $legend.Docking="Top"; $legend.Alignment="Center"; $chart.Legends.Add($legend); $panelCentral.Controls.Add($chart); $chart.BringToFront()

# --- PESTAÑA 3: DIAGNÓSTICO ---
$tabDiag.BackColor = [System.Drawing.Color]::WhiteSmoke
$panelDiagContainer = New-Object System.Windows.Forms.Panel; $panelDiagContainer.Dock="Fill"; $panelDiagContainer.Padding=New-Object System.Windows.Forms.Padding(10); $tabDiag.Controls.Add($panelDiagContainer)
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel; $tableLayout.Dock = "Fill"; $tableLayout.ColumnCount = 2; $tableLayout.RowCount = 1; $tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))); $tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))); $panelDiagContainer.Controls.Add($tableLayout)

$grpScan = New-Object System.Windows.Forms.GroupBox; $grpScan.Text = "Información de Adaptadores y Wi-Fi"; $grpScan.Dock="Fill"; $grpScan.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$txtScan = New-Object System.Windows.Forms.RichTextBox; $txtScan.Dock = "Fill"; $txtScan.BackColor = [System.Drawing.Color]::Black; $txtScan.ForeColor = [System.Drawing.Color]::Lime; $txtScan.Font = New-Object System.Drawing.Font("Consolas", 10)
$btnDoScan = New-Object System.Windows.Forms.Button; $btnDoScan.Text = "ACTUALIZAR ESCÁNER"; $btnDoScan.Dock = "Bottom"; $btnDoScan.Height = 40; $btnDoScan.BackColor = [System.Drawing.Color]::SteelBlue; $btnDoScan.ForeColor="White"
$grpScan.Controls.Add($txtScan); $grpScan.Controls.Add($btnDoScan); $tableLayout.Controls.Add($grpScan, 0, 0)

$grpTrace = New-Object System.Windows.Forms.GroupBox; $grpTrace.Text = "Herramienta de Traza (TraceRoute)"; $grpTrace.Dock="Fill"; $grpTrace.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$panelTraceTop = New-Object System.Windows.Forms.Panel; $panelTraceTop.Dock = "Top"; $panelTraceTop.Height = 50
$txtTraceIP = New-Object System.Windows.Forms.TextBox; $txtTraceIP.Text = "8.8.8.8"; $txtTraceIP.Location = New-Object System.Drawing.Point(10, 15); $txtTraceIP.Size = New-Object System.Drawing.Size(200, 25)
$btnTrace = New-Object System.Windows.Forms.Button; $btnTrace.Text = "INICIAR TRAZA"; $btnTrace.Location = New-Object System.Drawing.Point(220, 12); $btnTrace.Size = New-Object System.Drawing.Size(150, 30); $btnTrace.BackColor = [System.Drawing.Color]::DarkGreen; $btnTrace.ForeColor="White"
$panelTraceTop.Controls.AddRange(@($txtTraceIP, $btnTrace))
$txtTraceOut = New-Object System.Windows.Forms.RichTextBox; $txtTraceOut.Dock = "Fill"; $txtTraceOut.BackColor = [System.Drawing.Color]::Black; $txtTraceOut.ForeColor = [System.Drawing.Color]::Yellow; $txtTraceOut.Font = New-Object System.Drawing.Font("Consolas", 10); $txtTraceOut.ScrollBars = "Both"; $txtTraceOut.WordWrap = $false
$grpTrace.Controls.Add($txtTraceOut); $grpTrace.Controls.Add($panelTraceTop); $tableLayout.Controls.Add($grpTrace, 1, 0)

# ==================================================
# LÓGICA Y FUNCIONES
# ==================================================
$ActionScan = {
    $txtScan.Clear(); $txtScan.AppendText("=== RESUMEN GENERAL ===`nHostname : $env:COMPUTERNAME`nFecha    : $(Get-Date)`n`n")
    try {
        $netInfo = Get-NetIPConfiguration | Where-Object {$_.IPv4Address -ne $null}
        foreach($n in $netInfo){ $txtScan.AppendText("--- ADAPTADOR: $($n.InterfaceAlias) ---`nIPv4        : $($n.IPv4Address.IPAddress)`nGateway     : $($n.IPv4DefaultGateway.NextHop)`nDNS         : $($n.DNSServer.ServerAddresses -join ', ')`n`n") }
        $txtScan.AppendText("=== DETALLES WI-FI (NETSH) ===`n")
        $p = New-Object System.Diagnostics.Process; $p.StartInfo.UseShellExecute=$false; $p.StartInfo.RedirectStandardOutput=$true; $p.StartInfo.FileName="netsh"; $p.StartInfo.Arguments="wlan show interfaces"; $p.StartInfo.StandardOutputEncoding=[System.Text.Encoding]::UTF8; $p.Start()|Out-Null; $output=$p.StandardOutput.ReadToEnd(); $p.WaitForExit(); $txtScan.AppendText($output)
    } catch { $txtScan.AppendText("Error obteniendo datos.`n") }
}
$btnDoScan.Add_Click($ActionScan)

$ActionTrace = {
    $target = $txtTraceIP.Text
    $txtTraceOut.Clear(); $txtTraceOut.AppendText("Iniciando Traza a $target... Espere.`n"); $form.Refresh()
    try { tracert -d $target | ForEach-Object { $txtTraceOut.AppendText("$($_)`n"); $txtTraceOut.ScrollToCaret(); [System.Windows.Forms.Application]::DoEvents() }; $txtTraceOut.AppendText("`n--- TRAZA COMPLETADA ---") } catch { $txtTraceOut.AppendText("Error.") }
}
$btnTrace.Add_Click($ActionTrace)

function Enviar-Notificacion($destino, $ip){
    $tiempoTranscurrido = (Get-Date) - $script:UltimoCorreoEnviado
    if($CorreoConfig.Activar -and $tiempoTranscurrido.TotalMinutes -gt 15){
        try { 
            $secPass = ConvertTo-SecureString $CorreoConfig.Password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($CorreoConfig.Usuario, $secPass)
            Send-MailMessage -From $CorreoConfig.Desde -To $CorreoConfig.Para -Subject "ALERTA RED: $destino ($ip)" -Body "Caída detectada en $destino.`nEquipo: $Hostname" -SmtpServer $CorreoConfig.ServidorSMTP -Port $CorreoConfig.Puerto -UseSsl -Credential $cred -ErrorAction Stop
            $script:UltimoCorreoEnviado = Get-Date 
        } catch {}
    }
}

function Cargar-Escenario($tipo, $custom=$null){
    $script:TiempoGlobal = 0; $chart.Series.Clear(); $list.Items.Clear(); $flowPanel.Controls.Clear(); $script:SeriesDict.Clear(); $script:EstadoRed.Clear(); $script:ContadorFallosConsecutivos = 0
    if($tipo -eq "Servicios"){ $script:Destinos = $DestinosServicios } elseif($tipo -eq "Impresoras"){ 
        if ($custom) { $script:Destinos = $custom } else { return } 
    } else{ $script:Destinos = $custom }
    
    foreach($n in $script:Destinos.Keys){
        $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series; $s.Name = $n; $s.ChartType = "Line"; $s.BorderWidth = 2; $s.Color = $script:Destinos[$n].Color; $chart.Series.Add($s); $script:SeriesDict[$n] = $s
        $script:EstadoRed[$n] = @{ Historial = New-Object System.Collections.ArrayList; CaidaInicio = $null; EstaCaido = $false; Activo = $true }
        $cb = New-Object System.Windows.Forms.CheckBox; $cb.Text = $n; $cb.Checked = $true; $cb.ForeColor = $script:Destinos[$n].Color; $cb.AutoSize = $true; $cb.Add_CheckedChanged({ $script:EstadoRed[$this.Text].Activo = $this.Checked; $chart.Series[$this.Text].Enabled = $this.Checked }); $flowPanel.Controls.Add($cb)
    }
}

function Show-PrinterDialog {
    $dlg = New-Object System.Windows.Forms.Form; $dlg.Text = "Configurar Impresoras"; $dlg.Size = New-Object System.Drawing.Size(400, 400); $dlg.StartPosition = "CenterParent"
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Ingrese las impresoras (Nombre, IP) - Una por línea:"; $lbl.Dock = "Top"; $lbl.Padding = New-Object System.Windows.Forms.Padding(10); $lbl.Height = 40
    $txt = New-Object System.Windows.Forms.TextBox; $txt.Multiline = $true; $txt.Dock = "Fill"; $txt.ScrollBars = "Vertical"; $txt.Text = "HP Gerencia,192.168.1.50`r`nRicoh Bodega,192.168.1.51"
    $pnlBtn = New-Object System.Windows.Forms.Panel; $pnlBtn.Dock = "Bottom"; $pnlBtn.Height = 50
    $btnOK = New-Object System.Windows.Forms.Button; $btnOK.Text = "CARGAR"; $btnOK.DialogResult = "OK"; $btnOK.Dock="Right"; $btnOK.Width=100
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "CANCELAR"; $btnCancel.DialogResult = "Cancel"; $btnCancel.Dock="Left"; $btnCancel.Width=100
    $pnlBtn.Controls.Add($btnOK); $pnlBtn.Controls.Add($btnCancel); $dlg.Controls.Add($txt); $dlg.Controls.Add($lbl); $dlg.Controls.Add($pnlBtn)
    if ($dlg.ShowDialog() -eq "OK") { return $txt.Text } else { return $null }
}

$ActionManual = { 
    $raw = $txtCustom.Text
    if([string]::IsNullOrWhiteSpace($raw)){ [System.Windows.Forms.MessageBox]::Show("Ingrese una IP o Dominio.", "Vacio", "OK", "Warning"); return }
    $ips = $raw.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if($ips.Count -gt 0){ 
        $c=@([System.Drawing.Color]::Blue, [System.Drawing.Color]::Red, [System.Drawing.Color]::Green); $d=[Ordered]@{}; $i=0
        foreach($ip in $ips){ $d["Target-$($i)"]=@{IP=$ip;Color=$c[$i%3]};$i++ }
        Cargar-Escenario "Custom" $d 
    }
}

$PingSender = New-Object System.Net.NetworkInformation.Ping
$timer = New-Object System.Windows.Forms.Timer; $timer.Interval = $Config.IntervaloSegundos * 1000

$timer.Add_Tick({
    if($script:ModoTimer -ne "Infinito"){ 
        $script:SegundosRestantes--; $ts = [TimeSpan]::FromSeconds($script:SegundosRestantes); $lblTimer.Text="Restante: "+$ts.ToString("mm\:ss"); if($script:SegundosRestantes -le 0){ $timer.Stop(); [System.Windows.Forms.MessageBox]::Show("Fin."); return } 
    }
    $script:TiempoGlobal++; $ahora = Get-Date; $horaStr = $ahora.ToString("HH:mm:ss"); $EnMonitor = ($tabControl.SelectedTab -eq $tabMonitor)
    if($EnMonitor){ $list.BeginUpdate() }
    $HuboFallo = $false; $Sonar = $false

    foreach($n in $script:Destinos.Keys){
        if(-not $script:EstadoRed[$n].Activo){ continue }
        $ms=0; $st="N/A"; $col=[System.Drawing.Color]::Black; $bg=[System.Drawing.Color]::White; $ex=""; $pingSuccess = $false
        try {
            $res = $PingSender.Send($script:Destinos[$n].IP, $Config.TimeoutPing)
            if ($res.Status -eq "Success") { $pingSuccess = $true; $ms = $res.RoundtripTime } 
            else { Start-Sleep -Milliseconds 50; $res = $PingSender.Send($script:Destinos[$n].IP, $Config.TimeoutPing); if ($res.Status -eq "Success") { $pingSuccess = $true; $ms = $res.RoundtripTime } }
        } catch { $pingSuccess = $false }

        if($pingSuccess){
            if($script:EstadoRed[$n].EstaCaido){ $script:EstadoRed[$n].EstaCaido=$false; $ex="(REC)"; $st="OK"; $col=[System.Drawing.Color]::Green; if($script:ContadorFallosConsecutivos -gt 0){$script:ContadorFallosConsecutivos--} }
            else{ if($ms -gt $Config.UmbralLentitud){ $st="LENTO"; $col=[System.Drawing.Color]::Orange } else { $st="OK"; $col=$script:Destinos[$n].Color } }
            $script:SeriesDict[$n].Points.AddXY($script:TiempoGlobal, $ms); $script:EstadoRed[$n].Historial.Add($ms)|Out-Null
        } else {
            $ms=0; $st="CAÍDA"; $bg=[System.Drawing.Color]::MistyRose; $col=[System.Drawing.Color]::Red; $Sonar=$true; $HuboFallo=$true
            if(-not $script:EstadoRed[$n].EstaCaido){ $script:EstadoRed[$n].EstaCaido=$true; $script:EstadoRed[$n].CaidaInicio=$ahora }
            $idx = $script:SeriesDict[$n].Points.AddXY($script:TiempoGlobal, 0); $script:SeriesDict[$n].Points[$idx].IsEmpty = $true 
            if($script:ContadorFallosConsecutivos -ge 10){ Enviar-Notificacion $n $script:Destinos[$n].IP }
        }
        "$($ahora.ToString('yyyy-MM-dd'));$horaStr;$n;$($script:Destinos[$n].IP);$ms;$st;$ex" | Out-File $Config.RutaLog -Append -Encoding UTF8
        if($EnMonitor){
            $prom=0; if($script:EstadoRed[$n].Historial.Count -gt 0){ $sum=0; $script:EstadoRed[$n].Historial|ForEach-Object{$sum+=$_}; $prom=[Math]::Round($sum/$script:EstadoRed[$n].Historial.Count) }
            $it=New-Object System.Windows.Forms.ListViewItem($horaStr); $it.SubItems.Add($n); $it.SubItems.Add($script:Destinos[$n].IP); $it.SubItems.Add("$ms"); $it.SubItems.Add("$prom"); $it.SubItems.Add("$st $ex"); $it.ForeColor=$col; $it.BackColor=$bg
            $list.Items.Insert(0, $it)
        }
        if($script:EstadoRed[$n].Historial.Count -gt 60){ $script:EstadoRed[$n].Historial.RemoveAt(0) }
        if($script:SeriesDict[$n].Points.Count -gt $Config.MaxPuntosGrafica){ $script:SeriesDict[$n].Points.RemoveAt(0) }
    }
    
    if($HuboFallo){ $script:ContadorFallosConsecutivos++ }else{ $script:ContadorFallosConsecutivos=0 }
    if($EnMonitor){ if($list.Items.Count -gt 100){ for($i=$list.Items.Count-1;$i -ge 100;$i--){$list.Items.RemoveAt($i)} }; $list.EndUpdate(); $chart.ChartAreas[0].RecalculateAxesScale() }
    if($Sonar -and $Config.AlertaSonora){ [System.Console]::Beep(1500,100) }
})

$btnServ.Add_Click({ $timer.Stop(); Cargar-Escenario "Servicios"; $timer.Start() })
$btnImp.Add_Click({ 
    $timer.Stop()
    $raw = Show-PrinterDialog
    if($raw){
        $d = [Ordered]@{}; $cols = @([System.Drawing.Color]::Teal, [System.Drawing.Color]::DarkCyan, [System.Drawing.Color]::SaddleBrown, [System.Drawing.Color]::Purple, [System.Drawing.Color]::Red); $idx=0
        foreach($l in $raw.Split("`n")){ if($l -match ","){ $p=$l.Split(","); $d[$p[0].Trim()] = @{IP=$p[1].Trim(); Color=$cols[$idx%$cols.Count]}; $idx++ } }
        if($d.Count -gt 0){ Cargar-Escenario "Impresoras" $d; $timer.Start() }
    }
})
$btnManual.Add_Click($ActionManual)
$btnToggle.Add_Click({if($timer.Enabled){$timer.Stop();$btnToggle.Text="REANUDAR";$btnToggle.BackColor="Orange"}else{$timer.Start();$btnToggle.Text="PAUSAR";$btnToggle.BackColor="DarkGreen"}})
$btnInf.Add_CheckedChanged({$script:ModoTimer="Infinito"}); $btn5m.Add_CheckedChanged({$script:ModoTimer="5min";$script:SegundosRestantes=300}); $btn10m.Add_CheckedChanged({$script:ModoTimer="10min";$script:SegundosRestantes=600})

Cargar-Escenario "Servicios"
$timer.Start()
[void]$form.ShowDialog()
$timer.Stop()
$PingSender.Dispose()