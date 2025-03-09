# Fully optimized PowerShell GUI with enhanced features (dynamic button, failure counter, packet loss display, dynamic range, customizable intervals, data usage tracking, data usage limit, clear status updates, and error handling)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Form initialization
$form = New-Object System.Windows.Forms.Form
$form.Text = "Network Monitor"
$form.Size = New-Object System.Drawing.Size(800,650)
$form.StartPosition = "CenterScreen"

# Input for IP or website
$label = New-Object System.Windows.Forms.Label
$label.Text = "Enter IP or Website:"
$label.Location = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(150,20)
$textBox.Location = New-Object System.Drawing.Point(140,18)
$form.Controls.Add($textBox)

# Ping Interval input
$intervalLabel = New-Object System.Windows.Forms.Label
$intervalLabel.Text = "Interval (sec):"
$intervalLabel.Location = New-Object System.Drawing.Point(300,20)
$form.Controls.Add($intervalLabel)

$intervalInput = New-Object System.Windows.Forms.NumericUpDown
$intervalInput.Location = New-Object System.Drawing.Point(400,18)
$intervalInput.Minimum = 5
$intervalInput.Maximum = 3600
$intervalInput.Value = 60
$form.Controls.Add($intervalInput)

# Data limit input
$dataLimitLabel = New-Object System.Windows.Forms.Label
$dataLimitLabel.Text = "Data Limit (KB):"
$dataLimitLabel.Location = New-Object System.Drawing.Point(20,50)
$form.Controls.Add($dataLimitLabel)

$dataLimitInput = New-Object System.Windows.Forms.NumericUpDown
$dataLimitInput.Location = New-Object System.Drawing.Point(140,48)
$dataLimitInput.Minimum = 1
$dataLimitInput.Maximum = 10000
$dataLimitInput.Value = 500
$form.Controls.Add($dataLimitInput)

# Button to start/stop monitoring
$toggleButton = New-Object System.Windows.Forms.Button
$toggleButton.Text = "Start Monitoring"
$toggleButton.AutoSize = $true
$toggleButton.Location = New-Object System.Drawing.Point(600,45)
$form.Controls.Add($toggleButton)

# Labels for statistics
$failureCounter = New-Object System.Windows.Forms.Label
$failureCounter.Text = "Failures: 0"
$failureCounter.AutoSize = $true
$failureCounter.Location = New-Object System.Drawing.Point(20,510)
$form.Controls.Add($failureCounter)

$packetLossLabel = New-Object System.Windows.Forms.Label
$packetLossLabel.Text = "Packet Loss: 0%"
$packetLossLabel.AutoSize = $true
$packetLossLabel.Location = New-Object System.Drawing.Point(20,530)
$form.Controls.Add($packetLossLabel)

$dataUsageLabel = New-Object System.Windows.Forms.Label
$dataUsageLabel.Text = "Estimated Data Used: 0 KB"
$dataUsageLabel.AutoSize = $true
$dataUsageLabel.Location = New-Object System.Drawing.Point(20,550)
$form.Controls.Add($dataUsageLabel)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Idle"
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(20,570)
$form.Controls.Add($statusLabel)

# Chart setup
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chart.Size = New-Object System.Drawing.Size(750,400)
$chart.Location = New-Object System.Drawing.Point(20,80)
$form.Controls.Add($chart)

$chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartArea.AxisX.Title = "Time"
$chartArea.AxisY.Title = "Response Time (ms)"
$chart.ChartAreas.Add($chartArea)

# Series Online
$seriesOnline = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$seriesOnline.Name = "Online"
$seriesOnline.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
$seriesOnline.Color = [System.Drawing.Color]::Green
$chart.Series.Add($seriesOnline)

# Series Offline
$offlineSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$offlineSeries.Name = "Offline"
$offlineSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Point
$offlineSeries.Color = [System.Drawing.Color]::Red
$offlineSeries.MarkerSize = 10
$offlineSeries.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
$chart.Series.Add($offlineSeries)

# Timer setup
$timer = New-Object System.Windows.Forms.Timer

# Variables
$global:isMonitoring = $false
$global:failureCount = 0
$global:pingCount = 0
$global:dataUsedKB = 0

$toggleButton.Add_Click({
    if (-not $global:isMonitoring) {
        $global:target = if ([string]::IsNullOrWhiteSpace($textBox.Text)) { "localhost" } else { $textBox.Text }
        $global:startTime = Get-Date
        $global:failureCount = 0
        $global:pingCount = 0
        $global:dataUsedKB = 0

        $seriesOnline.Points.Clear()
        $offlineSeries.Points.Clear()
        $timer.Interval = [int]($intervalInput.Value * 1000)

        $statusLabel.Text = "Status: Monitoring started for $($global:target) at $(Get-Date -Format 'HH:mm:ss')"
        $toggleButton.Text = "Stop Monitoring"
        $timer.Start()
        $global:isMonitoring = $true
    }
    else {
        $timer.Stop()
        $statusLabel.Text = "Status: Monitoring stopped at $(Get-Date -Format 'HH:mm:ss')"
        $toggleButton.Text = "Start Monitoring"
        $global:isMonitoring = $false
    }
})

$timer.Add_Tick({
    $global:pingCount++
    $global:dataUsedKB += 0.1 # Approximate data per ping

    try {
        $pingResult = Test-Connection -ComputerName $global:target -Count 1 -ErrorAction Stop
        $responseTime = $pingResult.ResponseTime
        $seriesOnline.Points.AddXY((Get-Date -Format "HH:mm:ss"), $responseTime)
        $offlineSeries.Points.AddXY((Get-Date -Format "HH:mm:ss"), [double]::NaN)
    }
    catch {
        $global:failureCount++
        $responseTime = 0
        $offlineSeries.Points.AddXY((Get-Date -Format "HH:mm:ss"), 0)
        $failureCounter.Text = "Failures: $global:failureCount"
    }

    $packetLossPercent = [math]::Round(($global:failureCount / $global:pingCount) * 100, 2)
    $packetLossLabel.Text = "Packet Loss: $packetLossPercent%"
    $dataUsageLabel.Text = "Estimated Data Used: $([math]::Round($global:dataUsedKB,2)) KB"

    if ($global:dataUsedKB -ge $dataLimitInput.Value) {
        $timer.Stop()
        $statusLabel.Text = "Status: Data limit of $($dataLimitInput.Value) KB reached, monitoring stopped."
        $toggleButton.Text = "Start Monitoring"
        $global:isMonitoring = $false
    }

    if ($seriesOnline.Points.Count -gt 50) {
        $seriesOnline.Points.RemoveAt(0)
        $offlineSeries.Points.RemoveAt(0)
    }

    $chart.ResetAutoValues()
})

$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
