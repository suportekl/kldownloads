Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$url = "https://kl-quartz.com.br/tecnicokl"
$pastaDestino = "$env:USERPROFILE\Downloads\KL-Quartz"

if (-not (Test-Path $pastaDestino)) {
    New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null
}

# Janela
$form = New-Object Windows.Forms.Form
$form.Text = "📦 KL-Quartz Downloader"
$form.Size = New-Object Drawing.Size(600, 500) # Aumentar a altura para 500
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Função para criar botões estilizados
function Novo-Botao($texto, $x, $y) {
    $btn = New-Object Windows.Forms.Button
    $btn.Text = $texto
    $btn.Location = New-Object Drawing.Point($x, $y)
    $btn.Size = New-Object Drawing.Size(150, 40)
    $btn.BackColor = [System.Drawing.Color]::FromArgb(60, 120, 216)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatStyle = "Flat"
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    return $btn
}

# Botões principais
$btnDownloadPage = Novo-Botao "Lista de Downloads" 20 20
$btnIPsPage = Novo-Botao "Listar IP's" 180 20
$btnDownload = Novo-Botao "⬇ Baixar Arquivo" 20 400
$btnOpen = Novo-Botao "Abrir Pasta" 200 400
$btnSair = Novo-Botao "❌ Sair" 380 400

# Tela de Downloads
$listBox = New-Object Windows.Forms.ListBox
$listBox.Location = New-Object Drawing.Point(20, 60)
$listBox.Size = New-Object Drawing.Size(540, 250)
$listBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$listBox.BorderStyle = "FixedSingle"
$form.Controls.Add($listBox)

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Location = New-Object Drawing.Point(20, 320)
$progressBar.Size = New-Object Drawing.Size(540, 20)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

# Tela de IPs
$labelIPs = New-Object Windows.Forms.Label
$labelIPs.Text = "Lista de IPs e MACs"
$labelIPs.AutoSize = $true
$labelIPs.Location = New-Object Drawing.Point(20, 60)
$labelIPs.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)

$txtIPs = New-Object Windows.Forms.TextBox
$txtIPs.Multiline = $true
$txtIPs.Location = New-Object Drawing.Point(20, 100)
$txtIPs.Size = New-Object Drawing.Size(540, 220)
$txtIPs.Font = New-Object System.Drawing.Font("Consolas", 10)
$txtIPs.ScrollBars = "Vertical"

$progressBarIPs = New-Object Windows.Forms.ProgressBar
$progressBarIPs.Location = New-Object Drawing.Point(20, 330)
$progressBarIPs.Size = New-Object Drawing.Size(540, 20)
$progressBarIPs.Style = 'Continuous'

$labelIPStatus = New-Object Windows.Forms.Label
$labelIPStatus.Location = New-Object Drawing.Point(20, 360)
$labelIPStatus.Size = New-Object Drawing.Size(540, 20)
$labelIPStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$form.Controls.AddRange(@($labelIPs, $txtIPs, $progressBarIPs, $labelIPStatus))

# Função para obter lista de arquivos
function Get-Files {
    try {
        $html = Invoke-WebRequest -Uri $url -UseBasicParsing
        $pattern = '<a[^>]*href="(?<href>https[^"]+)"[^>]*>(?<nome>[^<]+)</a>'
        $matches = [regex]::Matches($html.Content, $pattern)

        $resultado = @()
        foreach ($m in $matches) {
            $href = $m.Groups["href"].Value
            $nome = $m.Groups["nome"].Value.Trim()
            if ($href -match "onedrive\.live\.com|\.exe$|\.zip$|\.pdf$|\.rar$|\.msi$|\.docx$|\.xlsx$") {
                $resultado += [PSCustomObject]@{ Nome = $nome; href = $href }
            }
        }
        return $resultado
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erro ao acessar a página: $_")
        return @()
    }
}

# Função para listar IPs da rede
function Listar-IPs {
    # Alternar visibilidade
    $listBox.Visible = $false
    $progressBar.Visible = $false
    $btnDownload.Visible = $false
    $btnOpen.Visible = $false
    $btnSair.Visible = $false

    $labelIPs.Visible = $true
    $txtIPs.Visible = $true
    $progressBarIPs.Visible = $true
    $labelIPStatus.Visible = $true

    $labelIPStatus.Text = "Buscando dispositivos..."
    $txtIPs.Text = ""
    [System.Windows.Forms.Application]::DoEvents()

    $arpTable = arp -a | Select-String "192.168"
    $total = $arpTable.Count
    $progressBarIPs.Maximum = $total
    $progressBarIPs.Value = 0

    $ipsList = ("{0,-18} {1,-20} {2}" -f "IP", "MAC", "Fabricante") + "`r`n"
    $ipsList += "-"*65 + "`r`n"

    $index = 0
    foreach ($entry in $arpTable) {
        $index++
        $progressBarIPs.Value = $index
        [System.Windows.Forms.Application]::DoEvents()

        $parts = $entry -split '\s+'
        if ($parts.Length -ge 3) {
            $deviceIp = $parts[1]
            $mac = $parts[2]

            if ($mac -ne "ff-ff-ff-ff-ff-ff" -and $mac -ne "00-00-00-00-00-00") {
                $macClean = $mac -replace '-', ':'
                try {
                    $vendor = Invoke-RestMethod -Uri "https://api.macvendors.com/$macClean" -TimeoutSec 3
                } catch {
                    $vendor = "Desconhecido"
                }

                $ipsList += ("{0,-18} {1,-20} {2}" -f $deviceIp, $mac, $vendor) + "`r`n"
            }
        }
    }

    $progressBarIPs.Visible = $false
    $labelIPStatus.Text = "Dispositivos encontrados:"
    $txtIPs.Text = $ipsList
}

# Ações dos botões
$btnDownloadPage.Add_Click({
    $listBox.Items.Clear()
    $arquivos = Get-Files
    $arquivos | ForEach-Object { $listBox.Items.Add($_.Nome) }

    $listBox.Visible = $true
    $progressBar.Visible = $true
    $labelIPs.Visible = $false
    $txtIPs.Visible = $false
    $progressBarIPs.Visible = $false
    $labelIPStatus.Visible = $false

    $btnDownload.Visible = $true
    $btnOpen.Visible = $true
    $btnSair.Visible = $true
})

$btnIPsPage.Add_Click({
    Listar-IPs
})

$btnDownload.Add_Click({
    if ($listBox.SelectedItem) {
        $nome = $listBox.SelectedItem
        $arquivo = $arquivos | Where-Object { $_.Nome -eq $nome }
        if ($arquivo) {
            $novoNome = $arquivo.Nome
            $caminho = Join-Path $pastaDestino $novoNome
            try {
                $response = Invoke-WebRequest -Uri $arquivo.href -Method Head -UseBasicParsing
                $totalBytes = [int]$response.Headers["Content-Length"]
                $progressBar.Maximum = $totalBytes

                $client = New-Object System.Net.WebClient
                $client.add_DownloadProgressChanged({
                    param($sender, $e)
                    $progressBar.Value = $e.BytesReceived
                })

                $client.DownloadFileAsync($arquivo.href, $caminho)
                while ($client.IsBusy) { Start-Sleep -Seconds 1 }

                [System.Windows.Forms.MessageBox]::Show("Download concluído para: $caminho", "Sucesso", "OK", "Information")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Erro ao baixar: $_")
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Selecione um arquivo.")
    }
})

$btnOpen.Add_Click({ Start-Process explorer.exe $pastaDestino })
$btnSair.Add_Click({ $form.Close() })

$form.Controls.AddRange(@($btnDownloadPage, $btnIPsPage, $btnDownload, $btnOpen, $btnSair))

# Exibir a janela
$form.Topmost = $true
$form.ShowDialog()
