Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$url = "https://kl-quartz.com.br/tecnicokl"
$pastaDestino = "$env:USERPROFILE\Downloads\KL-Quartz"

if (-not (Test-Path $pastaDestino)) {
    New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null
}

# Janela principal
$form = New-Object Windows.Forms.Form
$form.Text = "📦 KL-Quartz Downloader"
$form.Size = New-Object Drawing.Size(600, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

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

# Botões
$btnDownloadPage = Novo-Botao "Lista de Downloads" 20 20
$btnIPsPage = Novo-Botao "Listar IP's" 180 20
$btnIniciarVarredura = Novo-Botao "Iniciar Varredura" 340 20
$btnDownload = Novo-Botao "⬇ Baixar Arquivo" 20 450
$btnOpen = Novo-Botao "Abrir Pasta" 200 450
$btnSair = Novo-Botao "❌ Sair" 380 450

# Lista de arquivos
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
$txtIPs.Location = New-Object Drawing.Point(20, 140)
$txtIPs.Size = New-Object Drawing.Size(540, 200)
$txtIPs.Font = New-Object System.Drawing.Font("Consolas", 10)
$txtIPs.ScrollBars = "Vertical"

$progressBarIPs = New-Object Windows.Forms.ProgressBar
$progressBarIPs.Location = New-Object Drawing.Point(20, 350)
$progressBarIPs.Size = New-Object Drawing.Size(540, 20)
$progressBarIPs.Style = 'Continuous'

$labelIPStatus = New-Object Windows.Forms.Label
$labelIPStatus.Location = New-Object Drawing.Point(20, 380)
$labelIPStatus.Size = New-Object Drawing.Size(540, 20)
$labelIPStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Input de IP
$labelIPInput = New-Object Windows.Forms.Label
$labelIPInput.Text = "Digite o IP base (ex: 192.168.1):"
$labelIPInput.Location = New-Object Drawing.Point(20, 100)
$labelIPInput.AutoSize = $true

$inputIP = New-Object Windows.Forms.TextBox
$inputIP.Location = New-Object Drawing.Point(260, 97)
$inputIP.Size = New-Object Drawing.Size(150, 25)

$btnIniciarVarredura.Location = New-Object Drawing.Point(420, 95)
$btnIniciarVarredura.Size = New-Object Drawing.Size(140, 30)

$form.Controls.AddRange(@($labelIPs, $txtIPs, $progressBarIPs, $labelIPStatus, $labelIPInput, $inputIP, $btnIniciarVarredura))


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

function Listar-IPs {
    $ipBase = $inputIP.Text.Trim()
    if (-not ($ipBase -match '^\d{1,3}\.\d{1,3}\.\d{1,3}$')) {
        [System.Windows.Forms.MessageBox]::Show("Digite um IP base válido no formato: 192.168.1")
        return
    }

    $labelIPStatus.Text = "Buscando dispositivos em $ipBase.1 a $ipBase.254..."
    $txtIPs.Text = ""
    [System.Windows.Forms.Application]::DoEvents()

    $ipsList = ("{0,-18} {1,-20} {2}" -f "IP", "MAC", "Fabricante") + "`r`n"
    $ipsList += "-"*65 + "`r`n"

    $progressBarIPs.Maximum = 254
    $progressBarIPs.Value = 0

    # Primeiro verifica o gateway
    $gateway = (Get-NetRoute -DestinationPrefix "0.0.0.0/0").NextHop
    if ($gateway) {
        try {
            $arpInfo = arp -a $gateway | Select-String "$gateway\s+([-\w]+)"
            if ($arpInfo -match "$gateway\s+([-\w]+)") {
                $mac = $matches[1]
                $macClean = $mac -replace '-', ':'
                try {
                    $vendor = Invoke-RestMethod -Uri "https://api.macvendors.com/$macClean" -TimeoutSec 3
                } catch {
                    $vendor = "Desconhecido"
                }
                $ipsList += ("{0,-18} {1,-20} {2}" -f $gateway, $mac, $vendor) + "`r`n"
            }
        } catch {
            $ipsList += ("{0,-18} {1,-20} {2}" -f $gateway, "Erro", "Falha ao consultar") + "`r`n"
        }
        $progressBarIPs.Value = 1
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Verifica outros IPs com método compatível
    for ($i = 1; $i -le 254; $i++) {
        $ip = "$ipBase.$i"
        $progressBarIPs.Value = $i
        [System.Windows.Forms.Application]::DoEvents()

        if ($ip -ne $gateway) {
            # Método alternativo compatível com todas versões do PowerShell
            $ping = New-Object System.Net.NetworkInformation.Ping
            try {
                $reply = $ping.Send($ip, 1000) # Timeout de 1000ms (1 segundo)
                if ($reply.Status -eq 'Success') {
                    arp -a $ip | ForEach-Object {
                        if ($_ -match "$ip\s+([-\w]+)") {
                            $mac = $matches[1]
                            $macClean = $mac -replace '-', ':'
                            try {
                                $vendor = Invoke-RestMethod -Uri "https://api.macvendors.com/$macClean" -TimeoutSec 3
                            } catch {
                                $vendor = "Desconhecido"
                            }
                            $ipsList += ("{0,-18} {1,-20} {2}" -f $ip, $mac, $vendor) + "`r`n"
                        }
                    }
                }
            } catch {
                # Ignora erros de ping
            }
        }
    }

    $progressBarIPs.Visible = $false
    $labelIPStatus.Text = "Dispositivos encontrados:"
    $txtIPs.Text = $ipsList
}

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
    $labelIPInput.Visible = $false
    $inputIP.Visible = $false
    $btnIniciarVarredura.Visible = $false

    $btnDownload.Visible = $true
    $btnOpen.Visible = $true
    $btnSair.Visible = $true
})

$btnIPsPage.Add_Click({
    $listBox.Visible = $false
    $progressBar.Visible = $false
    $btnDownload.Visible = $false
    $btnOpen.Visible = $false
    $btnSair.Visible = $false

    $labelIPs.Visible = $true
    $txtIPs.Visible = $true
    $progressBarIPs.Visible = $true
    $labelIPStatus.Visible = $true
    $labelIPInput.Visible = $true
    $inputIP.Visible = $true
    $btnIniciarVarredura.Visible = $true
    
    $txtIPs.Text = "Digite o IP base (ex: 192.168.1) e clique em Iniciar Varredura"
    $labelIPStatus.Text = "Aguardando entrada do usuário..."
})

$btnIniciarVarredura.Add_Click({
    Listar-IPs
})

$btnDownload.Add_Click({
    if ($listBox.SelectedItem) {
        $nome = $listBox.SelectedItem
        $arquivo = $arquivos | Where-Object { $_.Nome -eq $nome }
        if ($arquivo) {
            $caminho = Join-Path $pastaDestino $arquivo.Nome
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

$form.Controls.AddRange(@($btnDownloadPage, $btnIPsPage, $btnIniciarVarredura, $btnDownload, $btnOpen, $btnSair))

# Exibir a janela
$form.Topmost = $true
$form.ShowDialog()
