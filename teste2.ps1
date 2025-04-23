Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$url = "https://kl-quartz.com.br/tecnicokl"
$pastaDestino = "$env:USERPROFILE\Downloads\KL-Quartz"

if (-not (Test-Path $pastaDestino)) {
    New-Item -ItemType Directory -Path $pastaDestino -Force | Out-Null
}

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

$btnDownloadPage = Novo-Botao "Lista de Downloads" 20 20
$btnIPsPage = Novo-Botao "Listar IP's" 180 20
$btnIniciarVarredura = Novo-Botao "Iniciar Varredura" 340 20
$btnDownload = Novo-Botao "⬇ Baixar Arquivo" 20 450
$btnOpen = Novo-Botao "Abrir Pasta" 200 450
$btnSair = Novo-Botao "❌ Sair" 380 450

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

    for ($i = 1; $i -le 254; $i++) {
        $ip = "$ipBase.$i"
        $progressBarIPs.Value = $i
        [System.Windows.Forms.Application]::DoEvents()

        if ($ip -ne $gateway) {
            $ping = New-Object System.Net.NetworkInformation.Ping
            try {
                $reply = $ping.Send($ip, 1000)
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
            }
        }
    }

    $progressBarIPs.Visible = $false
    $labelIPStatus.Text = "Dispositivos encontrados:"
    $txtIPs.Text = $ipsList
}

$btnDownloadPage.Add_Click({
    $listBox.Items.Clear()
    $global:arquivos = Get-Files
    $global:arquivos | ForEach-Object { $listBox.Items.Add($_.Nome) }

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
            # 🔽 Corrige a extensão do arquivo, mesmo que o nome esteja sem
            $extensao = [System.IO.Path]::GetExtension($arquivo.href)
            $nomeComExtensao = if ($arquivo.Nome -notmatch "\.[a-z0-9]{2,4}$") {
                "$($arquivo.Nome)$extensao"
            } else {
                $arquivo.Nome
            }

            $caminho = Join-Path $pastaDestino $nomeComExtensao

            try {
                $progressBar.Style = "Marquee"
                $progressBar.MarqueeAnimationSpeed = 50
                $form.Refresh()

                $client = New-Object System.Net.WebClient
                $global:downloadComplete = $false

                $client.Add_DownloadProgressChanged({
                    param($s, $e)
                    if ($progressBar.Style -ne "Continuous") {
                        $progressBar.Style = "Continuous"
                    }
                    $progressBar.Value = $e.ProgressPercentage
                    $form.Refresh()
                })

                $client.Add_DownloadFileCompleted({
                    param($s, $e)
                    $global:downloadComplete = $true
                    if ($e.Error) {
                        [System.Windows.Forms.MessageBox]::Show("Erro no download: $($e.Error.Message)", "Erro", "OK", "Error")
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Download concluído para: $caminho", "Sucesso", "OK", "Information")
                    }
                    $client.Dispose()
                })

                $client.DownloadFileAsync([Uri]$arquivo.href, $caminho)

                while (-not $global:downloadComplete) {
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 100
                }

            } catch {
                [System.Windows.Forms.MessageBox]::Show("Erro ao baixar: $_", "Erro", "OK", "Error")
            } finally {
                $progressBar.Style = "Continuous"
                $progressBar.Value = 0
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Selecione um arquivo antes de baixar.", "Aviso", "OK", "Warning")
    }
})


$btnOpen.Add_Click({ Start-Process explorer.exe $pastaDestino })
$btnSair.Add_Click({ $form.Close() })

$form.Controls.AddRange(@($btnDownloadPage, $btnIPsPage, $btnIniciarVarredura, $btnDownload, $btnOpen, $btnSair))

$form.Topmost = $true
$form.ShowDialog()
