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
$form.Size = New-Object Drawing.Size(600, 450)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Título estilizado
$label = New-Object Windows.Forms.Label
$label.Text = "Selecione um arquivo para baixar:"
$label.AutoSize = $true
$label.Location = New-Object Drawing.Point(20, 20)
$label.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($label)

# Lista com borda e fonte
$listBox = New-Object Windows.Forms.ListBox
$listBox.Location = New-Object Drawing.Point(20, 60)
$listBox.Size = New-Object Drawing.Size(540, 250)
$listBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$listBox.BorderStyle = "FixedSingle"
$form.Controls.Add($listBox)

# Barra de progresso
$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Location = New-Object Drawing.Point(20, 320)
$progressBar.Size = New-Object Drawing.Size(540, 30)
$progressBar.Style = "Continuous" # Barra de progresso contínua
$form.Controls.Add($progressBar)

# Função de download
function Get-Files {
    try {
        $html = Invoke-WebRequest -Uri $url -UseBasicParsing
        $pattern = '<a[^>]*href="(?<href>https[^"]+)"[^>]*>(?<nome>[^<]+)</a>'
        $matches = [regex]::Matches($html.Content, $pattern)

        $resultado = @()
        foreach ($m in $matches) {
            $href = $m.Groups["href"].Value
            $nome = $m.Groups["nome"].Value.Trim()

            # Incluir só se for um link de download real
            if ($href -match "onedrive\.live\.com|\.exe$|\.zip$|\.pdf$|\.rar$|\.msi$|\.docx$|\.xlsx$") {
                $resultado += [PSCustomObject]@{
                    Nome = $nome
                    href = $href
                }
            }
        }
        return $resultado
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erro ao acessar a página: $_")
        return @()
    }
}

$arquivos = Get-Files
$arquivos | ForEach-Object {
    $listBox.Items.Add($_.Nome)
}

# Estilo dos botões
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

$btnDownload = Novo-Botao "⬇ Baixar Arquivo" 20 360
$btnOpen = Novo-Botao "📁 Abrir Pasta" 200 360
$btnSair = Novo-Botao "❌ Sair" 380 360

$form.Controls.AddRange(@($btnDownload, $btnOpen, $btnSair))

# Ações dos botões
$btnDownload.Add_Click({
    if ($listBox.SelectedItem) {
        $nome = $listBox.SelectedItem
        $arquivo = $arquivos | Where-Object { $_.Nome -eq $nome }
        if ($arquivo) {
            # Garantir que o nome do arquivo termine com .exe
            $extensao = [System.IO.Path]::GetExtension($arquivo.Nome)
            if ($extensao -ne ".exe") {
                $novoNome = [System.IO.Path]::ChangeExtension($arquivo.Nome, ".exe")
            } else {
                $novoNome = $arquivo.Nome
            }

            $caminho = Join-Path $pastaDestino $novoNome
            try {
                # Inicia o download e calcula o tamanho total
                $response = Invoke-WebRequest -Uri $arquivo.href -Method Head -UseBasicParsing
                $totalBytes = [int]$response.Headers["Content-Length"]
                $progressBar.Maximum = $totalBytes

                # Cria o objeto WebClient
                $client = New-Object System.Net.WebClient

                # Associa o evento de progresso de maneira correta
                $client.add_DownloadProgressChanged({
                    param($sender, $e)
                    $progressBar.Value = $e.BytesReceived
                })

                # Baixa o arquivo de forma assíncrona
                $client.DownloadFileAsync($arquivo.href, $caminho)
                
                # Espera o download terminar
                while ($client.IsBusy) {
                    Start-Sleep -Seconds 1
                }

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

# Exibe a janela
$form.Topmost = $true
$form.ShowDialog()