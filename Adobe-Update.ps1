# Caminho para o arquivo com a lista de computadores
$computerListPath = "C:\Temp\computer_list.txt"
$computerDonePath = "C:\Temp\computer_done_acrobat.txt"

# Credenciais para a sessão remota
$user = "a-borjano-1"
$password = ConvertTo-SecureString "1@2l3l4a5N" -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential ($user, $password)

# URL para download da versão mais recente do Adobe Acrobat DC
$downloadUrl = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/2300320269/AcrobatDCx64Upd2300320269.msp"

# Versão mínima necessária
$requiredVersion = [Version]"23.003.20269"

# Lê os hostnames dos computadores
$computers = Get-Content -Path $computerListPath

# Lista para armazenar os computadores não processados
$computersNotProcessed = @()

# Loop através de cada computador
foreach ($computer in $computers) {
    $processed = $false
    Write-Host "Verificando conectividade com o computador: $computer"

    # Testa a conectividade com ping
    $pingable = Test-Connection -ComputerName $computer -Count 1 -Quiet
    if ($pingable) {
        Write-Host "Conectando-se ao computador: $computer"

        # Estabelece uma sessão remota com credenciais
        $session = New-PSSession -ComputerName $computer -Credential $credentials -ErrorAction SilentlyContinue

        if ($session) {
            # Obtém a versão atual do Adobe Acrobat DC
            $currentVersion = Invoke-Command -Session $session -ScriptBlock {
                $adobePath = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
                if (Test-Path -Path $adobePath) {
                    $fileVersion = (Get-Item $adobePath).VersionInfo.ProductVersion
                    return [Version]$fileVersion
                }
                return $null
            }

            if ($currentVersion) {
                Write-Host "Versão atual do Adobe Acrobat DC no computador ${computer}: $currentVersion"

                if ($currentVersion -lt $requiredVersion) {
                    Write-Host "Atualizando Adobe Acrobat DC de versão $currentVersion para $requiredVersion"
                    Invoke-Command -Session $session -ScriptBlock {
                        $tempPath = "C:\Temp"
                        if (-not (Test-Path -Path $tempPath)) {
                            New-Item -Path $tempPath -ItemType Directory
                        }
                        $installerPath = "$tempPath\AcrobatDCx64Upd2300320269.msp"
                        Invoke-WebRequest -Uri $using:downloadUrl -OutFile $installerPath

                        # Executa o instalador diretamente
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/p $installerPath /qn" -Wait

                        # Remove o arquivo do instalador
                        Remove-Item -Path $installerPath -Force
                    }

                    # Adiciona o computador à lista de concluídos
                    Add-Content -Path $computerDonePath -Value "$computer - Adobe Acrobat DC atualizado para versão $requiredVersion"
                    $processed = $true
                } else {
                    Write-Host "Adobe Acrobat DC já está atualizado na versão $currentVersion no computador: $computer"
                    $processed = $true
                }
            } else {
                Write-Host "Adobe Acrobat DC não encontrado no computador: $computer"
            }

            # Fecha a sessão remota
            Remove-PSSession -Session $session
        } else {
            Write-Host "Não foi possível estabelecer uma sessão remota com o computador: $computer"
        }
    } else {
        Write-Host "Não foi possível conectar-se ao computador: $computer"
    }

    # Adiciona o computador à lista de não processados se necessário
    if (-not $processed) {
        $computersNotProcessed += $computer
    }
}

# Atualiza o arquivo computer_list.txt com os computadores não processados
Set-Content -Path $computerListPath -Value $computersNotProcessed

Write-Host "Processamento concluído."
