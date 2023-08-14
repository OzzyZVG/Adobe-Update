$installerPath = "\\br01s-fs\PUBLIC\edge\MicrosoftEdgeSetup.exe"
$computerListPath = "C:\Temp\computer_list.txt"
$logPath = "C:\Temp\EdgeUpdateLog.txt"

# Ler a lista de computadores
$computers = Get-Content $computerListPath

foreach ($computer in $computers) {
    Write-Host "Verificando se a máquina $computer está online"
    if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
        Write-Host "Máquina online, acessando via hostname: $computer"

        # Habilitar WinRM
        Write-Host "Habilitando WinRM na máquina $computer"
        Invoke-Command -ComputerName $computer -ScriptBlock {
            Enable-PSRemoting -Force -SkipNetworkProfileCheck
        }

        # Copiar o instalador para a máquina remota
        Write-Host "Copiando o instalador para a máquina $computer"
        Copy-Item -Path $installerPath -Destination "\\$computer\C$\Temp" -ErrorAction SilentlyContinue

        # Fechar o Edge se estiver em execução
        Write-Host "Fechando o Edge na máquina $computer, se estiver em execução"
        Invoke-Command -ComputerName $computer -ScriptBlock {
            Stop-Process -Name msedge -Force -ErrorAction SilentlyContinue
        }

        # Instalar o Edge
        Write-Host "Instalando o Edge na máquina $computer"
        $installResult = Invoke-Command -ComputerName $computer -ScriptBlock {
            Start-Process "C:\Temp\MicrosoftEdgeSetup.exe" -ArgumentList "/silent /install" -Wait -PassThru
            return $LASTEXITCODE
        } -ErrorAction SilentlyContinue

        if ($installResult -ne 0) {
            Write-Host "Falha na instalação na máquina $computer com código de saída $installResult"
            continue
        }

        # Verificar se a atualização foi bem-sucedida
        Write-Host "Verificando se a atualização foi bem-sucedida na máquina $computer"
        $edgeVersion = Invoke-Command -ComputerName $computer -ScriptBlock {
            (Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe").VersionInfo.ProductVersion
        }
        if ($edgeVersion -eq "115.0.1901.203") {
            Write-Host "Atualização bem-sucedida na máquina $computer"
            Add-Content -Path $logPath -Value $computer
            $computers = $computers | Where-Object { $_ -ne $computer }
            Set-Content -Path $computerListPath -Value $computers
        } else {
            Write-Host "Falha na atualização na máquina $computer"
        }
    } else {
        Write-Host "Máquina $computer está offline"
    }
}