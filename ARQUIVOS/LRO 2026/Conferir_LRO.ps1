# --- CONFIGURAÇÃO INICIAL ---
$Hoje = Get-Date
$MesAtual = $Hoje.Month.ToString("00")
$AnoAtual = $Hoje.Year.ToString().Substring(2,2)
$HoraAtual = $Hoje.Hour

# --- MAPEAMENTO DE PASTAS ---
$MapaPastas = @{
    "01" = "1 - JANEIRO";   "02" = "2 - FEVEREIRO"; "03" = "3 - MARÇO";
    "04" = "4 - ABRIL";     "05" = "5 - MAIO";      "06" = "6 - JUNHO";
    "07" = "7 - JULHO";     "08" = "8 - AGOSTO";    "09" = "9 - SETEMBRO";
    "10" = "10 - OUTUBRO";  "11" = "11 - NOVEMBRO"; "12" = "12 - DEZEMBRO"
}

# === INÍCIO DO LOOP PRINCIPAL ===
do {
    Clear-Host
    Write-Host "=== Verificador de LRO ===" -ForegroundColor Cyan
    Write-Host "Local: Raiz | Data: $Hoje" -ForegroundColor DarkGray
    Write-Host ""

    # --- RESET DE VARIÁVEIS ---
    $ProblemasEncontrados = 0
    $RelatorioCobranca = @() 
    $ListaPendentesVerificacao = @()
    $ListaParaCriar = @() 
    $ListaAnoErrado = @()

    # --- INPUTS ---
    $InputMes = Read-Host "Digite o MÊS ou tecle Enter para usar o atual ($MesAtual)"
    if ($InputMes -eq "") { $Mes = $MesAtual } else { $Mes = $InputMes.PadLeft(2, '0') }

    $InputAno = Read-Host "Digite o ANO ou tecle Enter para usar o atual ($AnoAtual)"
    if ($InputAno -eq "") { $Ano = $AnoAtual } else { $Ano = $InputAno }
    $AnoErrado = ([int]$Ano - 1).ToString("00")

    # --- VALIDAÇÃO DA PASTA ---
    $NomePastaAlvo = $MapaPastas[$Mes]
    $CaminhoCompleto = ".\$NomePastaAlvo"

    if (-not (Test-Path $CaminhoCompleto)) {
        Write-Host "[ERRO] Pasta não encontrada: '$NomePastaAlvo'" -ForegroundColor Red
        $TentarNovamente = Read-Host "Deseja tentar outro mês? (S/N)"
        if ($TentarNovamente -eq "S" -or $TentarNovamente -eq "s") { continue } else { break }
    }

    # --- CÁLCULO DE DIAS (AGORA AUTOMÁTICO) ---
    Switch ($Mes) {
        {$_ -in '01','03','05','07','08','10','12'} { $QtdDias = 31 }
        {$_ -in '04','06','09','11'} { $QtdDias = 30 }
        '02' {
            # Converte '26' para '2026' para fazer o cálculo correto
            $AnoCompleto = 2000 + [int]$Ano
            
            if ([DateTime]::IsLeapYear($AnoCompleto)) {
                $QtdDias = 29
                Write-Host ">> Ano Bissexto detectado: Fevereiro terá 29 dias." -ForegroundColor DarkGray
            } else {
                $QtdDias = 28
                Write-Host ">> Ano Comum detectado: Fevereiro terá 28 dias." -ForegroundColor DarkGray
            }
        }
        Default { Write-Host "Mês inválido."; Start-Sleep -Seconds 2; continue }
    }
    
    # Trava de Data (Mês Atual)
    if ($Mes -eq $MesAtual -and $Ano -eq $AnoAtual) { $QtdDias = $Hoje.Day }

    Write-Host ""
    Write-Host ">> Acessando: '$NomePastaAlvo'" -ForegroundColor Yellow
    Write-Host ">> Analisando ($QtdDias dias)..." 
    Write-Host "--------------------------------------------------------"

    # --- LOOP DE VERIFICAÇÃO ---
    1..$QtdDias | ForEach-Object {
        $Dia = $_
        $DiaFormatado = "{0:D2}" -f $Dia
        $DataString = "$DiaFormatado$Mes$Ano"
        $DataStringErrada = "$DiaFormatado$Mes$AnoErrado"
        
        # Lógica de Turnos
        if ($Dia -lt $Hoje.Day -or ($Mes -ne $MesAtual)) { $TurnosParaVerificar = @(1, 2, 3) }
        elseif ($Dia -eq $Hoje.Day -and $Mes -eq $MesAtual) {
            if ($HoraAtual -lt 14) { $TurnosParaVerificar = @() }
            elseif ($HoraAtual -lt 21) { $TurnosParaVerificar = @(1) }
            else { $TurnosParaVerificar = @(1, 2) }
        }
        else { $TurnosParaVerificar = @() }

        # Checagem de Arquivos
        $TurnosParaVerificar | ForEach-Object {
            $Turno = $_
            $PadraoBusca = "${DataString}_${Turno}TURNO*"
            $ArquivosDoTurno = Get-ChildItem -Path "$CaminhoCompleto\$PadraoBusca" -ErrorAction SilentlyContinue

            $PadraoBuscaErrado = "${DataStringErrada}_${Turno}TURNO*"
            $ArquivoAnoErrado = Get-ChildItem -Path "$CaminhoCompleto\$PadraoBuscaErrado" -ErrorAction SilentlyContinue

            $TemOK = $ArquivosDoTurno | Where-Object { $_.Name -match "OK" }
            $TemFaltaLRO = $ArquivosDoTurno | Where-Object { $_.Name -match "FALTA LRO" }
            $TemFaltaAss = $ArquivosDoTurno | Where-Object { $_.Name -match "FALTA ASS" }
            $TemNovo = $ArquivosDoTurno | Where-Object { $_.Name -notmatch "OK" -and $_.Name -notmatch "FALTA" }

            if ($TemOK) { }
            elseif ($TemFaltaAss) {
                Write-Host "[!] FALTA ASSINATURA: $($TemFaltaAss.Name)" -ForegroundColor Yellow
                $RelatorioCobranca += "$($TemFaltaAss.Name)"
                $ProblemasEncontrados++
            }
            elseif ($TemNovo -and $TemFaltaLRO) {
                Write-Host "[!] LRO CONFECCIONADO (TXT AINDA PRESENTE): $($TemNovo.Name)" -ForegroundColor DarkYellow
                $ListaPendentesVerificacao += $TemNovo.FullName
                $ProblemasEncontrados++
            }
            elseif ($TemFaltaLRO) {
                Write-Host "[!] NÃO CONFECCIONADO: $($TemFaltaLRO.Name)" -ForegroundColor Magenta
                $RelatorioCobranca += "$($TemFaltaLRO.Name)"
                $ProblemasEncontrados++
            }
            elseif ($TemNovo) {
                Write-Host "[?] PENDENTE DE VERIFICAÇÃO: $($TemNovo.Name)" -ForegroundColor Cyan
                $ListaPendentesVerificacao += $TemNovo.FullName
                $ProblemasEncontrados++
            }
            elseif ($ArquivoAnoErrado) {
                 $ArqErrado = $ArquivoAnoErrado | Select-Object -First 1
                 Write-Host "[!] ANO INCORRETO ($AnoErrado): $($ArqErrado.Name)" -ForegroundColor DarkRed
                 $ListaAnoErrado += $ArqErrado
                 $ProblemasEncontrados++
            }
            else {
                Write-Host "[X] INEXISTENTE: Dia $DiaFormatado - ${Turno}º Turno" -ForegroundColor Red
                $RelatorioCobranca += "Dia $DiaFormatado - ${Turno}º Turno: NÃO ENCONTRADO"
                $ProblemasEncontrados++
                $ListaParaCriar += [PSCustomObject]@{
                    Dia = $DiaFormatado; Turno = $Turno; DataString = $DataString
                }
            }
        }
    }

    Write-Host "--------------------------------------------------------"

    # --- MENSAGEM DE SUCESSO ---
    if ($ProblemasEncontrados -eq 0) {
        Write-Host "Tudo em dia! Nenhuma pendência encontrada." -ForegroundColor Green
        Write-Host ""
    }

    # --- 1. CORREÇÃO DE ANO ---
    if ($ListaAnoErrado.Count -gt 0) {
        Write-Host "Existem" $ListaAnoErrado.Count "arquivos com o ano incorreto." -ForegroundColor DarkRed
        $CorrigirAno = Read-Host "Deseja corrigir os nomes agora? (S/N)"
        if ($CorrigirAno -eq "S" -or $CorrigirAno -eq "s") {
            foreach ($arquivo in $ListaAnoErrado) {
                Invoke-Item $arquivo.FullName
                $Confirmacao = Read-Host ">> Arquivo aberto. Renomear para ano $Ano? (S/N)"
                if ($Confirmacao -eq "S" -or $Confirmacao -eq "s") {
                    $NovoNome = $arquivo.Name.Replace($AnoErrado, $Ano)
                    try {
                        Rename-Item -Path $arquivo.FullName -NewName $NovoNome -ErrorAction Stop
                        Write-Host "   [V] Renomeado para: $NovoNome" -ForegroundColor Green
                    } catch {
                        Write-Host "   [!] Erro: Feche o arquivo e tente de novo." -ForegroundColor Red
                    }
                }
            }
        }
    }

    # --- 2. RELATÓRIO DE COBRANÇA ---
    if ($ProblemasEncontrados -gt 0 -and $RelatorioCobranca.Count -gt 0) {
        Write-Host "--- RESUMO PARA COPIAR E COBRAR ---" -ForegroundColor White -BackgroundColor DarkBlue
        foreach ($linha in $RelatorioCobranca) { Write-Host $linha }
        Write-Host "-----------------------------------"
        Write-Host ""
    }

    # --- 3. CRIAÇÃO AUTOMÁTICA DE TXT ---
    if ($ListaParaCriar.Count -gt 0) {
        Write-Host "Existem" $ListaParaCriar.Count "turnos inexistentes." -ForegroundColor Yellow
        $CriarTxt = Read-Host "Deseja criar os arquivos 'FALTA LRO ()' agora? (S/N)"
        if ($CriarTxt -eq "S" -or $CriarTxt -eq "s") {
            foreach ($item in $ListaParaCriar) {
                $NomeArquivo = "$($item.DataString)_$($item.Turno)TURNO FALTA LRO ().txt"
                $CaminhoFinal = Join-Path -Path $CaminhoCompleto -ChildPath $NomeArquivo
                New-Item -Path $CaminhoFinal -ItemType File -Value "Falta LRO" -Force | Out-Null
                Write-Host "   [V] Criado: $NomeArquivo" -ForegroundColor Green
            }
            Write-Host ""
        }
    }

    # --- 4. ABRIR ARQUIVOS PENDENTES (CARROSSEL COM RETRY) ---
    if ($ListaPendentesVerificacao.Count -gt 0) {
        $TotalPendentes = $ListaPendentesVerificacao.Count
        Write-Host "Existem $TotalPendentes livros para verificar visualmente." -ForegroundColor Cyan
        $Iniciar = Read-Host "Deseja iniciar a verificação UM POR UM agora? (S/N)"
        
        if ($Iniciar -eq "S" -or $Iniciar -eq "s") {
            $Contador = 1
            foreach ($caminhoArquivo in $ListaPendentesVerificacao) {
                $ItemArquivo = Get-Item $caminhoArquivo
                
                Write-Host ""
                Write-Host "--- VERIFICANDO ($Contador de $TotalPendentes) ---" -ForegroundColor Yellow
                Write-Host "Arquivo: $($ItemArquivo.Name)" -ForegroundColor Cyan
                
                Invoke-Item $ItemArquivo.FullName
                
                $Acao = Read-Host ">> Documento assinado? Adicionar 'OK'? (S = Sim / N = Não / Enter = Pular)"
                
                if ($Acao -eq "S" -or $Acao -eq "s") {
                    $NovoNome = $ItemArquivo.BaseName + " OK" + $ItemArquivo.Extension
                    
                    $SucessoRenomear = $false
                    do {
                        try {
                            Rename-Item -Path $ItemArquivo.FullName -NewName $NovoNome -ErrorAction Stop
                            Write-Host "   [V] Validado e renomeado para: $NovoNome" -ForegroundColor Green
                            $SucessoRenomear = $true
                        }
                        catch {
                            Write-Host ""
                            Write-Host "   [!] ERRO: O arquivo está aberto e não pode ser renomeado." -ForegroundColor Red
                            Write-Host "       Por favor, feche a janela do PDF." -ForegroundColor Yellow
                            $Retry = Read-Host "       Pressione Enter para tentar novamente (ou 'N' para desistir)"
                            if ($Retry -eq "N" -or $Retry -eq "n") { break }
                        }
                    } until ($SucessoRenomear)
                    
                } else {
                    Write-Host "   [-] Mantido como está." -ForegroundColor Gray
                }
                
                $Contador++
            }
            Write-Host "--- Fim da verificação visual ---" -ForegroundColor Yellow
        }
    }

    $Continuar = Read-Host "Deseja verificar outro mês? (S = Sim / Enter = Sair)"
    
} until ($Continuar -ne "S" -and $Continuar -ne "s")