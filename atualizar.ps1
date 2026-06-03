# ICC - Gerador de Relatorio Integrado (Inadimplencia + Fluxo de Caixa)
# Executa a cada 7 minutos via Task Scheduler

$REPO        = "C:\Users\MICRO1\Documents\icc-inadimplencia"
$HTML        = "$REPO\index.html"
$SNAPSHOT    = "$REPO\snapshot.json"
$HOJE        = Get-Date
$OMIE_CR     = "https://app.omie.com.br/api/v1/financas/contareceber/"
$OMIE_CL     = "https://app.omie.com.br/api/v1/geral/clientes/"
$META_VALOR  = 200000
$META_INICIO = 1031143.11  # valor no inicio do tracking

$EMPRESAS = @(
    @{ nome="Instituto"; cor="#2196F3"; grad="linear-gradient(135deg,#1e3a5f,#2196F3)"; app_key="3946880386449"; app_secret="0c15f825cded97455749c7d6b7558f1e" },
    @{ nome="Telecom";   cor="#00BCD4"; grad="linear-gradient(135deg,#004d5f,#00BCD4)"; app_key="4472437527558"; app_secret="eb030b4871537b1d984ff4078a469f75" },
    @{ nome="Medical";   emoji=""; cor="#4CAF50"; grad="linear-gradient(135deg,#1b4d1f,#4CAF50)"; app_key="7069173264153"; app_secret="9632f5b931f568b6b09accbf25f47496" }
)

# ─── API helpers ────────────────────────────────────────────────────────────

function Get-MapaClientes($app_key, $app_secret) {
    $mapa = @{}; $pag = 1
    do {
        $b = @{ call="ListarClientesResumido"; app_key=$app_key; app_secret=$app_secret
                param=@(@{ pagina=$pag; registros_por_pagina=500; apenas_importado_api="N" }) } | ConvertTo-Json -Depth 5
        try { $r = Invoke-RestMethod -Uri $OMIE_CL -Method Post -Body $b -ContentType "application/json" -TimeoutSec 30 }
        catch { break }
        foreach ($cl in $r.clientes_cadastro_resumido) {
            $n = if ($cl.nome_fantasia -and $cl.nome_fantasia.Trim()) { $cl.nome_fantasia } else { $cl.razao_social }
            $mapa[[string]$cl.codigo_cliente] = $n
        }
        $tot = [int]$r.total_de_paginas; $pag++
    } while ($pag -le $tot)
    return $mapa
}

function Get-Contas($app_key, $app_secret) {
    $all = @(); $pag = 1
    do {
        $b = @{ call="ListarContasReceber"; app_key=$app_key; app_secret=$app_secret
                param=@(@{ pagina=$pag; registros_por_pagina=500; apenas_importado_api="N"; filtrar_por_status="ATRASADO" }) } | ConvertTo-Json -Depth 5
        try { $r = Invoke-RestMethod -Uri $OMIE_CR -Method Post -Body $b -ContentType "application/json" -TimeoutSec 30 }
        catch { break }
        if ($r.faultstring) { break }
        if ($r.conta_receber_cadastro) { $all += $r.conta_receber_cadastro }
        $tot = [int]$r.total_de_paginas; $pag++
    } while ($pag -le $tot)
    return $all
}

function Get-ContasMes($app_key, $app_secret, $mesAno) {
    # Busca titulos com vencimento no mes/ano informado (formato MM/AAAA)
    $all = @(); $pag = 1
    $partes = $mesAno -split "/"
    $m = $partes[0]; $a = $partes[1]
    $d1 = "01/$m/$a"
    $ultimo = [datetime]::DaysInMonth([int]$a,[int]$m)
    $d2 = "$ultimo/$m/$a"
    do {
        $b = @{ call="ListarContasReceber"; app_key=$app_key; app_secret=$app_secret
                param=@(@{ pagina=$pag; registros_por_pagina=500; apenas_importado_api="N"
                           filtrar_por_status="ATRASADO"
                           filtrar_por_data_de=$d1; filtrar_por_data_ate=$d2
                           filtrar_por_tipo_data="V" }) } | ConvertTo-Json -Depth 5
        try { $r = Invoke-RestMethod -Uri $OMIE_CR -Method Post -Body $b -ContentType "application/json" -TimeoutSec 30 }
        catch { break }
        if ($r.faultstring) { break }
        if ($r.conta_receber_cadastro) { $all += $r.conta_receber_cadastro }
        $tot = [int]$r.total_de_paginas; $pag++
    } while ($pag -le $tot)
    return $all
}

function Get-DiasAtraso($dataVenc) {
    try { return [int]($HOJE - [datetime]::ParseExact($dataVenc,"dd/MM/yyyy",$null)).TotalDays }
    catch { return 0 }
}

function Fmt-BRL($v) {
    $ptBR = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")
    return "R$ " + ([double]$v).ToString("N2", $ptBR)
}

function Esc-Html($s) {
    return $s -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" -replace '"',"&quot;"
}

function Get-Grupo($dias) {
    if ($dias -le 30)  { return @{ g="caixa";      go=0; label="1-30 dias" } }
    if ($dias -le 60)  { return @{ g="transicao";  go=1; label="31-60 dias" } }
    if ($dias -le 90)  { return @{ g="negociacao"; go=2; label="61-90 dias" } }
    return @{ g="negociacao"; go=2; label="90+ dias" }
    # Recuperacao judicial vai para regua
}

function Get-AgingClass($dias) {
    if ($dias -le 30) { return "green" }
    if ($dias -le 60) { return "yellow" }
    if ($dias -le 90) { return "orange" }
    return "red"
}

# ─── Coleta principal ──────────────────────────────────────────────────────

Write-Host "[$(Get-Date -f 'HH:mm:ss')] Iniciando coleta..."

$dadosEmp = @()
$todosClientes = @()  # para priorizacao

foreach ($emp in $EMPRESAS) {
    Write-Host "[$(Get-Date -f 'HH:mm:ss')] $($emp.nome)..."
    $mapa   = Get-MapaClientes $emp.app_key $emp.app_secret
    $contas = Get-Contas $emp.app_key $emp.app_secret
    Write-Host "  $($contas.Count) titulos"

    $clMap = @{}
    $total=0.0; $f1=0.0; $f2=0.0; $f3=0.0; $f4=0.0
    $cnt1=0; $cnt2=0; $cnt3=0; $cnt4=0

    # Inadimplencia corrente (mes atual)
    $mesAtual = $HOJE.ToString("MM/yyyy")
    $currVal=0.0; $currCnt=0

    # Historico mensal (agrupar por mes de vencimento)
    $histMes = @{}

    foreach ($c in $contas) {
        $val  = [double]$c.valor_documento
        $dias = Get-DiasAtraso $c.data_vencimento
        $cod  = [string]$c.codigo_cliente_fornecedor
        $nome = if ($mapa.ContainsKey($cod)) { $mapa[$cod] } else { "Cod $cod" }

        $total += $val
        if     ($dias -le 30) { $f1 += $val; $cnt1++ }
        elseif ($dias -le 60) { $f2 += $val; $cnt2++ }
        elseif ($dias -le 90) { $f3 += $val; $cnt3++ }
        else                  { $f4 += $val; $cnt4++ }

        # Agrupar mes de vencimento
        try {
            $vencDate = [datetime]::ParseExact($c.data_vencimento,"dd/MM/yyyy",$null)
            $mk = $vencDate.ToString("MM/yyyy")
            if (-not $histMes.ContainsKey($mk)) { $histMes[$mk] = @{ val=0.0; cnt=0 } }
            $histMes[$mk].val += $val
            $histMes[$mk].cnt += 1
        } catch {}

        # Corrente
        if ($c.data_vencimento -and $c.data_vencimento.EndsWith("/" + $HOJE.Year)) {
            $mvenc = $c.data_vencimento.Substring(3,2) + "/" + $c.data_vencimento.Substring(6,4)
            if ($mvenc -eq $mesAtual) { $currVal += $val; $currCnt++ }
        }

        # Acumular clientes
        if (-not $clMap.ContainsKey($nome)) {
            $clMap[$nome] = @{ total=0.0; titulos=0; max_dias=0; nome=$nome; empresa="ICC $($emp.nome)" }
        }
        $clMap[$nome].total   += $val
        $clMap[$nome].titulos += 1
        if ($dias -gt $clMap[$nome].max_dias) { $clMap[$nome].max_dias = $dias }
    }

    # Top 30 clientes por valor
    $topCl = $clMap.GetEnumerator() | Sort-Object { $_.Value.total } -Descending | Select-Object -First 30

    # Para priorizacao global
    foreach ($cl in $clMap.GetEnumerator()) {
        $g = Get-Grupo $cl.Value.max_dias
        $isRJ = $cl.Key -match "RECUPER"
        if ($isRJ) { $g.g = "regua"; $g.go = 3 }
        $todosClientes += @{
            n  = $cl.Key
            c  = "ICC $($emp.nome)"
            v  = $cl.Value.total
            d  = $cl.Value.max_dias
            a  = $g.label
            g  = $g.g
            go = $g.go
            p  = 0
            cnt= $cl.Value.titulos
        }
    }

    $dadosEmp += @{
        nome=$emp.nome; cor=$emp.cor; grad=$emp.grad
        total=$total; titulos=$contas.Count; nclientes=$clMap.Count
        f1=$f1; f2=$f2; f3=$f3; f4=$f4
        cnt1=$cnt1; cnt2=$cnt2; cnt3=$cnt3; cnt4=$cnt4
        currVal=$currVal; currCnt=$currCnt
        topCl=$topCl; histMes=$histMes
    }
    Write-Host "  Total: $(Fmt-BRL $total)"
}

# ─── Totais consolidados ───────────────────────────────────────────────────

$totalGeral    = 0.0; $titulosGeral = 0; $clientesGeral = 0
$f4Geral = 0.0; $currValGeral = 0.0; $currCntGeral = 0
foreach ($e in $dadosEmp) {
    $totalGeral    += $e.total
    $titulosGeral  += $e.titulos
    $clientesGeral += $e.nclientes
    $f4Geral       += $e.f4
    $currValGeral  += $e.currVal
    $currCntGeral  += $e.currCnt
}
$casos90 = 0
foreach ($__e in $dadosEmp) { $casos90 += $__e.cnt4 }
$dataStr  = $HOJE.ToString("dd/MM/yyyy, HH:mm:ss")
$dataCurta = $HOJE.ToString("dd/MM/yyyy")
$mesNome  = (Get-Culture).DateTimeFormat.GetMonthName($HOJE.Month).Substring(0,3)
$mesNome  = $mesNome.Substring(0,1).ToUpper() + $mesNome.Substring(1)
$mesLabel = "$mesNome/$($HOJE.Year)"

# ─── Snapshot para trends (DIA/SEM/MES/TRI) ───────────────────────────────

$snap = $null
if (Test-Path $SNAPSHOT) {
    try { $snap = Get-Content $SNAPSHOT -Raw | ConvertFrom-Json } catch {}
}

function Fmt-Trend($curr, $prev) {
    if ($null -eq $prev -or $prev -eq 0) { return @{ val="&mdash;"; cls="flat" } }
    $pct = [math]::Round(($curr - $prev) / $prev * 100, 1)
    if ($pct -lt 0) { return @{ val="&darr; $pct%"; cls="good" } }
    if ($pct -gt 0) { return @{ val="&uarr; +$pct%"; cls="bad" } }
    return @{ val="0%"; cls="flat" }
}

$_sDia   = if ($snap) { $snap.dia_total } else { $null }
$_sSem   = if ($snap) { $snap.sem_total } else { $null }
$_sMes   = if ($snap) { $snap.mes_total } else { $null }
$_sTri   = if ($snap) { $snap.tri_total } else { $null }
$_sDiaC  = if ($snap) { $snap.dia_cli   } else { $null }
$_sSemC  = if ($snap) { $snap.sem_cli   } else { $null }
$_sDiaT  = if ($snap) { $snap.dia_tit   } else { $null }
$_sSemT  = if ($snap) { $snap.sem_tit   } else { $null }
$_sDia90 = if ($snap) { $snap.dia_90    } else { $null }
$_sSem90 = if ($snap) { $snap.sem_90    } else { $null }
$tDia   = Fmt-Trend $totalGeral    $_sDia
$tSem   = Fmt-Trend $totalGeral    $_sSem
$tMes   = Fmt-Trend $totalGeral    $_sMes
$tTri   = Fmt-Trend $totalGeral    $_sTri
$tDiaC  = Fmt-Trend $clientesGeral $_sDiaC
$tSemC  = Fmt-Trend $clientesGeral $_sSemC
$tDiaT  = Fmt-Trend $titulosGeral  $_sDiaT
$tSemT  = Fmt-Trend $titulosGeral  $_sSemT
$tDia90 = Fmt-Trend $casos90       $_sDia90
$tSem90 = Fmt-Trend $casos90       $_sSem90

# Salvar snapshot atual
$novoSnap = @{
    dia_total=$totalGeral; dia_cli=$clientesGeral; dia_tit=$titulosGeral; dia_90=$casos90
    sem_total=if ($snap -and $snap.sem_total) { $snap.sem_total } else { $null }
    mes_total=if ($snap -and $snap.mes_total) { $snap.mes_total } else { $null }
    tri_total=if ($snap -and $snap.tri_total) { $snap.tri_total } else { $null }
    sem_cli=if ($snap -and $snap.sem_cli) { $snap.sem_cli } else { $null }
    sem_tit=if ($snap -and $snap.sem_tit) { $snap.sem_tit } else { $null }
    sem_90=if ($snap -and $snap.sem_90) { $snap.sem_90 } else { $null }
    ts_dia=$HOJE.ToString("yyyy-MM-dd HH:mm")
    ts_sem=if ($snap -and $snap.ts_sem) { $snap.ts_sem } else { $null }
    ts_mes=if ($snap -and $snap.ts_mes) { $snap.ts_mes } else { $null }
    ts_tri=if ($snap -and $snap.ts_tri) { $snap.ts_tri } else { $null }
}

# Atualizar referencia semanal se passou 7 dias
if ($snap -and $snap.ts_sem) {
    try {
        $tsSem = [datetime]::ParseExact($snap.ts_sem,"yyyy-MM-dd HH:mm",$null)
        if (($HOJE - $tsSem).TotalDays -ge 7) {
            $novoSnap.sem_total=$snap.dia_total; $novoSnap.sem_cli=$snap.dia_cli
            $novoSnap.sem_tit=$snap.dia_tit; $novoSnap.sem_90=$snap.dia_90
            $novoSnap.ts_sem=$HOJE.ToString("yyyy-MM-dd HH:mm")
        }
    } catch {}
} else { $novoSnap.ts_sem=$HOJE.ToString("yyyy-MM-dd HH:mm") }

# Mensal
if ($snap -and $snap.ts_mes) {
    try {
        $tsMes = [datetime]::ParseExact($snap.ts_mes,"yyyy-MM-dd HH:mm",$null)
        if (($HOJE - $tsMes).TotalDays -ge 30) {
            $novoSnap.mes_total=$snap.dia_total
            $novoSnap.ts_mes=$HOJE.ToString("yyyy-MM-dd HH:mm")
        }
    } catch {}
} else { $novoSnap.ts_mes=$HOJE.ToString("yyyy-MM-dd HH:mm") }

# Trimestral
if ($snap -and $snap.ts_tri) {
    try {
        $tsTri = [datetime]::ParseExact($snap.ts_tri,"yyyy-MM-dd HH:mm",$null)
        if (($HOJE - $tsTri).TotalDays -ge 90) {
            $novoSnap.tri_total=$snap.dia_total
            $novoSnap.ts_tri=$HOJE.ToString("yyyy-MM-dd HH:mm")
        }
    } catch {}
} else { $novoSnap.ts_tri=$HOJE.ToString("yyyy-MM-dd HH:mm") }

$novoSnap | ConvertTo-Json | Set-Content $SNAPSHOT -Encoding UTF8

# ─── Pareto 80/20 nos clientes globais ────────────────────────────────────

$todosOrdenados = $todosClientes | Sort-Object { $_.v } -Descending
$acum = 0.0
for ($i=0; $i -lt $todosOrdenados.Count; $i++) {
    $acum += $todosOrdenados[$i].v
    if ($acum / $totalGeral -le 0.80) { $todosOrdenados[$i].p = 1 }
}

# ─── Historico mensal global (unir de todas as empresas) ──────────────────

$histGlobal = @{}
foreach ($emp in $dadosEmp) {
    foreach ($mk in $emp.histMes.Keys) {
        if (-not $histGlobal.ContainsKey($mk)) {
            $histGlobal[$mk] = @{ Instituto=@{val=0.0;cnt=0}; Telecom=@{val=0.0;cnt=0}; Medical=@{val=0.0;cnt=0} }
        }
        $histGlobal[$mk][$emp.nome].val += $emp.histMes[$mk].val
        $histGlobal[$mk][$emp.nome].cnt += $emp.histMes[$mk].cnt
    }
}

$mesesOrdenados = $histGlobal.Keys | Sort-Object {
    $p = $_ -split "/"
    [int]$p[1] * 100 + [int]$p[0]
} -Descending

function Fmt-MesNome($mk) {
    $p = $mk -split "/"; $m = [int]$p[0]; $y = $p[1]
    $nomes = @("","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez")
    return "$($nomes[$m])/$y"
}

# ─── Meta gauge ───────────────────────────────────────────────────────────

$range    = $META_INICIO - $META_VALOR
$reduzido = $META_INICIO - $totalGeral
$pctMeta  = if ($range -gt 0) { [math]::Round($reduzido / $range * 100, 1) } else { 0 }
$knobLeft = [math]::Max(0, [math]::Min(100, $pctMeta))
$dimWidth = [math]::Round(100 - $knobLeft, 1)

# dias restantes ate dez/2026
$metaFim  = [datetime]"2026-12-31"
$diasMeta = [int]($metaFim - $HOJE).TotalDays
$recuperarTotal = [math]::Round($totalGeral - $META_VALOR, 2)
$recDia   = if ($diasMeta -gt 0) { [math]::Round($recuperarTotal / $diasMeta, 2) } else { 0 }

$metaStatus = if ($pctMeta -ge 20) { @{ txt="No ritmo"; cor="#2e7d32"; bg="#2e7d3212"; brd="#2e7d3240"; ico="&#x2705;" } }
              elseif ($pctMeta -ge 5) { @{ txt="Atencao"; cor="#e65100"; bg="#e6510012"; brd="#e6510040"; ico="&#x26A0;&#xFE0F;" } }
              else { @{ txt="Critico"; cor="#c62828"; bg="#c6282812"; brd="#c6282840"; ico="&#x1F6A8;" } }

# ─── Resumo priorizacao ───────────────────────────────────────────────────

$gcaixa      = @($todosOrdenados | Where-Object { $_.g -eq "caixa" })
$gtransicao  = @($todosOrdenados | Where-Object { $_.g -eq "transicao" })
$gnegociacao = @($todosOrdenados | Where-Object { $_.g -eq "negociacao" })
$gregua      = @($todosOrdenados | Where-Object { $_.g -eq "regua" })

$vcaixa=0.0; foreach ($__x in $gcaixa)      { $vcaixa      += $__x.v }
$vtransicao=0.0; foreach ($__x in $gtransicao)  { $vtransicao  += $__x.v }
$vnegociacao=0.0; foreach ($__x in $gnegociacao) { $vnegociacao += $__x.v }
$vregua=0.0; foreach ($__x in $gregua)      { $vregua      += $__x.v }

# ─── Fetch Fluxo de Caixa HTML ────────────────────────────────────────────

Write-Host "[$(Get-Date -f 'HH:mm:ss')] Buscando fluxo de caixa..."

$fluxoUrl = "https://raw.githubusercontent.com/exactus-data-ia/icc-fluxo-de-caixa/main/index.html"
$fluxoBanco      = "Dados indisponiveis"
$fluxoAReceber   = "Dados indisponiveis"
$fluxoAPagar     = "Dados indisponiveis"
$fluxoProjetado  = "Dados indisponiveis"
$fluxoLastUpdate = "indisponivel"
$fluxoTabsHtml   = "<div style='padding:40px;text-align:center;color:#aaa'>Dados de fluxo de caixa indisponiveis no momento.</div>"
$fluxoScript     = "/* fluxo script indisponivel */"
$fluxoProjetadoPositivo = $true

try {
    $fluxoResp = Invoke-WebRequest -Uri $fluxoUrl -UseBasicParsing -TimeoutSec 30
    $fluxoRaw  = $fluxoResp.Content

    # Extrair KPIs do header consolidado
    # Formato: Banco: R$ X · A Receber: R$ X · A Pagar: R$ X · Saldo Projetado: R$ X
    if ($fluxoRaw -match 'Banco:\s*(R\$\s*[\d\.,]+)') { $fluxoBanco = $Matches[1].Trim() }
    if ($fluxoRaw -match 'A Receber:\s*(R\$\s*[\d\.,]+)') { $fluxoAReceber = $Matches[1].Trim() }
    if ($fluxoRaw -match 'A Pagar:\s*(R\$\s*[\d\.,]+)') { $fluxoAPagar = $Matches[1].Trim() }
    if ($fluxoRaw -match 'Saldo Projetado:\s*(R\$\s*[-\d\.,]+)') { $fluxoProjetado = $Matches[1].Trim() }

    # Parse sinal do saldo projetado
    $projetadoNum = $fluxoProjetado -replace 'R\$\s*','' -replace '\.','' -replace ',','.' -replace '\s',''
    try {
        $projetadoVal = [double]$projetadoNum
        $fluxoProjetadoPositivo = ($projetadoVal -ge 0)
    } catch { $fluxoProjetadoPositivo = $true }

    # Extrair timestamp de atualizacao
    if ($fluxoRaw -match '<div class="header-right"[^>]*>([\s\S]*?)</div>') {
        $hrContent = $Matches[1] -replace '<[^>]+>','' -replace '\s+',' '
        if ($hrContent -match '(\d{2}/\d{2}/\d{4}[\s,]+\d{2}:\d{2}(:\d{2})?)') {
            $fluxoLastUpdate = $Matches[1].Trim()
        }
    }

    # Extrair tab content: do primeiro <div id="consolidado" ate (nao incluindo) o ultimo <script>
    $lastScriptIdx = $fluxoRaw.LastIndexOf('<script')
    if ($lastScriptIdx -gt 0) {
        # Tab HTML: from first <div id="consolidado" up to last <script>
        $consolidadoIdx = $fluxoRaw.IndexOf('<div id="consolidado"')
        if ($consolidadoIdx -lt 0) { $consolidadoIdx = $fluxoRaw.IndexOf("<div id='consolidado'") }
        if ($consolidadoIdx -ge 0 -and $consolidadoIdx -lt $lastScriptIdx) {
            $fluxoTabsHtml = $fluxoRaw.Substring($consolidadoIdx, $lastScriptIdx - $consolidadoIdx).TrimEnd()
        }

        # Script content: content of the last <script> block
        $lastScriptEnd = $fluxoRaw.IndexOf('>', $lastScriptIdx) + 1
        $lastScriptClose = $fluxoRaw.LastIndexOf('</script>')
        if ($lastScriptEnd -gt 0 -and $lastScriptClose -gt $lastScriptEnd) {
            $fluxoScript = $fluxoRaw.Substring($lastScriptEnd, $lastScriptClose - $lastScriptEnd).Trim()
        }
    }

    Write-Host "  Fluxo OK - Banco: $fluxoBanco"
} catch {
    Write-Host "  AVISO: Falha ao buscar fluxo de caixa: $_"
}

# ─── Transform fluxo HTML to avoid conflicts ──────────────────────────────

$tabNames = @("consolidado","antecipacao","instituto","telecom","medical")

# Transform tab HTML
foreach ($t in $tabNames) {
    $fluxoTabsHtml = $fluxoTabsHtml.Replace("id=""$t""", "id=""flx_$t""")
    $fluxoTabsHtml = $fluxoTabsHtml.Replace("showTab('$t')", "showFluxoTab('flx_$t')")
    $fluxoTabsHtml = $fluxoTabsHtml.Replace("data-tab=""$t""", "data-tab=""flx_$t""")
    $fluxoTabsHtml = $fluxoTabsHtml.Replace("id=""chart_$t""", "id=""flx_chart_$t""")
    $fluxoTabsHtml = $fluxoTabsHtml.Replace("id=""chart_pag_$t""", "id=""flx_chart_pag_$t""")
}
$fluxoTabsHtml = $fluxoTabsHtml.Replace('class="tab-content', 'class="tab-content flx-tab')
$fluxoTabsHtml = $fluxoTabsHtml.Replace('class="tab-btn', 'class="tab-btn flx-tab-btn')

# Transform fluxo script
$fluxoScript = $fluxoScript.Replace("function showTab(", "function showFluxoTab(")
$fluxoScript = $fluxoScript.Replace("querySelectorAll('.tab-content')", "querySelectorAll('.flx-tab')")
$fluxoScript = $fluxoScript.Replace("querySelectorAll('.tab-btn')", "querySelectorAll('.flx-tab-btn')")
$fluxoScript = $fluxoScript.Replace("window.addEventListener('load'", "window.addEventListener('load_disabled'")

foreach ($t in $tabNames) {
    $fluxoScript = $fluxoScript.Replace("buildChart('$t')", "buildChart('flx_$t')")
    $fluxoScript = $fluxoScript.Replace("buildPagChart('$t')", "buildPagChart('flx_$t')")
    $fluxoScript = $fluxoScript.Replace("setView('$t'", "setView('flx_$t'")
    $fluxoScript = $fluxoScript.Replace("setPagView('$t'", "setPagView('flx_$t'")
    $fluxoScript = $fluxoScript.Replace("CHARTS['$t']", "CHARTS['flx_$t']")
    $fluxoScript = $fluxoScript.Replace("PAG_CHARTS['$t']", "PAG_CHARTS['flx_$t']")
    $fluxoScript = $fluxoScript.Replace("getElementById('chart_$t')", "getElementById('flx_chart_$t')")
    $fluxoScript = $fluxoScript.Replace("getElementById('chart_pag_$t')", "getElementById('flx_chart_pag_$t')")
}

# ─── Gerador de HTML ──────────────────────────────────────────────────────

$CSS = @'
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f0f2f5; color: #222; }
.header { background: linear-gradient(135deg, #0d1b2a, #1e3a5f); color: white; padding: 18px 30px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 12px rgba(0,0,0,0.3); }
.header-left h1 { font-size: 20px; font-weight: 700; }
.header-left p { font-size: 12px; opacity: 0.65; margin-top: 3px; }
.header-right { text-align: right; font-size: 12px; opacity: 0.75; line-height: 1.6; }
.main-nav{background:white;border-bottom:3px solid #e8e8e8;padding:0 20px;display:flex;gap:4px;box-shadow:0 2px 8px rgba(0,0,0,0.08);}
.main-tab-btn{padding:16px 32px;border:none;background:none;cursor:pointer;font-size:15px;font-weight:600;color:#888;border-bottom:4px solid transparent;margin-bottom:-3px;transition:all 0.2s;white-space:nowrap;}
.main-tab-btn:hover{color:#1e3a5f;background:#f8f9fa;}
.main-tab-btn.active{color:#1e3a5f;border-bottom-color:#1e3a5f;font-weight:800;background:#f0f4ff;}
.main-section{display:none;}
.main-section.active{display:block;}
.tabs { background: white; border-bottom: 2px solid #e8e8e8; padding: 0 20px; display: flex; gap: 2px; overflow-x: auto; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
.tab-btn { padding: 14px 24px; border: none; background: none; cursor: pointer; font-size: 13.5px; font-weight: 500; color: #666; border-bottom: 3px solid transparent; margin-bottom: -2px; transition: all 0.2s; white-space: nowrap; }
.tab-btn:hover { color: #1e3a5f; }
.tab-btn.active { color: #1e3a5f; border-bottom-color: #1e3a5f; font-weight: 700; }
.tab-content { display: none; padding: 24px; max-width: 1440px; margin: 0 auto; }
.tab-content.active { display: block; }
.company-header { color: white; padding: 22px 28px; border-radius: 12px; margin-bottom: 22px; box-shadow: 0 4px 16px rgba(0,0,0,0.18); }
.company-header h2 { font-size: 24px; margin-bottom: 4px; }
.company-header p { opacity: 0.8; font-size: 13px; }
.cards-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 14px; margin-bottom: 22px; }
.card { background: white; border-radius: 10px; padding: 20px 18px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.07); border-top: 4px solid #2196F3; transition: transform 0.15s; }
.card:hover { transform: translateY(-2px); }
.card.red { border-top-color: #e53935; }
.card.orange { border-top-color: #fb8c00; }
.card.green { border-top-color: #43a047; }
.card-value { font-size: 26px; font-weight: 800; color: #1e3a5f; margin-bottom: 5px; line-height: 1.1; }
.card.red .card-value { color: #c62828; }
.card.orange .card-value { color: #e65100; }
.card.green .card-value { color: #2e7d32; }
.card-value.green { color: #2e7d32; }
.card-value.red { color: #c62828; }
.card-label { font-size: 12px; color: #888; font-weight: 500; }
.card-trends { display: grid; grid-template-columns: repeat(4,1fr); gap: 2px; margin-top: 10px; padding-top: 8px; border-top: 1px solid #f0f0f0; }
.trend-item { display: flex; flex-direction: column; align-items: center; gap: 2px; }
.trend-period { font-size: 9px; color: #bbb; font-weight: 600; text-transform: uppercase; letter-spacing: 0.3px; }
.trend-val { font-size: 11px; font-weight: 800; }
.trend-val.bad  { color: #e53935; }
.trend-val.good { color: #43a047; }
.trend-val.flat { color: #9e9e9e; }
.section { background: white; border-radius: 10px; padding: 20px 22px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
.section h3 { font-size: 15px; font-weight: 700; margin-bottom: 16px; color: #1e3a5f; padding-bottom: 10px; border-bottom: 2px solid #f0f2f5; }
.aging-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }
.aging-box { border-radius: 10px; padding: 16px 12px; text-align: center; border: 1px solid; }
.aging-box.green  { background: #f1f8e9; border-color: #aed581; }
.aging-box.yellow { background: #fffde7; border-color: #ffe082; }
.aging-box.orange { background: #fff3e0; border-color: #ffcc80; }
.aging-box.red    { background: #ffebee; border-color: #ef9a9a; }
.aging-count { font-size: 36px; font-weight: 800; line-height: 1; }
.aging-box.green  .aging-count { color: #33691e; }
.aging-box.yellow .aging-count { color: #f57f17; }
.aging-box.orange .aging-count { color: #bf360c; }
.aging-box.red    .aging-count { color: #b71c1c; }
.aging-label { font-size: 12px; font-weight: 700; margin: 6px 0 4px; color: #555; }
.aging-value { font-size: 14px; font-weight: 800; margin-top: 8px; padding-top: 8px; border-top: 1px solid rgba(0,0,0,0.09); }
.aging-box.green  .aging-value { color: #33691e; }
.aging-box.yellow .aging-value { color: #f57f17; }
.aging-box.orange .aging-value { color: #bf360c; }
.aging-box.red    .aging-value { color: #b71c1c; }
.insight-card { display: flex; gap: 14px; align-items: flex-start; padding: 14px 16px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid; }
.insight-card.danger  { background: #fff5f5; border-color: #e53935; }
.insight-card.warning { background: #fffdf0; border-color: #fdd835; }
.insight-card.info    { background: #f0f7ff; border-color: #1e88e5; }
.insight-icon  { font-size: 22px; flex-shrink: 0; line-height: 1; margin-top: 1px; }
.insight-title { font-weight: 700; font-size: 13px; margin-bottom: 4px; }
.insight-text  { font-size: 12.5px; color: #555; line-height: 1.55; }
.action-group  { border: 1px solid #e8e8e8; border-radius: 8px; overflow: hidden; margin-bottom: 14px; }
.action-header { padding: 11px 16px; font-size: 13px; font-weight: 700; border-bottom: 1px solid #e8e8e8; }
.action-header.red    { background: #ffebee; color: #b71c1c; border-bottom-color: #ffcdd2; }
.action-header.orange { background: #fff3e0; color: #bf360c; border-bottom-color: #ffe0b2; }
.action-header.blue   { background: #e3f2fd; color: #0d47a1; border-bottom-color: #bbdefb; }
.action-item { padding: 9px 16px; font-size: 12.5px; border-bottom: 1px solid #fafafa; }
.action-item:last-child { border-bottom: none; }
.action-item.red    { background: #fff8f8; }
.action-item.orange { background: #fffaf5; }
.action-item.blue   { background: #f8fbff; }
.no-items { padding: 12px 16px; font-size: 12.5px; color: #aaa; font-style: italic; }
.table-container { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
thead th { background: #f5f7fa; padding: 10px 12px; text-align: left; font-weight: 700; color: #444; border-bottom: 2px solid #e0e4ea; white-space: nowrap; }
tbody td { padding: 9px 12px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }
tbody tr:hover td { background: #f9fafb; }
.badge { display: inline-block; padding: 3px 9px; border-radius: 20px; font-size: 11px; font-weight: 700; white-space: nowrap; }
.badge.red    { background: #ffebee; color: #c62828; }
.badge.orange { background: #fff3e0; color: #e65100; }
.badge.yellow { background: #fff8e1; color: #f57f17; }
.badge.green  { background: #e8f5e9; color: #2e7d32; }
.badge.blue   { background: #e3f2fd; color: #1565c0; }
.badge.gray   { background: #f5f5f5; color: #616161; border: 1px solid #e0e0e0; }
.company-breakdown { display: flex; flex-direction: column; gap: 10px; }
.company-row { display: flex; align-items: center; gap: 14px; padding: 14px 16px; background: #fafbfc; border-radius: 8px; border: 1px solid #ebebeb; }
.company-dot { width: 13px; height: 13px; border-radius: 50%; flex-shrink: 0; }
.company-info { flex: 1; }
.company-info strong { display: block; font-size: 14px; }
.company-info span   { font-size: 12px; color: #999; }
.pct-bar-wrap { flex: 2; }
.pct-bar-bg { background: #eee; border-radius: 4px; height: 6px; }
.pct-bar    { height: 6px; border-radius: 4px; }
.company-value { text-align: right; }
.company-value strong { display: block; font-size: 16px; color: #1e3a5f; font-weight: 800; }
.company-value span   { font-size: 11px; color: #aaa; }
.curr-card { background: white; border-radius: 10px; padding: 20px 22px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); border-top: 4px solid #7b1fa2; }
.curr-card-title { font-size: 14px; font-weight: 800; color: #1e3a5f; margin-bottom: 14px; }
.curr-rows { display: flex; flex-direction: column; gap: 0; }
.curr-company-row { display: flex; align-items: center; justify-content: space-between; padding: 11px 0; border-bottom: 1px solid #f5f5f5; }
.curr-company-row:last-child { border-bottom: none; }
.curr-company-name { display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 600; color: #333; }
.curr-dot   { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
.curr-value { font-size: 16px; font-weight: 800; color: #1e3a5f; }
.curr-count { font-size: 11px; color: #aaa; margin-left: 5px; }
details.curr-hist { margin-top: 14px; border-top: 1px solid #f0f0f0; padding-top: 12px; }
details.curr-hist summary { cursor: pointer; list-style: none; font-size: 12px; font-weight: 700; color: #1565c0; user-select: none; display: flex; align-items: center; gap: 6px; }
details.curr-hist summary::-webkit-details-marker { display: none; }
details.curr-hist summary::before { content: '\25B6'; font-size: 10px; transition: transform 0.2s; display: inline-block; }
details[open].curr-hist summary::before { transform: rotate(90deg); }
.curr-hist-wrap { overflow-x: auto; margin-top: 12px; }
.curr-hist-tbl { width: 100%; border-collapse: collapse; font-size: 12px; }
.curr-hist-tbl thead th { background: #f5f7fa; padding: 8px 12px; text-align: right; font-weight: 700; color: #555; border-bottom: 2px solid #e0e4ea; white-space: nowrap; }
.curr-hist-tbl thead th:first-child { text-align: left; }
.curr-hist-tbl tbody td { padding: 9px 12px; border-bottom: 1px solid #f5f5f5; text-align: right; vertical-align: middle; }
.curr-hist-tbl tbody td:first-child { text-align: left; font-weight: 600; white-space: nowrap; }
.curr-hist-tbl tbody tr:hover td { background: #fafbfc; }
.curr-hist-tbl tfoot td { background: #f0f4f8; font-weight: 800; border-top: 2px solid #dde3ed; padding: 9px 12px; text-align: right; }
.curr-hist-tbl tfoot td:first-child { text-align: left; }
.meta-card { background: white; border-radius: 10px; padding: 20px 22px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); border-top: 4px solid #1e3a5f; }
.meta-top  { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; gap: 12px; flex-wrap: wrap; }
.meta-title { font-size: 14px; font-weight: 800; color: #1e3a5f; margin-bottom: 3px; }
.meta-subtitle { font-size: 12px; color: #888; }
.meta-status-badge { padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 700; white-space: nowrap; border: 1px solid; }
.meta-gauge-area { margin-bottom: 18px; }
.meta-gauge-labels { display: flex; justify-content: space-between; font-size: 11px; color: #aaa; margin-bottom: 8px; }
.meta-gauge-track { position: relative; height: 32px; }
.meta-gauge-bar  { position: absolute; top: 50%; transform: translateY(-50%); left: 0; right: 0; height: 14px; border-radius: 7px; background: linear-gradient(to right, #ef5350 0%, #ffa726 35%, #ffee58 60%, #66bb6a 100%); }
.meta-gauge-dim  { position: absolute; top: 50%; transform: translateY(-50%); right: 0; height: 14px; background: rgba(0,0,0,0.30); border-radius: 0 7px 7px 0; }
.meta-gauge-knob { position: absolute; top: 50%; transform: translate(-50%,-50%); width: 24px; height: 24px; border-radius: 50%; background: white; border: 3px solid #1e3a5f; box-shadow: 0 2px 8px rgba(0,0,0,0.25); z-index: 3; }
.meta-gauge-knob::before { content: attr(data-label); position: absolute; bottom: 28px; left: 50%; transform: translateX(-50%); background: #1e3a5f; color: white; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 4px; white-space: nowrap; }
.meta-gauge-knob::after  { content: ''; position: absolute; bottom: 20px; left: 50%; transform: translateX(-50%); border: 5px solid transparent; border-top-color: #1e3a5f; }
.meta-pct-row  { display: flex; align-items: center; gap: 8px; margin-top: 8px; }
.meta-pct-bg   { flex: 1; height: 7px; background: #eee; border-radius: 4px; overflow: hidden; }
.meta-pct-fill { height: 100%; border-radius: 4px; }
.meta-pct-txt  { font-size: 12px; font-weight: 800; min-width: 44px; text-align: right; }
.meta-stats { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; border-top: 1px solid #f0f0f0; padding-top: 14px; }
.meta-stat-val { font-size: 17px; font-weight: 800; line-height: 1.2; }
.meta-stat-lbl { font-size: 11px; color: #999; margin-top: 3px; }
.prio-summary { display: grid; grid-template-columns: repeat(3,1fr); gap: 14px; margin-bottom: 20px; }
.prio-card { background: white; border-radius: 10px; padding: 18px 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }
.prio-card.caixa      { border-top: 4px solid #43a047; }
.prio-card.negociacao { border-top: 4px solid #e53935; }
.prio-card.regua      { border-top: 4px solid #9e9e9e; }
.prio-card-title { font-size: 11px; font-weight: 700; color: #777; text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 6px; }
.prio-card-value { font-size: 22px; font-weight: 800; color: #1e3a5f; line-height: 1.1; }
.prio-card-sub   { font-size: 11px; color: #aaa; margin-top: 4px; }
.prio-table-wrap { overflow-x: auto; }
.prio-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
.prio-table thead th { background: #f5f7fa; padding: 10px 12px; text-align: left; font-weight: 700; color: #444; border-bottom: 2px solid #e0e4ea; cursor: pointer; user-select: none; white-space: nowrap; }
.prio-table thead th:hover { background: #eef0f3; }
.prio-table thead th.sort-asc::after  { content: ' \2191'; color: #1e3a5f; }
.prio-table thead th.sort-desc::after { content: ' \2193'; color: #1e3a5f; }
.prio-table tbody td { padding: 9px 12px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }
.prio-table tbody tr:hover td { background: #f9fafb; }
.prio-table .group-sep td { background: #eef2f7; font-weight: 800; color: #1e3a5f; font-size: 12px; padding: 8px 12px; border-top: 2px solid #d0d8e8; border-bottom: 1px solid #dde3ed; }
.prio-table .pareto-row td { background: #fffdf0 !important; }
.prio-table .pareto-row td:first-child { border-left: 3px solid #fdd835; padding-left: 9px; }
.pareto-star { color: #f9a825; margin-left: 3px; font-size: 12px; }
.group-badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 10.5px; font-weight: 700; white-space: nowrap; }
.group-badge.caixa      { background: #e8f5e9; color: #2e7d32; }
.group-badge.transicao  { background: #fff3e0; color: #e65100; }
.group-badge.negociacao { background: #ffebee; color: #c62828; }
.group-badge.regua      { background: #f5f5f5; color: #757575; border: 1px solid #e0e0e0; }
.rec-table-wrap { overflow-x: auto; }
.rec-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.rec-table thead th { background: #1e3a5f; color: white; padding: 10px 14px; text-align: center; font-weight: 700; white-space: nowrap; }
.rec-table thead th:first-child { text-align: left; border-radius: 8px 0 0 0; }
.rec-table thead th:last-child  { border-radius: 0 8px 0 0; }
.rec-table tbody td { padding: 11px 14px; border-bottom: 1px solid #f0f0f0; }
.rec-cell { text-align: right; font-weight: 600; color: #2e7d32; font-variant-numeric: tabular-nums; }
.rec-company { font-weight: 600; color: #333; }
.rec-total-row td { background: #f0f7f0; font-weight: 800; border-top: 2px solid #a5d6a7; }
.rec-total-row .rec-cell { color: #1b5e20; font-size: 14px; }
.rec-zero { color: #bbb !important; font-weight: 400 !important; }
@media (max-width: 768px) {
  .aging-grid { grid-template-columns: repeat(2,1fr); }
  .cards-grid { grid-template-columns: repeat(2,1fr); }
  .prio-summary { grid-template-columns: 1fr; }
  .tab-content { padding: 14px; }
}
'@

# ─── Historico mensal HTML ────────────────────────────────────────────────

$linhasHist = ""
$totInst=0.0; $totTele=0.0; $totMed=0.0
foreach ($mk in $mesesOrdenados) {
    $eh = $histGlobal[$mk]
    $isAtual = ($mk -eq $mesAtual)
    $nomeMes = Fmt-MesNome $mk
    $badge = if ($isAtual) { " <span style='font-size:10px;background:#1565c0;color:white;padding:1px 5px;border-radius:3px;font-weight:700'>ATUAL</span>" } else { "" }

    $vi = $eh.Instituto.val; $vt = $eh.Telecom.val; $vm = $eh.Medical.val
    $totInst += $vi; $totTele += $vt; $totMed += $vm
    $vtotal = $vi + $vt + $vm

    function Cel-Hist($v, $c) {
        if ($v -eq 0) { return "<td style='color:#ddd'>&#x2014;</td>" }
        return "<td style='color:#1e3a5f;font-weight:600'>$(Fmt-BRL $v)<br><span style='font-size:10px;color:#bbb'>$c tit.</span></td>"
    }

    $linhasHist += "<tr><td>$nomeMes$badge</td>$(Cel-Hist $vi $eh.Instituto.cnt)$(Cel-Hist $vt $eh.Telecom.cnt)$(Cel-Hist $vm $eh.Medical.cnt)<td style='font-weight:800;color:#1e3a5f'>$(Fmt-BRL $vtotal)</td></tr>`n"
}

# ─── Bloco por empresa ────────────────────────────────────────────────────

function Gerar-TabEmpresa($emp, $empNome) {
    $p1 = if ($emp.total -gt 0) { [math]::Round($emp.f1/$emp.total*100,1) } else { 0 }
    $p2 = if ($emp.total -gt 0) { [math]::Round($emp.f2/$emp.total*100,1) } else { 0 }
    $p3 = if ($emp.total -gt 0) { [math]::Round($emp.f3/$emp.total*100,1) } else { 0 }
    $p4 = if ($emp.total -gt 0) { [math]::Round($emp.f4/$emp.total*100,1) } else { 0 }

    # Insights
    $ins = ""
    if ($emp.cnt4 -gt 0) {
        $ins += "<div class='insight-card danger'><div class='insight-icon'>&#x1F6A8;</div><div><div class='insight-title'>Risco Critico</div><div class='insight-text'>$($emp.cnt4) titulo(s) acima de 90 dias totalizando $(Fmt-BRL $emp.f4). Alta probabilidade de perda definitiva.</div></div></div>`n"
    }
    if ($emp.f4 / [Math]::Max($emp.total,1) -gt 0.5) {
        $ins += "<div class='insight-card warning'><div class='insight-icon'>&#x26A0;&#xFE0F;</div><div><div class='insight-title'>Concentracao Critica</div><div class='insight-text'>Mais de $p4% da inadimplencia de $empNome esta acima de 90 dias. Revisao urgente de politica de credito.</div></div></div>`n"
    }
    if ($emp.cnt1 -gt 0) {
        $ins += "<div class='insight-card info'><div class='insight-icon'>&#x1F4A1;</div><div><div class='insight-title'>Oportunidade Imediata</div><div class='insight-text'>$($emp.cnt1) titulo(s) com 1-30 dias em atraso ($(Fmt-BRL $emp.f1)). Cobranca ativa agora tem alta chance de recuperacao.</div></div></div>`n"
    }

    # Acoes
    $acs = ""
    if ($emp.cnt4 -gt 0) {
        $acs += "<div class='action-group'><div class='action-header red'>&#x1F6A8; URGENTE &mdash; 90+ dias</div>"
        $top90 = $emp.topCl | Where-Object { $_.Value.max_dias -gt 90 } | Select-Object -First 3
        foreach ($cl in $top90) { $acs += "<div class='action-item red'>Contato executivo: $(Esc-Html $cl.Key) ($(Fmt-BRL $cl.Value.total))</div>" }
        $acs += "</div>`n"
    }
    if ($emp.cnt1 -gt 0 -or $emp.cnt2 -gt 0) {
        $acs += "<div class='action-group'><div class='action-header orange'>&#x1F4DE; PRIORIDADE &mdash; 1-60 dias</div>"
        $top60 = $emp.topCl | Where-Object { $_.Value.max_dias -le 60 } | Select-Object -First 3
        foreach ($cl in $top60) { $acs += "<div class='action-item orange'>Cobranca ativa: $(Esc-Html $cl.Key) ($(Fmt-BRL $cl.Value.total))</div>" }
        if (-not $top60) { $acs += "<div class='no-items'>Nenhum cliente nesta faixa no momento.</div>" }
        $acs += "</div>`n"
    }

    # Tabela clientes
    $linCl = ""
    foreach ($cl in $emp.topCl) {
        $d = $cl.Value.max_dias
        $ac = Get-AgingClass $d
        $badgeMap = @{ green="green"; yellow="yellow"; orange="orange"; red="red" }
        $labelMap  = @{ green="1-30d"; yellow="31-60d"; orange="61-90d"; red="90+d" }
        $linCl += "<tr><td>$(Esc-Html $cl.Key)</td><td style='text-align:right;font-weight:700;color:#1e3a5f'>$(Fmt-BRL $cl.Value.total)</td><td style='text-align:center'>$($cl.Value.titulos)</td><td style='text-align:center'><span class='badge $($badgeMap[$ac])'>$d dias</span></td><td style='text-align:center'><span class='badge $($badgeMap[$ac])'>$($labelMap[$ac])</span></td></tr>`n"
    }

    return @"
<div class="tab-content" id="$($empNome.ToLower())">
  <div class="company-header" style="background:$($emp.grad)">
    <h2>ICC $empNome</h2>
    <p>Relatorio de Inadimplencia &mdash; $dataCurta</p>
  </div>
  <div class="cards-grid">
    <div class="card"><div class="card-value">$(Fmt-BRL $emp.total)</div><div class="card-label">Total Inadimplente</div></div>
    <div class="card"><div class="card-value">$($emp.nclientes)</div><div class="card-label">Clientes em Atraso</div></div>
    <div class="card"><div class="card-value">$($emp.titulos)</div><div class="card-label">Titulos em Aberto</div></div>
    <div class="card red"><div class="card-value">$($emp.cnt4)</div><div class="card-label">Casos Criticos (90+ dias)</div></div>
  </div>
  <div class="section">
    <h3>&#x23F0; Aging dos Titulos</h3>
    <div class="aging-grid">
      <div class="aging-box green"><div class="aging-count">$($emp.cnt1)</div><div class="aging-label">1-30 dias</div><div class="aging-value">$(Fmt-BRL $emp.f1)</div></div>
      <div class="aging-box yellow"><div class="aging-count">$($emp.cnt2)</div><div class="aging-label">31-60 dias</div><div class="aging-value">$(Fmt-BRL $emp.f2)</div></div>
      <div class="aging-box orange"><div class="aging-count">$($emp.cnt3)</div><div class="aging-label">61-90 dias</div><div class="aging-value">$(Fmt-BRL $emp.f3)</div></div>
      <div class="aging-box red"><div class="aging-count">$($emp.cnt4)</div><div class="aging-label">90+ dias</div><div class="aging-value">$(Fmt-BRL $emp.f4)</div></div>
    </div>
  </div>
  <div class="section"><h3>&#x1F4A1; Insights do Dia</h3>$ins</div>
  <div class="section"><h3>&#x2705; Acoes Recomendadas</h3>$acs</div>
  <div class="section">
    <h3>&#x1F4CB; Detalhamento por Cliente (Top 30)</h3>
    <div class="table-container">
      <table>
        <thead><tr><th>Cliente</th><th style="text-align:right">Valor em Aberto</th><th style="text-align:center">Titulos</th><th style="text-align:center">Maior Atraso</th><th style="text-align:center">Faixa</th></tr></thead>
        <tbody>$linCl</tbody>
      </table>
    </div>
  </div>
</div>
"@
}

$tabInstituto = Gerar-TabEmpresa $dadosEmp[0] "Instituto"
$tabTelecom   = Gerar-TabEmpresa $dadosEmp[1] "Telecom"
$tabMedical   = Gerar-TabEmpresa $dadosEmp[2] "Medical"

# ─── Priorizacao JSON para JS ──────────────────────────────────────────────

$prioJson = "["
$prioItems = @()
foreach ($cl in $todosOrdenados) {
    $prioItems += "{`"n`":`"$(Esc-Html $cl.n)`",`"c`":`"$($cl.c)`",`"v`":$($cl.v),`"d`":$($cl.d),`"a`":`"$($cl.a)`",`"g`":`"$($cl.g)`",`"go`":$($cl.go),`"p`":$($cl.p),`"cnt`":$($cl.cnt)}"
}
$prioJson += ($prioItems -join ",") + "]"

# ─── Company breakdown HTML ───────────────────────────────────────────────

$breakdownHtml = ""
foreach ($emp in $dadosEmp) {
    $pct = if ($totalGeral -gt 0) { [math]::Round($emp.total/$totalGeral*100,1) } else { 0 }
    $breakdownHtml += "<div class='company-row'><div class='company-dot' style='background:$($emp.cor)'></div><div class='company-info'><strong>ICC $($emp.nome)</strong><span>$($emp.titulos) titulos &nbsp;|&nbsp; $($emp.nclientes) clientes</span></div><div class='pct-bar-wrap'><div class='pct-bar-bg'><div class='pct-bar' style='width:$pct%;background:$($emp.cor)'></div></div></div><div class='company-value'><strong>$(Fmt-BRL $emp.total)</strong><span>$pct% do total</span></div></div>`n"
}

# ─── Corrente por empresa ─────────────────────────────────────────────────

$corrRows = ""
foreach ($emp in $dadosEmp) {
    $corrRows += "<div class='curr-company-row'><div class='curr-company-name'><div class='curr-dot' style='background:$($emp.cor)'></div>ICC $($emp.nome)</div><div><span class='curr-value'>$(Fmt-BRL $emp.currVal)</span><span class='curr-count'>$($emp.currCnt) titulo(s)</span></div></div>`n"
}

# ─── Fluxo KPI cards ──────────────────────────────────────────────────────

$projetadoCardClass = if ($fluxoProjetadoPositivo) { "green" } else { "red" }
$projetadoValClass  = if ($fluxoProjetadoPositivo) { "green" } else { "red" }

$fluxoKpiCards = @"
        <div class="card"><div class="card-value">$fluxoBanco</div><div class="card-label">Saldo em Banco</div></div>
        <div class="card"><div class="card-value">$fluxoAReceber</div><div class="card-label">A Receber (projetado)</div></div>
        <div class="card orange"><div class="card-value">$fluxoAPagar</div><div class="card-label">A Pagar (projetado)</div></div>
        <div class="card $projetadoCardClass"><div class="card-value $projetadoValClass">$fluxoProjetado</div><div class="card-label">Saldo Projetado</div></div>
"@

# ─── HTML completo ────────────────────────────────────────────────────────

$htmlContent = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="refresh" content="420">
<title>ICC Grupo &mdash; Relatorio Integrado</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
$CSS
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <h1>&#x1F4CA; Relatorio Integrado &mdash; ICC Grupo</h1>
    <p>Inadimplencia &middot; Fluxo de Caixa</p>
  </div>
  <div class="header-right">Ultima atualizacao<br><strong>$dataStr</strong></div>
</div>

<div class="main-nav">
  <button class="main-tab-btn active" data-section="inad" onclick="showMainSection('inad')">&#x1F4CA; Inadimplencia</button>
  <button class="main-tab-btn" data-section="fluxo" onclick="showMainSection('fluxo')">&#x1F4B0; Fluxo de Caixa</button>
</div>

<!-- SECAO: INADIMPLENCIA -->
<div id="main-inad" class="main-section active">

  <div style="padding:20px 24px 0">
    <div class="company-header" style="background:linear-gradient(135deg,#0d1b2a,#1e3a5f);margin-bottom:16px">
      <h2>Resumo Executivo &mdash; ICC Grupo</h2>
      <p>Inadimplencia e Fluxo de Caixa &mdash; $dataCurta</p>
    </div>

    <div style="margin-bottom:8px;font-size:12px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:1px;padding:0 4px">Inadimplencia</div>
    <div class="cards-grid" style="margin-bottom:20px">
      <div class="card">
        <div class="card-value">$(Fmt-BRL $totalGeral)</div>
        <div class="card-label">Total Consolidado em Aberto</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDia.cls)">$($tDia.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSem.cls)">$($tSem.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val $($tMes.cls)">$($tMes.val)</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val $($tTri.cls)">$($tTri.val)</span></div>
        </div>
      </div>
      <div class="card">
        <div class="card-value">$clientesGeral</div>
        <div class="card-label">Clientes Inadimplentes</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDiaC.cls)">$($tDiaC.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSemC.cls)">$($tSemC.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val flat">&mdash;</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val flat">&mdash;</span></div>
        </div>
      </div>
      <div class="card">
        <div class="card-value">$titulosGeral</div>
        <div class="card-label">Titulos em Aberto</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDiaT.cls)">$($tDiaT.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSemT.cls)">$($tSemT.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val flat">&mdash;</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val flat">&mdash;</span></div>
        </div>
      </div>
      <div class="card red">
        <div class="card-value">$casos90</div>
        <div class="card-label">Casos Criticos (90+ dias)</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDia90.cls)">$($tDia90.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSem90.cls)">$($tSem90.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val flat">&mdash;</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val flat">&mdash;</span></div>
        </div>
      </div>
    </div>

    <div style="margin-bottom:8px;font-size:12px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:1px;padding:0 4px">Fluxo de Caixa <span style="font-size:10px;color:#aaa;font-weight:400">(dados: $fluxoLastUpdate)</span></div>
    <div class="cards-grid" style="margin-bottom:20px">
$fluxoKpiCards
    </div>
  </div>

  <div class="tabs">
    <button class="tab-btn active" data-tab="consolidado" onclick="showTab('consolidado')">&#x1F4CA; Consolidado</button>
    <button class="tab-btn" data-tab="priorizacao" onclick="showTab('priorizacao')">&#x1F3AF; Priorizacao</button>
    <button class="tab-btn" data-tab="instituto" onclick="showTab('instituto')">&#x1F393; ICC Instituto</button>
    <button class="tab-btn" data-tab="telecom" onclick="showTab('telecom')">&#x1F4E1; ICC Telecom</button>
    <button class="tab-btn" data-tab="medical" onclick="showTab('medical')">&#x1F3E5; ICC Medical</button>
  </div>

  <!-- TAB: CONSOLIDADO -->
  <div class="tab-content active" id="consolidado">
    <div class="company-header" style="background:linear-gradient(135deg,#0d1b2a,#1e3a5f)">
      <h2>Resumo Executivo Consolidado &mdash; ICC Grupo</h2>
      <p>Visao Gerencial de Inadimplencia &mdash; $dataCurta</p>
    </div>
    <div class="cards-grid">
      <div class="card">
        <div class="card-value">$(Fmt-BRL $totalGeral)</div>
        <div class="card-label">Total Consolidado em Aberto</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDia.cls)">$($tDia.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSem.cls)">$($tSem.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val $($tMes.cls)">$($tMes.val)</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val $($tTri.cls)">$($tTri.val)</span></div>
        </div>
      </div>
      <div class="card">
        <div class="card-value">$clientesGeral</div>
        <div class="card-label">Clientes Inadimplentes</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDiaC.cls)">$($tDiaC.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSemC.cls)">$($tSemC.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val flat">&mdash;</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val flat">&mdash;</span></div>
        </div>
      </div>
      <div class="card">
        <div class="card-value">$titulosGeral</div>
        <div class="card-label">Titulos em Aberto</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDiaT.cls)">$($tDiaT.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSemT.cls)">$($tSemT.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val flat">&mdash;</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val flat">&mdash;</span></div>
        </div>
      </div>
      <div class="card red">
        <div class="card-value">$casos90</div>
        <div class="card-label">Casos Criticos (90+ dias)</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDia90.cls)">$($tDia90.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSem90.cls)">$($tSem90.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val flat">&mdash;</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val flat">&mdash;</span></div>
        </div>
      </div>
    </div>

    <div class="curr-card">
      <div class="curr-card-title">&#x1F4C5; INADIMPLENCIA CORRENTE &mdash; $mesLabel</div>
      <div class="curr-rows">$corrRows</div>
      <details class="curr-hist">
        <summary>Ver historico mensal completo</summary>
        <div class="curr-hist-wrap">
          <table class="curr-hist-tbl">
            <thead><tr><th>Mes</th><th>Instituto</th><th>Telecom</th><th>Medical</th><th>Total</th></tr></thead>
            <tbody>$linhasHist</tbody>
            <tfoot><tr><td>Total acumulado</td><td>$(Fmt-BRL $totInst)</td><td>$(Fmt-BRL $totTele)</td><td>$(Fmt-BRL $totMed)</td><td>$(Fmt-BRL $totalGeral)</td></tr></tfoot>
          </table>
        </div>
      </details>
    </div>

    <div class="meta-card">
      <div class="meta-top">
        <div>
          <div class="meta-title">&#x1F3AF; META DE INADIMPLENCIA &mdash; ICC GRUPO</div>
          <div class="meta-subtitle">Reduzir para <strong>$(Fmt-BRL $META_VALOR)</strong> ate dez/2026 &nbsp;&middot;&nbsp; Definida em 01/06/2026</div>
        </div>
        <div class="meta-status-badge" style="color:$($metaStatus.cor);border-color:$($metaStatus.brd);background:$($metaStatus.bg)">
          $($metaStatus.ico) $($metaStatus.txt)
        </div>
      </div>
      <div class="meta-gauge-area">
        <div class="meta-gauge-labels">
          <span>Inicio&nbsp;(01/06/2026):&nbsp;<strong>$(Fmt-BRL $META_INICIO)</strong></span>
          <span><strong>$(Fmt-BRL $META_VALOR)</strong>&nbsp;&#x1F3C1;&nbsp;Meta&nbsp;dez/26</span>
        </div>
        <div class="meta-gauge-track">
          <div class="meta-gauge-bar"></div>
          <div class="meta-gauge-dim" style="width:$dimWidth%"></div>
          <div class="meta-gauge-knob" style="left:$($knobLeft)%" data-label="$(Fmt-BRL $totalGeral)"></div>
        </div>
        <div class="meta-pct-row">
          <div class="meta-pct-bg"><div class="meta-pct-fill" style="width:$pctMeta%;background:#e53935"></div></div>
          <span class="meta-pct-txt" style="color:#e53935">$pctMeta%</span>
        </div>
      </div>
      <div class="meta-stats">
        <div><div class="meta-stat-val" style="color:#1e3a5f">$(Fmt-BRL $totalGeral)</div><div class="meta-stat-lbl">Inadimplencia atual</div></div>
        <div><div class="meta-stat-val" style="color:#43a047">$(Fmt-BRL $reduzido)</div><div class="meta-stat-lbl">Ja reduzido desde inicio</div></div>
        <div><div class="meta-stat-val" style="color:#e53935">$(Fmt-BRL $recuperarTotal)</div><div class="meta-stat-lbl">Falta reduzir</div></div>
        <div><div class="meta-stat-val">$(Fmt-BRL $recDia)</div><div class="meta-stat-lbl">Recuperar / dia (necessario)</div></div>
      </div>
    </div>

    <div class="section">
      <h3>Inadimplencia por Empresa</h3>
      <div class="company-breakdown">$breakdownHtml</div>
    </div>
  </div>

  <!-- TAB: PRIORIZACAO -->
  <div class="tab-content" id="priorizacao">
    <div class="company-header" style="background:linear-gradient(135deg,#1a1a2e,#16213e)">
      <h2>&#x1F3AF; Priorizacao de Cobranca</h2>
      <p>Ordem de ataque para maxima recuperacao de caixa &mdash; $dataCurta</p>
    </div>
    <div class="prio-summary">
      <div class="prio-card caixa">
        <div class="prio-card-title">&#x1F4B5; CAIXA RAPIDO (1-30d)</div>
        <div class="prio-card-value">$(Fmt-BRL $vcaixa)</div>
        <div class="prio-card-sub">$($gcaixa.Count) cliente(s) &middot; maior chance de recuperacao</div>
      </div>
      <div class="prio-card negociacao">
        <div class="prio-card-title">&#x1F91D; NEGOCIACAO (31-90d)</div>
        <div class="prio-card-value">$(Fmt-BRL ($vtransicao + $vnegociacao))</div>
        <div class="prio-card-sub">$($gtransicao.Count + $gnegociacao.Count) cliente(s) &middot; exige contato executivo</div>
      </div>
      <div class="prio-card regua">
        <div class="prio-card-title">&#x2696;&#xFE0F; REGUA / JURIDICO (90+d)</div>
        <div class="prio-card-value">$(Fmt-BRL $vregua)</div>
        <div class="prio-card-sub">$($gregua.Count) cliente(s) &middot; recuperacao judicial/protesto</div>
      </div>
    </div>
    <div class="section">
      <h3>Lista Completa de Priorizacao</h3>
      <div class="prio-table-wrap">
        <table class="prio-table" id="prioTable">
          <thead>
            <tr>
              <th onclick="sortPrio(0)">Cliente</th>
              <th onclick="sortPrio(1)">Empresa</th>
              <th onclick="sortPrio(2)" style="text-align:right">Valor</th>
              <th onclick="sortPrio(3)" style="text-align:center">Dias</th>
              <th onclick="sortPrio(4)" style="text-align:center">Faixa</th>
              <th onclick="sortPrio(5)" style="text-align:center">Grupo</th>
              <th style="text-align:center">Titulos</th>
            </tr>
          </thead>
          <tbody id="prioBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  $tabInstituto
  $tabTelecom
  $tabMedical

</div>

<!-- SECAO: FLUXO DE CAIXA -->
<div id="main-fluxo" class="main-section">

  <div class="tabs">
    <button class="tab-btn flx-tab-btn active" data-tab="flx_consolidado" onclick="showFluxoTab('flx_consolidado')">&#x1F4CA; Consolidado</button>
    <button class="tab-btn flx-tab-btn" data-tab="flx_antecipacao" onclick="showFluxoTab('flx_antecipacao')">&#x1F4C5; Antecipacao</button>
    <button class="tab-btn flx-tab-btn" data-tab="flx_instituto" onclick="showFluxoTab('flx_instituto')">&#x1F393; ICC Instituto</button>
    <button class="tab-btn flx-tab-btn" data-tab="flx_telecom" onclick="showFluxoTab('flx_telecom')">&#x1F4E1; ICC Telecom</button>
    <button class="tab-btn flx-tab-btn" data-tab="flx_medical" onclick="showFluxoTab('flx_medical')">&#x1F3E5; ICC Medical</button>
  </div>

  $fluxoTabsHtml

</div>

<!-- MAIN JS -->
<script>
function showMainSection(id) {
  document.querySelectorAll('.main-section').forEach(function(s){s.classList.remove('active');});
  document.querySelectorAll('.main-tab-btn').forEach(function(b){b.classList.remove('active');});
  document.getElementById('main-' + id).classList.add('active');
  document.querySelectorAll('[data-section="' + id + '"]').forEach(function(b){b.classList.add('active');});
  if (id === 'fluxo') {
    setTimeout(function() {
      try { showFluxoTab('flx_consolidado'); } catch(e) {}
    }, 50);
  }
}

function showTab(id) {
  document.querySelectorAll('#main-inad .tab-content').forEach(function(t){ t.classList.remove('active'); });
  document.querySelectorAll('#main-inad .tab-btn').forEach(function(b){ b.classList.remove('active'); });
  var tab = document.getElementById(id);
  if (tab) tab.classList.add('active');
  document.querySelectorAll('#main-inad .tab-btn').forEach(function(btn){
    if (btn.getAttribute('data-tab') === id) btn.classList.add('active');
  });
}

var D=$prioJson;

var sortCol=-1, sortDir=1;
function sortPrio(col){
  if(sortCol===col) sortDir*=-1; else { sortCol=col; sortDir=1; }
  document.querySelectorAll('.prio-table thead th').forEach(function(th,i){
    th.classList.remove('sort-asc','sort-desc');
    if(i===col) th.classList.add(sortDir===1?'sort-asc':'sort-desc');
  });
  renderPrio();
}

function renderPrio(){
  var sorted=D.slice().sort(function(a,b){
    var va,vb;
    if(sortCol===0){va=a.n;vb=b.n;}
    else if(sortCol===1){va=a.c;vb=b.c;}
    else if(sortCol===2){va=a.v;vb=b.v;}
    else if(sortCol===3){va=a.d;vb=b.d;}
    else if(sortCol===4){va=a.a;vb=b.a;}
    else if(sortCol===5){va=a.go;vb=b.go;}
    else{va=a.go*1e12-a.v;vb=b.go*1e12-b.v;}
    if(va<vb)return -1*sortDir;if(va>vb)return 1*sortDir;return 0;
  });
  var html=''; var lastG='';
  sorted.forEach(function(r){
    if(r.g!==lastG){
      var lbl={'caixa':'&#x1F4B5; CAIXA RAPIDO (1-30 dias)','transicao':'&#x23F0; TRANSICAO (31-60 dias)','negociacao':'&#x1F91D; NEGOCIACAO (61-90+ dias)','regua':'&#x2696;&#xFE0F; REGUA / JURIDICO'}[r.g]||r.g;
      html+='<tr class="group-sep"><td colspan="7">'+lbl+'</td></tr>';
      lastG=r.g;
    }
    var rowCls=r.p?'pareto-row':'';
    var star=r.p?'<span class="pareto-star">&#x2605;</span>':'';
    var bval=r.v.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
    var gcls={'caixa':'caixa','transicao':'transicao','negociacao':'negociacao','regua':'regua'}[r.g]||'gray';
    html+='<tr class="'+rowCls+'"><td>'+r.n+star+'</td><td>'+r.c+'</td><td style="text-align:right;font-weight:700;color:#1e3a5f">'+bval+'</td><td style="text-align:center;font-weight:700">'+r.d+'d</td><td style="text-align:center">'+r.a+'</td><td style="text-align:center"><span class="group-badge '+gcls+'">'+r.g+'</span></td><td style="text-align:center;color:#888">'+r.cnt+'</td></tr>';
  });
  document.getElementById('prioBody').innerHTML=html;
}
renderPrio();
</script>

<!-- FLUXO SCRIPT -->
<script>
$fluxoScript
</script>

</body>
</html>
"@

# ─── Salvar e publicar ────────────────────────────────────────────────────

Write-Host ""
Write-Host "[$(Get-Date -f 'HH:mm:ss')] Salvando HTML ($([math]::Round($htmlContent.Length/1024,1)) KB)..."
[System.IO.File]::WriteAllText($HTML, $htmlContent, [System.Text.Encoding]::UTF8)

Write-Host "[$(Get-Date -f 'HH:mm:ss')] Publicando no GitHub..."
Set-Location $REPO
$env:GIT_TERMINAL_PROMPT = "0"
git add index.html snapshot.json 2>&1 | Out-Null
$commitOut = git commit -m "Auto: $dataStr" 2>&1
Write-Host "  git commit: $commitOut"
$pushOut = git push 2>&1
Write-Host "  git push: $pushOut"
Write-Host "[$(Get-Date -f 'HH:mm:ss')] OK: https://exactus-data-ia.github.io/icc-inadimplencia/"
