# ICC - Gerador de Relatorio Integrado (Inadimplencia + Fluxo de Caixa)
# Executa a cada 7 minutos via Task Scheduler

$REPO        = "C:\Users\MICRO1\Documents\icc-inadimplencia"
$HTML        = "$REPO\index.html"
$SNAPSHOT    = "$REPO\snapshot.json"
$HOJE        = Get-Date
$OMIE_CR     = "https://app.omie.com.br/api/v1/financas/contareceber/"
$OMIE_CL     = "https://app.omie.com.br/api/v1/geral/clientes/"
$OMIE_CP     = "https://app.omie.com.br/api/v1/financas/contapagar/"
$META_VALOR  = 200000
$META_INICIO = 1031143.11  # valor no inicio do tracking
$HIST_FILE   = "$REPO\historico.json"
$EMAIL_CFG   = "$REPO\email-config.json"
$EMAIL_STAMP = "$REPO\email-sent.txt"

$EMPRESAS = @(
    @{ nome="Instituto"; cor="#2196F3"; grad="linear-gradient(135deg,#1e3a5f,#2196F3)"; app_key="3946880386449"; app_secret="0c15f825cded97455749c7d6b7558f1e" },
    @{ nome="Telecom";   cor="#00BCD4"; grad="linear-gradient(135deg,#004d5f,#00BCD4)"; app_key="4472437527558"; app_secret="eb030b4871537b1d984ff4078a469f75" },
    @{ nome="Medical";   emoji=""; cor="#4CAF50"; grad="linear-gradient(135deg,#1b4d1f,#4CAF50)"; app_key="7069173264153"; app_secret="9632f5b931f568b6b09accbf25f47496" }
)

# â”€â”€â”€ API helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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
        $r = $null
        for ($tentativa = 1; $tentativa -le 3; $tentativa++) {
            try {
                $r = Invoke-RestMethod -Uri $OMIE_CR -Method Post -Body $b -ContentType "application/json" -TimeoutSec 60
                break
            } catch {
                Write-Host "  [AVISO] pag $pag tentativa $tentativa falhou: $($_.Exception.Message)"
                if ($tentativa -lt 3) { Start-Sleep -Seconds 5 }
            }
        }
        if ($null -eq $r) { Write-Host "  [ERRO] API nao respondeu apos 3 tentativas (pag $pag). Abortando."; break }
        if ($r.faultstring) { Write-Host "  [ERRO API] $($r.faultstring) (pag $pag)"; break }
        if ($r.conta_receber_cadastro) { $all += $r.conta_receber_cadastro }
        $tot = [int]$r.total_de_paginas; $pag++
    } while ($pag -le $tot)
    # Garantir que apenas status ATRASADO seja considerado inadimplente
    # (a API pode retornar RECEBIDO/CANCELADO misturados no mesmo resultado)
    return @($all | Where-Object { $_.status_titulo -eq "ATRASADO" })
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

function Fmt-Situacao($status) {
    switch ($status) {
        "ATRASADO" { return @{ txt="ATRASADO"; cls="red"   } }
        "ABERTO"   { return @{ txt="A VENCER";  cls="blue"  } }
        "RECEBIDO" { return @{ txt="RECEBIDO"; cls="green" } }
        "PAGO"     { return @{ txt="PAGO";     cls="green" } }
        default    { return @{ txt=$status;    cls="gray"  } }
    }
}

function Get-ContasReceberAberto($app_key, $app_secret) {
    # Contas a receber NAO vencidas (status ABERTO = a vencer)
    $all = @(); $pag = 1
    do {
        $b = @{ call="ListarContasReceber"; app_key=$app_key; app_secret=$app_secret
                param=@(@{ pagina=$pag; registros_por_pagina=500; apenas_importado_api="N"; filtrar_por_status="ABERTO" }) } | ConvertTo-Json -Depth 5
        $r = $null
        for ($t=1; $t -le 3; $t++) {
            try { $r = Invoke-RestMethod -Uri $OMIE_CR -Method Post -Body $b -ContentType "application/json" -TimeoutSec 60; break }
            catch { if ($t -lt 3) { Start-Sleep -Seconds 5 } }
        }
        if ($null -eq $r -or $r.faultstring) { break }
        if ($r.conta_receber_cadastro) { $all += $r.conta_receber_cadastro }
        $tot = [int]$r.total_de_paginas; $pag++
    } while ($pag -le $tot)
    return $all
}

function Get-ContasPagar($app_key, $app_secret) {
    # Contas a pagar pendentes: ABERTO (a vencer) + ATRASADO (vencido)
    $all = @()
    foreach ($status in @("ABERTO","ATRASADO")) {
        $pag = 1
        do {
            $b = @{ call="ListarContasPagar"; app_key=$app_key; app_secret=$app_secret
                    param=@(@{ pagina=$pag; registros_por_pagina=500; apenas_importado_api="N"; filtrar_por_status=$status }) } | ConvertTo-Json -Depth 5
            $r = $null
            for ($t=1; $t -le 3; $t++) {
                try { $r = Invoke-RestMethod -Uri $OMIE_CP -Method Post -Body $b -ContentType "application/json" -TimeoutSec 60; break }
                catch { if ($t -lt 3) { Start-Sleep -Seconds 5 } }
            }
            if ($null -eq $r) { break }
            if ($r.faultstring) { Write-Host "  [ERRO CP $status] $($r.faultstring)"; break }
            if ($r.conta_pagar_cadastro) { $all += $r.conta_pagar_cadastro }
            $tot = [int]$r.total_de_paginas; $pag++
        } while ($pag -le $tot)
    }
    return $all
}


function Send-InsightEmail($insights, $emailCfg, $dadosEmp, $totalGeral, $totalInad31, $totalAtraso, $_totalBaixas, $_baixasSorted, $casos90, $pctMeta, $META_VALOR, $recuperarTotal, $recDia, $dataStr, $dataCurta) {
    $ptBR = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")

    $insightsHtml = ""
    foreach ($ins in $insights) {
        $bg  = switch ($ins.tipo) { "danger"{"#fff5f5"} "good"{"#f0fff4"} "warning"{"#fffdf0"} default{"#f0f7ff"} }
        $brd = switch ($ins.tipo) { "danger"{"#e53935"} "good"{"#43a047"} "warning"{"#fdd835"} default{"#1e88e5"} }
        $insightsHtml += "<div style='border-left:4px solid $brd;background:$bg;padding:12px 16px;margin-bottom:10px;border-radius:0 6px 6px 0'>
            <div style='font-weight:700;font-size:14px;margin-bottom:4px'>$($ins.icone) $($ins.titulo)</div>
            <div style='font-size:13px;color:#555;line-height:1.5'>$($ins.texto)</div>
        </div>"
    }

    $empHtml = ""
    foreach ($emp in $dadosEmp) {
        $inad31emp = [math]::Round($emp.total - $emp.f1, 2)
        $empHtml += "<tr><td style='padding:8px 12px;font-weight:600'>ICC $($emp.nome)</td>
            <td style='padding:8px 12px;text-align:right;color:#e65100'>R`$ $(([double]$emp.f1).ToString('N2',$ptBR))</td>
            <td style='padding:8px 12px;text-align:right;color:#c62828'>R`$ $(([double]$inad31emp).ToString('N2',$ptBR))</td>
            <td style='padding:8px 12px;text-align:right;font-weight:800'>R`$ $(([double]$emp.total).ToString('N2',$ptBR))</td></tr>"
    }

    $body = @"
<!DOCTYPE html><html><head><meta charset='UTF-8'></head>
<body style='font-family:Arial,sans-serif;background:#f0f2f5;margin:0;padding:20px'>
<div style='max-width:680px;margin:0 auto;background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.1)'>
  <div style='background:linear-gradient(135deg,#0d1b2a,#1e3a5f);color:white;padding:28px 32px'>
    <h1 style='margin:0;font-size:22px'>&#x1F4CA; Relatorio Diario de Inadimplencia</h1>
    <p style='margin:6px 0 0;opacity:0.75;font-size:13px'>ICC Grupo &mdash; $dataStr</p>
  </div>
  <div style='padding:28px 32px'>
    <h2 style='font-size:16px;color:#1e3a5f;margin:0 0 16px;padding-bottom:10px;border-bottom:2px solid #f0f2f5'>&#x1F4C8; Posicao Atual</h2>
    <table style='width:100%;border-collapse:collapse;margin-bottom:24px'>
      <thead><tr style='background:#f5f7fa'>
        <th style='padding:10px 12px;text-align:left;font-size:12px;color:#666;border-bottom:2px solid #e0e4ea'>Empresa</th>
        <th style='padding:10px 12px;text-align:right;font-size:12px;color:#e65100;border-bottom:2px solid #e0e4ea'>Em Atraso (ate 30d)</th>
        <th style='padding:10px 12px;text-align:right;font-size:12px;color:#c62828;border-bottom:2px solid #e0e4ea'>Inadimplente (31+d)</th>
        <th style='padding:10px 12px;text-align:right;font-size:12px;color:#1e3a5f;border-bottom:2px solid #e0e4ea'>Total</th>
      </tr></thead>
      <tbody>$empHtml</tbody>
      <tfoot><tr style='background:#f0f4f8;font-weight:800'>
        <td style='padding:10px 12px;border-top:2px solid #dde3ed'>TOTAL GRUPO</td>
        <td style='padding:10px 12px;text-align:right;color:#e65100;border-top:2px solid #dde3ed'>R`$ $(([double]$totalAtraso).ToString('N2',$ptBR))</td>
        <td style='padding:10px 12px;text-align:right;color:#c62828;border-top:2px solid #dde3ed'>R`$ $(([double]$totalInad31).ToString('N2',$ptBR))</td>
        <td style='padding:10px 12px;text-align:right;font-size:18px;border-top:2px solid #dde3ed'>R`$ $(([double]$totalGeral).ToString('N2',$ptBR))</td>
      </tr></tfoot>
    </table>
    <h2 style='font-size:16px;color:#1e3a5f;margin:0 0 16px;padding-bottom:10px;border-bottom:2px solid #f0f2f5'>&#x1F9E0; Insights &amp; Alertas</h2>
    $insightsHtml
    <div style='background:#f5f7fa;border-radius:8px;padding:16px 20px;margin-top:20px;border-left:4px solid #1e3a5f'>
      <div style='font-weight:700;font-size:13px;color:#1e3a5f;margin-bottom:8px'>&#x1F3AF; Meta de Inadimplencia &mdash; Dez/2026</div>
      <div style='font-size:13px;color:#555'>Atual: <strong>R`$ $(([double]$totalGeral).ToString('N2',$ptBR))</strong> &nbsp;|&nbsp; Meta: <strong>R`$ $(([double]$META_VALOR).ToString('N2',$ptBR))</strong> &nbsp;|&nbsp; Progresso: <strong>$pctMeta%</strong></div>
      <div style='font-size:12px;color:#888;margin-top:4px'>Faltam <strong>R`$ $(([double]$recuperarTotal).ToString('N2',$ptBR))</strong> para a meta &nbsp;|&nbsp; Necessario: <strong>R`$ $(([double]$recDia).ToString('N2',$ptBR))/dia</strong></div>
    </div>
    <div style='margin-top:16px;font-size:12px;color:#888'>
      &#x2705; <strong>$($_baixasSorted.Count) titulos baixados</strong> nos ultimos 60 dias &nbsp;&middot;&nbsp; Total recuperado: <strong>R`$ $(([double]$_totalBaixas).ToString('N2',$ptBR))</strong>
    </div>
  </div>
  <div style='background:#f5f7fa;padding:16px 32px;font-size:11px;color:#aaa;border-top:1px solid #e8ecf0;line-height:1.8'>
    Gerado automaticamente &mdash; $dataStr &nbsp;|&nbsp; ICC Analytics<br>
    Painel completo: <a href='https://exactus-data-ia.github.io/icc-inadimplencia/' style='color:#1565c0'>https://exactus-data-ia.github.io/icc-inadimplencia/</a>
  </div>
</div>
</body></html>
"@

    $smtpClient = New-Object System.Net.Mail.SmtpClient($emailCfg.SmtpServer, [int]$emailCfg.SmtpPort)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($emailCfg.De, $emailCfg.AppPassword)

    $msg = New-Object System.Net.Mail.MailMessage
    $msg.From = New-Object System.Net.Mail.MailAddress($emailCfg.De, "ICC Analytics")
    foreach ($dest in $emailCfg.Para) { $msg.To.Add($dest) }
    $msg.Subject = "ICC Inadimplencia $dataCurta — Inad: R`$ $(([double]$totalInad31).ToString('N2',$ptBR)) | Atraso: R`$ $(([double]$totalAtraso).ToString('N2',$ptBR))"
    $msg.Body = $body
    $msg.IsBodyHtml = $true
    $msg.BodyEncoding = [System.Text.Encoding]::UTF8
    $msg.SubjectEncoding = [System.Text.Encoding]::UTF8

    try {
        $smtpClient.Send($msg)
        Write-Host "  [EMAIL] Enviado para $($emailCfg.Para -join ', ')"
        (Get-Date).ToString("yyyy-MM-dd HH:mm") | Set-Content $EMAIL_STAMP -Encoding UTF8
    } catch {
        Write-Host "  [EMAIL ERRO] $($_.Exception.Message)"
    } finally {
        $smtpClient.Dispose()
        $msg.Dispose()
    }
}

function Get-ContasLiquidadas($app_key, $app_secret) {
    # Busca titulos com status RECEBIDO modificados nos ultimos 60 dias (info.dAlt = data do pagamento)
    # O filtro de data sem filtrar_por_tipo_data aplica em info.dAlt (ultima alteracao)
    $d1  = $HOJE.AddDays(-60).ToString("dd/MM/yyyy")
    $d2  = $HOJE.ToString("dd/MM/yyyy")
    $all = @(); $pag = 1
    do {
        $b = @{ call="ListarContasReceber"; app_key=$app_key; app_secret=$app_secret
                param=@(@{ pagina=$pag; registros_por_pagina=500; apenas_importado_api="N"
                           filtrar_por_data_de=$d1; filtrar_por_data_ate=$d2 }) } | ConvertTo-Json -Depth 5
        $r = $null
        for ($t=1; $t -le 3; $t++) {
            try { $r = Invoke-RestMethod -Uri $OMIE_CR -Method Post -Body $b -ContentType "application/json" -TimeoutSec 60; break }
            catch { if ($t -lt 3) { Start-Sleep -Seconds 5 } }
        }
        if ($null -eq $r -or $r.faultstring) { break }
        if ($r.conta_receber_cadastro) { $all += $r.conta_receber_cadastro }
        $tot = [int]$r.total_de_paginas; $pag++
    } while ($pag -le $tot)
    # Retornar apenas os RECEBIDO (pagos), adicionando _dl = data do pagamento (info.dAlt)
    return @($all | Where-Object { $_.status_titulo -eq "RECEBIDO" } | ForEach-Object {
        $_ | Add-Member -NotePropertyName "_dl" -NotePropertyValue ($_.info.dAlt) -Force -PassThru
    })
}

function Build-MovTable($abertos, $liquidados, $mapa) {
    $ptBR = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")
    $html = ""

    # ── Em Aberto ─────────────────────────────────────────────────────────
    $abSorted = @($abertos | Sort-Object { [double]$_.valor_documento } -Descending)
    $totalAb  = [double]($abSorted | Measure-Object { [double]$_.valor_documento } -Sum).Sum
    $html += "<div class='mov-section'>"
    $html += "<div class='mov-header em-aberto'>&#x1F534; EM ABERTO &mdash; $($abSorted.Count) titulo(s) &middot; R$ $($totalAb.ToString('N2',$ptBR))</div>"
    if ($abSorted.Count -eq 0) {
        $html += "<p style='color:#aaa;padding:12px;font-style:italic'>Nenhum titulo em aberto.</p>"
    } else {
        $html += "<table class='mov-table'><thead><tr><th>Cliente</th><th>Vencimento</th><th>Atraso</th><th style='text-align:right'>Valor</th></tr></thead><tbody>"
        foreach ($c in $abSorted) {
            $cod  = [string]$c.codigo_cliente_fornecedor
            $nome = if ($mapa.ContainsKey($cod)) { $mapa[$cod] } else { "Cod $cod" }
            $dias = Get-DiasAtraso $c.data_vencimento
            $acls = if ($dias -le 30) { "orange" } elseif ($dias -le 60) { "yellow" } elseif ($dias -le 90) { "orange" } else { "red" }
            $val  = [double]$c.valor_documento
            $html += "<tr><td>$(Esc-Html $nome)</td><td style='text-align:center'>$($c.data_vencimento)</td>"
            $html += "<td style='text-align:center'><span class='badge $acls'>$dias dias</span></td>"
            $html += "<td style='text-align:right;font-weight:700;color:#1e3a5f'>R$ $($val.ToString('N2',$ptBR))</td></tr>"
        }
        $html += "</tbody></table>"
    }
    $html += "</div>"

    # ── Dado Baixa ────────────────────────────────────────────────────────
    $liqSorted = @($liquidados | Where-Object { $_._dl -and $_._dl.Trim().Length -gt 2 } |
                   Sort-Object { try { [datetime]::ParseExact($_._dl,"dd/MM/yyyy",$null) } catch { [datetime]::MinValue } } -Descending)
    $totalLiq  = [double]($liqSorted | Measure-Object { [double]$_.valor_documento } -Sum).Sum
    $html += "<div class='mov-section' style='margin-top:16px'>"
    $html += "<div class='mov-header dado-baixa'>&#x2705; DADO BAIXA &mdash; $($liqSorted.Count) titulo(s) &middot; R$ $($totalLiq.ToString('N2',$ptBR))</div>"
    if ($liqSorted.Count -eq 0) {
        $html += "<p style='color:#aaa;padding:12px;font-style:italic'>Nenhuma baixa registrada nos ultimos 60 dias.</p>"
    } else {
        # Agrupar por data de liquidacao
        $grupos = @{}
        foreach ($c in $liqSorted) {
            $dk = $c._dl
            if (-not $grupos.ContainsKey($dk)) { $grupos[$dk] = @() }
            $grupos[$dk] += $c
        }
        $datasOrdenadas = $grupos.Keys | Sort-Object { try { [datetime]::ParseExact($_,"dd/MM/yyyy",$null) } catch { [datetime]::MinValue } } -Descending
        foreach ($data in $datasOrdenadas) {
            $items    = $grupos[$data]
            $subTotal = [double]($items | Measure-Object { [double]$_.valor_documento } -Sum).Sum
            $html += "<div class='mov-data-grupo'>&#x1F4C5; $data &nbsp;&middot;&nbsp; R$ $($subTotal.ToString('N2',$ptBR))</div>"
            $html += "<table class='mov-table'><thead><tr><th>Cliente</th><th>Vencimento</th><th style='text-align:right'>Valor</th></tr></thead><tbody>"
            foreach ($c in ($items | Sort-Object { [double]$_.valor_documento } -Descending)) {
                $cod  = [string]$c.codigo_cliente_fornecedor
                $nome = if ($mapa.ContainsKey($cod)) { $mapa[$cod] } else { "Cod $cod" }
                $val  = [double]$c.valor_documento
                $html += "<tr><td>$(Esc-Html $nome)</td><td style='text-align:center'>$($c.data_vencimento)</td>"
                $html += "<td style='text-align:right;font-weight:700;color:#2e7d32'>R$ $($val.ToString('N2',$ptBR))</td></tr>"
            }
            $html += "</tbody></table>"
        }
    }
    $html += "</div>"
    return $html
}

function Build-CRTable($rows, $mapaClientes) {
    if (-not $rows -or $rows.Count -eq 0) { return "<p style='color:#aaa;font-style:italic;padding:12px'>Nenhum titulo pendente.</p>" }
    $sorted = @($rows | Sort-Object { [double]$_.valor_documento } -Descending | Select-Object -First 200)
    $html = "<table style='width:100%;border-collapse:collapse;font-size:12.5px'>"
    $html += "<thead><tr style='background:#f5f7fa'><th style='padding:9px 12px;text-align:left;border-bottom:2px solid #e0e4ea'>Cliente</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Situacao</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Vencimento</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Previsao Pgto</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Ult. Recebimento</th><th style='padding:9px 12px;text-align:right;border-bottom:2px solid #e0e4ea'>Valor</th></tr></thead><tbody>"
    foreach ($c in $sorted) {
        $cod  = [string]$c.codigo_cliente_fornecedor
        $nome = if ($mapaClientes.ContainsKey($cod)) { $mapaClientes[$cod] } else { "Cod $cod" }
        $sit  = Fmt-Situacao $c.status_titulo
        $prev = if ($c.data_previsao -and $c.data_previsao.Trim() -and $c.data_previsao.Length -gt 2) { $c.data_previsao } else { "&mdash;" }
        $rec  = if ($c.data_recebimento -and $c.data_recebimento.Trim() -and $c.data_recebimento.Length -gt 2) { $c.data_recebimento } else { "&mdash;" }
        $html += "<tr style='border-bottom:1px solid #f0f0f0'>"
        $html += "<td style='padding:9px 12px'>$(Esc-Html $nome)</td>"
        $html += "<td style='padding:9px 12px;text-align:center'><span class='badge $($sit.cls)'>$($sit.txt)</span></td>"
        $html += "<td style='padding:9px 12px;text-align:center'>$($c.data_vencimento)</td>"
        $html += "<td style='padding:9px 12px;text-align:center'>$prev</td>"
        $html += "<td style='padding:9px 12px;text-align:center'>$rec</td>"
        $html += "<td style='padding:9px 12px;text-align:right;font-weight:700;color:#1e3a5f'>$(Fmt-BRL ([double]$c.valor_documento))</td>"
        $html += "</tr>"
    }
    $html += "</tbody></table>"
    return $html
}

function Build-CPTable($rows, $mapaClientes) {
    if (-not $rows -or $rows.Count -eq 0) { return "<p style='color:#aaa;font-style:italic;padding:12px'>Nenhum titulo pendente.</p>" }
    $sorted = @($rows | Sort-Object { [double]$_.valor_documento } -Descending | Select-Object -First 200)
    $html = "<table style='width:100%;border-collapse:collapse;font-size:12.5px'>"
    $html += "<thead><tr style='background:#f5f7fa'><th style='padding:9px 12px;text-align:left;border-bottom:2px solid #e0e4ea'>Fornecedor</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Situacao</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Vencimento</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Previsao Pgto</th><th style='padding:9px 12px;text-align:center;border-bottom:2px solid #e0e4ea'>Ult. Pagamento</th><th style='padding:9px 12px;text-align:right;border-bottom:2px solid #e0e4ea'>Valor</th></tr></thead><tbody>"
    foreach ($c in $sorted) {
        $cod  = [string]$c.codigo_cliente_fornecedor
        $nome = if ($mapaClientes.ContainsKey($cod)) { $mapaClientes[$cod] } else { "Cod $cod" }
        $sit  = Fmt-Situacao $c.status_titulo
        $prev = if ($c.data_previsao -and $c.data_previsao.Trim() -and $c.data_previsao.Length -gt 2) { $c.data_previsao } else { "&mdash;" }
        $pag  = if ($c.data_pagamento -and $c.data_pagamento.Trim() -and $c.data_pagamento.Length -gt 2) { $c.data_pagamento } else { "&mdash;" }
        $html += "<tr style='border-bottom:1px solid #f0f0f0'>"
        $html += "<td style='padding:9px 12px'>$(Esc-Html $nome)</td>"
        $html += "<td style='padding:9px 12px;text-align:center'><span class='badge $($sit.cls)'>$($sit.txt)</span></td>"
        $html += "<td style='padding:9px 12px;text-align:center'>$($c.data_vencimento)</td>"
        $html += "<td style='padding:9px 12px;text-align:center'>$prev</td>"
        $html += "<td style='padding:9px 12px;text-align:center'>$pag</td>"
        $html += "<td style='padding:9px 12px;text-align:right;font-weight:700;color:#c62828'>$(Fmt-BRL ([double]$c.valor_documento))</td>"
        $html += "</tr>"
        }
    $html += "</tbody></table>"
    return $html
}

# â”€â”€â”€ Coleta principal â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

Write-Host "[$(Get-Date -f 'HH:mm:ss')] Iniciando coleta..."

$dadosEmp = @()
$todosClientes = @()  # para priorizacao

foreach ($emp in $EMPRESAS) {
    Write-Host "[$(Get-Date -f 'HH:mm:ss')] $($emp.nome)..."
    $mapa        = Get-MapaClientes $emp.app_key $emp.app_secret
    $contas      = Get-Contas $emp.app_key $emp.app_secret
    $crAbertas   = Get-ContasReceberAberto $emp.app_key $emp.app_secret
    $cpPendentes = Get-ContasPagar $emp.app_key $emp.app_secret
    $liquidados  = Get-ContasLiquidadas $emp.app_key $emp.app_secret
    Write-Host "  CR atrasado: $($contas.Count) | CR a vencer: $($crAbertas.Count) | CP pendente: $($cpPendentes.Count) | Liquidados: $($liquidados.Count)"

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

    # Totais diretos da API (CR atrasado + CR aberto = total pendente a receber)
    $totalAReceber = 0.0
    foreach ($__c in $contas)    { $totalAReceber += [double]$__c.valor_documento }
    foreach ($__c in $crAbertas) { $totalAReceber += [double]$__c.valor_documento }

    $totalAPagar = 0.0
    foreach ($__c in $cpPendentes) { $totalAPagar += [double]$__c.valor_documento }

    # Gerar tabelas HTML com nomes resolvidos
    $allCR = @($contas) + @($crAbertas)
    $crTableHtml = Build-CRTable $allCR $mapa
    $cpTableHtml = Build-CPTable $cpPendentes $mapa

    $ncli1  = ($clMap.Values | Where-Object { $_.max_dias -le 30 }).Count
    $ncli31 = ($clMap.Values | Where-Object { $_.max_dias -gt 30 }).Count
    $movHtml = Build-MovTable $contas $liquidados $mapa

    $dadosEmp += @{
        nome=$emp.nome; cor=$emp.cor; grad=$emp.grad
        total=$total; titulos=$contas.Count; nclientes=$clMap.Count
        f1=$f1; f2=$f2; f3=$f3; f4=$f4
        cnt1=$cnt1; cnt2=$cnt2; cnt3=$cnt3; cnt4=$cnt4
        ncli1=$ncli1; ncli31=$ncli31
        currVal=$currVal; currCnt=$currCnt
        topCl=$topCl; histMes=$histMes
        totalAReceber=$totalAReceber; totalAPagar=$totalAPagar
        crTableHtml=$crTableHtml; cpTableHtml=$cpTableHtml
        crCount=($contas.Count + $crAbertas.Count); cpCount=$cpPendentes.Count
        movHtml=$movHtml; liquidados=$liquidados; contasAbertas=$contas; mapaClientes=$mapa
    }
    Write-Host "  Inadimplente: $(Fmt-BRL $total) | A Receber: $(Fmt-BRL $totalAReceber) | A Pagar: $(Fmt-BRL $totalAPagar)"
}

# â”€â”€â”€ Totais consolidados â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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

# Separacao gerencial: atraso (1-30 dias) vs inadimplente (31+ dias)
$totalAtraso   = 0.0; $titulosAtraso   = 0; $clientesAtraso   = 0
$totalInad31   = 0.0; $titulosInad31   = 0; $clientesInad31   = 0
foreach ($__e in $dadosEmp) {
    $totalAtraso   += $__e.f1
    $titulosAtraso += $__e.cnt1
    $clientesAtraso += $__e.ncli1
    $totalInad31   += ($__e.total - $__e.f1)
    $titulosInad31 += ($__e.cnt2 + $__e.cnt3 + $__e.cnt4)
    $clientesInad31 += $__e.ncli31
}
$dataStr  = $HOJE.ToString("dd/MM/yyyy, HH:mm:ss")
$dataCurta = $HOJE.ToString("dd/MM/yyyy")
$mesNome  = (Get-Culture).DateTimeFormat.GetMonthName($HOJE.Month).Substring(0,3)
$mesNome  = $mesNome.Substring(0,1).ToUpper() + $mesNome.Substring(1)
$mesLabel = "$mesNome/$($HOJE.Year)"

# â”€â”€â”€ Snapshot para trends (DIA/SEM/MES/TRI) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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

# ─── Histórico diário (90 dias) ──────────────────────────────────────────────

$histData = @()
if (Test-Path $HIST_FILE) {
    try { $histData = @(Get-Content $HIST_FILE -Raw | ConvertFrom-Json) } catch {}
}
$histData = @($histData | Where-Object { $_.data -ne $HOJE.ToString(“yyyy-MM-dd”) })
$histData += [PSCustomObject]@{
    data    = $HOJE.ToString(“yyyy-MM-dd”)
    inad31  = [math]::Round($totalInad31, 2)
    atraso  = [math]::Round($totalAtraso, 2)
    cliInad = $clientesInad31
    cliAtr  = $clientesAtraso
    titInad = $titulosInad31
    titAtr  = $titulosAtraso
    casos90 = $casos90
}
if ($histData.Count -gt 90) { $histData = @($histData | Select-Object -Last 90) }
$histData | ConvertTo-Json | Set-Content $HIST_FILE -Encoding UTF8

# Referências do histórico para insights
$ref1d  = $histData | Where-Object { $_.data -eq $HOJE.AddDays(-1).ToString(“yyyy-MM-dd”) } | Select-Object -Last 1
$ref7d  = $histData | Where-Object { $_.data -eq $HOJE.AddDays(-7).ToString(“yyyy-MM-dd”) } | Select-Object -Last 1
$ref30d = $histData | Where-Object { $_.data -eq $HOJE.AddDays(-30).ToString(“yyyy-MM-dd”) } | Select-Object -Last 1

function Calc-VarPct($curr, $prev) {
    if ($null -eq $prev -or $prev -eq 0) { return $null }
    return [math]::Round(($curr - $prev) / [math]::Abs($prev) * 100, 1)
}

$var1d  = Calc-VarPct $totalInad31 $ref1d.inad31
$var7d  = Calc-VarPct $totalInad31 $ref7d.inad31
$var30d = Calc-VarPct $totalInad31 $ref30d.inad31

# â”€â”€â”€ Pareto 80/20 nos clientes globais â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

$todosOrdenados = $todosClientes | Sort-Object { $_.v } -Descending
$acum = 0.0
for ($i=0; $i -lt $todosOrdenados.Count; $i++) {
    $acum += $todosOrdenados[$i].v
    if ($acum / $totalGeral -le 0.80) { $todosOrdenados[$i].p = 1 }
}

# â”€â”€â”€ Historico mensal global (unir de todas as empresas) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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

# â”€â”€â”€ Meta gauge â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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

# â”€â”€â”€ Resumo priorizacao â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

$gcaixa      = @($todosOrdenados | Where-Object { $_.g -eq "caixa" })
$gtransicao  = @($todosOrdenados | Where-Object { $_.g -eq "transicao" })
$gnegociacao = @($todosOrdenados | Where-Object { $_.g -eq "negociacao" })
$gregua      = @($todosOrdenados | Where-Object { $_.g -eq "regua" })

$vcaixa=0.0; foreach ($__x in $gcaixa)      { $vcaixa      += $__x.v }
$vtransicao=0.0; foreach ($__x in $gtransicao)  { $vtransicao  += $__x.v }
$vnegociacao=0.0; foreach ($__x in $gnegociacao) { $vnegociacao += $__x.v }
$vregua=0.0; foreach ($__x in $gregua)      { $vregua      += $__x.v }

# â”€â”€â”€ Gerador de HTML â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

$CSS = @'
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f0f2f5; color: #222; }
.header { background: linear-gradient(135deg, #0d1b2a, #1e3a5f); color: white; padding: 18px 30px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 12px rgba(0,0,0,0.3); }
.header-left h1 { font-size: 20px; font-weight: 700; }
.header-left p { font-size: 12px; opacity: 0.65; margin-top: 3px; }
.header-right { text-align: right; font-size: 12px; opacity: 0.75; line-height: 1.6; }
.main-nav{background:#ffffff;padding:0 36px;display:flex;align-items:stretch;gap:0;border-bottom:1px solid #e8ecf0;box-shadow:0 1px 0 rgba(0,0,0,0.04);}
.main-tab-btn{position:relative;padding:0 30px;height:46px;border:none;background:none;cursor:pointer;font-size:11px;font-weight:500;letter-spacing:1.8px;text-transform:uppercase;color:#aab2bc;border-bottom:2px solid transparent;margin-bottom:-1px;transition:color 0.2s ease,border-color 0.2s ease;white-space:nowrap;display:flex;align-items:center;}
.main-tab-btn:hover{color:#1e3a5f;}
.main-tab-btn.active{color:#1e3a5f;border-bottom-color:#1e3a5f;font-weight:600;}
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
.card.orange-soft { border-top-color: #fb8c00; background: #fff8f0; }
.card-value { font-size: 26px; font-weight: 800; color: #1e3a5f; margin-bottom: 5px; line-height: 1.1; }
.card.red .card-value { color: #c62828; }
.card.orange .card-value { color: #e65100; }
.card.green .card-value { color: #2e7d32; }
.card-value.green { color: #2e7d32; }
.card-value.red { color: #c62828; }
.card-label { font-size: 12px; color: #888; font-weight: 500; }
.section-label-atraso { font-size: 11px; font-weight: 700; color: #e65100; text-transform: uppercase; letter-spacing: 1.2px; margin: 18px 0 8px; padding: 6px 12px; background: #fff3e0; border-left: 3px solid #fb8c00; border-radius: 0 4px 4px 0; display: inline-block; }
.section-label-inad { font-size: 11px; font-weight: 700; color: #c62828; text-transform: uppercase; letter-spacing: 1.2px; margin: 18px 0 8px; padding: 6px 12px; background: #ffebee; border-left: 3px solid #e53935; border-radius: 0 4px 4px 0; display: inline-block; }
.mov-panel { border: 1px solid #e0e4ea; border-radius: 8px; overflow: hidden; }
.mov-panel > summary { cursor: pointer; padding: 14px 18px; font-size: 13px; font-weight: 700; color: #1e3a5f; background: #f5f7fa; user-select: none; list-style: none; display: flex; align-items: center; gap: 8px; }
.mov-panel > summary::-webkit-details-marker { display: none; }
.mov-panel > summary::before { content: "▶"; font-size: 10px; transition: transform 0.2s; }
.mov-panel[open] > summary::before { transform: rotate(90deg); }
.mov-body { padding: 16px 18px; }
.mov-section {}
.mov-header { font-size: 12px; font-weight: 700; padding: 7px 12px; border-radius: 4px; margin-bottom: 10px; display: inline-block; }
.mov-header.em-aberto { background: #fff3e0; color: #e65100; border-left: 3px solid #fb8c00; }
.mov-header.dado-baixa { background: #e8f5e9; color: #2e7d32; border-left: 3px solid #43a047; }
.mov-table { width: 100%; border-collapse: collapse; font-size: 12.5px; margin-bottom: 6px; }
.mov-table thead tr { background: #f5f7fa; }
.mov-table th { padding: 8px 10px; text-align: left; border-bottom: 2px solid #e0e4ea; font-size: 11px; color: #666; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
.mov-table td { padding: 8px 10px; border-bottom: 1px solid #f0f0f0; }
.mov-table tr:last-child td { border-bottom: none; }
.mov-data-grupo { font-size: 12px; font-weight: 700; color: #2e7d32; background: #f1f8e9; padding: 6px 10px; margin: 12px 0 6px; border-radius: 4px; }
.baixas-panel { border-top: 1px solid #e8ecf0; margin-top: 16px; padding-top: 4px; }
.baixas-panel > summary { cursor: pointer; font-size: 12px; font-weight: 600; color: #2e7d32; padding: 10px 0 6px; list-style: none; display: flex; align-items: center; gap: 6px; user-select: none; }
.baixas-panel > summary::-webkit-details-marker { display: none; }
.baixas-panel > summary::before { content: "▶"; font-size: 9px; color: #43a047; transition: transform 0.2s; }
.baixas-panel[open] > summary::before { transform: rotate(90deg); }
.baixas-body { padding: 8px 0 4px; max-height: 360px; overflow-y: auto; }
.baixas-table { width: 100%; border-collapse: collapse; font-size: 12px; }
.baixas-table thead tr { background: #f5f7fa; }
.baixas-table th { padding: 7px 10px; text-align: left; border-bottom: 2px solid #e0e4ea; font-size: 11px; color: #888; font-weight: 600; text-transform: uppercase; letter-spacing: 0.4px; }
.baixas-table td { padding: 7px 10px; border-bottom: 1px solid #f5f5f5; }
.baixas-table tr:last-child td { border-bottom: none; }
.baixas-table tr:hover td { background: #f9fffe; }
.baixas-empty { color: #aaa; font-style: italic; font-size: 12px; padding: 8px 0; }
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
.insight-card.good    { background: #f0fff4; border-color: #43a047; }
.insights-panel { background: white; border-radius: 10px; padding: 20px 22px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-top: 4px solid #7b1fa2; }
.insights-panel h3 { font-size: 15px; font-weight: 700; margin-bottom: 16px; color: #1e3a5f; padding-bottom: 10px; border-bottom: 2px solid #f0f2f5; }
.trend-bar-wrap { margin-top: 20px; padding-top: 16px; border-top: 1px solid #f0f0f0; }
.trend-bar-title { font-size: 12px; font-weight: 700; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 10px; }
.trend-bar-row { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
.trend-bar-label { font-size: 11px; color: #666; width: 100px; flex-shrink: 0; }
.trend-bar-bg { flex: 1; height: 8px; background: #eee; border-radius: 4px; overflow: hidden; }
.trend-bar-fill { height: 100%; border-radius: 4px; transition: width 0.5s; }
.trend-bar-fill.up   { background: #e53935; }
.trend-bar-fill.down { background: #43a047; }
.trend-bar-val { font-size: 11px; font-weight: 700; min-width: 55px; text-align: right; }
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
.block-insight { border-radius: 8px; padding: 11px 14px; margin: 12px 0 4px; border-left: 4px solid; font-size: 12.5px; line-height: 1.55; }
.block-insight.good    { background:#f0fff4; border-color:#43a047; color:#1b5e20; }
.block-insight.danger  { background:#fff5f5; border-color:#e53935; color:#b71c1c; }
.block-insight.warning { background:#fffdf0; border-color:#fdd835; color:#5d4037; }
.block-insight.info    { background:#f0f7ff; border-color:#1e88e5; color:#0d47a1; }
.block-insight strong  { font-weight:800; }
.block-insight-wrap    { margin-bottom: 16px; }
'@

# â”€â”€â”€ Historico mensal HTML â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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

# â”€â”€â”€ Bloco por empresa â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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
    <p>Relatorio de Inadimplencia &middot; Contas a Receber &middot; Contas a Pagar &mdash; $dataCurta</p>
  </div>
  <div class="section-label-atraso" style="margin-top:0">&#x23F3; CONTAS A RECEBER EM ATRASO &mdash; 1 a 30 dias</div>
  <div class="cards-grid" style="margin-bottom:20px">
    <div class="card orange-soft"><div class="card-value" style="color:#e65100">$(Fmt-BRL $emp.f1)</div><div class="card-label">Total em Atraso (1-30 dias)</div></div>
    <div class="card orange-soft"><div class="card-value" style="color:#e65100">$($emp.ncli1)</div><div class="card-label">Clientes em Atraso</div></div>
    <div class="card orange-soft"><div class="card-value" style="color:#e65100">$($emp.cnt1)</div><div class="card-label">Titulos em Atraso</div></div>
  </div>
  <div class="section-label-inad">&#x1F6A8; INADIMPLENTE &mdash; 31 dias ou mais</div>
  <div class="cards-grid" style="margin-bottom:20px">
    <div class="card red"><div class="card-value">$(Fmt-BRL ($emp.total - $emp.f1))</div><div class="card-label">Total Inadimplente (31+ dias)</div></div>
    <div class="card red"><div class="card-value">$($emp.ncli31)</div><div class="card-label">Clientes Inadimplentes</div></div>
    <div class="card red"><div class="card-value">$($emp.cnt2 + $emp.cnt3 + $emp.cnt4)</div><div class="card-label">Titulos Inadimplentes</div></div>
    <div class="card green"><div class="card-value" style="color:#2e7d32">$(Fmt-BRL $emp.totalAReceber)</div><div class="card-label">A Receber (pendente)</div></div>
    <div class="card orange"><div class="card-value" style="color:#c62828">$(Fmt-BRL $emp.totalAPagar)</div><div class="card-label">A Pagar (pendente)</div></div>
  </div>
  <div class="section">
    <h3>&#x23F0; Aging da Inadimplencia</h3>
    <div class="aging-grid">
      <div class="aging-box green"><div class="aging-count">$($emp.cnt1)</div><div class="aging-label">1-30 dias</div><div class="aging-value">$(Fmt-BRL $emp.f1)</div></div>
      <div class="aging-box yellow"><div class="aging-count">$($emp.cnt2)</div><div class="aging-label">31-60 dias</div><div class="aging-value">$(Fmt-BRL $emp.f2)</div></div>
      <div class="aging-box orange"><div class="aging-count">$($emp.cnt3)</div><div class="aging-label">61-90 dias</div><div class="aging-value">$(Fmt-BRL $emp.f3)</div></div>
      <div class="aging-box red"><div class="aging-count">$($emp.cnt4)</div><div class="aging-label">90+ dias</div><div class="aging-value">$(Fmt-BRL $emp.f4)</div></div>
    </div>
  </div>
  <div class="section">
    <details class="mov-panel">
      <summary>&#x1F4CB; Movimentacao de Titulos &mdash; Em Aberto e Dado Baixa (ultimos 60 dias)</summary>
      <div class="mov-body">$($emp.movHtml)</div>
    </details>
  </div>
  <div class="section"><h3>&#x1F4A1; Insights do Dia</h3>$ins</div>
  <div class="section"><h3>&#x2705; Acoes Recomendadas</h3>$acs</div>
  <div class="section">
    <h3>&#x1F4CB; Inadimplencia por Cliente (Top 30)</h3>
    <div class="table-container">
      <table>
        <thead><tr><th>Cliente</th><th style="text-align:right">Valor em Atraso</th><th style="text-align:center">Titulos</th><th style="text-align:center">Maior Atraso</th><th style="text-align:center">Faixa</th></tr></thead>
        <tbody>$linCl</tbody>
      </table>
    </div>
  </div>
  <div class="section">
    <h3>&#x1F4B0; Contas a Receber &mdash; Todas Pendentes ($($emp.crCount) titulos &middot; $(Fmt-BRL $emp.totalAReceber))</h3>
    <div class="table-container">$($emp.crTableHtml)</div>
  </div>
  <div class="section">
    <h3>&#x1F4B3; Contas a Pagar &mdash; Todas Pendentes ($($emp.cpCount) titulos &middot; $(Fmt-BRL $emp.totalAPagar))</h3>
    <div class="table-container">$($emp.cpTableHtml)</div>
  </div>
</div>
"@
}

$tabInstituto = Gerar-TabEmpresa $dadosEmp[0] "Instituto"
$tabTelecom   = Gerar-TabEmpresa $dadosEmp[1] "Telecom"
$tabMedical   = Gerar-TabEmpresa $dadosEmp[2] "Medical"

# â”€â”€â”€ Priorizacao JSON para JS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

$prioJson = "["
$prioItems = @()
foreach ($cl in $todosOrdenados) {
    $prioItems += "{`"n`":`"$(Esc-Html $cl.n)`",`"c`":`"$($cl.c)`",`"v`":$($cl.v),`"d`":$($cl.d),`"a`":`"$($cl.a)`",`"g`":`"$($cl.g)`",`"go`":$($cl.go),`"p`":$($cl.p),`"cnt`":$($cl.cnt)}"
}
$prioJson += ($prioItems -join ",") + "]"

# â”€â”€â”€ Company breakdown HTML â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

$breakdownHtml = ""
foreach ($emp in $dadosEmp) {
    $pct = if ($totalGeral -gt 0) { [math]::Round($emp.total/$totalGeral*100,1) } else { 0 }
    $breakdownHtml += "<div class='company-row'><div class='company-dot' style='background:$($emp.cor)'></div><div class='company-info'><strong>ICC $($emp.nome)</strong><span>$($emp.titulos) titulos &nbsp;|&nbsp; $($emp.nclientes) clientes</span></div><div class='pct-bar-wrap'><div class='pct-bar-bg'><div class='pct-bar' style='width:$pct%;background:$($emp.cor)'></div></div></div><div class='company-value'><strong>$(Fmt-BRL $emp.total)</strong><span>$pct% do total</span></div></div>`n"
}

# â”€â”€â”€ Corrente por empresa â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

$corrRows = ""
foreach ($emp in $dadosEmp) {
    $corrRows += "<div class='curr-company-row'><div class='curr-company-name'><div class='curr-dot' style='background:$($emp.cor)'></div>ICC $($emp.nome)</div><div><span class='curr-value'>$(Fmt-BRL $emp.currVal)</span><span class='curr-count'>$($emp.currCnt) titulo(s)</span></div></div>`n"
}

# Movimentacao consolidada (todas as empresas)
$_allAbertos    = @(); $dadosEmp | ForEach-Object { $_allAbertos    += @($_.contasAbertas) }
$_allLiquidados = @(); $dadosEmp | ForEach-Object { $_allLiquidados += @($_.liquidados) }
$_mapaGeral = @{}; $dadosEmp | ForEach-Object { $_.mapaClientes.GetEnumerator() | ForEach-Object { $_mapaGeral[$_.Key] = $_.Value } }
$movHtmlConsolidado = Build-MovTable $_allAbertos $_allLiquidados $_mapaGeral

# Lista de baixas para o meta-card (titulos pagos, mais novo primeiro)
$ptBR2 = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")
$_baixasSorted = @($_allLiquidados | Where-Object { $_._dl -and $_._dl.Trim().Length -gt 2 } |
    Sort-Object { try { [datetime]::ParseExact($_._dl,"dd/MM/yyyy",$null) } catch { [datetime]::MinValue } } -Descending)
$_totalBaixas = [double]($_baixasSorted | Measure-Object { [double]$_.valor_documento } -Sum).Sum

# ─── Insights correlacionados inline (bloco Inadimplente, Corrente, Meta) ─────

# -- Insight 1: bloco Inadimplente (abaixo dos cards inad31) --
$_ptBRi = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")
$insightInadHtml = "<div class='block-insight-wrap'>"

# Tendência semanal
if ($null -ne $var7d -and $null -ne $ref7d) {
    $delta7v = [math]::Round($totalInad31 - [double]$ref7d.inad31, 2)
    $delta7Abs = [math]::Abs($delta7v)
    $delta7Fmt = "R$ " + $delta7Abs.ToString("N2", $_ptBRi)
    $novos7  = if ($ref7d.cliInad) { [int]$clientesInad31 - [int]$ref7d.cliInad } else { 0 }
    if ($var7d -gt 0) {
        $direcao7 = "cresceu"
        $cls7 = "danger"
        $novosTexto = if ($novos7 -gt 0) { "$novos7 novo(s) cliente(s)" } elseif ($novos7 -lt 0) { "$([math]::Abs($novos7)) cliente(s) saiu(ram)" } else { "sem variacao de clientes" }
        $insightInadHtml += "<div class='block-insight $cls7'>Inadimplencia <strong>$direcao7 $var7d%</strong> vs semana passada (+$delta7Fmt). $novosTexto inadimplentes.</div>`n"
    } elseif ($var7d -lt 0) {
        $direcao7 = "caiu"
        $cls7 = "good"
        $novosTexto = if ($novos7 -lt 0) { "$([math]::Abs($novos7)) cliente(s) regularizou(aram)" } elseif ($novos7 -gt 0) { "$novos7 novo(s) cliente(s) entraram" } else { "sem variacao de clientes" }
        $insightInadHtml += "<div class='block-insight $cls7'>Inadimplencia <strong>$direcao7 $([math]::Abs($var7d))%</strong> vs semana passada (-$delta7Fmt). $novosTexto.</div>`n"
    } else {
        $insightInadHtml += "<div class='block-insight info'>Inadimplencia <strong>estavel</strong> em relacao a semana passada. Manter ritmo de cobranca.</div>`n"
    }
}

# Recuperar 1-30 dias vs meta restante
if ($recuperarTotal -gt 0 -and $totalAtraso -gt 0) {
    $pctAtrasoMeta = [math]::Round($totalAtraso / $recuperarTotal * 100, 1)
    $insightInadHtml += "<div class='block-insight warning'>Se recuperar apenas os 1-30 dias (<strong>$(Fmt-BRL $totalAtraso)</strong>), isso representa <strong>$pctAtrasoMeta%</strong> da meta restante de reducao ($(Fmt-BRL $recuperarTotal)).</div>`n"
} elseif ($recuperarTotal -le 0) {
    $insightInadHtml += "<div class='block-insight good'>Meta atingida! Inadimplencia abaixo do alvo estabelecido.</div>`n"
}

# Casos 90+ vs meta
if ($casos90 -gt 0) {
    $val90i = 0.0; foreach ($__e in $dadosEmp) { $val90i += $__e.f4 }
    if ($recuperarTotal -gt 0) {
        $novoRecuperar = $recuperarTotal + $val90i
        $pctDificil = [math]::Round(($novoRecuperar - $recuperarTotal) / $recuperarTotal * 100, 1)
        $insightInadHtml += "<div class='block-insight danger'>Casos criticos 90+ (<strong>$(Fmt-BRL $val90i)</strong>) &mdash; se nao recuperados em 30 dias, a meta fica <strong>$pctDificil% mais dificil</strong> de atingir.</div>`n"
    } else {
        $insightInadHtml += "<div class='block-insight danger'>Atencao: <strong>$casos90 titulos acima de 90 dias</strong> ($(Fmt-BRL $val90i)) podem virar perda definitiva. Acionar protocolo urgente.</div>`n"
    }
}

$insightInadHtml += "</div>"

# -- Insight 2: bloco Inadimplência Corrente --
$insightCorrenteHtml = "<div class='block-insight-wrap'>"

$currValGeralI = 0.0; $currCntGeralI = 0
foreach ($__e in $dadosEmp) { $currValGeralI += $__e.currVal; $currCntGeralI += $__e.currCnt }

# Ticket médio
if ($currCntGeralI -gt 0) {
    $ticketMedio = [math]::Round($currValGeralI / $currCntGeralI, 2)
    $insightCorrenteHtml += "<div class='block-insight info'>Ticket medio dos titulos vencidos no mes: <strong>$(Fmt-BRL $ticketMedio)</strong> ($currCntGeralI titulo(s)).</div>`n"
}

# Projeção de recuperação
$_baixas60Valor = $_totalBaixas
if ($_baixas60Valor -gt 0 -and $currValGeralI -gt 0) {
    $recMes = [math]::Round($_baixas60Valor / 2, 2)   # media mensal baseada nos 60 dias
    if ($recMes -gt 0) {
        $diasParaRecuperar = [math]::Round($currValGeralI / ($recMes / 30), 0)
        $insightCorrenteHtml += "<div class='block-insight warning'>No ritmo atual de recuperacao (baixas ultimos 60 dias = <strong>$(Fmt-BRL $_baixas60Valor)</strong> / 2 meses = <strong>$(Fmt-BRL $recMes)/mes</strong>), o mes corrente seria recuperado em aproximadamente <strong>$diasParaRecuperar dias</strong>.</div>`n"
    }
}

# Corrente vs meta restante
if ($currValGeralI -gt 0 -and $recuperarTotal -gt 0) {
    $pctCorrenteMeta = [math]::Round($currValGeralI / $recuperarTotal * 100, 1)
    $insightCorrenteHtml += "<div class='block-insight danger'><strong>$(Fmt-BRL $currValGeralI)</strong> vencido neste mes representa <strong>$pctCorrenteMeta%</strong> do que ainda falta para a meta.</div>`n"
} elseif ($currValGeralI -eq 0) {
    $insightCorrenteHtml += "<div class='block-insight good'>Nenhum titulo venceu neste mes ainda. Bom controle de vencimentos.</div>`n"
}

$insightCorrenteHtml += "</div>"

# -- Insight 3: bloco Meta de Inadimplência (chips: chance de pagar + risco alto) --

# Construir set de clientes com historico de pagamento recente
$codsComHistPagto = @{}
foreach ($liq in $_allLiquidados) { $codsComHistPagto[[string]$liq.codigo_cliente_fornecedor] = $true }

# Construir mapa de inadimplentes (codigo -> {nome, valor, max_dias, emp})
$delinqMap = @{}
foreach ($empd in $dadosEmp) {
    foreach ($ct in $empd.contasAbertas) {
        $cod = [string]$ct.codigo_cliente_fornecedor
        $val = [double]$ct.valor_documento
        $dias = Get-DiasAtraso $ct.data_vencimento
        $nomeD = if ($empd.mapaClientes.ContainsKey($cod)) { $empd.mapaClientes[$cod] } else { "Cod $cod" }
        if (-not $delinqMap.ContainsKey($cod)) {
            $delinqMap[$cod] = @{ nome=$nomeD; valor=0.0; max_dias=0; emp="ICC $($empd.nome)" }
        }
        $delinqMap[$cod].valor += $val
        if ($dias -gt $delinqMap[$cod].max_dias) { $delinqMap[$cod].max_dias = $dias }
    }
}

# Melhor chance: esta em codsComHistPagto E em delinqMap, maior valor
$bestChance = $null
$bestChanceVal = 0.0
foreach ($kv in $delinqMap.GetEnumerator()) {
    if ($codsComHistPagto.ContainsKey($kv.Key) -and $kv.Value.valor -gt $bestChanceVal) {
        $bestChanceVal = $kv.Value.valor
        $bestChance = $kv.Value
    }
}

# Maior risco: NAO esta em codsComHistPagto E max_dias > 90, maior valor
$highRisk = $null
$highRiskVal = 0.0
foreach ($kv in $delinqMap.GetEnumerator()) {
    if ((-not $codsComHistPagto.ContainsKey($kv.Key)) -and $kv.Value.max_dias -gt 90 -and $kv.Value.valor -gt $highRiskVal) {
        $highRiskVal = $kv.Value.valor
        $highRisk = $kv.Value
    }
}

$insightMetaHtml = "<div class='block-insight-wrap'>"

if ($bestChance) {
    $nomeChance = Esc-Html $bestChance.nome
    $insightMetaHtml += "<div class='block-insight good'>&#x2B50; <strong>Priorize: $nomeChance</strong> &mdash; pagou recentemente e ainda tem <strong>$(Fmt-BRL $bestChance.valor)</strong> em aberto. Alta chance de recuperacao.</div>`n"
} else {
    $insightMetaHtml += "<div class='block-insight info'>Nenhum cliente inadimplente com historico de pagamento recente identificado. Focar nos de menor atraso.</div>`n"
}

if ($highRisk) {
    $nomeRisco = Esc-Html $highRisk.nome
    $insightMetaHtml += "<div class='block-insight danger'>&#x26A0; <strong>Risco: $nomeRisco</strong> ($($highRisk.emp)) &mdash; <strong>$($highRisk.max_dias) dias</strong> sem pagamento, <strong>$(Fmt-BRL $highRisk.valor)</strong> inadimplente. Sem historico de pagamento recente &mdash; acionar protocolo urgente.</div>`n"
} else {
    $insightMetaHtml += "<div class='block-insight warning'>Nenhum caso critico (90+ dias) sem historico de pagamento identificado no momento.</div>`n"
}

$insightMetaHtml += "</div>"

# ─── Geração de insights (aqui: todas as variáveis disponíveis) ───────────────

$ptBR_ins = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")
$insightsList = @()

# Posição atual — sempre exibida
$pctAtraso = if ($totalGeral -gt 0) { [math]::Round($totalAtraso / $totalGeral * 100, 1) } else { 0 }
$pctInad31 = if ($totalGeral -gt 0) { [math]::Round($totalInad31 / $totalGeral * 100, 1) } else { 0 }
$insightsList += @{ tipo="info"; icone="&#x1F4CA;"; titulo="Posicao Atual: $(Fmt-BRL $totalGeral) em aberto (total)";
    texto="Em atraso ate 30 dias: $(Fmt-BRL $totalAtraso) ($pctAtraso% do total) | Inadimplente 31+: $(Fmt-BRL $totalInad31) ($pctInad31% do total). $clientesInad31 cliente(s) inadimplente(s) com $titulosInad31 titulo(s)." }

# Tendência semanal (disponível após 7 dias de histórico)
if ($null -ne $var7d) {
    $delta7 = [math]::Round($totalInad31 - $ref7d.inad31, 2)
    $deltaFmt7 = "R$ " + ([math]::Abs($delta7)).ToString("N2", $ptBR_ins)
    if ($var7d -gt 5) {
        $insightsList += @{ tipo="danger"; icone="&#x1F6A8;"; titulo="Inadimplencia em Alta esta Semana (+$var7d%)";
            texto="Inadimplencia (31+ dias) cresceu $var7d% nos ultimos 7 dias (+$deltaFmt7). Acao imediata necessaria no pipeline de cobranca." }
    } elseif ($var7d -lt -5) {
        $insightsList += @{ tipo="good"; icone="&#x2705;"; titulo="Inadimplencia Caindo esta Semana ($var7d%)";
            texto="Inadimplencia (31+ dias) reduziu $([math]::Abs($var7d))% nos ultimos 7 dias (-$deltaFmt7). Bom ritmo de recuperacao." }
    } else {
        $insightsList += @{ tipo="info"; icone="&#x2194;"; titulo="Inadimplencia Estavel na Semana ($var7d%)";
            texto="Variacao de apenas $var7d% em 7 dias. Mantenha o ritmo de cobranca ativo para nao deixar acumular." }
    }
}

# Tendência mensal (disponível após 30 dias)
if ($null -ne $var30d) {
    $delta30 = [math]::Round($totalInad31 - $ref30d.inad31, 2)
    $deltaFmt30 = "R$ " + ([math]::Abs($delta30)).ToString("N2", $ptBR_ins)
    if ($var30d -gt 0) {
        $insightsList += @{ tipo="warning"; icone="&#x26A0;&#xFE0F;"; titulo="Tendencia 30 dias: Alta de $var30d%";
            texto="Em relacao a 30 dias atras, inadimplencia aumentou $var30d% (+$deltaFmt30). Revisar politica de credito e frequencia de cobranca." }
    } else {
        $insightsList += @{ tipo="good"; icone="&#x1F4C9;"; titulo="Tendencia 30 dias: Queda de $([math]::Abs($var30d))%";
            texto="Em relacao a 30 dias atras, inadimplencia reduziu $([math]::Abs($var30d))% (-$deltaFmt30). Trajetoria positiva de recuperacao." }
    }
}

# Casos críticos 90+
if ($casos90 -gt 0) {
    $val90emp = 0.0; foreach ($__e in $dadosEmp) { $val90emp += $__e.f4 }
    $pct90emp = if ($totalInad31 -gt 0) { [math]::Round($val90emp / $totalInad31 * 100, 1) } else { 0 }
    $insightsList += @{ tipo="danger"; icone="&#x1F534;"; titulo="$casos90 Caso(s) Critico(s) — Acima de 90 dias";
        texto="$(Fmt-BRL $val90emp) concentrados em titulos acima de 90 dias ($pct90emp% da inadimplencia). Risco elevado de perda definitiva. Acionar protocolo de negociacao ou juridico imediatamente." }
}

# Recuperação (60 dias)
if ($_totalBaixas -gt 0) {
    $taxaRec = if (($totalInad31 + $_totalBaixas) -gt 0) { [math]::Round($_totalBaixas / ($totalInad31 + $_totalBaixas) * 100, 1) } else { 0 }
    $insightsList += @{ tipo="good"; icone="&#x1F4B0;"; titulo="Recuperacao: $(Fmt-BRL $_totalBaixas) nos ultimos 60 dias";
        texto="$($_baixasSorted.Count) titulos baixados. Taxa de recuperacao: $taxaRec% do portfolio inadimplente total. Monitorar se ritmo e suficiente para atingir a meta." }
}

# Progresso meta — sempre exibido
$insightsList += @{
    tipo   = if ($pctMeta -ge 20) { "good" } elseif ($pctMeta -ge 5) { "warning" } else { "danger" }
    icone  = if ($pctMeta -ge 20) { "&#x1F3AF;" } elseif ($pctMeta -ge 5) { "&#x26A0;&#xFE0F;" } else { "&#x1F6A8;" }
    titulo = "Meta Dez/2026: $pctMeta% concluida ($diasMeta dias restantes)"
    texto  = "Faltam $(Fmt-BRL $recuperarTotal) para a meta de $(Fmt-BRL $META_VALOR). Ritmo necessario: $(Fmt-BRL $recDia)/dia. $(if($pctMeta -lt 5){'Ritmo atual insuficiente — revisar estrategia com urgencia.'}elseif($pctMeta -lt 20){'Monitorar ritmo semanalmente.'}else{'No caminho certo. Manter pressao de cobranca.'})"
}

# Concentração — sempre calculado
$top1client = $todosClientes | Sort-Object { $_.v } -Descending | Select-Object -First 1
if ($top1client -and $totalInad31 -gt 0) {
    $pctTop1 = [math]::Round($top1client.v / $totalInad31 * 100, 1)
    if ($pctTop1 -gt 15) {
        $insightsList += @{ tipo="warning"; icone="&#x1F4CC;"; titulo="Concentracao: $($top1client.n) = $pctTop1% da inadimplencia";
            texto="Dependencia alta de um unico devedor ($(Fmt-BRL $top1client.v)). Negociacao com este cliente tem impacto desproporcional na reducao da inadimplencia total." }
    }
}

# ─── HTML do bloco de insights ────────────────────────────────────────────────

$insightsPanelHtml = ""
foreach ($ins in $insightsList) {
    $cls = switch ($ins.tipo) { "danger"{"danger"} "good"{"good"} "warning"{"warning"} default{"info"} }
    $insightsPanelHtml += "<div class='insight-card $cls'><div class='insight-icon'>$($ins.icone)</div><div><div class='insight-title'>$($ins.titulo)</div><div class='insight-text'>$($ins.texto)</div></div></div>`n"
}

# Barra de tendência histórica
$trendBarHtml = ""
if ($histData.Count -ge 2) {
    $ultimos = @($histData | Select-Object -Last 7)
    $maxInad = ($ultimos | Measure-Object { [double]$_.inad31 } -Maximum).Maximum
    if ($maxInad -gt 0) {
        $trendBarHtml = "<div class='trend-bar-wrap'><div class='trend-bar-title'>&#x1F4C5; Evolucao Diaria &mdash; Inadimplencia 31+ dias (ultimos $($ultimos.Count) dias)</div>"
        foreach ($snap_h in $ultimos) {
            $pctBar = [math]::Max(4, [math]::Round([double]$snap_h.inad31 / $maxInad * 100, 0))
            $isHoje = ($snap_h.data -eq $HOJE.ToString("yyyy-MM-dd"))
            $corBar = if ($isHoje) { "#1e3a5f" } else { "#90a4ae" }
            $valFmt = "R$ " + ([double]$snap_h.inad31).ToString("N0", $ptBR_ins)
            $dataLbl = try { ([datetime]::ParseExact($snap_h.data,"yyyy-MM-dd",$null)).ToString("dd/MM") } catch { $snap_h.data }
            $trendBarHtml += "<div class='trend-bar-row'>"
            $trendBarHtml += "<div class='trend-bar-label'>$dataLbl$(if($isHoje){" <strong>hoje</strong>"})</div>"
            $trendBarHtml += "<div class='trend-bar-bg'><div class='trend-bar-fill' style='width:$pctBar%;background:$corBar'></div></div>"
            $trendBarHtml += "<div class='trend-bar-val' style='color:$corBar'>$valFmt</div>"
            $trendBarHtml += "</div>"
        }
        $trendBarHtml += "</div>"
    }
}
$insightsPanelHtml += $trendBarHtml

$baixasHtml = ""
if ($_baixasSorted.Count -eq 0) {
    $baixasHtml = "<p class='baixas-empty'>Nenhuma baixa registrada nos ultimos 60 dias.</p>"
} else {
    $baixasHtml = "<table class='baixas-table'><thead><tr><th>Data Pgto</th><th>Empresa</th><th>Cliente</th><th>Vencimento</th><th style='text-align:right'>Valor</th></tr></thead><tbody>"
    foreach ($c in $_baixasSorted) {
        $cod   = [string]$c.codigo_cliente_fornecedor
        $nome  = if ($_mapaGeral.ContainsKey($cod)) { $_mapaGeral[$cod] } else { "Cod $cod" }
        # Identificar empresa pelo app_key — nao disponivel aqui, usar nome do mapa de cada empresa
        $empNome = ""
        foreach ($__e in $dadosEmp) {
            if ($__e.mapaClientes.ContainsKey($cod)) { $empNome = "ICC $($__e.nome)"; break }
        }
        $val   = [double]$c.valor_documento
        $baixasHtml += "<tr>"
        $baixasHtml += "<td style='white-space:nowrap;font-weight:700;color:#2e7d32'>$($c._dl)</td>"
        $baixasHtml += "<td style='font-size:11px;color:#666'>$empNome</td>"
        $baixasHtml += "<td>$(Esc-Html $nome)</td>"
        $baixasHtml += "<td style='text-align:center;color:#999;font-size:11px'>$($c.data_vencimento)</td>"
        $baixasHtml += "<td style='text-align:right;font-weight:700;color:#2e7d32'>R`$ $($val.ToString('N2',$ptBR2))</td>"
        $baixasHtml += "</tr>"
    }
    $baixasHtml += “</tbody></table>”
}

# ─── HTML dos insights para o painel ─────────────────────────────────────────

$insightsPanelHtml = “”
foreach ($ins in $insightsList) {
    $cls = switch ($ins.tipo) { “danger”{“danger”} “good”{“good”} “warning”{“warning”} default{“info”} }
    $insightsPanelHtml += “<div class='insight-card $cls'><div class='insight-icon'>$($ins.icone)</div><div><div class='insight-title'>$($ins.titulo)</div><div class='insight-text'>$($ins.texto)</div></div></div>`n”
}

# Barra de tendência histórica (últimos 7 snapshots)
$trendBarHtml = “”
if ($histData.Count -ge 2) {
    $ultimos = @($histData | Select-Object -Last 7)
    $maxInad = ($ultimos | Measure-Object { [double]$_.inad31 } -Maximum).Maximum
    if ($maxInad -gt 0) {
        $trendBarHtml = “<div class='trend-bar-wrap'><div class='trend-bar-title'>&#x1F4C5; Evolucao Diaria — Inadimplencia 31+ dias (ultimos $($ultimos.Count) dias)</div>”
        foreach ($snap_h in $ultimos) {
            $pctBar = [math]::Round([double]$snap_h.inad31 / $maxInad * 100, 0)
            $isHoje = ($snap_h.data -eq $HOJE.ToString(“yyyy-MM-dd”))
            $corBar = if ($isHoje) { “#1e3a5f” } else { “#90a4ae” }
            $ptBR_b = [System.Globalization.CultureInfo]::GetCultureInfo(“pt-BR”)
            $valFmt = “R$ “ + ([double]$snap_h.inad31).ToString(“N0”, $ptBR_b)
            $dataLbl = try { ([datetime]::ParseExact($snap_h.data,”yyyy-MM-dd”,$null)).ToString(“dd/MM”) } catch { $snap_h.data }
            $trendBarHtml += “<div class='trend-bar-row'>”
            $trendBarHtml += “<div class='trend-bar-label'>$dataLbl$(if($isHoje){' <strong>(hoje)</strong>'})</div>”
            $trendBarHtml += “<div class='trend-bar-bg'><div class='trend-bar-fill' style='width:$pctBar%;background:$corBar'></div></div>”
            $trendBarHtml += “<div class='trend-bar-val' style='color:$corBar'>$valFmt</div>”
            $trendBarHtml += “</div>”
        }
        $trendBarHtml += “</div>”
    }
}
$insightsPanelHtml += $trendBarHtml

# â”€â”€â”€ HTML completo â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

$htmlContent = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="refresh" content="4500">
<title>ICC Grupo &mdash; Relatorio Integrado</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
$CSS
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <h1>&#x1F4CA; Relatorio de Inadimplencia &mdash; ICC Grupo</h1>
    <p>Inadimplencia &mdash; ICC Instituto &middot; Telecom &middot; Medical</p>
  </div>
  <div class="header-right">Ultima atualizacao<br><strong>$dataStr</strong></div>
</div>

<!-- SECAO: INADIMPLENCIA -->
<div id="main-inad" class="main-section active">

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
    <div class="section-label-atraso">&#x23F3; CONTAS A RECEBER EM ATRASO &mdash; 1 a 30 dias</div>
    <div class="cards-grid" style="margin-bottom:20px">
      <div class="card orange-soft">
        <div class="card-value" style="color:#e65100">$(Fmt-BRL $totalAtraso)</div>
        <div class="card-label">Total em Atraso (1-30 dias)</div>
      </div>
      <div class="card orange-soft">
        <div class="card-value" style="color:#e65100">$clientesAtraso</div>
        <div class="card-label">Clientes em Atraso</div>
      </div>
      <div class="card orange-soft">
        <div class="card-value" style="color:#e65100">$titulosAtraso</div>
        <div class="card-label">Titulos em Atraso</div>
      </div>
    </div>

    <div class="section-label-inad">&#x1F6A8; INADIMPLENTE &mdash; 31 dias ou mais</div>
    <div class="cards-grid" style="margin-bottom:20px">
      <div class="card red">
        <div class="card-value">$(Fmt-BRL $totalInad31)</div>
        <div class="card-label">Total Inadimplente (31+ dias)</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDia.cls)">$($tDia.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSem.cls)">$($tSem.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val $($tMes.cls)">$($tMes.val)</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val $($tTri.cls)">$($tTri.val)</span></div>
        </div>
      </div>
      <div class="card red">
        <div class="card-value">$clientesInad31</div>
        <div class="card-label">Clientes Inadimplentes</div>
        <div class="card-trends">
          <div class="trend-item"><span class="trend-period">Dia</span><span class="trend-val $($tDiaC.cls)">$($tDiaC.val)</span></div>
          <div class="trend-item"><span class="trend-period">Sem</span><span class="trend-val $($tSemC.cls)">$($tSemC.val)</span></div>
          <div class="trend-item"><span class="trend-period">Mes</span><span class="trend-val flat">&mdash;</span></div>
          <div class="trend-item"><span class="trend-period">Tri</span><span class="trend-val flat">&mdash;</span></div>
        </div>
      </div>
      <div class="card red">
        <div class="card-value">$titulosInad31</div>
        <div class="card-label">Titulos Inadimplentes</div>
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
    $insightInadHtml

    <div class="curr-card">
      <div class="curr-card-title">&#x1F4C5; INADIMPLENCIA CORRENTE &mdash; $mesLabel</div>
      <div class="curr-rows">$corrRows</div>
      $insightCorrenteHtml
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
      $insightMetaHtml
      <details class="baixas-panel">
        <summary>&#x2705; Titulos Baixados &mdash; $($_baixasSorted.Count) pagamentos &middot; $(Fmt-BRL $_totalBaixas) (ultimos 60 dias)</summary>
        <div class="baixas-body">$baixasHtml</div>
      </details>
    </div>

    <div class="section">
      <h3>Inadimplencia por Empresa</h3>
      <div class="company-breakdown">$breakdownHtml</div>
    </div>

    <div class="insights-panel">
      <h3>&#x1F9E0; Smart Insights &amp; Alertas</h3>
      $insightsPanelHtml
    </div>

    <div class="section">
      <details class="mov-panel">
        <summary>&#x1F4CB; Movimentacao de Titulos &mdash; Consolidado (Em Aberto e Dado Baixa, ultimos 60 dias)</summary>
        <div class="mov-body">$movHtmlConsolidado</div>
      </details>
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

<!-- MAIN JS -->
<script>

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


</body>
</html>
“@

# ─── Envio de e-mail (uma vez por dia, primeiro ciclo >= HoraEnvio) ──────────

if (Test-Path $EMAIL_CFG) {
    try {
        $emailCfg = Get-Content $EMAIL_CFG -Raw | ConvertFrom-Json
        $horaEnvio = if ($emailCfg.HoraEnvio) { [int]$emailCfg.HoraEnvio } else { 8 }
        $ultimoEnvio = $null
        if (Test-Path $EMAIL_STAMP) {
            try { $ultimoEnvio = [datetime]::ParseExact((Get-Content $EMAIL_STAMP -Raw).Trim(),”yyyy-MM-dd HH:mm”,$null) } catch {}
        }
        $jaEnviouHoje = $ultimoEnvio -and $ultimoEnvio.Date -eq $HOJE.Date
        $alertaCritico = (-not $jaEnviouHoje) -and ($null -ne $var1d -and $var1d -gt 10)
        $envioRotina   = (-not $jaEnviouHoje) -and ($HOJE.Hour -ge $horaEnvio)

        if ($emailCfg.AppPassword -and $emailCfg.AppPassword -ne “COLE_AQUI_APP_PASSWORD”) {
            if ($envioRotina -or $alertaCritico) {
                Write-Host “[$(Get-Date -f 'HH:mm:ss')] Enviando e-mail de insights...”
                Send-InsightEmail $insightsList $emailCfg $dadosEmp $totalGeral $totalInad31 $totalAtraso $_totalBaixas $_baixasSorted $casos90 $pctMeta $META_VALOR $recuperarTotal $recDia $dataStr $dataCurta
            } else {
                Write-Host “[$(Get-Date -f 'HH:mm:ss')] E-mail ja enviado hoje (proximo envio amanha >= $($horaEnvio)h).”
            }
        } else {
            Write-Host “[$(Get-Date -f 'HH:mm:ss')] [EMAIL] Configure email-config.json com suas credenciais para ativar envio.”
        }
    } catch {
        Write-Host “[$(Get-Date -f 'HH:mm:ss')] [EMAIL ERRO] $($_.Exception.Message)”
    }
}

# â”€â”€â”€ Salvar e publicar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

Write-Host ""
Write-Host "[$(Get-Date -f 'HH:mm:ss')] Salvando HTML ($([math]::Round($htmlContent.Length/1024,1)) KB)..."
[System.IO.File]::WriteAllText($HTML, $htmlContent, [System.Text.Encoding]::UTF8)

Set-Location $REPO
$netlifyCmd = "$env:APPDATA\npm\netlify.cmd"
if (Test-Path $netlifyCmd) {
    # --- Deploy via Netlify ---
    Write-Host "[$(Get-Date -f 'HH:mm:ss')] Publicando no Netlify..."
    $deployOut = & $netlifyCmd deploy --dir . --prod 2>&1
    Write-Host "  netlify deploy: $deployOut"
    $siteUrl = ($deployOut | Select-String "Website URL:" | ForEach-Object { ($_ -replace ".*Website URL:\s*","").Trim() } | Select-Object -First 1)
    if (-not $siteUrl) { $siteUrl = "https://app.netlify.com" }
    Write-Host "[$(Get-Date -f 'HH:mm:ss')] OK (Netlify): $siteUrl"
} else {
    # --- Fallback: GitHub Pages ---
    Write-Host "[$(Get-Date -f 'HH:mm:ss')] Publicando no GitHub Pages (Netlify nao configurado)..."
    $env:GIT_TERMINAL_PROMPT = "0"
    git add index.html snapshot.json historico.json 2>&1 | Out-Null
    $commitOut = git commit -m "Auto: $dataStr" 2>&1
    Write-Host "  git commit: $commitOut"
    $pushOut = git push 2>&1
    Write-Host "  git push: $pushOut"
    Write-Host "[$(Get-Date -f 'HH:mm:ss')] OK: https://exactus-data-ia.github.io/icc-inadimplencia/"
}


