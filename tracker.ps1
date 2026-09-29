# ============================================================
# 可转债自动跟踪脚本 tracker.ps1
# 由 Windows 计划任务每 5 分钟运行一次（也可双击 run-now.bat 手动运行）
# 功能：抓取东财可转债实时行情 -> 过滤 -> 综合评分 ->
#       保存盘中快照 -> 生成网页看板 dashboard.html -> 收盘后冻结当日最终名单
# 数据来源：东方财富公开行情接口（可转债板块 MK0354）
# ============================================================
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# PowerShell 5.1 需显式启用 TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$RootDir  = $PSScriptRoot
$DataDir  = Join-Path $RootDir 'data'
$SnapDir  = Join-Path $DataDir 'snapshots'
$FinalDir = Join-Path $DataDir 'final'
$LogFile  = Join-Path $DataDir 'runlog.txt'
$DashFile = Join-Path $RootDir $(if ($env:CB_DASH_FILE) { $env:CB_DASH_FILE } else { 'dashboard.html' })
$PyInitFile = Join-Path $RootDir 'pyinit.txt'   # 拼音首字母码表（搜索功能用，U+4E00 起按码点索引）

. (Join-Path $RootDir 'config.ps1')

foreach ($d in @($DataDir, $SnapDir, $FinalDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:Curl = if (Get-Command curl.exe -ErrorAction SilentlyContinue) { 'curl.exe' } else { 'curl' }

function Write-Log([string]$Msg) {
    $line = ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Msg)
    [System.IO.File]::AppendAllText($LogFile, $line + "`r`n", $Utf8NoBom)
    Write-Output $line
    try {
        if ((Get-Item $LogFile).Length -gt 512KB) {
            $tail = [System.IO.File]::ReadAllLines($LogFile, $Utf8NoBom) | Select-Object -Last 2000
            [System.IO.File]::WriteAllLines($LogFile, $tail, $Utf8NoBom)
        }
    } catch { }
}

function Convert-Num($v) {
    if ($null -eq $v -or $v -eq '-') { return $null }
    try { return [double]$v } catch { return $null }
}

function Format-Num($v, $digits) {
    if ($null -eq $v) { return '' }
    return ([double]$v).ToString('F' + $digits, [Globalization.CultureInfo]::InvariantCulture)
}

function Get-Phase {
    $now = Get-Date
    if ($now.DayOfWeek -eq [System.DayOfWeek]::Saturday -or $now.DayOfWeek -eq [System.DayOfWeek]::Sunday) { return 'weekend' }
    $t = $now.TimeOfDay
    $am1   = [TimeSpan]::FromMinutes(555)   # 09:15
    $am2   = [TimeSpan]::FromMinutes(695)   # 11:35
    $pm1   = [TimeSpan]::FromMinutes(775)   # 12:55
    $pm2   = [TimeSpan]::FromMinutes(910)   # 15:10
    $close = [TimeSpan]::FromMinutes(905)   # 15:05
    if (($t -ge $am1 -and $t -le $am2) -or ($t -ge $pm1 -and $t -le $pm2)) { return 'intraday' }
    if ($t -ge $close) { return 'afterclose' }
    return 'early'
}

$script:UsedHost = $null
$script:UsedVia  = $null

function Get-QuoteDiff {
    # 备用行情源单页最多返回 100 条，需分页抓取全部可转债
    $pageSize = 100
    $best = @()
    foreach ($h in $config.Hosts) {
        $all = @()
        $total = 0
        $ok = $false
        for ($pn = 1; $pn -le 20; $pn++) {
            $path = ('/api/qt/clist/get?pn={0}&pz={1}&po=1&np=1&fltt=2&invt=2&fid=f6&fs=b:MK0354&fields={2}' -f $pn, $pageSize, $config.Fields)
            $url = $h + $path
            $page = $null
            try {
                $resp = Invoke-RestMethod -Uri $url -TimeoutSec 6 -UserAgent 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
                if ($resp -and $resp.data) {
                    $total = [int]$resp.data.total
                    $page = @($resp.data.diff)
                }
            } catch {
                Write-Log ('抓取失败 [{0}] 第{1}页 (IRM): {2}' -f $h, $pn, $_.Exception.Message)
            }
            if (-not $page -or $page.Count -eq 0) {
                $tmpF = Join-Path ([IO.Path]::GetTempPath()) ('cbq-' + [Guid]::NewGuid().ToString('N') + '.json')
                try {
                    & $script:Curl -s -m 10 -A 'Mozilla/5.0' $url -o $tmpF 2>$null
                    if ((Test-Path $tmpF) -and ((Get-Item $tmpF).Length -gt 0)) {
                        $json = ([System.IO.File]::ReadAllText($tmpF, [System.Text.Encoding]::UTF8)) -replace '^\uFEFF', ''
                        $obj = $json | ConvertFrom-Json
                        if ($obj -and $obj.data) {
                            $total = [int]$obj.data.total
                            $page = @($obj.data.diff)
                        }
                    }
                } catch {
                    Write-Log ('抓取失败 [{0}] 第{1}页 (curl): {2}' -f $h, $pn, $_.Exception.Message)
                } finally {
                    if (Test-Path $tmpF) { Remove-Item $tmpF -Force -ErrorAction SilentlyContinue }
                }
            }
            if (-not $page -or $page.Count -eq 0) { break }
            $all += $page
            if ($total -gt 0 -and $all.Count -ge $total) { $ok = $true; break }
        }
        if ($all.Count -gt 0) {
            $script:UsedHost = $h
            if ($ok) { return $all }
            if ($all.Count -gt $best.Count) { $best = $all }  # 记录最佳部分结果
            Write-Log ('数据源 [{0}] 仅返回部分数据 {1}/{2}' -f $h, $all.Count, $total)
        }
    }
    if ($best.Count -gt 0) { return $best }
    return $null
}

function New-BoardItems($List) {
    $idx = 0
    @($List | ForEach-Object {
        $idx++
        [PSCustomObject]@{
            rank     = $idx
            code     = $_.Code
            name     = $_.Name
            price    = $_.Price
            pct      = $_.Pct
            speed    = $_.Speed
            amount   = $_.Amount
            turnover = $_.Turnover
            rankAmt  = $_.RankAmt
            rankSpd  = $_.RankSpd
            score    = $_.Score
            flow5     = $_.Flow5
            flowPct5  = $_.FlowPct5
            flowToday = $_.FlowToday
            flowXL    = $_.FlowXL
            flowL     = $_.FlowL
            flowM     = $_.FlowM
            flowS     = $_.FlowS
            volRatio  = $_.VolRatio
            amplitude = $_.Amplitude
        }
    })
}

function Write-Dashboard($payload, $statusText, $srcLabel, $cfgJson) {
    $template = @'
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<meta http-equiv="refresh" content="60">
<title>可转债跟踪工作台</title>
<link rel="manifest" href="./manifest.webmanifest">
<meta name="theme-color" content="#2f6fed">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="可转债">
<link rel="apple-touch-icon" href="./icon-192.png">
<link rel="icon" href="./icon-192.png" type="image/png">
<style>
:root{--up:#e03131;--down:#0ca678;--bg:#f2f5f9;--card:#ffffff;--border:#e4eaf1;--text:#222b36;--muted:#7b8694;--accent:#2f6fed;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Segoe UI","Microsoft YaHei",sans-serif;background:var(--bg);color:var(--text);padding:14px;}
.wrap{max-width:1100px;margin:0 auto;}
header{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:16px 20px;margin-bottom:12px;}
h1{font-size:19px;font-weight:700;}
.meta{display:flex;flex-wrap:wrap;gap:6px 16px;margin-top:10px;font-size:13px;color:var(--muted);align-items:center;}
.badge{padding:2px 10px;border-radius:99px;font-size:12px;color:#fff;font-weight:600;}
.status-live{background:#f08c00;}
.status-closed{background:#94a3b8;}
.tabs{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;}
.tab{border:1px solid var(--border);background:var(--card);padding:7px 16px;border-radius:8px;cursor:pointer;font-size:14px;color:var(--muted);}
.tab:hover{color:var(--accent);border-color:var(--accent);}
.tab.active{background:var(--accent);border-color:var(--accent);color:#fff;}
.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:12px;overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:13px;white-space:nowrap;}
th,td{padding:7px 10px;text-align:right;border-bottom:1px solid var(--border);}
th{color:var(--muted);font-weight:600;background:#f8fafc;}
th.l,td.l{text-align:left;}
tr:hover td{background:#f4f8ff;}
.up{color:var(--up);}
.down{color:var(--down);}
.rank{display:inline-block;min-width:22px;text-align:center;font-weight:700;}
.rank.top{background:var(--accent);color:#fff;border-radius:4px;}
.score{font-weight:700;}
.hist-date{font-weight:700;margin:14px 0 4px;font-size:14px;}
.hist-date:first-child{margin-top:0;}
.note{font-size:12px;color:var(--muted);margin-top:6px;line-height:1.7;}
footer{text-align:center;font-size:12px;color:var(--muted);margin:6px 0 20px;line-height:1.9;}
.refresh{border:1px solid var(--border);background:var(--card);padding:3px 12px;border-radius:6px;cursor:pointer;font-size:12px;color:var(--muted);}
.refresh:hover{color:var(--accent);border-color:var(--accent);}
.cards{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px;}
.stat{flex:1 1 130px;background:var(--card);border:1px solid var(--border);border-radius:10px;padding:10px 14px;text-align:center;}
.stat-v{font-size:16px;font-weight:700;}
.stat-l{font-size:12px;color:var(--muted);margin-top:3px;}
.status-strip{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:8px 16px;margin-bottom:12px;font-size:13px;color:var(--muted);display:flex;flex-wrap:wrap;gap:4px 10px;align-items:center;}
.status-strip a{color:var(--accent);text-decoration:none;}
.logbox{font-size:11px;}
.logbox summary{cursor:pointer;color:var(--accent);}
.logbox pre{background:#f8fafc;padding:8px;border-radius:6px;overflow-x:auto;white-space:pre-wrap;}
.topsearch{margin-top:12px;}
#top-search-input{width:100%;padding:11px 14px;border:2px solid var(--accent);border-radius:10px;font-size:16px;font-weight:600;color:var(--text);background:#f0f5ff;outline:none;box-shadow:0 1px 4px rgba(47,111,237,.15);}
#top-search-input::placeholder{color:#9db4dd;font-weight:400;}
#top-search-input:focus{border-color:#1d5bd6;background:#fff;box-shadow:0 0 0 3px rgba(47,111,237,.18);}
.search-hint{font-size:12px;color:var(--muted);margin-top:6px;}
#top-search-result{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:12px;overflow-x:auto;}
</style>
</head>
<body>
<div class="wrap">
<header>
<h1>可转债跟踪工作台</h1>
<div class="meta">
<span>更新时间：<b>%%GENERATED_AT%%</b></span>
<span class="badge %%STATUS_CLASS%%">%%STATUS_TEXT%%</span>
<span>%%SOURCE%%</span>
<span id="cnt"></span>
<span>每 60 秒自动刷新</span>
<button class="refresh" onclick="location.reload()">立即刷新</button>
</div>
<div class="note" id="cfg-note"></div>
<div class="topsearch">
<input id="top-search-input" type="text" placeholder="搜索转债：代码 / 名称 / 拼音首字母（如 123261 / 华峰 / hf）">
<div class="search-hint">支持拼音首字母连打（如 hf、hfzz → 华峰转债）；按 / 键快速聚焦搜索框</div>
</div>
</header>
<div class="card" id="top-search-result" style="display:none"></div>
<div class="cards">
<div class="stat"><div class="stat-v" id="ov-up">-</div><div class="stat-l">上涨家数</div></div>
<div class="stat"><div class="stat-v" id="ov-down">-</div><div class="stat-l">下跌家数</div></div>
<div class="stat"><div class="stat-v" id="ov-total">-</div><div class="stat-l">全市场总成交额</div></div>
<div class="stat"><div class="stat-v" id="ov-avg">-</div><div class="stat-l">平均涨跌幅</div></div>
<div class="stat"><div class="stat-v" id="ov-maxup">-</div><div class="stat-l">涨幅最大</div></div>
<div class="stat"><div class="stat-v" id="ov-maxdown">-</div><div class="stat-l">跌幅最大</div></div>
<div class="stat"><div class="stat-v" id="ov-flow5">-</div><div class="stat-l">5日主力净流入(合计)</div></div>
<div class="stat"><div class="stat-v" id="ov-flowtoday">-</div><div class="stat-l">今日主力净流入(合计)</div></div>
</div>
<div class="status-strip" id="status-strip"></div>
<div class="tabs">
<button class="tab active" data-tab="composite">综合榜</button>
<button class="tab" data-tab="amount">成交额榜</button>
<button class="tab" data-tab="speed">涨速榜</button>
<button class="tab" data-tab="active">活跃榜</button>
<button class="tab" data-tab="flow">资金流向</button>
</div>
<div class="card" id="composite"></div>
<div class="card" id="amount" style="display:none"></div>
<div class="card" id="speed" style="display:none"></div>
<div class="card" id="flow" style="display:none"></div>
<div class="card" id="active" style="display:none"></div>
<footer>
数据来源：东方财富公开行情接口（可转债板块）。涨速为东财口径（最近几分钟价格变化速度），行情可能存在延迟。<br>
本看板仅供个人学习参考，不构成任何投资建议。股市有风险，投资需谨慎。
</footer>
</div>
<script>
var DATA = %%DATA%%;
var CFG = %%CFG%%;
(function () {
  'use strict';
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    window.addEventListener('load', function () {
      try { navigator.serviceWorker.register('./sw.js', { updateViaCache: 'none' }); } catch (e) {}
    });
  }
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function num(v, d) {
    if (v == null || v === '-') { return '-'; }
    return (+v).toFixed(d == null ? 2 : d);
  }
  function pct(v) {
    if (v == null || v === '-') { return '-'; }
    v = +v;
    return (v > 0 ? '+' : '') + v.toFixed(2) + '%';
  }
  function turn(v) {
    if (v == null || v === '-') { return '-'; }
    return (+v).toFixed(2) + '%';
  }
  function cls(v) { v = +v; return v > 0 ? 'up' : (v < 0 ? 'down' : ''); }
  function amt(v) {
    if (v == null || v === '-') { return '-'; }
    v = +v;
    var neg = v < 0 ? '-' : '';
    v = Math.abs(v);
    if (v >= 1e8) { return neg + (v / 1e8).toFixed(2) + '亿'; }
    if (v >= 1e4) { return neg + (v / 1e4).toFixed(0) + '万'; }
    return neg + Math.round(v);
  }
  function rankHtml(n) {
    var c = (n >= 1 && n <= 3) ? 'rank top' : 'rank';
    return '<span class="' + c + '">' + n + '</span>';
  }
  function head(cols) {
    return '<thead><tr>' + cols.map(function (c) { return '<th>' + c + '</th>'; }).join('') + '</tr></thead>';
  }
  function body(items, fn) {
    return '<tbody>' + items.map(fn).join('') + '</tbody>';
  }
  function compRow(it) {
    return '<tr>' +
      '<td class="l">' + rankHtml(it.rank) + '</td>' +
      '<td class="l">' + esc(it.code) + '</td>' +
      '<td class="l">' + esc(it.name) + '</td>' +
      '<td>' + num(it.price, 3) + '</td>' +
      '<td class="' + cls(it.pct) + '">' + pct(it.pct) + '</td>' +
      '<td class="' + cls(it.speed) + '">' + pct(it.speed) + '</td>' +
      '<td class="' + cls(it.flowToday) + '">' + amt(it.flowToday) + '</td>' +
      '<td class="' + cls(it.flow5) + '">' + amt(it.flow5) + '</td>' +
      '<td>' + it.rankAmt + '</td>' +
      '<td>' + it.rankSpd + '</td>' +
      '<td class="score">' + num(it.score, 1) + '</td>' +
      '</tr>';
  }
  function simpleRow(it) {
    return '<tr>' +
      '<td class="l">' + rankHtml(it.rank) + '</td>' +
      '<td class="l">' + esc(it.code) + '</td>' +
      '<td class="l">' + esc(it.name) + '</td>' +
      '<td>' + num(it.price, 3) + '</td>' +
      '<td class="' + cls(it.pct) + '">' + pct(it.pct) + '</td>' +
      '<td class="' + cls(it.speed) + '">' + pct(it.speed) + '</td>' +
      '<td>' + amt(it.amount) + '</td>' +
      '<td>' + turn(it.turnover) + '</td>' +
      '<td class="score">' + num(it.score, 1) + '</td>' +
      '</tr>';
  }
  var compCols = ['排名', '代码', '名称', '最新价', '涨跌幅', '涨速', '今日主力净流入', '5日主力净流入', '成交额排名', '涨速排名', '综合分'];
  var simpleCols = ['排名', '代码', '名称', '最新价', '涨跌幅', '涨速', '成交额', '换手率', '综合分'];
  var composite = [].concat(DATA.composite || []);
  var amount = [].concat(DATA.amount || []);
  var speed = [].concat(DATA.speed || []);
  document.getElementById('composite').innerHTML = '<table>' + head(compCols) + body(composite, compRow) + '</table>';
  document.getElementById('amount').innerHTML = '<table>' + head(simpleCols) + body(amount, simpleRow) + '</table>';
  document.getElementById('speed').innerHTML = '<table>' + head(simpleCols) + body(speed, simpleRow) + '</table>';
  var flow = [].concat(DATA.flow || []);
  var flowOut = [].concat(DATA.flowOut || []);
  var flowCols = ['排名', '代码', '名称', '最新价', '涨跌幅', '5日主力净流入', '5日净占比', '今日主力净流入', '今日超大单', '今日大单', '今日中单', '今日小单'];
  function flowRow(it) {
    return '<tr>' +
      '<td class="l">' + rankHtml(it.rank) + '</td>' +
      '<td class="l">' + esc(it.code) + '</td>' +
      '<td class="l">' + esc(it.name) + '</td>' +
      '<td>' + num(it.price, 3) + '</td>' +
      '<td class="' + cls(it.pct) + '">' + pct(it.pct) + '</td>' +
      '<td class="' + cls(it.flow5) + '">' + amt(it.flow5) + '</td>' +
      '<td class="' + cls(it.flowPct5) + '">' + pct(it.flowPct5) + '</td>' +
      '<td class="' + cls(it.flowToday) + '">' + amt(it.flowToday) + '</td>' +
      '<td class="' + cls(it.flowXL) + '">' + amt(it.flowXL) + '</td>' +
      '<td class="' + cls(it.flowL) + '">' + amt(it.flowL) + '</td>' +
      '<td class="' + cls(it.flowM) + '">' + amt(it.flowM) + '</td>' +
      '<td class="' + cls(it.flowS) + '">' + amt(it.flowS) + '</td>' +
      '</tr>';
  }
  var flowHtml = '';
  if (flow.length > 0) {
    flowHtml += '<div class="hist-date">5日主力净流入 Top ' + flow.length + '</div>';
    flowHtml += '<table>' + head(flowCols) + body(flow, flowRow) + '</table>';
  } else {
    flowHtml += '<div class="note">暂无资金流向数据（下一个交易日运行后自动出现）。</div>';
  }
  if (flowOut.length > 0) {
    flowHtml += '<div class="hist-date">5日主力净流出提醒 Top ' + flowOut.length + '</div>';
    flowHtml += '<table>' + head(flowCols) + body(flowOut, flowRow) + '</table>';
  }
  flowHtml += '<div class="note">5日主力净流入 = 最近 5 个交易日主力资金（超大单+大单）净流入，东方财富资金流向口径；红色为净流入、绿色为净流出；超大单/大单/中单/小单列为今日净流入。</div>';
  document.getElementById('flow').innerHTML = flowHtml;
  var active = [].concat(DATA.active || []);
  var activeCols = ['排名', '代码', '名称', '量比', '振幅', '换手率', '涨跌幅', '成交额', '涨速'];
  function activeRow(it) {
    return '<tr>' +
      '<td class="l">' + rankHtml(it.rank) + '</td>' +
      '<td class="l">' + esc(it.code) + '</td>' +
      '<td class="l">' + esc(it.name) + '</td>' +
      '<td>' + num(it.volRatio, 2) + '</td>' +
      '<td>' + pct(it.amplitude) + '</td>' +
      '<td>' + turn(it.turnover) + '</td>' +
      '<td class="' + cls(it.pct) + '">' + pct(it.pct) + '</td>' +
      '<td>' + amt(it.amount) + '</td>' +
      '<td class="' + cls(it.speed) + '">' + pct(it.speed) + '</td>' +
      '</tr>';
  }
  if (active.length > 0) {
    document.getElementById('active').innerHTML = '<table>' + head(activeCols) + body(active, activeRow) + '</table>' +
      '<div class="note">量比 = 当日每分钟平均成交量 ÷ 过去 5 个交易日每分钟平均成交量，大于 1 表示比近期更活跃；振幅 =（最高 − 最低）÷ 昨收。</div>';
  } else {
    document.getElementById('active').innerHTML = '<div class="note">暂无活跃度数据（下一个交易日运行后自动出现）。</div>';
  }
  document.getElementById('cnt').textContent = '共 ' + DATA.total + ' 只 · 有效 ' + DATA.valid + ' 只';
  document.getElementById('cfg-note').innerHTML =
    '综合分 = 成交额得分×' + Math.round(CFG.weightAmount * 100) + '% + 涨速得分×' + Math.round(CFG.weightSpeed * 100) + '%；' +
    '成交额 ≥ ' + amt(CFG.minAmount) + ' 的转债计入综合榜；涨速按 ±' + CFG.speedClamp + '% 截断归一化。' +
    '（上市首日、停牌、无成交的转债不参与排名）';
  // ===== 工作台：市场总览 =====
  if (DATA.overview) {
    var ov = DATA.overview;
    document.getElementById('ov-up').textContent = ov.upCount + ' 只';
    document.getElementById('ov-down').textContent = ov.downCount + ' 只';
    document.getElementById('ov-up').className = 'stat-v up';
    document.getElementById('ov-down').className = 'stat-v down';
    document.getElementById('ov-total').textContent = amt(ov.totalAmt);
    var avgEl = document.getElementById('ov-avg');
    avgEl.textContent = (ov.avgPct > 0 ? '+' : '') + ov.avgPct + '%';
    avgEl.className = 'stat-v ' + cls(ov.avgPct);
    var upEl = document.getElementById('ov-maxup');
    upEl.textContent = esc(ov.maxGainName) + ' +' + ov.maxGainPct + '%';
    upEl.className = 'stat-v up';
    var dnEl = document.getElementById('ov-maxdown');
    dnEl.textContent = esc(ov.maxLossName) + ' ' + ov.maxLossPct + '%';
    dnEl.className = 'stat-v down';
    var f5El = document.getElementById('ov-flow5');
    f5El.textContent = amt(ov.flow5Total);
    f5El.className = 'stat-v ' + cls(ov.flow5Total);
    var ftEl = document.getElementById('ov-flowtoday');
    ftEl.textContent = amt(ov.flowTodayTotal);
    ftEl.className = 'stat-v ' + cls(ov.flowTodayTotal);
  }

  // ===== 工作台：平台状态栏 =====
  (function () {
    var st = DATA.status || {};
    var html = '<span class="badge ' + (DATA.marketStatus === '交易中' ? 'status-live' : 'status-closed') + '">' + esc(DATA.marketStatus) + '</span>';
    html += ' 数据更新 ' + esc(DATA.generatedAt);
    if (!navigator.onLine) { html += ' · <b style="color:#f08c00">离线数据</b>'; }
    if (st.source) { html += ' · ' + esc(st.source); }
    html += ' · 手机发布：' + (st.lastPublishOk ? (esc(st.lastPublish) + ' 成功') : '尚未发布');
    if (st.siteUrl) { html += ' · 网址 <a href="' + esc(st.siteUrl) + '" target="_blank" rel="noopener">' + esc(st.siteUrl) + '</a>'; }
    if (st.logTail && st.logTail.length) {
      html += ' <details class="logbox"><summary>最近日志</summary><pre>' + esc(st.logTail.join('\n')) + '</pre></details>';
    }
    document.getElementById('status-strip').innerHTML = html;
  })();

  // ===== 工作台：首页顶部全市场搜索（代码 / 名称 / 拼音首字母） =====
  var allItems = [].concat(DATA.all || []);
  var PY_INIT = "%%PY_INIT%%";   // 拼音首字母码表，U+4E00 起按码点索引（由 tracker.ps1 注入）
  var PY_BASE = 0x4E00;
  function pyInitials(s) {
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      var d = c - PY_BASE;
      var ch = (d >= 0 && d < PY_INIT.length) ? PY_INIT.charAt(d) : '';
      out += (ch >= 'a' && ch <= 'z') ? ch : s.charAt(i).toLowerCase();
    }
    return out;
  }
  (function () {
    var input = document.getElementById('top-search-input');
    var box = document.getElementById('top-search-result');
    var cols = ['代码', '名称', '拼音缩写', '最新价', '涨跌幅', '涨速', '成交额', '换手率', '成交额排名', '涨速排名', '综合分'];
    function doSearch() {
      var q = (input.value || '').trim().toLowerCase();
      if (!q) { box.style.display = 'none'; box.innerHTML = ''; return; }
      var hits = [];
      for (var i = 0; i < allItems.length; i++) {
        var it = allItems[i];
        var py = pyInitials(it.name);
        if (it.code.indexOf(q) >= 0 || it.name.toLowerCase().indexOf(q) >= 0 || py.indexOf(q) >= 0) {
          hits.push({ it: it, py: py });
          if (hits.length >= 30) { break; }
        }
      }
      if (!hits.length) {
        box.style.display = '';
        box.innerHTML = '<div class="note">没有找到匹配的转债（试试代码 / 名称 / 拼音首字母，如 123261 / 华峰 / hf）</div>';
        return;
      }
      var rows = hits.map(function (h) {
        var it = h.it;
        return '<tr>' +
          '<td class="l">' + esc(it.code) + '</td>' +
          '<td class="l">' + esc(it.name) + '</td>' +
          '<td class="l">' + esc(h.py) + '</td>' +
          '<td>' + num(it.price, 3) + '</td>' +
          '<td class="' + cls(it.pct) + '">' + pct(it.pct) + '</td>' +
          '<td class="' + cls(it.speed) + '">' + pct(it.speed) + '</td>' +
          '<td>' + amt(it.amount) + '</td>' +
          '<td>' + turn(it.turnover) + '</td>' +
          '<td>' + it.rankAmt + '</td>' +
          '<td>' + it.rankSpd + '</td>' +
          '<td class="score">' + num(it.score, 1) + '</td>' +
          '</tr>';
      }).join('');
      box.style.display = '';
      box.innerHTML = '<table>' + head(cols) + '<tbody>' + rows + '</tbody></table>' +
        '<div class="note">显示前 ' + hits.length + ' 条匹配（最多 30 条），拼音首字母可连打，如 hf、hfzz</div>';
    }
    input.addEventListener('input', function () {
      try { sessionStorage.setItem('cbq', input.value); } catch (e) {}
      doSearch();
    });
    var saved = '';
    try { saved = sessionStorage.getItem('cbq') || ''; } catch (e) {}
    if (saved) { input.value = saved; doSearch(); }
    document.addEventListener('keydown', function (e) {
      var tag = (document.activeElement && document.activeElement.tagName) || '';
      if (e.key === '/' && document.activeElement !== input && !/input|textarea/i.test(tag)) {
        e.preventDefault();
        input.focus();
      }
    });
  })();

  var tabs = document.querySelectorAll('.tab');
  var ids = ['composite', 'amount', 'speed', 'active', 'flow'];
  for (var k = 0; k < tabs.length; k++) {
    tabs[k].addEventListener('click', function () {
      for (var j = 0; j < tabs.length; j++) { tabs[j].classList.remove('active'); }
      this.classList.add('active');
      var tab = this.getAttribute('data-tab');
      for (var m = 0; m < ids.length; m++) {
        document.getElementById(ids[m]).style.display = (ids[m] === tab) ? '' : 'none';
      }
    });
  }
})();
</script>
</body>
</html>
'@
    $dataJson = ($payload | ConvertTo-Json -Depth 8).Replace('</', '<\/')
    $statusClass = if ($statusText -eq '交易中') { 'status-live' } else { 'status-closed' }
    $html = $template.Replace('%%DATA%%', $dataJson)
    $html = $html.Replace('%%CFG%%', $cfgJson)
    $html = $html.Replace('%%GENERATED_AT%%', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $html = $html.Replace('%%STATUS_TEXT%%', $statusText)
    $html = $html.Replace('%%STATUS_CLASS%%', $statusClass)
    $html = $html.Replace('%%SOURCE%%', $srcLabel)
    $pyInit = ''
    if (Test-Path $PyInitFile) {
        $pyInit = [System.IO.File]::ReadAllText($PyInitFile, [System.Text.Encoding]::ASCII).TrimEnd()
    }
    $html = $html.Replace('%%PY_INIT%%', $pyInit)
    [System.IO.File]::WriteAllText($DashFile, $html, $Utf8NoBom)
}

# ============================================================
# 主流程
# ============================================================
$phase = Get-Phase
if ($phase -eq 'weekend' -or $phase -eq 'early') { exit 0 }

$today = Get-Date -Format 'yyyy-MM-dd'
$finalFile = Join-Path $FinalDir ($today + '.json')

$diff = Get-QuoteDiff
if (-not $diff -or $diff.Count -eq 0) {
    Write-Log '抓取失败：所有数据源均不可用'
    exit 1
}

# 行情时间戳（f124，unix 秒）用于交易日 / 休市自检
$maxTs = 0
foreach ($d in $diff) {
    $ts = Convert-Num $d.f124
    if ($ts -and $ts -gt $maxTs) { $maxTs = [int64]$ts }
}
$dataDate = $null
$ageMin = -1
if ($maxTs -gt 0) {
    $dataTime = [DateTimeOffset]::FromUnixTimeSeconds($maxTs).ToLocalTime().DateTime
    $dataDate = $dataTime.ToString('yyyy-MM-dd')
    $ageMin = ((Get-Date) - $dataTime).TotalMinutes
}

if ($phase -eq 'intraday') {
    if ($dataDate -ne $today -or $ageMin -gt [double]$config.StaleMinutes) {
        Write-Log '休市：行情未更新（节假日或非交易时段），跳过'
        exit 0
    }
    $statusText = '交易中'
} else {
    # 已收盘：仅当今日最终名单尚未冻结时执行，冻结后后续运行直接跳过
    if ($dataDate -ne $today) { Write-Log '休市：最新行情非今天（节假日），跳过'; exit 0 }
    if (Test-Path $finalFile) { exit 0 }
    $statusText = '已收盘'
}

# ---------- 过滤 ----------
$records = @()
foreach ($d in $diff) {
    $name = [string]$d.f14
    if (-not $name) { continue }
    if ($name.StartsWith('N')) { continue }          # 上市首日（N 前缀）
    $prev = Convert-Num $d.f18
    if ($prev -and $prev -eq 100) { continue }       # 上市首日（昨收=发行价 100）
    $amt = Convert-Num $d.f6
    if (-not $amt -or $amt -le 0) { continue }       # 停牌 / 无成交
    $speed = Convert-Num $d.f22
    if ($null -eq $speed) { continue }               # 涨速缺失
    $records += [PSCustomObject]@{
        Code     = [string]$d.f12
        Name     = $name
        Price    = Convert-Num $d.f2
        Pct      = Convert-Num $d.f3
        Speed    = $speed
        Amount   = $amt
        Turnover = Convert-Num $d.f8
        High     = Convert-Num $d.f15
        Low      = Convert-Num $d.f16
        Open     = Convert-Num $d.f17
        Prev     = $prev
        Flow5    = Convert-Num $d.f164
        FlowPct5 = Convert-Num $d.f165
        FlowToday = Convert-Num $d.f62
        FlowXL  = Convert-Num $d.f66
        FlowL   = Convert-Num $d.f72
        FlowM   = Convert-Num $d.f78
        FlowS   = Convert-Num $d.f84
        VolRatio  = Convert-Num $d.f10
        Amplitude = Convert-Num $d.f7
    }
}
if ($records.Count -eq 0) {
    Write-Log '过滤后无有效数据'
    exit 1
}

# ---------- 综合评分 ----------
$wAmt   = [double]$config.WeightAmount
$wSpd   = [double]$config.WeightSpeed
$minAmt = [double]$config.MinAmount
$clamp  = [double]$config.SpeedClamp
$topN   = [int]$config.TopN

$logAmts = @($records | ForEach-Object { [Math]::Log10($_.Amount) })
$minLog = ($logAmts | Measure-Object -Minimum).Minimum
$maxLog = ($logAmts | Measure-Object -Maximum).Maximum
$rangeLog = $maxLog - $minLog
if ($rangeLog -le 0) { $rangeLog = 1 }

$spdAll = @($records | ForEach-Object { [Math]::Max(-$clamp, [Math]::Min($clamp, $_.Speed)) })
$minSpd = ($spdAll | Measure-Object -Minimum).Minimum
$maxSpd = ($spdAll | Measure-Object -Maximum).Maximum
$rangeSpd = $maxSpd - $minSpd
if ($rangeSpd -le 0) { $rangeSpd = 1 }

foreach ($r in $records) {
    $amtScore = ([Math]::Log10($r.Amount) - $minLog) / $rangeLog * 100
    $clamped  = [Math]::Max(-$clamp, [Math]::Min($clamp, $r.Speed))
    $spdScore = ($clamped - $minSpd) / $rangeSpd * 100
    $r | Add-Member -NotePropertyName AmtScore -NotePropertyValue ([Math]::Round($amtScore, 1))
    $r | Add-Member -NotePropertyName SpdScore -NotePropertyValue ([Math]::Round($spdScore, 1))
    $r | Add-Member -NotePropertyName Score    -NotePropertyValue ([Math]::Round($wAmt * $amtScore + $wSpd * $spdScore, 1))
    $r | Add-Member -NotePropertyName Eligible -NotePropertyValue ($r.Amount -ge $minAmt)
}

$sortedAmt = @($records | Sort-Object -Property Amount -Descending)
$sortedSpd = @($records | Sort-Object -Property Speed  -Descending)
$rankAmt = @{}; $rankSpd = @{}
for ($i = 0; $i -lt $sortedAmt.Count; $i++) { $rankAmt[$sortedAmt[$i].Code] = $i + 1 }
for ($i = 0; $i -lt $sortedSpd.Count; $i++) { $rankSpd[$sortedSpd[$i].Code] = $i + 1 }
foreach ($r in $records) {
    $r | Add-Member -NotePropertyName RankAmt -NotePropertyValue ([int]$rankAmt[$r.Code])
    $r | Add-Member -NotePropertyName RankSpd -NotePropertyValue ([int]$rankSpd[$r.Code])
}

$compBoard = @($records | Where-Object { $_.Eligible } | Sort-Object -Property Score -Descending | Select-Object -First $topN)
$amtBoard  = @($sortedAmt | Select-Object -First $topN)
$spdBoard  = @($sortedSpd | Select-Object -First $topN)
$compItems = New-BoardItems $compBoard
$amtItems  = New-BoardItems $amtBoard
$spdItems  = New-BoardItems $spdBoard

# ---------- 资金流向榜（东财资金流向口径：主力 = 超大单 + 大单） ----------
$flowBoard    = @($records | Sort-Object -Property Flow5 -Descending | Select-Object -First $topN)
$flowOutBoard = @($records | Where-Object { $_.Flow5 -lt 0 } | Sort-Object -Property Flow5 | Select-Object -First 10)
$flowItems    = New-BoardItems $flowBoard
$flowOutItems = New-BoardItems $flowOutBoard

# ---------- 活跃榜（量比排序，振幅/换手率为辅助观察） ----------
$activeBoard = @($records | Sort-Object -Property VolRatio -Descending | Select-Object -First $topN)
$activeItems = New-BoardItems $activeBoard

# ---------- 收盘后：冻结当日最终名单 ----------
if ($phase -eq 'afterclose') {
    $finalObj = [PSCustomObject]@{
        date        = $today
        generatedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        total       = $diff.Count
        valid       = $records.Count
        board       = $compItems
    }
    [System.IO.File]::WriteAllText($finalFile, ($finalObj | ConvertTo-Json -Depth 8), $Utf8NoBom)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('排名,代码,名称,最新价,涨跌幅%,涨速%,成交额(元),换手率%,成交额排名,涨速排名,综合分')
    foreach ($it in $compItems) {
        $vals = @($it.rank, $it.code, $it.name,
            (Format-Num $it.price 3), (Format-Num $it.pct 2), (Format-Num $it.speed 2),
            [Math]::Round([double]$it.amount, 0), (Format-Num $it.turnover 2),
            $it.rankAmt, $it.rankSpd, (Format-Num $it.score 1))
        [void]$sb.AppendLine(($vals -join ','))
    }
    $Utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Join-Path $FinalDir ($today + '.csv')), $sb.ToString(), $Utf8Bom)
    Write-Log ('已冻结当日收盘最终名单 {0}（Top {1}，有效 {2} 只）' -f $today, $compItems.Count, $records.Count)
}

# ---------- 输出 ----------
$srcLabel = if ($script:UsedHost -like '*push2delay*') { '延时行情 push2delay.eastmoney.com（约滞后几分钟）' } elseif ($script:UsedHost) { '实时行情 ' + ([Uri]$script:UsedHost).Host } else { '未知' }

# ---------- 工作台：市场总览 ----------
$upCount   = @($records | Where-Object { $_.Pct -gt 0 }).Count
$downCount = @($records | Where-Object { $_.Pct -lt 0 }).Count
$flatCount = $records.Count - $upCount - $downCount
$totalAmt  = ($records | Measure-Object -Property Amount -Sum).Sum
$avgPct    = ($records | Measure-Object -Property Pct -Average).Average
$maxGain   = $records | Sort-Object -Property Pct -Descending | Select-Object -First 1
$maxLoss   = $records | Sort-Object -Property Pct | Select-Object -First 1
$flow5Sum  = 0.0
$flowTodaySum = 0.0
foreach ($r in $records) {
    if ($r.Flow5)    { $flow5Sum += [double]$r.Flow5 }
    if ($r.FlowToday) { $flowTodaySum += [double]$r.FlowToday }
}
$overview = [PSCustomObject]@{
    upCount     = $upCount
    downCount   = $downCount
    flatCount   = $flatCount
    totalAmt    = [Math]::Round($totalAmt, 0)
    avgPct      = [Math]::Round($avgPct, 2)
    maxGainName = $maxGain.Name
    maxGainPct  = $maxGain.Pct
    maxGainCode = $maxGain.Code
    maxLossName = $maxLoss.Name
    maxLossPct  = $maxLoss.Pct
    maxLossCode = $maxLoss.Code
    flow5Total  = [Math]::Round($flow5Sum, 0)
    flowTodayTotal = [Math]::Round($flowTodaySum, 0)
}

# ---------- 工作台：平台状态 ----------
$siteUrl = ''
if (Test-Path (Join-Path $RootDir 'publish-config.ps1')) { . (Join-Path $RootDir 'publish-config.ps1') }
$lastPub = ''
$lastPubOk = $false
try {
    $ps = [System.IO.File]::ReadAllText((Join-Path $DataDir 'publish-state.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    $lastPub = [string]$ps.time
    $lastPubOk = $true
} catch { }
$logTail = @()
try {
    $logTail = @([System.IO.File]::ReadAllLines($LogFile, $Utf8NoBom) | Select-Object -Last 5 | ForEach-Object { [string]$_ })
} catch { }
$statusInfo = [PSCustomObject]@{
    generatedAt  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    marketStatus = $statusText
    source       = $srcLabel
    lastPublish  = $lastPub
    lastPublishOk = $lastPubOk
    siteUrl      = $siteUrl
    logTail      = $logTail
}

$allSorted = @($records | Sort-Object -Property Score -Descending)

$payload = [PSCustomObject]@{
    generatedAt  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    date         = $today
    marketStatus = $statusText
    source       = $srcLabel
    total        = $diff.Count
    valid        = $records.Count
    composite    = $compItems
    amount       = $amtItems
    speed        = $spdItems
    flow         = $flowItems
    flowOut      = $flowOutItems
    active       = $activeItems
    overview     = $overview
    all          = (New-BoardItems $allSorted)
    status       = $statusInfo
}
[System.IO.File]::WriteAllText((Join-Path $DataDir 'latest.json'), ($payload | ConvertTo-Json -Depth 8), $Utf8NoBom)

$daySnapDir = Join-Path $SnapDir $today
if (-not (Test-Path $daySnapDir)) { New-Item -ItemType Directory -Path $daySnapDir -Force | Out-Null }
$snapObj = [PSCustomObject]@{
    generatedAt  = $payload.generatedAt
    date         = $today
    marketStatus = $statusText
    source       = $srcLabel
    total        = $diff.Count
    valid        = $records.Count
    all          = (New-BoardItems $allSorted)
}
[System.IO.File]::WriteAllText((Join-Path $daySnapDir ((Get-Date -Format 'HHmm') + '.json')), ($snapObj | ConvertTo-Json -Depth 8), $Utf8NoBom)

$cfgObj = [PSCustomObject]@{
    weightAmount = $wAmt
    weightSpeed  = $wSpd
    minAmount    = $minAmt
    topN         = $topN
    speedClamp   = $clamp
}
$cfgJson = ($cfgObj | ConvertTo-Json -Compress).Replace('</', '<\/')

Write-Dashboard $payload $statusText $srcLabel $cfgJson

Write-Log ('更新完成 [{0}] 共 {1} 只 / 有效 {2} 只 / 数据源 {3}' -f $statusText, $diff.Count, $records.Count, $srcLabel)
exit 0
