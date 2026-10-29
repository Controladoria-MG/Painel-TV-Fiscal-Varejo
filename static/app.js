// Painel de TV — Escrita Fiscal Varejo: só placares e gráficos, SEM nome de responsável nem
// de cliente (a planilha tem essas colunas, mas nada aqui lê
// UsuarioResponsavel/UsuarioBaixa/RazaoSocial/CodCliente/Coordenador).

// ── Configuração deste painel ─────────────────────────────
// Departamento comparado sem acento (a planilha às vezes chega com o acento
// corrompido). O valor é o segmento exibido nos placares, ou null para não
// dividir (aí o quadro "Concluído no mês" divide por classificação).
const PAINEL = {
  nome: 'Escrita Fiscal Varejo',
  sub:  'Painel de tarefas',
  departamentos: { 'EF - VAREJO': null },
};

// Mesma planilha que o Controle de Tarefas lê no servidor do TI: quando a base
// é atualizada lá ("Atualizar base"), o painel pega sozinho — não tem base própria.
// Fora do servidor da MG (teste local) lê a cópia do GitHub, porque o servidor
// do TI não libera leitura a partir de outro endereço.
const NO_SERVIDOR_MG = location.hostname.endsWith('mgcontecnica.com.br');
const UNIDADE = {
  nome:    'São Paulo',
  arquivo: NO_SERVIDOR_MG
    ? 'https://tarefasctr.mgcontecnica.com.br/dados/base/sao_paulo.xlsx'
    : 'https://controladoria-mg.github.io/Portal-das-tarefas/dados/base/sao_paulo.xlsx',
};

const ATUALIZAR_DADOS_MIN   = 10;  // busca a planilha de novo
const RECARREGAR_PAGINA_H   = 6;   // recarga completa (evita acúmulo de memória na TV)

const COR = {
  marca:    '#C00000',
  aberto:   '#F87171',
  vencehoje:'#DC2626',
  atraso:   '#991B1B',
  trilho:   '#F0F1F5',
  texto:    '#1F2328',
  muted:    '#6B7280',
  grade:    '#ECEDF2',
};

let linhas = [];
let dataBase = '';
let mesRef = null;       // 1º dia do mês de vencimento que a base cobre
let jaCarregou = false;
const graficos = {};

// ── Utilitários ───────────────────────────────────────────
function semAcento(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\uFFFD/g, 'A').toUpperCase().trim();
}
function fmt(n) { return n.toLocaleString('pt-BR'); }
function pct(v, t) { return t ? Math.round((v / t) * 100) : 0; }
// "% do total" sem mostrar 0% quando ainda existe alguma tarefa no status.
// Concluído com 1 casa e arredondado pra baixo: com poucas pendências, "99%"
// ou "100%" escondem o que falta — 6.348 de 6.359 aparece como 99,8%.
function pctConcluido(v, t) { return t ? Math.floor((v / t) * 1000) / 10 : 0; }
function fmtPct(p) { return `${p === 100 || p === 0 ? p : p.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`; }
function pctTexto(v, t) { const p = pct(v, t); return v && !p ? '<1%' : `${p}%`; }
function normData(d) { const n = new Date(d); n.setHours(0, 0, 0, 0); return n; }
function rem(mult) { return parseFloat(getComputedStyle(document.documentElement).fontSize) * mult; }

// Mesmo critério do portal (app.js): "Em Aberto" que vence hoje vira "Vencendo Hoje".
function statusEfetivo(r, hoje) {
  if (r.Status === 'Em Aberto' && r.DataVencimento instanceof Date &&
      normData(r.DataVencimento).getTime() === hoje.getTime()) return 'Vencendo Hoje';
  return r.Status;
}

function contar(rows) {
  const c = { total: rows.length, baixado: 0, aberto: 0, vencehoje: 0, atraso: 0 };
  for (const r of rows) {
    if      (r.st === 'Baixado')       c.baixado++;
    else if (r.st === 'Em Aberto')     c.aberto++;
    else if (r.st === 'Vencendo Hoje') c.vencehoje++;
    else if (r.st === 'Em Atraso')     c.atraso++;
  }
  return c;
}

// ── Carregamento ──────────────────────────────────────────
async function carregar() {
  const resp = await fetch(UNIDADE.arquivo, { cache: 'no-store' });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const lastMod = resp.headers.get('Last-Modified');
  if (lastMod) {
    const d = new Date(lastMod);
    dataBase = `${d.toLocaleDateString('pt-BR')} às ${d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`;
  }
  const wb = XLSX.read(await resp.arrayBuffer(), { type: 'array', cellDates: true });
  const ws = wb.Sheets['Pendencias'];
  if (!ws) throw new Error('aba "Pendencias" não encontrada');

  const hoje = normData(new Date());
  const alvo = PAINEL.departamentos;
  const saida = [];
  const porMes = {};
  for (const r of XLSX.utils.sheet_to_json(ws, { raw: true })) {
    const dep = semAcento(r.Departamento);
    if (!(dep in alvo)) continue;
    if (r.DataVencimento instanceof Date) {
      const k = r.DataVencimento.getFullYear() * 12 + r.DataVencimento.getMonth();
      porMes[k] = (porMes[k] || 0) + 1;
    }
    // Guarda só o necessário — nada que identifique pessoa ou cliente.
    saida.push({
      st:    statusEfetivo(r, hoje),
      seg:   alvo[dep],
      cls:   String(r.Classificacao || 'Sem classificação').trim(),
      tit:   String(r.Titulo || 'Sem título').trim(),   // tipo de rotina, não identifica ninguém
      baixa: r.DataBaixa instanceof Date ? r.DataBaixa : null,
    });
  }
  linhas = saida;

  // Mês de referência vem da própria base (vencimento mais comum), não do
  // relógio: na virada do mês a base ainda é a do mês anterior até o robô rodar.
  const k = Object.entries(porMes).sort((a, b) => b[1] - a[1])[0]?.[0];
  mesRef = k !== undefined ? new Date(Math.floor(k / 12), k % 12, 1) : new Date(hoje.getFullYear(), hoje.getMonth(), 1);
}

// ── Cabeçalho e relógio ───────────────────────────────────
function mesRefTexto() {
  return mesRef.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
}

function atualizarRelogio() {
  const agora = new Date();
  document.getElementById('tv-hora').textContent = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  const data = agora.toLocaleDateString('pt-BR', { weekday: 'long', day: '2-digit', month: 'long' });
  document.getElementById('tv-data').textContent = data.charAt(0).toUpperCase() + data.slice(1);
}

// ── Placares ──────────────────────────────────────────────
function renderPlacares() {
  const c = contar(linhas);
  const segs = [...new Set(linhas.map(r => r.seg).filter(Boolean))].sort();
  const porSeg = Object.fromEntries(segs.map(s => [s, contar(linhas.filter(r => r.seg === s))]));

  const defs = [
    { classe: 'total',     chave: 'total',     label: 'Tarefas do mês', desc: `vencimento em ${mesRefTexto()}` },
    { classe: 'baixado',   chave: 'baixado',   label: 'Baixadas',       desc: `${fmtPct(pctConcluido(c.baixado, c.total))} do total` },
    { classe: 'aberto',    chave: 'aberto',    label: 'Em Aberto',      desc: `${pctTexto(c.aberto, c.total)} do total` },
    { classe: 'vencehoje', chave: 'vencehoje', label: 'Vencendo Hoje',  desc: `${pctTexto(c.vencehoje, c.total)} do total` },
    { classe: 'atraso',    chave: 'atraso',    label: 'Em Atraso',      desc: `${pctTexto(c.atraso, c.total)} do total` },
  ];

  // Régua do placar: parte do total (no "Tarefas do mês", quanto já foi concluído).
  // Mínimo de 1,5% pra um status com poucas tarefas não sumir da régua.
  const medidor = chave => {
    const v = chave === 'total' ? c.baixado : c[chave];
    return v && c.total ? Math.max((v / c.total) * 100, 1.5) : 0;
  };

  document.getElementById('placares').innerHTML = defs.map(d => `
    <div class="placar ${d.classe}">
      <div class="placar-label">${d.label}</div>
      <div class="placar-valor">${fmt(c[d.chave])}</div>
      <div class="placar-desc">${d.desc}</div>
      ${segs.length > 1 ? `<div class="placar-split">${segs.map(s => `<div>${s}<b>${fmt(porSeg[s][d.chave])}</b></div>`).join('')}</div>` : ''}
      <div class="placar-medidor"><span style="width:${medidor(d.chave)}%"></span></div>
    </div>`).join('');
}

// ── Gráficos ──────────────────────────────────────────────
function recriar(id, config) {
  if (graficos[id]) graficos[id].destroy();
  graficos[id] = new Chart(document.getElementById(id), config);
}

const semInteracao = { events: [], plugins: { tooltip: { enabled: false } } };

// Escreve o valor de cada barra (TV não tem mouse, então nada de tooltip).
const rotulosBarras = {
  id: 'rotulosBarras',
  afterDatasetsDraw(chart, _a, opts) {
    const { ctx } = chart;
    ctx.save();
    let tamanho = rem(opts.tamanho || .85);
    ctx.font = `700 ${tamanho}px Inter, Segoe UI`;
    if (!opts.dentro) {
      // Barras verticais: ou todos os rótulos aparecem (mesma fonte), ou nenhum —
      // mostrar só os que cabem deixaria os números pequenos e sumiria com os grandes.
      const vaga = chart.scales.x.width / Math.max(chart.data.labels.length, 1) - rem(.1);
      const maior = Math.max(0, ...chart.data.datasets.flatMap(ds => ds.data.map(v => ctx.measureText(fmt(v)).width)));
      if (maior > vaga) tamanho *= vaga / maior;
      if (tamanho < rem(.5)) { ctx.restore(); return; }
      ctx.font = `700 ${tamanho}px Inter, Segoe UI`;
    }
    chart.data.datasets.forEach((ds, i) => {
      chart.getDatasetMeta(i).data.forEach((bar, j) => {
        const v = ds.data[j];
        if (!v) return;
        const texto = fmt(v);
        if (opts.dentro) {
          // Linha com uma cor só: o total escrito no fim já diz tudo.
          if (chart.data.datasets.filter(d => d.data[j]).length < 2) return;
          const largura = Math.abs(bar.x - bar.base);
          if (largura < ctx.measureText(texto).width + rem(.8)) return;
          ctx.fillStyle = i === 0 ? COR.texto : '#FFFFFF';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'middle';
          ctx.fillText(texto, (bar.x + bar.base) / 2, bar.y);
        } else {
          ctx.fillStyle = COR.texto;
          ctx.textAlign = 'center';
          ctx.textBaseline = 'bottom';
          ctx.fillText(texto, bar.x, bar.y - rem(.25));
        }
      });
    });

    // Barras horizontais empilhadas: total no fim de cada linha. Sem isso, uma
    // tarefa com 1 ou 2 pendências vira um risco sem número nenhum.
    if (opts.dentro) {
      ctx.font = `800 ${rem(opts.tamanho || .85)}px Inter, Segoe UI`;
      ctx.fillStyle = COR.texto;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      chart.data.labels.forEach((_l, j) => {
        let total = 0, fim = chart.scales.x.left;
        chart.data.datasets.forEach((ds, i) => {
          total += ds.data[j] || 0;
          const bar = chart.getDatasetMeta(i).data[j];
          if (ds.data[j]) fim = Math.max(fim, bar.x);
        });
        if (total) ctx.fillText(fmt(total), fim + rem(.4), chart.getDatasetMeta(0).data[j].y);
      });
    }
    ctx.restore();
  },
};

function renderProgresso() {
  const c = contar(linhas);
  const p = pctConcluido(c.baixado, c.total);

  recriar('grafico-progresso', {
    type: 'doughnut',
    data: { datasets: [{
      data: [c.baixado, Math.max(c.total - c.baixado, 0)],
      backgroundColor: ctx => {
        if (ctx.dataIndex === 1) return 'transparent';
        const a = ctx.chart.chartArea;
        if (!a) return COR.marca;
        const g = ctx.chart.ctx.createLinearGradient(a.left, a.top, a.right, a.bottom);
        g.addColorStop(0, '#EF4444');
        g.addColorStop(.5, COR.marca);
        g.addColorStop(1, '#7A0000');
        return g;
      },
      borderWidth: 0,
      borderRadius: rem(1),
    }] },
    options: { ...semInteracao, responsive: true, maintainAspectRatio: false, cutout: '80%', animation: { duration: 900 } },
    plugins: [{
      // Trilho como anel inteiro por baixo: com a parte "que falta" como fatia,
      // as pontas arredondadas viravam uma bolinha solta quando falta pouco.
      id: 'trilho',
      beforeDatasetsDraw(chart) {
        const arco = chart.getDatasetMeta(0).data[0];
        if (!arco) return;
        const { x, y, innerRadius, outerRadius } = arco;
        const ctx = chart.ctx;
        ctx.save();
        ctx.beginPath();
        ctx.arc(x, y, (innerRadius + outerRadius) / 2, 0, Math.PI * 2);
        ctx.lineWidth = outerRadius - innerRadius;
        ctx.strokeStyle = COR.trilho;
        ctx.stroke();
        ctx.restore();
      },
    }, {
      id: 'centro',
      afterDraw(chart) {
        const { ctx, chartArea: { left, right, top, bottom } } = chart;
        const cx = (left + right) / 2, cy = (top + bottom) / 2;
        ctx.save();
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillStyle = COR.texto;
        ctx.font = `900 ${rem(3.2)}px Inter, Segoe UI`;
        ctx.fillText(fmtPct(p), cx, cy - rem(.5));
        ctx.fillStyle = COR.muted;
        ctx.font = `600 ${rem(.9)}px Inter, Segoe UI`;
        ctx.fillText(`${fmt(c.baixado)} de ${fmt(c.total)} baixadas`, cx, cy + rem(1.8));
        ctx.restore();
      },
    }],
  });

  // Abaixo do anel: Indústria x Varejo (Contábil/Fiscal) ou por classificação (DP).
  const chave = linhas.some(r => r.seg) ? 'seg' : 'cls';
  const grupos = [...new Set(linhas.map(r => r[chave]))].sort();
  document.getElementById('segmentos').innerHTML = grupos.map(g => {
    const cg = contar(linhas.filter(r => r[chave] === g));
    const pg = pctConcluido(cg.baixado, cg.total);
    return `
      <div>
        <div class="segmento-topo"><span>${g}</span><b>${fmtPct(pg)}</b></div>
        <div class="segmento-trilho"><div class="segmento-barra" style="width:${pg}%"></div></div>
      </div>`;
  }).join('');
}

function renderDias() {
  const hoje = normData(new Date());
  const ini  = new Date(mesRef);
  const fimMes = new Date(ini.getFullYear(), ini.getMonth() + 1, 0);
  const fim  = hoje < fimMes ? hoje : fimMes;
  const porDia = {};
  let antes = 0;
  for (const r of linhas) {
    if (!r.baixa) continue;
    const d = normData(r.baixa);
    // Tarefa do mês baixada antes do dia 1 (adiantada) não tem barra — vai
    // escrita no título, junto com o total de hoje.
    if (d < ini) { antes++; continue; }
    if (d > fim) continue;
    porDia[d.getDate()] = (porDia[d.getDate()] || 0) + 1;
  }

  // Dias úteis sempre aparecem; sábado/domingo só se houve baixa.
  const dias = [];
  for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
    const fimDeSemana = d.getDay() === 0 || d.getDay() === 6;
    if (!fimDeSemana || porDia[d.getDate()]) dias.push(new Date(d));
  }
  const ehHoje = d => d.getTime() === hoje.getTime();
  const labels  = dias.map(d => String(d.getDate()).padStart(2, '0'));
  const valores = dias.map(d => porDia[d.getDate()] || 0);
  const iHoje   = dias.findIndex(ehHoje);
  // Degradê de cima pra baixo; a barra de hoje fica no vermelho mais escuro.
  const cor = ctx => {
    const { chart, dataIndex } = ctx;
    const area = chart.chartArea;
    if (!area) return COR.marca;
    const g = chart.ctx.createLinearGradient(0, area.top, 0, area.bottom);
    if (dataIndex === iHoje) { g.addColorStop(0, COR.atraso); g.addColorStop(1, '#6B0F0F'); }
    else { g.addColorStop(0, '#E11D1D'); g.addColorStop(1, 'rgba(192,0,0,.55)'); }
    return g;
  };

  document.getElementById('titulo-dias').textContent = `Baixadas por dia · ${mesRef.toLocaleDateString('pt-BR', { month: 'long' })}`;
  const hojeQtd = porDia[hoje.getDate()] && hoje >= ini && hoje <= fimMes ? porDia[hoje.getDate()] : 0;
  const extras = [];
  if (antes) extras.push(`<span class="chip">Antes do dia 1:<b>${fmt(antes)}</b></span>`);
  if (hoje >= ini && hoje <= fimMes) extras.push(`<span class="chip hoje">Hoje:<b>${fmt(hojeQtd)}</b></span>`);
  document.getElementById('dias-hoje').innerHTML = extras.join('');

  recriar('grafico-dias', {
    type: 'bar',
    data: { labels, datasets: [{ data: valores, backgroundColor: cor, borderRadius: rem(.35), borderSkipped: 'bottom', maxBarThickness: rem(2.4) }] },
    options: {
      ...semInteracao,
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: rem(1.4), left: rem(.6) } },
      plugins: { ...semInteracao.plugins, legend: { display: false }, rotulosBarras: { tamanho: .8 } },
      scales: {
        x: { grid: { display: false }, ticks: { color: COR.muted, font: { size: rem(.8), family: 'Inter' }, maxRotation: 0, autoSkip: window.innerWidth < 700, maxTicksLimit: window.innerWidth < 700 ? 10 : undefined } },
        y: { beginAtZero: true, grid: { color: COR.grade }, border: { display: false, dash: [4, 4] }, ticks: { color: COR.muted, font: { size: rem(.75), family: 'Inter' }, maxTicksLimit: 5 } },
      },
    },
    plugins: [rotulosBarras],
  });
}

const TOP_TAREFAS = 6;

// Tipos de tarefa em aberto que vencem hoje (status "Vencendo Hoje").
function renderPendencias() {
  const pend = linhas.filter(r => r.st === 'Vencendo Hoje');
  const porTit = {};
  for (const r of pend) porTit[r.tit] = (porTit[r.tit] || 0) + 1;
  const titulos = Object.entries(porTit).sort((a, b) => b[1] - a[1]).slice(0, TOP_TAREFAS).map(([t]) => t);
  const valores = titulos.map(t => porTit[t]);
  // Quebra em até 2 linhas (Chart.js aceita array como rótulo) em vez de cortar,
  // senão "Rebto planilha/Resumo de Folha (30)" e "(28)" ficam idênticas.
  const quebrar = t => {
    if (t.length <= 22) return t;
    const i = t.lastIndexOf(' ', 22);
    const corte = i > 8 ? i : 22;
    const resto = t.slice(corte).trim();
    return [t.slice(0, corte).trim(), resto.length > 24 ? resto.slice(0, 23) + '…' : resto];
  };
  const classes = titulos.map(quebrar);

  document.getElementById('titulo-pendencias').textContent = `Principais tarefas em aberto vencendo hoje · ${fmt(pend.length)} no total`;

  const vazio = document.getElementById('pendencias-vazio');
  vazio.classList.toggle('hidden', pend.length > 0);
  if (!pend.length) {
    if (graficos['grafico-pendencias']) { graficos['grafico-pendencias'].destroy(); delete graficos['grafico-pendencias']; }
    return;
  }

  recriar('grafico-pendencias', {
    type: 'bar',
    data: {
      labels: classes,
      datasets: [{
        label: 'Vencendo Hoje', data: valores, borderSkipped: false, borderRadius: rem(.35), barPercentage: .68, maxBarThickness: rem(3),
        backgroundColor: ctx => {
          const a = ctx.chart.chartArea;
          if (!a) return COR.vencehoje;
          const g = ctx.chart.ctx.createLinearGradient(a.left, 0, a.right, 0);
          g.addColorStop(0, COR.vencehoje);
          g.addColorStop(1, COR.atraso);
          return g;
        },
      }],
    },
    options: {
      ...semInteracao,
      indexAxis: 'y',
      layout: { padding: { left: rem(.6) } },
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        ...semInteracao.plugins,
        legend: { display: false },
        rotulosBarras: { dentro: true, tamanho: .85 },
      },
      scales: {
        x: { stacked: true, beginAtZero: true, grace: '12%', grid: { color: COR.grade }, border: { display: false, dash: [4, 4] }, ticks: { color: COR.muted, font: { size: rem(.75), family: 'Inter' }, maxTicksLimit: 5 } },
        y: { stacked: true, grid: { display: false }, ticks: { color: COR.texto, font: { size: rem(.8), weight: 600, family: 'Inter' } } },
      },
    },
    plugins: [rotulosBarras],
  });
}

function renderTudo() {
  renderPlacares();
  renderProgresso();
  renderDias();
  renderPendencias();
}

function rodape(erro) {
  const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  document.getElementById('tv-rodape').innerHTML =
    (dataBase ? `Base atualizada em ${dataBase} · ` : '') +
    `Painel conferido às ${agora}` +
    (erro ? ` · <span class="erro">Não foi possível buscar dados novos (${erro}) — exibindo a última leitura</span>` : '');
}

const TENTAR_DE_NOVO_SEG = 60;   // se a 1ª carga falhar (TV ligou sem rede, etc.)

// Versão do visual que está rodando (o ?v= do app.js no index.html). Quando o
// index.html publicado tiver outra versão, a página se recarrega sozinha —
// ninguém precisa apertar F5 nem limpar cache na TV depois de uma mudança.
const VERSAO_ATUAL = (document.querySelector('script[src*="app.js"]')?.src.match(/[?&]v=([^&]+)/) || [])[1] || '';

async function conferirVersaoNova() {
  if (!VERSAO_ATUAL) return;
  try {
    const html = await (await fetch(`index.html?t=${Date.now()}`, { cache: 'no-store' })).text();
    const publicada = (html.match(/app\.js\?v=([^"'&]+)/) || [])[1];
    // Recarrega pedindo "?v=<nova>": o GitHub guarda a página por até 10 min, e
    // um reload simples poderia voltar com a versão velha e ficar em loop.
    if (publicada && publicada !== VERSAO_ATUAL) location.replace(`${location.pathname}?v=${encodeURIComponent(publicada)}`);
  } catch (_) { /* sem rede: tenta de novo na próxima rodada */ }
}

async function atualizar() {
  // Só nas rodadas seguintes, nunca logo ao abrir — assim uma página recém-
  // carregada nunca recarrega de novo em seguida.
  if (jaCarregou) await conferirVersaoNova();
  try {
    await carregar();
    renderTudo();
    rodape();
    jaCarregou = true;
    document.getElementById('carregando').classList.add('hidden');
  } catch (e) {
    console.error('Erro ao atualizar o painel:', e);
    if (jaCarregou) { rodape(e.message); return; }
    // Nunca carregou: em vez de mostrar placares vazios, fica na tela de espera
    // com o motivo e tenta de novo sozinho.
    document.getElementById('carregando-texto').textContent =
      `Não foi possível carregar os dados (${e.message}). Tentando de novo em ${TENTAR_DE_NOVO_SEG} segundos...`;
    setTimeout(atualizar, TENTAR_DE_NOVO_SEG * 1000);
  }
}

// ── Início ────────────────────────────────────────────────
async function iniciar() {
  document.title = `${PAINEL.nome} - Painel de Tarefas`;
  document.getElementById('tv-titulo').textContent = PAINEL.nome;
  document.getElementById('tv-sub').textContent = `${PAINEL.sub} · ${UNIDADE.nome}`;
  atualizarRelogio();
  setInterval(atualizarRelogio, 15 * 1000);

  document.getElementById('painel').classList.remove('hidden');
  // Os gráficos medem o texto na hora de desenhar: se a Inter ainda não
  // carregou, medem com a fonte reserva (mais estreita) e os rótulos saem cortados.
  try {
    await Promise.all(['600 16px Inter', '700 16px Inter', '800 16px Inter', '900 16px Inter'].map(f => document.fonts.load(f)));
    await document.fonts.ready;
  } catch (_) { /* sem a fonte, segue com a reserva */ }
  await atualizar();

  setInterval(() => { if (jaCarregou) atualizar(); }, ATUALIZAR_DADOS_MIN * 60 * 1000);
  setTimeout(() => location.reload(), RECARREGAR_PAGINA_H * 60 * 60 * 1000);

  let timerResize = null;
  window.addEventListener('resize', () => {
    clearTimeout(timerResize);
    timerResize = setTimeout(() => { if (jaCarregou) renderTudo(); }, 300);
  });
}

iniciar();
