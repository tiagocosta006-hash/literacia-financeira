// ============================================================================
// SIMULADOR DE LITERACIA FINANCEIRA — NÍVEIS 1 A 5 (Fidelidade Total ao Excel)
// ============================================================================

const fmtEUR = (val) =>
  new Intl.NumberFormat('pt-PT', { style: 'currency', currency: 'EUR' }).format(
    Number.isFinite(val) ? val : 0
  );

const fmtPct = (val, decimals = 1) =>
  new Intl.NumberFormat('pt-PT', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(Number.isFinite(val) ? val * 100 : 0) + '%';

const fmtNum = (val, decimals = 1) =>
  new Intl.NumberFormat('pt-PT', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(Number.isFinite(val) ? val : 0);

// Chart instances
let chartN1 = null;
let chartN2Salario = null;
let chartN2Poupanca = null;
let chartN3Salario = null;
let chartN3Iva = null;
let chartN4Juros = null;
let chartN4Credito = null;
let chartN5Aloc = null;
let chartN5Evol = null;

// ============================================================================
// DADOS PREDEFINIDOS DO EXCEL
// ============================================================================

const EXCEL_DEFAULTS = {
  n1: {
    rendimento: 900,
    poupanca: 100,
    rows: [
      { categoria: 'Roupa', tipo: 'Desejo', valor: 100, notas: '' },
      { categoria: 'Luz', tipo: 'Necessidade', valor: 2, notas: '' },
      { categoria: 'Água', tipo: 'Necessidade', valor: 22, notas: '' },
      { categoria: 'Alimentação', tipo: 'Necessidade', valor: 2, notas: '' },
      { categoria: 'Renda', tipo: 'Necessidade', valor: 22, notas: '' },
      { categoria: 'Internet', tipo: 'Desejo', valor: 0, notas: '' },
      { categoria: 'Desporto', tipo: 'Desejo', valor: 0, notas: '' },
      { categoria: 'Transportes', tipo: 'Necessidade', valor: 0, notas: '' },
    ],
  },
  n1_realista: {
    rendimento: 920,
    poupanca: 150,
    rows: [
      { categoria: 'Roupa', tipo: 'Desejo', valor: 35, notas: 'Compras pontuais' },
      { categoria: 'Luz', tipo: 'Necessidade', valor: 45, notas: 'Fatura mensal' },
      { categoria: 'Água', tipo: 'Necessidade', valor: 20, notas: 'Serviços municipalizados' },
      { categoria: 'Alimentação', tipo: 'Necessidade', valor: 180, notas: 'Supermercado essencial' },
      { categoria: 'Renda', tipo: 'Necessidade', valor: 300, notas: 'Quarto / Habitação partilhada' },
      { categoria: 'Internet', tipo: 'Desejo', valor: 30, notas: 'Pacote telecomunicações' },
      { categoria: 'Desporto', tipo: 'Desejo', valor: 25, notas: 'Ginásio / atividade física' },
      { categoria: 'Transportes', tipo: 'Necessidade', valor: 40, notas: 'Passe mensal' },
    ],
  },
  n2: {
    rendimento: 920,
    despesas: 720,
    poupanca: 100,
    inflacao: 2.3,
    aumento: 1.0,
    rendibilidade: 2.5,
    cafeHoje: 0.8,
  },
  n3: {
    bruto: 920,
    ssTrab: 11,
    irsTaxa: 0,
    ssEmp: 23.75,
    ivaRows: [
      {
        categoria: 'Alimentação (bens essenciais no supermercado)',
        valorComIva: 180,
        taxa: 6,
        motivo: 'Taxa reduzida: pão, leite, fruta, legumes, medicamentos.',
      },
      {
        categoria: 'Refeições fora de casa (restauração)',
        valorComIva: 40,
        taxa: 13,
        motivo: 'Taxa intermédia. Bebidas alcoólicas e refrigerantes ficam a 23%.',
      },
      {
        categoria: 'Roupa e calçado',
        valorComIva: 30,
        taxa: 23,
        motivo: 'Taxa normal — não consta das Listas I ou II do CIVA.',
      },
      {
        categoria: 'Eletricidade',
        valorComIva: 45,
        taxa: 23,
        motivo: 'Taxa mista: 6% até 200 kWh/30 dias com potência até 6,9 kVA; o resto a 23%.',
      },
      {
        categoria: 'Internet e telemóvel',
        valorComIva: 30,
        taxa: 23,
        motivo: 'Taxa normal.',
      },
      {
        categoria: 'Transportes públicos',
        valorComIva: 40,
        taxa: 6,
        motivo: 'Taxa reduzida — transporte de passageiros.',
      },
      {
        categoria: 'Ginásio / desporto',
        valorComIva: 25,
        taxa: 23,
        motivo: 'Taxa normal.',
      },
      {
        categoria: 'Renda de casa',
        valorComIva: 300,
        taxa: 0,
        motivo: 'Isento de IVA (arrendamento para habitação, art. 9.º do CIVA).',
      },
    ],
    vpt: 100000,
    imiTaxa: 0.3,
    dep: 0,
    lucro: 40000,
    ircPme: 15,
    ircGeral: 19,
  },
  n4: {
    capIni: 1000,
    reforco: 50,
    taxaJuro: 4.0,
    credValor: 5000,
    credTan: 8.0,
    credPrazo: 60,
    credPoupTaxa: 2.5,
    dpMontante: 1000,
    dpTaxa: 1.5,
    caTaxa: 2.5,
    irsCap: 28,
    inflacao: 2.3,
    negInv: 3000,
    negFixos: 500,
    negPreco: 8,
    negVar: 3,
    negUnid: 200,
    seguros: [
      {
        tipo: 'Responsabilidade civil automóvel',
        obrigatorio: 'Sim, por lei',
        serve: 'Paga os danos que causares a terceiros.',
        premio: 300,
      },
      {
        tipo: 'Multirriscos habitação',
        obrigatorio: 'Só com crédito à habitação',
        serve: 'Cobre incêndio, inundação e outros danos na casa.',
        premio: 150,
      },
      {
        tipo: 'Seguro de vida',
        obrigatorio: 'Exigido pelo banco no crédito',
        serve: 'Liquida a dívida se o titular falecer.',
        premio: 200,
      },
      {
        tipo: 'Seguro de saúde',
        obrigatorio: 'Não',
        serve: 'Reduz o custo e o tempo de espera de consultas e exames.',
        premio: 250,
      },
    ],
  },
  n5: {
    capIni: 1000,
    poupMensal: 150,
    despesas: 720,
    inflacao: 2.3,
    instrumentos: [
      {
        nome: 'Fundo de emergência (depósito à ordem/prazo)',
        pesoPct: 20,
        retornoPct: 1.0,
        risco: 'Muito baixo',
        liquidez: 'Imediata',
      },
      {
        nome: 'Certificados de Aforro (Série F)',
        pesoPct: 15,
        retornoPct: 2.5,
        risco: 'Muito baixo',
        liquidez: 'Resgate após 3 meses',
      },
      {
        nome: 'PPR — Plano Poupança Reforma',
        pesoPct: 15,
        retornoPct: 3.0,
        risco: 'Baixo a médio',
        liquidez: 'Penalizado antes da reforma',
      },
      {
        nome: 'ETF de ações global',
        pesoPct: 35,
        retornoPct: 6.0,
        risco: 'Médio a alto',
        liquidez: 'Alta (dias)',
      },
      {
        nome: 'Ações individuais',
        pesoPct: 15,
        retornoPct: 7.0,
        risco: 'Alto',
        liquidez: 'Alta (dias)',
      },
    ],
  },
};

// State
let stateN1Rows = JSON.parse(JSON.stringify(EXCEL_DEFAULTS.n1.rows));
let stateN3IvaRows = JSON.parse(JSON.stringify(EXCEL_DEFAULTS.n3.ivaRows));
let stateN4Seguros = JSON.parse(JSON.stringify(EXCEL_DEFAULTS.n4.seguros));
let stateN5Inst = JSON.parse(JSON.stringify(EXCEL_DEFAULTS.n5.instrumentos));

// ============================================================================
// NAVEGAÇÃO ENTRE TABS
// ============================================================================

function switchTab(tabId) {
  document.querySelectorAll('.tab-panel').forEach((panel) => {
    panel.classList.remove('active');
  });
  document.querySelectorAll('.nav-btn').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.tab === tabId);
  });
  const target = document.getElementById(`tab-${tabId}`);
  if (target) target.classList.add('active');

  // Redesenhar gráficos do tab ativo para garantir dimensões perfeitas
  setTimeout(() => {
    if (tabId === 'nivel1') recalcN1();
    if (tabId === 'nivel2') recalcN2();
    if (tabId === 'nivel3') recalcN3();
    if (tabId === 'nivel4') recalcN4();
    if (tabId === 'nivel5') recalcN5();
  }, 30);
}

// ============================================================================
// TAB 0: DEMO LEGENDA («Como Usar»)
// ============================================================================

function updateDemoLegend() {
  const bruto = parseFloat(document.getElementById('demo-bruto').value) || 0;
  const ss = bruto * 0.11;
  const liq = bruto - ss;
  document.getElementById('demo-ss').textContent = fmtEUR(ss);
  document.getElementById('demo-liquido').textContent = fmtEUR(liq);
}

// ============================================================================
// NÍVEL 1 — NECESSIDADES VS DESEJOS
// ============================================================================

function renderN1Table() {
  const tbody = document.getElementById('n1-tbody');
  tbody.innerHTML = '';
  stateN1Rows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>
        <input type="text" class="excel-input" value="${row.categoria}"
          oninput="stateN1Rows[${idx}].categoria = this.value; recalcN1(false);" />
      </td>
      <td>
        <select class="excel-select" onchange="stateN1Rows[${idx}].tipo = this.value; recalcN1(false);">
          <option value="Necessidade" ${row.tipo === 'Necessidade' ? 'selected' : ''}>Necessidade</option>
          <option value="Desejo" ${row.tipo === 'Desejo' ? 'selected' : ''}>Desejo</option>
        </select>
      </td>
      <td class="num">
        <input type="number" class="excel-input" style="text-align:right" value="${row.valor}" step="5"
          oninput="stateN1Rows[${idx}].valor = parseFloat(this.value) || 0; recalcN1(false);" />
      </td>
      <td class="num td-calc" id="n1-row-pct-${idx}">0,0%</td>
      <td>
        <input type="text" class="excel-input" style="background:#fff; border-color:#cbd5e1; color:#334155; font-weight:500;"
          placeholder="Nota opcional..." value="${row.notas || ''}"
          oninput="stateN1Rows[${idx}].notas = this.value;" />
      </td>
      <td style="text-align:center;">
        <button class="btn-small danger" onclick="removeN1Row(${idx})" title="Remover linha">✕</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function addN1Row() {
  stateN1Rows.push({ categoria: 'Nova Despesa', tipo: 'Desejo', valor: 25, notas: '' });
  renderN1Table();
  recalcN1(false);
}

function removeN1Row(idx) {
  if (stateN1Rows.length <= 1) return;
  stateN1Rows.splice(idx, 1);
  renderN1Table();
  recalcN1(false);
}

function loadN1Preset(preset) {
  const src = preset === 'realista' ? EXCEL_DEFAULTS.n1_realista : EXCEL_DEFAULTS.n1;
  document.getElementById('n1-rendimento').value = src.rendimento;
  document.getElementById('n1-poupanca').value = src.poupanca;
  stateN1Rows = JSON.parse(JSON.stringify(src.rows));
  renderN1Table();
  recalcN1(false);
}

function recalcN1(rebuildTable = false) {
  if (rebuildTable) renderN1Table();

  const rend = parseFloat(document.getElementById('n1-rendimento').value) || 0;
  const poup = parseFloat(document.getElementById('n1-poupanca').value) || 0;

  let totNec = 0;
  let totDes = 0;

  stateN1Rows.forEach((r, idx) => {
    const val = Number(r.valor) || 0;
    if (r.tipo === 'Necessidade') totNec += val;
    else totDes += val;

    const pct = rend > 0 ? val / rend : 0;
    const cellPct = document.getElementById(`n1-row-pct-${idx}`);
    if (cellPct) cellPct.textContent = fmtPct(pct, 1);
  });

  const totDesp = totNec + totDes;
  const saldoAposDesp = rend - totDesp;
  const saldoFinal = rend - totDesp - poup;

  const pctNec = rend > 0 ? totNec / rend : 0;
  const pctDes = rend > 0 ? totDes / rend : 0;
  const pctPoup = rend > 0 ? poup / rend : 0;
  const pctAloc = rend > 0 ? (totDesp + poup) / rend : 0;
  const pctLivre = rend > 0 ? saldoAposDesp / rend : 0;

  // KPIs
  document.getElementById('n1-kpi-nec').textContent = fmtEUR(totNec);
  document.getElementById('n1-kpi-nec-pct').textContent = `${fmtPct(pctNec, 1)} do rendimento (Meta ≤ 50%)`;

  document.getElementById('n1-kpi-des').textContent = fmtEUR(totDes);
  document.getElementById('n1-kpi-des-pct').textContent = `${fmtPct(pctDes, 1)} do rendimento (Meta ≤ 30%)`;

  document.getElementById('n1-kpi-tot').textContent = fmtEUR(totDesp);
  document.getElementById('n1-kpi-saldo-desp').textContent = `Saldo após despesas: ${fmtEUR(saldoAposDesp)}`;

  const kpiFinal = document.getElementById('n1-kpi-saldo-final');
  kpiFinal.textContent = fmtEUR(saldoFinal);
  const kpiStatus = document.getElementById('n1-kpi-status');
  if (saldoFinal < 0) {
    kpiStatus.textContent = '⚠️ Atenção: Orçamento negativo! Reduz desejos ou despesas.';
    kpiFinal.style.color = '#dc2626';
  } else {
    kpiStatus.textContent = '✅ Orçamento equilibrado (sem saldo negativo)';
    kpiFinal.style.color = '';
  }

  // Resumo N1 Table
  document.getElementById('n1-res-rend').textContent = fmtEUR(rend);
  document.getElementById('n1-res-nec').textContent = fmtEUR(totNec);
  document.getElementById('n1-res-nec-pct').textContent = fmtPct(pctNec, 1);
  document.getElementById('n1-res-des').textContent = fmtEUR(totDes);
  document.getElementById('n1-res-des-pct').textContent = fmtPct(pctDes, 1);
  document.getElementById('n1-res-poup').textContent = fmtEUR(poup);
  document.getElementById('n1-res-poup-pct').textContent = fmtPct(pctPoup, 1);
  document.getElementById('n1-res-aloc-eur').textContent = fmtEUR(totDesp + poup);
  document.getElementById('n1-res-aloc-pct').textContent = fmtPct(pctAloc, 1);
  document.getElementById('n1-res-livre').textContent = fmtEUR(saldoAposDesp);
  document.getElementById('n1-res-livre-pct').textContent = fmtPct(pctLivre, 1);

  // Barras 50/30/20
  document.getElementById('n1-bar-nec-lbl').textContent = fmtPct(pctNec, 1);
  document.getElementById('n1-bar-nec').style.width = `${Math.min(100, Math.max(0, pctNec * 100))}%`;
  document.getElementById('n1-bar-nec').style.background = pctNec <= 0.5 ? '#0d9488' : '#dc2626';

  document.getElementById('n1-bar-des-lbl').textContent = fmtPct(pctDes, 1);
  document.getElementById('n1-bar-des').style.width = `${Math.min(100, Math.max(0, pctDes * 100))}%`;
  document.getElementById('n1-bar-des').style.background = pctDes <= 0.3 ? '#f59e0b' : '#dc2626';

  document.getElementById('n1-bar-poup-lbl').textContent = fmtPct(pctPoup, 1);
  document.getElementById('n1-bar-poup').style.width = `${Math.min(100, Math.max(0, pctPoup * 100))}%`;
  document.getElementById('n1-bar-poup').style.background = pctPoup >= 0.2 ? '#16a34a' : '#2563eb';

  // Doughnut Chart N1 (Idêntico à folha «Resumo & gráfico N1»: Necessidades, Desejos, Poupança, Saldo livre)
  const ctx = document.getElementById('chart-n1').getContext('2d');
  const dataVals = [
    Math.max(0, totNec),
    Math.max(0, totDes),
    Math.max(0, poup),
    Math.max(0, saldoFinal),
  ];

  if (chartN1) {
    chartN1.data.datasets[0].data = dataVals;
    chartN1.update();
  } else {
    chartN1 = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['Necessidades', 'Desejos', 'Poupança (planeada)', 'Saldo livre (após poupança)'],
        datasets: [
          {
            data: dataVals,
            backgroundColor: ['#0d9488', '#f59e0b', '#2563eb', '#94a3b8'],
            borderWidth: 2,
            borderColor: '#ffffff',
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom' },
          title: { display: true, text: 'Distribuição do Orçamento Mensal (€)' },
        },
      },
    });
  }
}

// ============================================================================
// NÍVEL 2 — IMPLICAÇÕES DO TEMPO E DA INFLAÇÃO
// ============================================================================

function syncFromN1ToN2() {
  const rendN1 = parseFloat(document.getElementById('n1-rendimento').value) || 900;
  const poupN1 = parseFloat(document.getElementById('n1-poupanca').value) || 100;
  const totDespN1 = stateN1Rows.reduce((acc, r) => acc + (Number(r.valor) || 0), 0);

  document.getElementById('n2-rendimento').value = rendN1;
  document.getElementById('n2-despesas').value = totDespN1;
  document.getElementById('n2-poupanca').value = poupN1;
  recalcN2();
}

function setZeroWageGrowthN2() {
  document.getElementById('n2-aumento').value = 0;
  recalcN2();
}

function resetN2() {
  const d = EXCEL_DEFAULTS.n2;
  document.getElementById('n2-rendimento').value = d.rendimento;
  document.getElementById('n2-despesas').value = d.despesas;
  document.getElementById('n2-poupanca').value = d.poupanca;
  document.getElementById('n2-inflacao').value = d.inflacao;
  document.getElementById('n2-aumento').value = d.aumento;
  document.getElementById('n2-rendibilidade').value = d.rendibilidade;
  document.getElementById('n2-cafe-hoje').value = d.cafeHoje;
  recalcN2();
}

function recalcN2() {
  const b7 = parseFloat(document.getElementById('n2-rendimento').value) || 0;
  const b8 = parseFloat(document.getElementById('n2-despesas').value) || 0;
  const b9 = parseFloat(document.getElementById('n2-poupanca').value) || 0;
  const b10 = (parseFloat(document.getElementById('n2-inflacao').value) || 0) / 100;
  const b11 = (parseFloat(document.getElementById('n2-aumento').value) || 0) / 100;
  const b12 = (parseFloat(document.getElementById('n2-rendibilidade').value) || 0) / 100;
  const b36 = parseFloat(document.getElementById('n2-cafe-hoje').value) || 0;

  const years = [];
  const salNominalArr = [];
  const custoCabazArr = [];
  const salHojeArr = [];
  const poupSemJurosArr = [];
  const poupComJurosArr = [];

  const tbody = document.getElementById('n2-tbody');
  tbody.innerHTML = '';

  let e26 = 0;
  let f26 = 0;
  let g26 = 0;

  for (let a = 0; a <= 10; a++) {
    // Fórmulas exatas do Excel Nível 2:
    // B(a) = B7*(1+B11)^a
    const salNom = b7 * Math.pow(1 + b11, a);
    // C(a) = B8*(1+B10)^a
    const custoCab = b8 * Math.pow(1 + b10, a);
    // D(a) = B(a) - C(a)
    const sobra = salNom - custoCab;
    // E(a) = B(a)/(1+B10)^a
    const salHoje = Math.pow(1 + b10, a) !== 0 ? salNom / Math.pow(1 + b10, a) : 0;
    // F(a) = B9*12*a
    const poupSem = b9 * 12 * a;
    // G(a) = IF(B12=0, B9*12*a, B9*12*((1+B12)^a - 1)/B12)
    const poupCom = b12 === 0 ? b9 * 12 * a : (b9 * 12 * (Math.pow(1 + b12, a) - 1)) / b12;
    // H(a) = G(a)/(1+B10)^a
    const poupHoje = Math.pow(1 + b10, a) !== 0 ? poupCom / Math.pow(1 + b10, a) : 0;

    if (a === 10) {
      e26 = salHoje;
      f26 = poupSem;
      g26 = poupCom;
    }

    years.push(`Ano ${a}`);
    salNominalArr.push(Number(salNom.toFixed(2)));
    custoCabazArr.push(Number(custoCab.toFixed(2)));
    salHojeArr.push(Number(salHoje.toFixed(2)));
    poupSemJurosArr.push(Number(poupSem.toFixed(2)));
    poupComJurosArr.push(Number(poupCom.toFixed(2)));

    const tr = document.createElement('tr');
    if (a === 10) tr.className = 'row-total';
    tr.innerHTML = `
      <td><strong>Ano ${a}</strong></td>
      <td class="num td-calc">${fmtEUR(salNom)}</td>
      <td class="num td-calc">${fmtEUR(custoCab)}</td>
      <td class="num td-calc" style="color:${sobra < 0 ? '#dc2626' : ''}">${fmtEUR(sobra)}</td>
      <td class="num td-calc">${fmtEUR(salHoje)}</td>
      <td class="num td-calc">${fmtEUR(poupSem)}</td>
      <td class="num td-calc">${fmtEUR(poupCom)}</td>
      <td class="num td-result">${fmtEUR(poupHoje)}</td>
    `;
    tbody.appendChild(tr);
  }

  // 3. Conclusões
  const perdaPoderCompraEur = b7 - e26;
  const perdaPoderCompraPct = b7 !== 0 ? 1 - e26 / b7 : 0;
  const ganhoJuros10 = g26 - f26;
  const juroReal = 1 + b10 !== 0 ? (1 + b12) / (1 + b10) - 1 : 0;

  document.getElementById('n2-kpi-perda-eur').textContent = `${fmtEUR(perdaPoderCompraEur)}/mês`;
  document.getElementById('n2-kpi-perda-pct').textContent = fmtPct(perdaPoderCompraPct, 1);
  document.getElementById('n2-kpi-ganho-juros').textContent = fmtEUR(ganhoJuros10);
  document.getElementById('n2-kpi-juro-real').textContent =
    (juroReal >= 0 ? '+' : '') + fmtPct(juroReal, 2);

  // 4. Exemplo Prático — Preço do Café
  const cafe5 = b36 * Math.pow(1 + b10, 5);
  const cafe10 = b36 * Math.pow(1 + b10, 10);
  const cafe20 = b36 * Math.pow(1 + b10, 20);
  const cafesHoje10Eur = b36 > 0 ? 10 / b36 : 0;
  const cafes20Anos10Eur = cafe20 > 0 ? 10 / cafe20 : 0;

  document.getElementById('n2-cafe-5').textContent = fmtEUR(cafe5);
  document.getElementById('n2-cafe-10').textContent = fmtEUR(cafe10);
  document.getElementById('n2-cafe-20').textContent = fmtEUR(cafe20);
  document.getElementById('n2-cafes-hoje').textContent = `${fmtNum(cafesHoje10Eur, 1)} cafés`;
  document.getElementById('n2-cafes-20').textContent = `${fmtNum(cafes20Anos10Eur, 1)} cafés`;

  // Ilustração visual das chávenas de café
  const cupsDiv = document.getElementById('n2-coffee-cups');
  const fullHoje = Math.min(30, Math.round(cafesHoje10Eur));
  const full20 = Math.min(30, Math.round(cafes20Anos10Eur));
  cupsDiv.innerHTML = `
    <div style="width:100%; font-size:0.8rem; font-weight:700; color:#92400e;">
      Hoje com 10 € (${fmtNum(cafesHoje10Eur, 1)} cafés): ${'☕'.repeat(fullHoje)}
    </div>
    <div style="width:100%; font-size:0.8rem; font-weight:700; color:#b91c1c; margin-top:0.25rem;">
      Daqui a 20 anos com os mesmos 10 € (${fmtNum(cafes20Anos10Eur, 1)} cafés): ${'☕'.repeat(full20)}
    </div>
  `;

  // Gráficos N2
  const ctxSal = document.getElementById('chart-n2-salario').getContext('2d');
  if (chartN2Salario) {
    chartN2Salario.data.labels = years;
    chartN2Salario.data.datasets[0].data = salNominalArr;
    chartN2Salario.data.datasets[1].data = custoCabazArr;
    chartN2Salario.data.datasets[2].data = salHojeArr;
    chartN2Salario.update();
  } else {
    chartN2Salario = new Chart(ctxSal, {
      type: 'line',
      data: {
        labels: years,
        datasets: [
          {
            label: 'Salário nominal (€)',
            data: salNominalArr,
            borderColor: '#2563eb',
            backgroundColor: 'rgba(37, 99, 235, 0.1)',
            tension: 0.25,
            borderWidth: 2.5,
          },
          {
            label: 'Custo do cabaz (€)',
            data: custoCabazArr,
            borderColor: '#dc2626',
            backgroundColor: 'rgba(220, 38, 38, 0.1)',
            tension: 0.25,
            borderWidth: 2.5,
          },
          {
            label: 'Salário a preços de hoje (€)',
            data: salHojeArr,
            borderColor: '#f59e0b',
            borderDash: [6, 4],
            tension: 0.25,
            borderWidth: 2.5,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
      },
    });
  }

  const ctxPoup = document.getElementById('chart-n2-poupanca').getContext('2d');
  if (chartN2Poupanca) {
    chartN2Poupanca.data.labels = years;
    chartN2Poupanca.data.datasets[0].data = poupSemJurosArr;
    chartN2Poupanca.data.datasets[1].data = poupComJurosArr;
    chartN2Poupanca.update();
  } else {
    chartN2Poupanca = new Chart(ctxPoup, {
      type: 'line',
      data: {
        labels: years,
        datasets: [
          {
            label: 'Poupança sem juros (€)',
            data: poupSemJurosArr,
            borderColor: '#64748b',
            borderDash: [5, 5],
            tension: 0.2,
            borderWidth: 2,
          },
          {
            label: 'Poupança com juros (€)',
            data: poupComJurosArr,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16, 185, 129, 0.12)',
            fill: true,
            tension: 0.25,
            borderWidth: 3,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
      },
    });
  }
}

// ============================================================================
// NÍVEL 3 — DESMISTIFICAR OS IMPOSTOS E O PAPEL DO ESTADO
// ============================================================================

function renderN3IvaTable() {
  const tbody = document.getElementById('n3-iva-tbody');
  tbody.innerHTML = '';
  stateN3IvaRows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${row.categoria}</strong></td>
      <td class="num">
        <input type="number" class="excel-input" style="text-align:right; max-width:120px;" value="${row.valorComIva}" step="5"
          oninput="stateN3IvaRows[${idx}].valorComIva = parseFloat(this.value) || 0; recalcN3(false);" />
      </td>
      <td class="num">
        <input type="number" class="excel-input" style="text-align:right; max-width:95px;" value="${row.taxa}" step="1"
          oninput="stateN3IvaRows[${idx}].taxa = parseFloat(this.value) || 0; recalcN3(false);" />
      </td>
      <td class="num td-calc" id="n3-iva-sem-${idx}">0,00 €</td>
      <td class="num td-calc" id="n3-iva-imp-${idx}">0,00 €</td>
      <td style="font-size:0.79rem; color:var(--text-muted);">${row.motivo}</td>
    `;
    tbody.appendChild(tr);
  });
}

function resetN3() {
  const d = EXCEL_DEFAULTS.n3;
  document.getElementById('n3-bruto').value = d.bruto;
  document.getElementById('n3-ss-trab').value = d.ssTrab;
  document.getElementById('n3-irs-taxa').value = d.irsTaxa;
  document.getElementById('n3-ss-emp').value = d.ssEmp;
  document.getElementById('n3-vpt').value = d.vpt;
  document.getElementById('n3-imi-taxa').value = d.imiTaxa;
  document.getElementById('n3-dep').value = d.dep;
  document.getElementById('n3-lucro').value = d.lucro;
  document.getElementById('n3-irc-pme').value = d.ircPme;
  document.getElementById('n3-irc-geral').value = d.ircGeral;
  stateN3IvaRows = JSON.parse(JSON.stringify(d.ivaRows));
  renderN3IvaTable();
  recalcN3(false);
}

function recalcN3(rebuildTable = false) {
  if (rebuildTable) renderN3IvaTable();

  // 1. Bruto -> Líquido
  const bruto = parseFloat(document.getElementById('n3-bruto').value) || 0;
  const ssTrab = (parseFloat(document.getElementById('n3-ss-trab').value) || 0) / 100;
  const irsTaxa = (parseFloat(document.getElementById('n3-irs-taxa').value) || 0) / 100;
  const ssEmp = (parseFloat(document.getElementById('n3-ss-emp').value) || 0) / 100;

  const b10 = bruto * ssTrab;
  const b11 = bruto * irsTaxa;
  const b12 = bruto - b10 - b11;
  const b14 = bruto * (1 + ssEmp);
  const b15 = b14 > 0 ? 1 - b12 / b14 : 0;

  document.getElementById('n3-ss-eur').textContent = fmtEUR(b10);
  document.getElementById('n3-irs-eur').textContent = fmtEUR(b11);
  document.getElementById('n3-liquido-eur').textContent = fmtEUR(b12);
  document.getElementById('n3-custo-emp').textContent = fmtEUR(b14);
  document.getElementById('n3-diff-emp').textContent = fmtPct(b15, 1);

  // 2. IVA
  let totComIva = 0;
  let totSemIva = 0;
  let totIva = 0;
  const ivaLabels = [];
  const ivaValues = [];

  stateN3IvaRows.forEach((r, idx) => {
    const comIva = Number(r.valorComIva) || 0;
    const taxa = (Number(r.taxa) || 0) / 100;
    const semIva = 1 + taxa !== 0 ? comIva / (1 + taxa) : comIva;
    const ivaPago = comIva - semIva;

    totComIva += comIva;
    totSemIva += semIva;
    totIva += ivaPago;

    const elSem = document.getElementById(`n3-iva-sem-${idx}`);
    const elImp = document.getElementById(`n3-iva-imp-${idx}`);
    if (elSem) elSem.textContent = fmtEUR(semIva);
    if (elImp) elImp.textContent = fmtEUR(ivaPago);

    ivaLabels.push(r.categoria.split(' (')[0]);
    ivaValues.push(Number(ivaPago.toFixed(2)));
  });

  const ivaAno = totIva * 12;
  const ivaPctLiq = b12 > 0 ? totIva / b12 : 0;

  document.getElementById('n3-iva-tot-com').textContent = fmtEUR(totComIva);
  document.getElementById('n3-iva-tot-sem').textContent = fmtEUR(totSemIva);
  document.getElementById('n3-iva-tot-imposto').textContent = fmtEUR(totIva);
  document.getElementById('n3-iva-resumo-extra').textContent = `IVA pago num ano: ${fmtEUR(ivaAno)} (${fmtPct(ivaPctLiq, 1)} do salário líquido)`;

  // 3. IMI
  const vpt = parseFloat(document.getElementById('n3-vpt').value) || 0;
  const imiTaxa = (parseFloat(document.getElementById('n3-imi-taxa').value) || 0) / 100;
  const dep = parseInt(document.getElementById('n3-dep').value, 10) || 0;

  const imiBruto = vpt * imiTaxa;
  const imiDed = dep >= 3 ? 140 : dep === 2 ? 70 : dep === 1 ? 30 : 0;
  const imiPagar = Math.max(0, imiBruto - imiDed);

  document.getElementById('n3-imi-bruto').textContent = fmtEUR(imiBruto);
  document.getElementById('n3-imi-ded').textContent = fmtEUR(imiDed);
  document.getElementById('n3-imi-pagar').textContent = fmtEUR(imiPagar);

  // 4. IRC
  const lucro = parseFloat(document.getElementById('n3-lucro').value) || 0;
  const ircPme = (parseFloat(document.getElementById('n3-irc-pme').value) || 0) / 100;
  const ircGeral = (parseFloat(document.getElementById('n3-irc-geral').value) || 0) / 100;

  // =MIN(B41,50000)*B42 + MAX(B41-50000,0)*B43
  const ircPagar = Math.min(Math.max(0, lucro), 50000) * ircPme + Math.max(lucro - 50000, 0) * ircGeral;
  const ircEfetiva = lucro > 0 ? ircPagar / lucro : 0;

  document.getElementById('n3-irc-pagar').textContent = fmtEUR(ircPagar);
  document.getElementById('n3-irc-efetiva').textContent = fmtPct(ircEfetiva, 1);

  // 5. Total entregue ao Estado por mês (=B10+B11+E28+B38/12)
  const totalEstadoMes = b10 + b11 + totIva + imiPagar / 12;
  const totalEstadoPctBruto = bruto > 0 ? totalEstadoMes / bruto : 0;

  // KPIs topo N3
  document.getElementById('n3-kpi-liquido').textContent = fmtEUR(b12);
  document.getElementById('n3-kpi-custo-emp').textContent = fmtEUR(b14);
  document.getElementById('n3-kpi-cunha').textContent = `Diferença empresa vs líquido: ${fmtPct(b15, 1)}`;
  document.getElementById('n3-kpi-iva-mes').textContent = `${fmtEUR(totIva)}/mês`;
  document.getElementById('n3-kpi-iva-ano').textContent = `${fmtEUR(ivaAno)}/ano (${fmtPct(ivaPctLiq, 1)} do líquido)`;
  document.getElementById('n3-kpi-total-estado').textContent = fmtEUR(totalEstadoMes);
  document.getElementById('n3-kpi-total-estado-pct').textContent = `${fmtPct(totalEstadoPctBruto, 1)} do salário bruto (SS + IRS + IVA + IMI/12)`;

  // Gráficos N3
  const ctxSal = document.getElementById('chart-n3-salario').getContext('2d');
  const salData = [Math.max(0, b12), Math.max(0, b10), Math.max(0, b11)];
  if (chartN3Salario) {
    chartN3Salario.data.datasets[0].data = salData;
    chartN3Salario.update();
  } else {
    chartN3Salario = new Chart(ctxSal, {
      type: 'doughnut',
      data: {
        labels: ['Salário líquido (€)', 'Segurança Social (€)', 'IRS (€)'],
        datasets: [
          {
            data: salData,
            backgroundColor: ['#10b981', '#f59e0b', '#ef4444'],
            borderWidth: 2,
            borderColor: '#ffffff',
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom' },
          title: { display: true, text: 'Composição do Salário Bruto (€/mês)' },
        },
      },
    });
  }

  const ctxIva = document.getElementById('chart-n3-iva').getContext('2d');
  if (chartN3Iva) {
    chartN3Iva.data.labels = ivaLabels;
    chartN3Iva.data.datasets[0].data = ivaValues;
    chartN3Iva.update();
  } else {
    chartN3Iva = new Chart(ctxIva, {
      type: 'bar',
      data: {
        labels: ivaLabels,
        datasets: [
          {
            label: 'IVA pago (€/mês)',
            data: ivaValues,
            backgroundColor: '#1f4e79',
            borderRadius: 6,
          },
        ],
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
      },
    });
  }
}

function saveReflectionN3() {
  try {
    const payload = {
      tipo: document.getElementById('n3-estado-tipo').value,
      justifica: document.getElementById('n3-justifica').value,
      troca: document.getElementById('n3-troca').value,
    };
    localStorage.setItem('finlab_n3_reflexao', JSON.stringify(payload));
  } catch (_) {}
}

// ============================================================================
// NÍVEL 4 — EMPREENDEDORISMO E SISTEMA BANCÁRIO
// ============================================================================

function renderN4SegurosTable() {
  const tbody = document.getElementById('n4-seg-tbody');
  tbody.innerHTML = '';
  stateN4Seguros.forEach((seg, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${seg.tipo}</strong></td>
      <td>${seg.obrigatorio}</td>
      <td style="font-size:0.8rem; color:var(--text-muted);">${seg.serve}</td>
      <td class="num">
        <input type="number" class="excel-input" style="text-align:right; max-width:110px;" value="${seg.premio}" step="10"
          oninput="stateN4Seguros[${idx}].premio = parseFloat(this.value) || 0; recalcN4(false);" />
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function resetN4() {
  const d = EXCEL_DEFAULTS.n4;
  document.getElementById('n4-cap-ini').value = d.capIni;
  document.getElementById('n4-reforco').value = d.reforco;
  document.getElementById('n4-taxa-juro').value = d.taxaJuro;
  document.getElementById('n4-cred-valor').value = d.credValor;
  document.getElementById('n4-cred-tan').value = d.credTan;
  document.getElementById('n4-cred-prazo').value = d.credPrazo;
  document.getElementById('n4-cred-poup-taxa').value = d.credPoupTaxa;
  document.getElementById('n4-dp-montante').value = d.dpMontante;
  document.getElementById('n4-dp-taxa').value = d.dpTaxa;
  document.getElementById('n4-ca-taxa').value = d.caTaxa;
  document.getElementById('n4-irs-cap').value = d.irsCap;
  document.getElementById('n4-inflacao').value = d.inflacao;
  document.getElementById('n4-neg-inv').value = d.negInv;
  document.getElementById('n4-neg-fixos').value = d.negFixos;
  document.getElementById('n4-neg-preco').value = d.negPreco;
  document.getElementById('n4-neg-var').value = d.negVar;
  document.getElementById('n4-neg-unid').value = d.negUnid;
  stateN4Seguros = JSON.parse(JSON.stringify(d.seguros));
  renderN4SegurosTable();
  recalcN4(false);
}

function recalcN4(rebuildTable = false) {
  if (rebuildTable) renderN4SegurosTable();

  // 1. Juros Simples vs Compostos (20 anos)
  const b7 = parseFloat(document.getElementById('n4-cap-ini').value) || 0;
  const b8 = parseFloat(document.getElementById('n4-reforco').value) || 0;
  const b9 = (parseFloat(document.getElementById('n4-taxa-juro').value) || 0) / 100;

  const years = [];
  const arrInv = [];
  const arrSimp = [];
  const arrComp = [];

  const jTbody = document.getElementById('n4-juros-tbody');
  jTbody.innerHTML = '';

  let inv20 = 0;
  let simp20 = 0;
  let comp20 = 0;
  let diff20 = 0;

  for (let a = 0; a <= 20; a++) {
    // B(a) = B7 + B8*12*a
    const totInv = b7 + b8 * 12 * a;
    // C(a) = B7*(1+B9*a) + B8*12*a
    const jSimples = b7 * (1 + b9 * a) + b8 * 12 * a;
    // D(a) = B7*(1+B9)^a + IF(B9=0, B8*12*a, B8*12*((1+B9)^a-1)/B9)
    const jCompostos =
      b7 * Math.pow(1 + b9, a) +
      (b9 === 0 ? b8 * 12 * a : (b8 * 12 * (Math.pow(1 + b9, a) - 1)) / b9);
    const diff = jCompostos - jSimples;

    if (a === 20) {
      inv20 = totInv;
      simp20 = jSimples;
      comp20 = jCompostos;
      diff20 = diff;
    }

    years.push(`Ano ${a}`);
    arrInv.push(Number(totInv.toFixed(2)));
    arrSimp.push(Number(jSimples.toFixed(2)));
    arrComp.push(Number(jCompostos.toFixed(2)));

    const tr = document.createElement('tr');
    if (a === 20) tr.className = 'row-total';
    tr.innerHTML = `
      <td><strong>Ano ${a}</strong></td>
      <td class="num td-calc">${fmtEUR(totInv)}</td>
      <td class="num td-calc">${fmtEUR(jSimples)}</td>
      <td class="num td-calc">${fmtEUR(jCompostos)}</td>
      <td class="num td-result">${fmtEUR(diff)}</td>
    `;
    jTbody.appendChild(tr);
  }

  document.getElementById('n4-inv-20').textContent = fmtEUR(inv20);
  document.getElementById('n4-simp-20').textContent = fmtEUR(simp20);
  document.getElementById('n4-comp-20').textContent = fmtEUR(comp20);
  document.getElementById('n4-diff-20').textContent = fmtEUR(diff20);
  document.getElementById('n4-kpi-composto').textContent = `+${fmtEUR(diff20)}`;

  // 2. Crédito Pessoal — Quanto custa mesmo?
  const b36 = parseFloat(document.getElementById('n4-cred-valor').value) || 0;
  const b37 = (parseFloat(document.getElementById('n4-cred-tan').value) || 0) / 100;
  const b38 = parseInt(document.getElementById('n4-cred-prazo').value, 10) || 1;
  const b39 = (parseFloat(document.getElementById('n4-cred-poup-taxa').value) || 0) / 100;

  // B40 = -PMT(B37/12, B38, B36)
  const rMes = b37 / 12;
  const pmt =
    b38 <= 0
      ? 0
      : rMes === 0
      ? b36 / b38
      : (b36 * rMes) / (1 - Math.pow(1 + rMes, -b38));

  const b41 = pmt * b38;
  const b42 = b41 - b36;
  const b43 = b36 > 0 ? b42 / b36 : 0;

  // B44 = IF(B39=0, B40*B38, B40*(((1+B39/12)^B38-1)/(B39/12)))
  const rPoupMes = b39 / 12;
  const b44 =
    rPoupMes === 0
      ? pmt * b38
      : pmt * ((Math.pow(1 + rPoupMes, b38) - 1) / rPoupMes);
  const b45 = b44 - b36;

  document.getElementById('n4-cred-pmt').textContent = fmtEUR(pmt);
  document.getElementById('n4-cred-total').textContent = fmtEUR(b41);
  document.getElementById('n4-cred-juros').textContent = fmtEUR(b42);
  document.getElementById('n4-cred-pct').textContent = fmtPct(b43, 1);
  document.getElementById('n4-cred-poup-acum').textContent = fmtEUR(b44);
  document.getElementById('n4-cred-diff').textContent = fmtEUR(b45);

  document.getElementById('n4-kpi-custo-cred').textContent = fmtEUR(b42);
  document.getElementById('n4-kpi-pressa').textContent = `Preço da pressa vs poupar: ${fmtEUR(b45)}`;

  // 3. Depósito a Prazo vs Certificados de Aforro
  const b48 = parseFloat(document.getElementById('n4-dp-montante').value) || 0;
  const b49 = (parseFloat(document.getElementById('n4-dp-taxa').value) || 0) / 100;
  const b50 = (parseFloat(document.getElementById('n4-ca-taxa').value) || 0) / 100;
  const b51 = (parseFloat(document.getElementById('n4-irs-cap').value) || 0) / 100;
  const b52 = (parseFloat(document.getElementById('n4-inflacao').value) || 0) / 100;

  const prodRows = [
    { nome: 'Depósito a prazo', taxa: b49 },
    { nome: 'Certificados de Aforro (Série F)', taxa: b50 },
  ];

  const dpTbody = document.getElementById('n4-dp-tbody');
  dpTbody.innerHTML = '';
  prodRows.forEach((p) => {
    const jBruto = b48 * p.taxa;
    const irsRet = jBruto * b51;
    const jLiq = jBruto - irsRet;
    const taxaLiq = b48 > 0 ? jLiq / b48 : 0;
    const taxaReal = 1 + b52 !== 0 ? (1 + taxaLiq) / (1 + b52) - 1 : 0;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${p.nome}</strong></td>
      <td class="num td-calc">${fmtEUR(jBruto)}</td>
      <td class="num td-calc">${fmtEUR(irsRet)}</td>
      <td class="num td-calc">${fmtEUR(jLiq)}</td>
      <td class="num td-calc">${fmtPct(taxaLiq, 2)}</td>
      <td class="num td-result" style="color:${taxaReal < 0 ? '#dc2626' : '#166534'}">${(taxaReal >= 0 ? '+' : '') + fmtPct(taxaReal, 2)}</td>
    `;
    dpTbody.appendChild(tr);
  });

  // 4. Ter um Negócio — Ponto de Equilíbrio
  const b59 = parseFloat(document.getElementById('n4-neg-inv').value) || 0;
  const b60 = parseFloat(document.getElementById('n4-neg-fixos').value) || 0;
  const b61 = parseFloat(document.getElementById('n4-neg-preco').value) || 0;
  const b62 = parseFloat(document.getElementById('n4-neg-var').value) || 0;
  const b63 = parseFloat(document.getElementById('n4-neg-unid').value) || 0;

  const b64 = b61 - b62;
  const b65 = b64 > 0 ? b60 / b64 : 0;
  const b66 = b63 * b64 - b60;
  const b67 = b66 <= 0 ? 'Nunca' : `${fmtNum(b59 / b66, 1)} meses`;

  document.getElementById('n4-neg-margem').textContent = fmtEUR(b64);
  document.getElementById('n4-neg-be').textContent = `${fmtNum(b65, 0)} unidades`;
  document.getElementById('n4-neg-lucro').textContent = fmtEUR(b66);
  document.getElementById('n4-neg-payback').textContent = b67;

  document.getElementById('n4-kpi-breakeven').textContent = `${fmtNum(b65, 0)} unid./mês`;
  document.getElementById('n4-kpi-lucro-neg').textContent = `Lucro mensal: ${fmtEUR(b66)} (Payback: ${b67})`;

  // 5. Seguros
  const totSegAnual = stateN4Seguros.reduce((acc, s) => acc + (Number(s.premio) || 0), 0);
  const segMensal = totSegAnual / 12;

  document.getElementById('n4-seg-total').textContent = fmtEUR(totSegAnual);
  document.getElementById('n4-seg-mensal').textContent = fmtEUR(segMensal);
  document.getElementById('n4-kpi-seguros').textContent = `${fmtEUR(segMensal)}/mês`;
  document.getElementById('n4-kpi-seguros-ano').textContent = `Total anual: ${fmtEUR(totSegAnual)}/ano`;

  // Gráficos N4
  const ctxJuros = document.getElementById('chart-n4-juros').getContext('2d');
  if (chartN4Juros) {
    chartN4Juros.data.labels = years;
    chartN4Juros.data.datasets[0].data = arrInv;
    chartN4Juros.data.datasets[1].data = arrSimp;
    chartN4Juros.data.datasets[2].data = arrComp;
    chartN4Juros.update();
  } else {
    chartN4Juros = new Chart(ctxJuros, {
      type: 'line',
      data: {
        labels: years,
        datasets: [
          {
            label: 'Total investido (€)',
            data: arrInv,
            borderColor: '#64748b',
            borderDash: [4, 4],
            tension: 0.2,
          },
          {
            label: 'Juros simples (€)',
            data: arrSimp,
            borderColor: '#2563eb',
            tension: 0.2,
            borderWidth: 2,
          },
          {
            label: 'Juros compostos (€)',
            data: arrComp,
            borderColor: '#ea580c',
            backgroundColor: 'rgba(234, 88, 12, 0.12)',
            fill: true,
            tension: 0.25,
            borderWidth: 3,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
      },
    });
  }

  const ctxCred = document.getElementById('chart-n4-credito').getContext('2d');
  const credVals = [Number(b36.toFixed(2)), Number(b41.toFixed(2)), Number(b44.toFixed(2))];
  if (chartN4Credito) {
    chartN4Credito.data.datasets[0].data = credVals;
    chartN4Credito.update();
  } else {
    chartN4Credito = new Chart(ctxCred, {
      type: 'bar',
      data: {
        labels: [
          'Valor pedido ao banco',
          'Total pago ao banco (c/ juros)',
          'Se poupasse a mesma prestação',
        ],
        datasets: [
          {
            label: 'Montante (€)',
            data: credVals,
            backgroundColor: ['#64748b', '#dc2626', '#10b981'],
            borderRadius: 8,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
      },
    });
  }
}

// ============================================================================
// NÍVEL 5 — CRIAÇÃO DE UM PATRIMÓNIO / PORTFÓLIO
// ============================================================================

function renderN5InstTable() {
  const tbody = document.getElementById('n5-inst-tbody');
  tbody.innerHTML = '';
  stateN5Inst.forEach((inst, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${inst.nome}</strong></td>
      <td class="num">
        <input type="number" class="excel-input" style="text-align:right; max-width:95px;" value="${inst.pesoPct}" step="5"
          oninput="stateN5Inst[${idx}].pesoPct = parseFloat(this.value) || 0; recalcN5(false);" />
      </td>
      <td class="num">
        <input type="number" class="excel-input" style="text-align:right; max-width:95px;" value="${inst.retornoPct}" step="0.5"
          oninput="stateN5Inst[${idx}].retornoPct = parseFloat(this.value) || 0; recalcN5(false);" />
      </td>
      <td>${inst.risco}</td>
      <td style="font-size:0.8rem; color:var(--text-muted);">${inst.liquidez}</td>
      <td class="num td-calc" id="n5-inst-mes-${idx}">0,00 €</td>
      <td class="num td-calc" id="n5-inst-10a-${idx}">0,00 €</td>
      <td class="num td-calc" id="n5-inst-20a-${idx}">0,00 €</td>
    `;
    tbody.appendChild(tr);
  });
}

function syncFromN1ToN5() {
  const poupN1 = parseFloat(document.getElementById('n1-poupanca').value) || 150;
  const totDespN1 = stateN1Rows.reduce((acc, r) => acc + (Number(r.valor) || 0), 0);
  document.getElementById('n5-poup-mensal').value = poupN1;
  if (totDespN1 > 0) document.getElementById('n5-despesas').value = totDespN1;
  recalcN5(false);
}

function resetN5() {
  const d = EXCEL_DEFAULTS.n5;
  document.getElementById('n5-cap-ini').value = d.capIni;
  document.getElementById('n5-poup-mensal').value = d.poupMensal;
  document.getElementById('n5-despesas').value = d.despesas;
  document.getElementById('n5-inflacao').value = d.inflacao;
  stateN5Inst = JSON.parse(JSON.stringify(d.instrumentos));
  renderN5InstTable();
  recalcN5(false);
}

function recalcN5(rebuildTable = false) {
  if (rebuildTable) renderN5InstTable();

  const b7 = parseFloat(document.getElementById('n5-cap-ini').value) || 0;
  const b8 = parseFloat(document.getElementById('n5-poup-mensal').value) || 0;
  const b9 = parseFloat(document.getElementById('n5-despesas').value) || 0;
  const b10 = (parseFloat(document.getElementById('n5-inflacao').value) || 0) / 100;

  let sumPeso = 0;
  let weightedRet = 0;
  let sumEurMes = 0;
  let sum10a = 0;
  let sum20a = 0;

  const alocLabels = [];
  const alocEurVals = [];

  stateN5Inst.forEach((inst, idx) => {
    const bRow = (Number(inst.pesoPct) || 0) / 100;
    const cRow = (Number(inst.retornoPct) || 0) / 100;
    const fRow = b8 * bRow;
    const gRow =
      cRow === 0 ? fRow * 12 * 10 : (fRow * 12 * (Math.pow(1 + cRow, 10) - 1)) / cRow;
    const hRow =
      cRow === 0 ? fRow * 12 * 20 : (fRow * 12 * (Math.pow(1 + cRow, 20) - 1)) / cRow;

    sumPeso += bRow;
    weightedRet += bRow * cRow;
    sumEurMes += fRow;
    sum10a += gRow;
    sum20a += hRow;

    const elMes = document.getElementById(`n5-inst-mes-${idx}`);
    const el10 = document.getElementById(`n5-inst-10a-${idx}`);
    const el20 = document.getElementById(`n5-inst-20a-${idx}`);
    if (elMes) elMes.textContent = fmtEUR(fRow);
    if (el10) el10.textContent = fmtEUR(gRow);
    if (el20) el20.textContent = fmtEUR(hRow);

    alocLabels.push(inst.nome.split(' (')[0]);
    alocEurVals.push(Number(fRow.toFixed(2)));
  });

  document.getElementById('n5-tot-pct').textContent = fmtPct(sumPeso, 1);
  document.getElementById('n5-tot-ret').textContent = fmtPct(weightedRet, 2);
  document.getElementById('n5-tot-eur').textContent = fmtEUR(sumEurMes);
  document.getElementById('n5-tot-10a').textContent = fmtEUR(sum10a);
  document.getElementById('n5-tot-20a').textContent = fmtEUR(sum20a);

  const badge = document.getElementById('n5-aloc-badge');
  if (Math.abs(sumPeso - 1) < 0.0001) {
    badge.className = 'alert-badge ok';
    badge.textContent = 'OK — 100% da poupança alocada';
  } else {
    badge.className = 'alert-badge warn';
    badge.textContent = `ATENÇÃO: a soma das percentagens tem de ser 100% (atual: ${fmtPct(sumPeso, 1)})`;
  }

  // 3. Fundo de emergência
  const feRec = b9 * 6;
  const f15 = b8 * ((Number(stateN5Inst[0].pesoPct) || 0) / 100);
  const mesesFe = f15 <= 0 ? 'Sem verba' : `${fmtNum(Math.max(0, (feRec - b7) / f15), 1)} meses`;

  document.getElementById('n5-fe-rec').textContent = fmtEUR(feRec);
  document.getElementById('n5-fe-mensal').textContent = fmtEUR(f15);
  document.getElementById('n5-fe-tempo').textContent = mesesFe;

  document.getElementById('n5-kpi-fe-meses').textContent = mesesFe;
  document.getElementById('n5-kpi-fe-meta').textContent = `Meta recomendada: ${fmtEUR(feRec)} (${fmtEUR(f15)}/mês)`;

  // 4. Evolução do Património (Ano 0 a 20)
  const years = [];
  const arrParado = [];
  const arrInvestido = [];

  const evolTbody = document.getElementById('n5-evol-tbody');
  evolTbody.innerHTML = '';

  let b51 = 0;
  let c51 = 0;
  let e51 = 0;

  for (let a = 0; a <= 20; a++) {
    // B(a) = B7 + B8*12*a
    const parado = b7 + b8 * 12 * a;
    // C(a) = B7*(1+C20)^a + IF(C20=0, B8*12*a, B8*12*((1+C20)^a-1)/C20)
    const investido =
      b7 * Math.pow(1 + weightedRet, a) +
      (weightedRet === 0
        ? b8 * 12 * a
        : (b8 * 12 * (Math.pow(1 + weightedRet, a) - 1)) / weightedRet);
    const diff = investido - parado;
    const hoje = Math.pow(1 + b10, a) !== 0 ? investido / Math.pow(1 + b10, a) : 0;

    if (a === 20) {
      b51 = parado;
      c51 = investido;
      e51 = hoje;
    }

    years.push(`Ano ${a}`);
    arrParado.push(Number(parado.toFixed(2)));
    arrInvestido.push(Number(investido.toFixed(2)));

    const tr = document.createElement('tr');
    if (a === 20) tr.className = 'row-total';
    tr.innerHTML = `
      <td><strong>Ano ${a}</strong></td>
      <td class="num td-calc">${fmtEUR(parado)}</td>
      <td class="num td-calc">${fmtEUR(investido)}</td>
      <td class="num td-result">${fmtEUR(diff)}</td>
      <td class="num td-calc">${fmtEUR(hoje)}</td>
    `;
    evolTbody.appendChild(tr);
  }

  const rendGerados = c51 - b51;
  const propRend = c51 > 0 ? 1 - b51 / c51 : 0;

  document.getElementById('n5-kpi-patrimonio').textContent = fmtEUR(c51);
  document.getElementById('n5-kpi-bolso').textContent = `Do teu bolso saíram: ${fmtEUR(b51)}`;
  document.getElementById('n5-kpi-rendimentos').textContent = fmtEUR(rendGerados);
  document.getElementById('n5-kpi-prop').textContent = `${fmtPct(propRend, 1)} do património vem dos rendimentos`;
  document.getElementById('n5-kpi-hoje').textContent = fmtEUR(e51);

  // Gráficos N5
  const ctxAloc = document.getElementById('chart-n5-aloc').getContext('2d');
  if (chartN5Aloc) {
    chartN5Aloc.data.labels = alocLabels;
    chartN5Aloc.data.datasets[0].data = alocEurVals;
    chartN5Aloc.update();
  } else {
    chartN5Aloc = new Chart(ctxAloc, {
      type: 'doughnut',
      data: {
        labels: alocLabels,
        datasets: [
          {
            data: alocEurVals,
            backgroundColor: ['#0d9488', '#2563eb', '#6366f1', '#f59e0b', '#e11d48'],
            borderWidth: 2,
            borderColor: '#ffffff',
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
      },
    });
  }

  const ctxEvol = document.getElementById('chart-n5-evol').getContext('2d');
  if (chartN5Evol) {
    chartN5Evol.data.labels = years;
    chartN5Evol.data.datasets[0].data = arrParado;
    chartN5Evol.data.datasets[1].data = arrInvestido;
    chartN5Evol.update();
  } else {
    chartN5Evol = new Chart(ctxEvol, {
      type: 'line',
      data: {
        labels: years,
        datasets: [
          {
            label: 'Dinheiro parado (€)',
            data: arrParado,
            borderColor: '#64748b',
            borderDash: [5, 5],
            tension: 0.2,
          },
          {
            label: 'Carteira investida (€)',
            data: arrInvestido,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16, 185, 129, 0.14)',
            fill: true,
            tension: 0.25,
            borderWidth: 3,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
      },
    });
  }
}

// ============================================================================
// REPOR TODOS OS VALORES DO EXCEL E INICIALIZAÇÃO
// ============================================================================

function resetAllToExcelDefaults() {
  loadN1Preset('excel');
  resetN2();
  resetN3();
  resetN4();
  resetN5();
  document.getElementById('demo-bruto').value = 920;
  updateDemoLegend();
}

window.addEventListener('DOMContentLoaded', () => {
  updateDemoLegend();
  recalcN1(true);
  recalcN2();
  recalcN3(true);
  recalcN4(true);
  recalcN5(true);

  try {
    const saved = localStorage.getItem('finlab_n3_reflexao');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed.tipo) document.getElementById('n3-estado-tipo').value = parsed.tipo;
      if (parsed.justifica) document.getElementById('n3-justifica').value = parsed.justifica;
      if (parsed.troca) document.getElementById('n3-troca').value = parsed.troca;
    }
  } catch (_) {}
});
