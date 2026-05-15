// ============================================================
//  DASHBOARD FROTA NORTETECH — Apps Script (Code.gs)
//  Cole este código no editor do Apps Script da sua planilha.
//  Acesse: Extensões > Apps Script
// ============================================================

// Nome da aba da planilha com os dados
var SHEET_NAME = "historico_manutenções";

// Horas disponíveis por veículo por mês (ajuste conforme sua operação)
var HORAS_OPERACAO_DIA = 12;

// ------------------------------------------------------------------
// MENU PERSONALIZADO
// ------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Dashboard Frota")
    .addItem("Abrir Dashboard", "abrirDashboard")
    .addToUi();
}

// ------------------------------------------------------------------
// ABRE O DASHBOARD EM UMA JANELA MODAL GRANDE
// ------------------------------------------------------------------
function abrirDashboard() {
  var html = HtmlService.createHtmlOutputFromFile("Dashboard_Frota")
    .setWidth(1400)
    .setHeight(900)
    .setTitle("Dashboard de Manutenção da Frota");
  SpreadsheetApp.getUi().showModalDialog(html, "Dashboard de Manutenção da Frota");
}

// ------------------------------------------------------------------
// LEITURA PRINCIPAL DE DADOS DA PLANILHA
// ------------------------------------------------------------------
function getDadosBrutos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Aba '" + SHEET_NAME + "' não encontrada.");

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    rows.push(obj);
  }
  return rows;
}

// ------------------------------------------------------------------
// HELPER: parse de data flexível (Date obj, string ou número serial)
// ------------------------------------------------------------------
function parseDate(val) {
  if (!val || val === "") return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  var d = new Date(val);
  if (!isNaN(d.getTime())) return d;
  return null;
}

// ------------------------------------------------------------------
// HELPER: diferença em horas entre duas datas (ignora negativos)
// ------------------------------------------------------------------
function diffHours(d1, d2) {
  if (!d1 || !d2) return 0;
  var diff = (d2 - d1) / 3600000;
  return diff > 0 ? diff : 0;
}

// ------------------------------------------------------------------
// HELPER: mesma data civil?
// ------------------------------------------------------------------
function mesmaData(d1, d2) {
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

// ------------------------------------------------------------------
// RETORNA OS FILTROS DISPONÍVEIS (meses, oficinas, modelos)
// ------------------------------------------------------------------
function getFiltros() {
  var rows = getDadosBrutos();
  var meses = {}, oficinas = {}, modelos = {};
  rows.forEach(function(r) {
    if (r.ref_manutencao) meses[r.ref_manutencao] = true;
    if (r.oficina && r.oficina !== "") oficinas[r.oficina] = true;
    if (r.modelo_veiculo && r.modelo_veiculo !== "") modelos[r.modelo_veiculo] = true;
  });
  return {
    meses: Object.keys(meses).sort().reverse(),
    oficinas: ["Todas"].concat(Object.keys(oficinas).sort()),
    modelos: ["Todos"].concat(Object.keys(modelos).sort())
  };
}

// ------------------------------------------------------------------
// FUNÇÃO PRINCIPAL: calcula todos os KRs com os filtros aplicados
// ------------------------------------------------------------------
function calcularKRs(filtroMes, filtroOficina, filtroModelo) {
  var rows = getDadosBrutos();

  // --- Filtragem ---
  var filtrado = rows.filter(function(r) {
    var ok = true;
    if (filtroMes && filtroMes !== "Todos") ok = ok && (r.ref_manutencao === filtroMes);
    if (filtroOficina && filtroOficina !== "Todas") ok = ok && (r.oficina === filtroOficina);
    if (filtroModelo && filtroModelo !== "Todos") ok = ok && (r.modelo_veiculo === filtroModelo);
    return ok;
  });

  // ----------------------------------------------------------------
  // KR1 — Taxa de Disponibilidade da Frota
  // ----------------------------------------------------------------
  var placasUnicas = {};
  filtrado.forEach(function(r) { if (r.placa) placasUnicas[r.placa] = true; });
  var totalPlacas = Object.keys(placasUnicas).length;

  // Dias no mês filtrado (usa o primeiro mês do filtro ou 30 padrão)
  var diasMes = 30;
  if (filtroMes && filtroMes !== "Todos") {
    var partes = filtroMes.split("-");
    diasMes = new Date(parseInt(partes[0]), parseInt(partes[1]), 0).getDate();
  }
  var horasDisponiveisTotais = totalPlacas * diasMes * HORAS_OPERACAO_DIA;

  var totalDowntimeHoras = 0;
  filtrado.forEach(function(r) {
    var ab = parseDate(r.data_abertura);
    var sa = parseDate(r.data_saida);
    totalDowntimeHoras += diffHours(ab, sa);
  });
  // Deduplica downtime por OS (mesma OS pode ter múltiplas linhas de itens)
  var osDt = {};
  filtrado.forEach(function(r) {
    var cod = r.cod_osm;
    if (!cod) return;
    var ab = parseDate(r.data_abertura);
    var sa = parseDate(r.data_saida);
    var h = diffHours(ab, sa);
    if (!osDt[cod] || h > osDt[cod]) osDt[cod] = h;
  });
  var downtimeReal = Object.values(osDt).reduce(function(a, b) { return a + b; }, 0);

  var taxaDisponibilidade = horasDisponiveisTotais > 0
    ? Math.min(100, Math.max(0, ((horasDisponiveisTotais - downtimeReal) / horasDisponiveisTotais) * 100))
    : 0;

  // ----------------------------------------------------------------
  // KR2 — Downtime Total (horas) vs mês anterior
  // ----------------------------------------------------------------
  var downtimeMesAtual = downtimeReal;
  var downtimeMesAnterior = 0;
  if (filtroMes && filtroMes !== "Todos") {
    var partesMes = filtroMes.split("-");
    var ano = parseInt(partesMes[0]), mes = parseInt(partesMes[1]);
    var mesAnt = mes === 1 ? (ano - 1) + "-12" : ano + "-" + String(mes - 1).padStart(2, "0");
    var rowsAnt = getDadosBrutos().filter(function(r) { return r.ref_manutencao === mesAnt; });
    var osDtAnt = {};
    rowsAnt.forEach(function(r) {
      var cod = r.cod_osm;
      var ab = parseDate(r.data_abertura);
      var sa = parseDate(r.data_saida);
      var h = diffHours(ab, sa);
      if (!osDtAnt[cod] || h > osDtAnt[cod]) osDtAnt[cod] = h;
    });
    downtimeMesAnterior = Object.values(osDtAnt).reduce(function(a, b) { return a + b; }, 0);
  }

  // ----------------------------------------------------------------
  // KR3 — Top 5 Veículos Ofensores
  // ----------------------------------------------------------------
  var placaDowntime = {};
  Object.keys(osDt).forEach(function(cod) {
    // encontra a placa desta OS
    var linha = filtrado.find(function(r) { return String(r.cod_osm) === String(cod); });
    if (linha && linha.placa) {
      placaDowntime[linha.placa] = (placaDowntime[linha.placa] || 0) + osDt[cod];
    }
  });
  var top5 = Object.entries(placaDowntime)
    .sort(function(a, b) { return b[1] - a[1]; })
    .slice(0, 5)
    .map(function(e) { return { placa: e[0], horas: Math.round(e[1] * 10) / 10 }; });

  // ----------------------------------------------------------------
  // KR4 — MTTR por Oficina (Pesada, Lanchas, Funilaria)
  // ----------------------------------------------------------------
  var gruposPesados = ["PESADA", "LANCHAS", "FUNILARIA"];
  var mttrPorGrupo = {};
  gruposPesados.forEach(function(g) {
    var osGrupo = {};
    filtrado.forEach(function(r) {
      if (!r.oficina || !String(r.oficina).toUpperCase().includes(g)) return;
      var cod = r.cod_osm;
      var ab = parseDate(r.data_abertura);
      var sa = parseDate(r.data_saida);
      var h = diffHours(ab, sa);
      if (!osGrupo[cod]) osGrupo[cod] = { horas: 0, nome: g };
      if (h > osGrupo[cod].horas) osGrupo[cod].horas = h;
    });
    var vals = Object.values(osGrupo).map(function(o) { return o.horas; });
    var media = vals.length > 0 ? vals.reduce(function(a, b) { return a + b; }, 0) / vals.length : 0;
    mttrPorGrupo[g] = { mttr: Math.round(media * 10) / 10, total: vals.length };
  });
  // Também mostra LEVE no KR4 como referência
  var osLeve = {};
  filtrado.forEach(function(r) {
    if (!r.oficina || !String(r.oficina).toUpperCase().includes("LEVE")) return;
    var cod = r.cod_osm;
    var ab = parseDate(r.data_abertura);
    var sa = parseDate(r.data_saida);
    var h = diffHours(ab, sa);
    if (!osLeve[cod] || h > osLeve[cod]) osLeve[cod] = h;
  });
  var valsLeve = Object.values(osLeve);
  var mttrLeve = valsLeve.length > 0 ? valsLeve.reduce(function(a, b) { return a + b; }, 0) / valsLeve.length : 0;
  mttrPorGrupo["LEVE"] = { mttr: Math.round(mttrLeve * 10) / 10, total: valsLeve.length };

  // ----------------------------------------------------------------
  // KR5 — % Same-Day Oficina Leve
  // ----------------------------------------------------------------
  var osLeveTotal = {}, osLeveSameDay = {};
  filtrado.forEach(function(r) {
    if (!r.oficina || !String(r.oficina).toUpperCase().includes("LEVE")) return;
    var cod = r.cod_osm;
    var ab = parseDate(r.data_abertura);
    var sa = parseDate(r.data_saida);
    osLeveTotal[cod] = true;
    if (ab && sa && mesmaData(ab, sa)) osLeveSameDay[cod] = true;
  });
  var totalOsLeve = Object.keys(osLeveTotal).length;
  var sameDayCount = Object.keys(osLeveSameDay).length;
  var pctSameDay = totalOsLeve > 0 ? (sameDayCount / totalOsLeve) * 100 : 0;

  // ----------------------------------------------------------------
  // KR6 — Preventiva vs Corretiva por mês (últimos 6 meses)
  // ----------------------------------------------------------------
  var allRows = getDadosBrutos();
  var mesesUnicos = [];
  if (filtroMes && filtroMes !== "Todos") {
    var p = filtroMes.split("-");
    var a = parseInt(p[0]), m = parseInt(p[1]);
    for (var i = 5; i >= 0; i--) {
      var mi = m - i;
      var ai = a;
      while (mi <= 0) { mi += 12; ai--; }
      mesesUnicos.push(ai + "-" + String(mi).padStart(2, "0"));
    }
  } else {
    mesesUnicos = Array.from(new Set(allRows.map(function(r) { return r.ref_manutencao; }).filter(Boolean)))
      .sort().slice(-6);
  }

  var prevCorr = mesesUnicos.map(function(mes) {
    var osPreventiva = {}, osCorretiva = {}, osOther = {};
    allRows.forEach(function(r) {
      if (r.ref_manutencao !== mes) return;
      if (filtroOficina && filtroOficina !== "Todas" && r.oficina !== filtroOficina) return;
      if (filtroModelo && filtroModelo !== "Todos" && r.modelo_veiculo !== filtroModelo) return;
      var cod = r.cod_osm;
      var tipo = String(r.tipo_manutencao || "").toUpperCase();
      if (tipo.includes("PREVENTIVA")) osPreventiva[cod] = true;
      else if (tipo.includes("CORRETIVA")) osCorretiva[cod] = true;
      else if (cod) osOther[cod] = true;
    });
    return {
      mes: mes,
      preventiva: Object.keys(osPreventiva).length,
      corretiva: Object.keys(osCorretiva).length,
      outros: Object.keys(osOther).length
    };
  });

  // ----------------------------------------------------------------
  // KR7 — Taxa de Encerramento de OS
  // ----------------------------------------------------------------
  var osUnicas = {}, osFinalizadas = {};
  filtrado.forEach(function(r) {
    var cod = r.cod_osm;
    if (!cod) return;
    osUnicas[cod] = true;
    var sit = String(r.situacao_osm || "").toLowerCase();
    if (sit.includes("finalizada")) osFinalizadas[cod] = true;
  });
  var totalOsUnicas = Object.keys(osUnicas).length;
  var totalFinalizadas = Object.keys(osFinalizadas).length;
  var taxaEncerramento = totalOsUnicas > 0 ? (totalFinalizadas / totalOsUnicas) * 100 : 0;

  // ----------------------------------------------------------------
  // KR8 — Backlog (OS abertas há mais de 5 dias)
  // ----------------------------------------------------------------
  var hoje = new Date();
  var backlog = [];
  var osVistas = {};
  filtrado.forEach(function(r) {
    var cod = r.cod_osm;
    if (!cod || osVistas[cod]) return;
    var sit = String(r.situacao_osm || "").toLowerCase();
    var emAberto = !sit.includes("finalizada") && !sit.includes("cancelado");
    if (!emAberto) return;
    var ab = parseDate(r.data_abertura);
    if (!ab) return;
    var diasAberto = Math.floor((hoje - ab) / 86400000);
    if (diasAberto >= 5) {
      backlog.push({
        cod_osm: cod,
        placa: r.placa,
        modelo: r.modelo_veiculo,
        oficina: String(r.oficina || "").replace("NORTE TECH SERVICOS EM ENERGIA LTDA - ", ""),
        dias: diasAberto,
        situacao: r.situacao_osm,
        abertura: ab ? Utilities.formatDate(ab, Session.getScriptTimeZone(), "dd/MM/yyyy") : ""
      });
      osVistas[cod] = true;
    }
  });
  backlog.sort(function(a, b) { return b.dias - a.dias; });

  // ----------------------------------------------------------------
  // KR9 — Qualidade de Lançamento de Peças
  // ----------------------------------------------------------------
  var osComItens = {}, osSemItens = {};
  filtrado.forEach(function(r) {
    var cod = r.cod_osm;
    if (!cod) return;
    if (r.cod_manu_item && r.cod_manu_item !== "") osComItens[cod] = true;
    else if (!osComItens[cod]) osSemItens[cod] = true;
  });
  // Remove as que têm itens do conjunto sem-itens
  Object.keys(osComItens).forEach(function(k) { delete osSemItens[k]; });
  var totalComItens = Object.keys(osComItens).length;
  var totalSemItens = Object.keys(osSemItens).length;

  // ----------------------------------------------------------------
  // RETORNA TUDO
  // ----------------------------------------------------------------
  return {
    filtroMes: filtroMes,
    filtroOficina: filtroOficina,
    filtroModelo: filtroModelo,
    kr1: {
      taxaDisponibilidade: Math.round(taxaDisponibilidade * 10) / 10,
      totalPlacas: totalPlacas,
      horasDisponiveisTotais: Math.round(horasDisponiveisTotais),
      downtimeHoras: Math.round(downtimeReal * 10) / 10
    },
    kr2: {
      downtimeMesAtual: Math.round(downtimeMesAtual * 10) / 10,
      downtimeMesAnterior: Math.round(downtimeMesAnterior * 10) / 10,
      variacaoPct: downtimeMesAnterior > 0
        ? Math.round(((downtimeMesAtual - downtimeMesAnterior) / downtimeMesAnterior) * 1000) / 10
        : 0
    },
    kr3: { top5: top5 },
    kr4: { mttrPorGrupo: mttrPorGrupo },
    kr5: {
      totalOsLeve: totalOsLeve,
      sameDayCount: sameDayCount,
      pctSameDay: Math.round(pctSameDay * 10) / 10
    },
    kr6: { prevCorr: prevCorr },
    kr7: {
      totalOs: totalOsUnicas,
      finalizadas: totalFinalizadas,
      abertas: totalOsUnicas - totalFinalizadas,
      taxaEncerramento: Math.round(taxaEncerramento * 10) / 10
    },
    kr8: { backlog: backlog.slice(0, 20) },
    kr9: {
      totalOs: totalOsUnicas,
      comItens: totalComItens,
      semItens: totalSemItens,
      pctLancamento: totalOsUnicas > 0
        ? Math.round((totalComItens / totalOsUnicas) * 1000) / 10
        : 0
    }
  };
}
