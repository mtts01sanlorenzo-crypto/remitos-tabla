/* Remitos Betowin (.xls) -> filas para pegar en TABLA GENERAL.
   Port a JavaScript de betowin_parser.py + tabla_general.py + exportar.py.
   Todo corre en el navegador; no se sube nada a ningún servidor. */
(function (root) {
  "use strict";

  const HEADERS = root.HEADERS_TG || (typeof require !== "undefined" ? require("./headers.js") : []);

  // ---------- utilidades de columnas ----------
  const colIdx = (L) => L.split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0); // "A"->1
  const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  // ---------- mapeos ----------
  const INSUMO_MAP = {
    "P 67": "P67", "CH 67": "P67", "P 57": "P 57", "P 89": "P 89", "ARE 1": "ARE 1", "ARE 2": "ARE 2",
    "CM-MS": "CM-MS", "CM-MS2": "CM-MS", "CEM HS": "CEM HS", "CEM T1": "CEM T1", "CEM TV": "CEM TV",
    "AG": "AGUA", "SP4": "PLASTSP4", "R-38": "R-38", "R 2027": "R 2027",
  };
  const COLS = {
    "P67": ["AE", "AU"], "P 57": ["AF", "AV"], "P 89": ["AG", "AW"], "ARE 1": ["AH", "AX"], "ARE 2": ["AI", "AY"],
    "CM-MS": ["AJ", "AZ"], "CEM HS": ["AK", "BA"], "CEM T1": ["AL", "BC"], "CEM TV": ["AM", "BB"],
    "AGUA": ["AN", "BD"], "PLASTSP4": ["AO", "BE"], "R-38": ["AP", "BF"], "R 2027": ["AQ", "BG"],
  };
  const FROM_REMITO = {
    "A": "Informe total datos de ciclo – N.º", "B": "Cliente (tal cual: '000003 - NOMBRE')",
    "F": "Hora PRINCIPIO ('Ora  hh:mm:ss')", "H": "Hora FIN ('Ora  hh:mm:ss')", "L": "Fecha PRINCIPIO",
    "M": "'Codigo: …' de la fórmula", "Q": "'Nombre fórmula: …'", "AD": "m³ Totales",
  };
  Object.entries(COLS).forEach(([ins, [t, d]]) => {
    FROM_REMITO[t] = `Columna 'Cor.' de ${ins} (suma del remito, kg)`;
    FROM_REMITO[d] = `Columna 'Dosif.' de ${ins} (suma del remito, kg)`;
  });
  Object.assign(FROM_REMITO, {
    "O": "Derivado del nombre de la fórmula (resistencia + cemento + slump). FORMATO POR CONFIRMAR",
    "P": "Derivado del nombre de la fórmula (B=BOMBEADO, D=DIRECTO). Vacío si el nombre no trae B/D",
    "S": "No viene en el remito: lo escribes tú en 'Planta' al procesar",
  });
  const CANDIDATE = { "U": "Posible: 'Albarán n.º' (confirmar si es tu N° de guía)" };
  const FORMULA_COLS = "G I J K N R AC AR AS AT BH BI BJ BK BL BM BN BO BP BQ BR BS BT BU BW BX BY BZ CA CB CC CD CF CG CH CI CO CP".split(" ");

  const CEMENTO_TXT = { "TI": "TIPO I", "T1": "TIPO I", "MS": "TIPO MS", "TV": "TIPO V", "HS": "TIPO HS" };
  const SERVICIO_TXT = { "B": "BOMBEADO", "D": "DIRECTO" };

  // ---------- parser Betowin ----------
  const RE_NUM = /-?\d+(?:[.,]\d+)?/;
  const norm = (v) => String(v).replace(/ /g, " ").trim().toLowerCase();
  const str = (v) => (v === undefined || v === null ? "" : String(v));

  function num(v) { // -> [valor|null, manual, fuera]
    if (typeof v === "number") return [v, false, false];
    const s = str(v).trim();
    if (!s) return [null, false, false];
    const manual = s.toUpperCase().startsWith("M ");
    const fuera = s.startsWith("*");
    const m = s.match(RE_NUM);
    return [m ? parseFloat(m[0].replace(",", ".")) : null, manual, fuera];
  }
  function fecha(s) { // -> {y,m,d} | null
    const m = str(s).match(/(\d{2})\/(\d{2})\/(\d{2,4})/);
    if (!m) return null;
    let y = parseInt(m[3], 10);
    if (m[3].length === 2) y += 2000;
    return { y, m: parseInt(m[2], 10), d: parseInt(m[1], 10) };
  }
  const fechaKey = (f) => (f ? `${f.y}-${f.m}-${f.d}` : "");
  const fechaSerial = (f) => Date.UTC(f.y, f.m - 1, f.d) / 86400000 + 25569;
  function hora(s) { const m = str(s).match(/(\d{2}:\d{2}:\d{2})/); return m ? m[1] : null; }
  function after(s, label) { s = str(s); const i = s.indexOf(label); return i >= 0 ? s.slice(i + label.length).trim() : ""; }

  function parseSheet(grid, nombreArchivo) {
    const nrows = grid.length;
    const ncols = grid.reduce((a, r) => Math.max(a, r.length), 0);
    const cell = (r, c) => (r < nrows && c < grid[r].length && grid[r][c] !== undefined && grid[r][c] !== null ? grid[r][c] : "");
    const remitos = [], consumos = [], ciclos = [];

    const starts = [];
    for (let r = 0; r < nrows; r++) if (str(cell(r, 0)).startsWith("Informe")) starts.push(r);

    starts.forEach((r0, bi) => {
      const r1 = bi + 1 < starts.length ? starts[bi + 1] : nrows;
      const titulo = str(cell(r0, 0));
      const esRemito = titulo.startsWith("Informe total");
      const informeId = titulo.slice(titulo.lastIndexOf("-") + 1).trim();
      const d = { archivo: nombreArchivo, informe: informeId };

      for (let r = r0; r < r1; r++) {
        if (str(cell(r, 0)).startsWith("Albar")) {
          const v = str(cell(r + 1, 0));
          d.albaran = v !== "" ? parseInt(v, 10) : null;
          if (Number.isNaN(d.albaran)) d.albaran = null;
          const cli = str(cell(r + 1, 1));
          const k = cli.indexOf(" - ");
          d.cliente_codigo = k >= 0 ? cli.slice(0, k) : cli;
          d.cliente = k >= 0 ? cli.slice(k + 3) : "";
          d.cod_hormigon = str(cell(r + 1, 3)).replace(".0", "");
          break;
        }
      }
      for (let r = r0; r < r1; r++) {
        const a = str(cell(r, 0)), b = str(cell(r, 1));
        if (a === "Subciclo") {
          d.m3_subciclo = num(str(cell(r + 1, 0)).replace("m3", ""))[0];
          d.m3 = num(str(cell(r + 1, 1)).replace("m3", ""))[0];
        } else if (a.startsWith("Pt.carga")) {
          d.punto_carga = after(a, ":");
        } else if (a === "PRINCIPIO") {
          d.fecha = fecha(cell(r + 1, 0));
          d.hora_inicio = hora(cell(r + 2, 0));
          d.hora_fin = hora(cell(r + 2, 1));
        } else if (a.startsWith("Nr.")) {
          d.diseno_codigo = after(b, "Codigo:");
          d.diseno = after(str(cell(r, 2)), "Nombre fórmula:");
        } else if (a === "CICLO TERMINADO EN MANUAL") {
          d.terminado_manual = true;
        }
      }
      d.tot_masa_cor = null; d.tot_masa_dosif = null;
      const alertas = [];
      let hdr = -1;
      for (let r = r0; r < r1; r++) if (str(cell(r, 0)) === "Codigo") { hdr = r; break; }
      if (hdr < 0) {
        alertas.push("no encontré la tabla de componentes");
      } else {
        const col = {};
        for (let c = 0; c < ncols; c++) { const t = str(cell(hdr, c)).trim(); if (t) col[norm(t)] = c; }
        const cNom = col["nombre componente"] ?? 1, cUm = col["um"] ?? 2;
        const cTeo = col["teór."] ?? col["teor."], cCor = col["cor."], cDos = col["dosif."];
        const cDif = col["dif."], cDev = col["% desv."];
        if (cCor === undefined || cDos === undefined) {
          alertas.push("la tabla no trae columnas 'Cor.' y/o 'Dosif.' (formato distinto al esperado)");
        } else {
          let vioMasa = false;
          for (let r = hdr + 1; r < r1; r++) {
            const nombre = str(cell(r, cNom));
            if (nombre.startsWith("Tot. masa")) {
              d.tot_masa_cor = num(cell(r, cCor))[0];
              d.tot_masa_dosif = num(cell(r, cDos))[0];
              vioMasa = true; break;
            }
            if (nombre.startsWith("Tot. áridos") || !str(cell(r, 0)).trim()) continue;
            const [dosif, manual] = num(cell(r, cDos));
            const [desv, , fuera] = cDev !== undefined ? num(cell(r, cDev)) : [null, false, false];
            const teo = cTeo !== undefined ? num(cell(r, cTeo))[0] : null;
            const cor = num(cell(r, cCor))[0];
            if (!(dosif || 0) && !(cor || 0) && !(teo || 0)) continue;
            consumos.push({
              archivo: nombreArchivo, informe: informeId, codigo: str(cell(r, 0)).trim(), insumo: nombre.trim(),
              um: str(cell(r, cUm)).trim(), teorico: teo, corregido: cor, dosificado: dosif,
              dif: cDif !== undefined ? num(cell(r, cDif))[0] : null, desv_pct: desv, fuera_tolerancia: fuera, dosif_manual: manual,
            });
          }
          if (!vioMasa) alertas.push("no encontré la fila 'Tot. masa' para cuadrar los totales");
        }
      }
      if (esRemito) {
        if (!d.m3) alertas.push("sin m³ (Subciclo/Totales)");
        if (!d.albaran) alertas.push("sin Albarán n.º");
        if (!d.cliente) alertas.push("sin cliente (¿cliente creado a último momento?)");
        if (!d.diseno) alertas.push("sin nombre de fórmula");
      }
      if (!d.fecha || !d.hora_inicio) alertas.push("sin fecha/hora de principio");
      d.alertas = alertas.join("; ");
      (esRemito ? remitos : ciclos).push(d);
    });
    return { remitos, consumos, ciclos };
  }

  // ---------- clasificar nombre de fórmula ----------
  function clasificarFormula(nombre) {
    const n = str(nombre);
    const m = n.match(/^\s*(F.?C|MR)\s*=\s*(\d+)/i);
    const toks = n.toUpperCase().split(/[-\s]+/);
    const cem = (toks.find((t) => CEMENTO_TXT[t]) || null);
    const srv = (toks.find((t) => SERVICIO_TXT[t]) || null);
    const sl = n.toUpperCase().match(/A(\d+)\s*["”″']/);
    const pref = m ? m[1].toUpperCase().replace(/´/g, "'").replace(/Ć/g, "C") : "";
    const tipo = pref.replace("FC", "F'C");
    const extra = n.match(/\(([^)]*)\)/);
    return {
      extra: extra ? extra[1].trim() : "", prefijo: tipo, fc: m ? m[2] : "",
      cemento: cem ? CEMENTO_TXT[cem] : "", servicio: srv ? SERVICIO_TXT[srv] : "", slump: sl ? sl[1] : "",
    };
  }
  function textoTipoConcreto(c) {
    const partes = [`${c.prefijo} ${c.fc}`.trim(), c.cemento, c.slump ? `SL ${c.slump}"` : ""];
    if (c.extra) partes.push(`(${c.extra})`);
    return partes.filter(Boolean).join(" - ");
  }

  // ---------- arma las filas de la tabla ----------
  function filasTabla(remitos, consumos, ciclos) {
    const avisos = [], avisosSrv = new Set();
    // pivote (archivo|informe) -> insumo -> {t,d}
    const piv = new Map();
    for (const c of consumos) {
      const ins = INSUMO_MAP[c.codigo];
      if (!ins) continue;
      const k = c.archivo + "|" + c.informe;
      if (!piv.has(k)) piv.set(k, {});
      const o = piv.get(k);
      if (!o[ins]) o[ins] = { t: null, d: null };
      if (c.corregido !== null && c.corregido !== undefined) o[ins].t = (o[ins].t || 0) + c.corregido;
      if (c.dosificado !== null && c.dosificado !== undefined) o[ins].d = (o[ins].d || 0) + c.dosificado;
    }
    const base = [...remitos.map((r) => ({ ...r, _manual: false })), ...ciclos.map((r) => ({ ...r, _manual: true }))]
      .map((r, i) => ({ r, i, o: parseInt(r.informe, 10) }))
      .sort((a, b) => (Number.isNaN(a.o) ? 1e18 : a.o) - (Number.isNaN(b.o) ? 1e18 : b.o) || a.i - b.i)
      .map((x) => x.r);

    const out = [], motivos = [];
    for (const r of base) {
      const mot = [];
      if (str(r.alertas).trim()) mot.push(r.alertas);
      const row = new Array(HEADERS.length).fill(null);
      const put = (L, v) => { row[colIdx(L) - 1] = v; };
      const inf = parseInt(r.informe, 10);
      if (Number.isNaN(inf)) mot.push("N.º de informe no numérico: " + r.informe); else put("A", inf);
      if (r._manual) { put("B", "MANUAL"); put("M", "MANUAL"); put("Q", "MANUAL"); }
      else {
        put("B", `${r.cliente_codigo} - ${r.cliente}`);
        put("M", `Codigo: ${r.diseno_codigo}`);
        put("Q", `Nombre fórmula: ${r.diseno}`);
        put("AD", r.m3 ?? null);
        const cl = clasificarFormula(r.diseno);
        put("O", textoTipoConcreto(cl) || null);
        put("P", cl.servicio || null);
        if (!cl.servicio) avisosSrv.add(r.diseno);
      }
      if (r.planta) put("S", r.planta);
      put("F", r.hora_inicio ? `Ora  ${r.hora_inicio}` : null);
      put("H", r.hora_fin ? `Ora  ${r.hora_fin}` : null);
      put("L", r.fecha ? { serial: fechaSerial(r.fecha) } : null);
      const key = r.archivo + "|" + r.informe;
      let sumT = 0, sumD = 0;
      const p = piv.get(key) || {};
      for (const [ins, [tc, dc]] of Object.entries(COLS)) {
        const o = p[ins];
        if (!o) continue;
        if (o.t !== null && o.t !== 0) { put(tc, Math.round(o.t * 100) / 100); sumT += o.t; }
        if (o.d !== null && o.d !== 0) { put(dc, Math.round(o.d * 100) / 100); sumD += o.d; }
      }
      const raros = new Set();
      for (const c of consumos) {
        if (c.archivo === r.archivo && c.informe === r.informe && !INSUMO_MAP[c.codigo] && ((c.dosificado || 0) !== 0 || (c.corregido || 0) !== 0)) raros.add(c.codigo);
      }
      if (raros.size) mot.push("insumo(s) sin columna en la tabla: " + [...raros].sort().join(", "));
      for (const [nom, suma, tot] of [["Cor.", sumT, r.tot_masa_cor], ["Dosif.", sumD, r.tot_masa_dosif]]) {
        if (tot !== null && tot !== undefined && Math.abs(suma - tot) > Math.max(1.0, 0.0005 * Math.abs(tot))) {
          mot.push(`no cuadra ${nom}: cargado ${fmt1(suma)} vs Tot. masa ${fmt1(tot)}`);
        }
      }
      out.push(row); motivos.push(mot.join("; "));
    }
    const nRev = motivos.filter(Boolean).length;
    if (nRev) avisos.push(`${nRev} fila(s) con problemas: NO se incluyen en la hoja para pegar; están en la hoja 'REVISAR (no pegar)'.`);
    if (avisosSrv.size) avisos.push("Fórmulas sin B/D en el nombre (TIPO DE DESPACHO queda vacío): " + [...avisosSrv].sort().join(", "));
    return { filas: out, motivos, avisos };
  }
  const fmt1 = (x) => x.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });

  function cobertura() {
    return HEADERS.map((h, i) => {
      const L = colLetter(i + 1);
      let est, det;
      if (FROM_REMITO[L]) { est = "✅ Lo da el remito"; det = FROM_REMITO[L]; }
      else if (CANDIDATE[L]) { est = "🟡 Posible"; det = CANDIDATE[L]; }
      else if (FORMULA_COLS.includes(L)) { est = "⚙️ Fórmula de la tabla"; det = "Se calcula sola al pegar las filas"; }
      else { est = "❌ No viene en el remito"; det = "Hay que llenarla a mano / de otra fuente"; }
      return [L, h, est, det];
    });
  }

  // ---------- Excel de salida ----------
  function construirXlsx(XLSX, filas, motivos) {
    const wb = XLSX.utils.book_new();
    const hojaFilas = (lista, conMotivo) => {
      const aoa = [HEADERS.concat(conMotivo ? ["MOTIVO (revisar)"] : [])];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      lista.forEach(({ row, mot }, i) => {
        row.forEach((v, j) => {
          if (v === null || v === undefined) return;
          const ref = XLSX.utils.encode_cell({ r: i + 1, c: j });
          if (typeof v === "object") ws[ref] = { t: "n", v: v.serial, z: "dd/mm/yyyy" };
          else if (typeof v === "number") ws[ref] = { t: "n", v };
          else ws[ref] = { t: "s", v: String(v) };
        });
        if (conMotivo) ws[XLSX.utils.encode_cell({ r: i + 1, c: HEADERS.length })] = { t: "s", v: mot };
      });
      ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(lista.length, 1), c: HEADERS.length - 1 + (conMotivo ? 1 : 0) } });
      ws["!cols"] = HEADERS.map((_, j) => ({ wch: j === 1 ? 38 : j === 12 ? 26 : j === 16 ? 40 : 14 })).concat(conMotivo ? [{ wch: 70 }] : []);
      return ws;
    };
    const ok = [], mal = [];
    filas.forEach((row, i) => (motivos[i] ? mal : ok).push({ row, mot: motivos[i] }));
    XLSX.utils.book_append_sheet(wb, hojaFilas(ok, false), "Pegar en TABLA GENERAL");
    if (mal.length) XLSX.utils.book_append_sheet(wb, hojaFilas(mal, true), "REVISAR (no pegar)");
    const wc = XLSX.utils.aoa_to_sheet([["Col.", "Columna de la tabla", "Estado", "Detalle"], ...cobertura()]);
    wc["!cols"] = [{ wch: 7 }, { wch: 42 }, { wch: 26 }, { wch: 62 }];
    XLSX.utils.book_append_sheet(wb, wc, "Cobertura");
    return XLSX.write(wb, { type: "array", bookType: "xlsx" });
  }

  // ---------- punto de entrada ----------
  // files: [{name, data: ArrayBuffer|Uint8Array}]
  function preparar(XLSX, files, planta) {
    planta = str(planta).trim().toUpperCase();
    let R = [], C = [], X = [];
    const errores = [];
    for (const f of files) {
      try {
        const wb = XLSX.read(f.data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true });
        const p = parseSheet(grid, f.name);
        p.remitos.forEach((r) => (r.planta = planta)); p.ciclos.forEach((r) => (r.planta = planta));
        R = R.concat(p.remitos); C = C.concat(p.consumos); X = X.concat(p.ciclos);
      } catch (e) {
        errores.push(`No pude leer "${f.name}" (${e.message}). Revisa que sea el .xls original de Betowin.`);
      }
    }
    if (!R.length && !X.length) return { vacio: true, avisos: errores.concat(["No encontré remitos en los archivos subidos."]) };

    const avisos = errores.slice();
    const n0 = R.length + X.length;
    const dedupe = (L) => { const seen = new Set(); return L.filter((r) => { const k = r.informe + "|" + fechaKey(r.fecha); if (seen.has(k)) return false; seen.add(k); return true; }); };
    R = dedupe(R); X = dedupe(X);
    const quitados = n0 - R.length - X.length;
    if (quitados) avisos.push(`Había ${quitados} remito(s)/ciclo(s) repetidos entre archivos; se dejó uno solo.`);

    const { filas, motivos, avisos: av } = filasTabla(R, C, X);
    avisos.push(...av);
    if (!planta) avisos.push("No escribiste la planta: la columna PLANTA (S) queda vacía.");

    const revisar = filas.map((row, i) => [row[0], motivos[i]]).filter((x) => x[1]);
    const fechas = R.map((r) => r.fecha).filter(Boolean).map(fechaSerial);
    const xlsx = construirXlsx(XLSX, filas, motivos);
    return {
      vacio: false, filas, motivos, revisar, avisos, xlsx,
      n_remitos: R.length, n_ciclos: X.length, m3: R.reduce((a, r) => a + (r.m3 || 0), 0),
      f_ini: fechas.length ? Math.min(...fechas) : null, f_fin: fechas.length ? Math.max(...fechas) : null,
    };
  }

  const API = { preparar, parseSheet, filasTabla, clasificarFormula, textoTipoConcreto, cobertura, colIdx, colLetter, fechaSerial };
  if (typeof module !== "undefined" && module.exports) module.exports = API; else root.Logica = API;
})(typeof window !== "undefined" ? window : globalThis);
