import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { runCleaningAgent } from "./cleaning-agent.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema, } from "@modelcontextprotocol/sdk/types.js";
import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";
import XLSX from "xlsx";
import ExcelJS from "exceljs";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType, } from "docx";
// ── Directorios base ──────────────────────────────────────────────────────────
const DOCS_DIR = "C:\\Users\\ferna\\Desktop\\UCB\\sis2\\DocumentosEmpresa";
const REPORTS_DIR = "C:\\Users\\ferna\\Desktop\\UCB\\sis2\\DocumentosEmpresa\\Reportes";
function ensureReportsDir() {
    if (!fs.existsSync(REPORTS_DIR))
        fs.mkdirSync(REPORTS_DIR, { recursive: true });
}
// ── Helpers de lectura ────────────────────────────────────────────────────────
function findExcelFiles(dir, files = []) {
    if (!fs.existsSync(dir))
        return files;
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory())
            findExcelFiles(full, files);
        else if (/\.(xlsx|xls|xlsm|csv)$/i.test(entry.name))
            files.push(full);
    }
    return files;
}
function sheetToRows(sheet) {
    return XLSX.utils.sheet_to_json(sheet, { defval: null });
}
function resolveSafePath(filePath) {
    const resolved = path.resolve(filePath.startsWith("~")
        ? filePath.replace("~", os.homedir())
        : path.join(DOCS_DIR, filePath));
    const ALLOWED_ROOTS = [
        os.homedir(),
        "C:\\Users\\ferna\\Desktop\\UCB\\hackaton\\DocumentosEmpresa",
    ];
    //const ALLOWED_ROOTS = [
    //DOCS_DIR,
    //os.homedir(),
    //];
    if (!ALLOWED_ROOTS.some((root) => resolved.startsWith(root))) {
        throw new Error(`Acceso denegado: el archivo debe estar dentro de una carpeta permitida`);
    }
    return resolved;
}
function assertFileExists(filePath) {
    if (!fs.existsSync(filePath))
        throw new Error(`Archivo no encontrado: ${filePath}`);
}
function round(n, decimals = 2) {
    return Math.round(n * 10 ** decimals) / 10 ** decimals;
}
// ── Generador Excel (.xlsx) ───────────────────────────────────────────────────
async function buildExcelReport(params) {
    ensureReportsDir();
    const wb = new ExcelJS.Workbook();
    wb.creator = "Excel MCP Server";
    wb.created = new Date();
    const ws = wb.addWorksheet(params.sheet_name || "Reporte");
    // ── Título ──
    ws.mergeCells("A1", `${colLetter(params.columns.length)}1`);
    const titleCell = ws.getCell("A1");
    titleCell.value = params.title;
    titleCell.font = { name: "Calibri", size: 16, bold: true, color: { argb: "FFFFFFFF" } };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F3864" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(1).height = 36;
    // ── Fecha ──
    ws.mergeCells("A2", `${colLetter(params.columns.length)}2`);
    const dateCell = ws.getCell("A2");
    dateCell.value = `Generado: ${new Date().toLocaleString("es-BO")}`;
    dateCell.font = { name: "Calibri", size: 10, italic: true, color: { argb: "FF666666" } };
    dateCell.alignment = { horizontal: "right" };
    ws.getRow(2).height = 18;
    // ── Encabezados ──
    const headerRow = ws.getRow(4);
    params.columns.forEach((col, i) => {
        const cell = headerRow.getCell(i + 1);
        cell.value = col.label;
        cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E75B6" } };
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.border = {
            bottom: { style: "medium", color: { argb: "FF1F4E79" } },
            right: { style: "thin", color: { argb: "FF9DC3E6" } },
        };
    });
    headerRow.height = 28;
    // ── Datos ──
    params.rows.forEach((row, ri) => {
        const dataRow = ws.getRow(ri + 5);
        const isEven = ri % 2 === 0;
        params.columns.forEach((col, ci) => {
            const cell = dataRow.getCell(ci + 1);
            const val = row[col.key];
            cell.value = val;
            cell.font = { name: "Calibri", size: 10 };
            cell.fill = {
                type: "pattern", pattern: "solid",
                fgColor: { argb: isEven ? "FFD6E4F0" : "FFFFFFFF" },
            };
            cell.alignment = { vertical: "middle", horizontal: col.type === "number" || col.type === "currency" || col.type === "percent" ? "right" : "left" };
            cell.border = { bottom: { style: "hair", color: { argb: "FFBDD7EE" } } };
            if (col.type === "currency" && typeof val === "number") {
                cell.numFmt = '"Bs. "#,##0.00';
            }
            else if (col.type === "percent" && typeof val === "number") {
                cell.numFmt = '0.00"%"';
            }
            else if (col.type === "number" && typeof val === "number") {
                cell.numFmt = "#,##0.00";
            }
        });
        dataRow.height = 20;
    });
    // ── Autofit columnas ──
    params.columns.forEach((col, i) => {
        const colObj = ws.getColumn(i + 1);
        const maxLen = Math.max(col.label.length, ...params.rows.map((r) => String(r[col.key] ?? "").length));
        colObj.width = Math.min(Math.max(maxLen + 4, 12), 40);
    });
    // ── Resumen (si viene) ──
    if (params.summary && params.summary.length > 0) {
        const sumStartRow = params.rows.length + 6;
        ws.getRow(sumStartRow).getCell(1).value = "Resumen";
        ws.getRow(sumStartRow).getCell(1).font = { bold: true, size: 11, color: { argb: "FF1F3864" } };
        params.summary.forEach((s, i) => {
            const r = ws.getRow(sumStartRow + 1 + i);
            r.getCell(1).value = s.label;
            r.getCell(1).font = { name: "Calibri", size: 10, bold: true };
            r.getCell(2).value = s.value;
            r.getCell(2).font = { name: "Calibri", size: 10 };
        });
    }
    // ── Hoja de datos para grafica ──
    if (params.chart) {
        const chartWs = wb.addWorksheet("Datos_Grafica");
        const labelCol = params.columns.find((c) => c.key === params.chart.label_key);
        const chartHeaders = [
            labelCol?.label ?? params.chart.label_key,
            ...params.chart.series.map((s) => s.name),
        ];
        const hRow = chartWs.addRow(chartHeaders);
        hRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E75B6" } };
        });
        params.rows.forEach((row) => {
            chartWs.addRow([
                row[params.chart.label_key],
                ...params.chart.series.map((s) => row[s.data_key] ?? 0),
            ]);
        });
        chartWs.addRow([]);
        const instrRow = chartWs.addRow([
            `Selecciona A1:${colLetter(chartHeaders.length)}${params.rows.length + 1} -> Insertar -> Grafico ${params.chart.type}`,
        ]);
        instrRow.getCell(1).font = { italic: true, color: { argb: "FF888888" }, size: 10 };
        chartWs.getColumn(1).width = 70;
    }
    const filename = `${sanitize(params.title)}_${datestamp()}.xlsx`;
    const outPath = path.join(REPORTS_DIR, filename);
    await wb.xlsx.writeFile(outPath);
    return outPath;
}
// ── Generador Word (.docx) ────────────────────────────────────────────────────
async function buildWordReport(params) {
    ensureReportsDir();
    const children = [];
    // Portada
    children.push(new Paragraph({
        text: params.title,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
    }));
    if (params.subtitle) {
        children.push(new Paragraph({
            children: [new TextRun({ text: params.subtitle, italics: true, size: 24, color: "555555" })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
        }));
    }
    children.push(new Paragraph({
        children: [new TextRun({ text: `Generado: ${new Date().toLocaleString("es-BO")}`, size: 18, color: "888888" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
    }));
    // Resumen ejecutivo
    if (params.summary && params.summary.length > 0) {
        children.push(new Paragraph({ text: "Resumen ejecutivo", heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 120 } }));
        const summaryTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: params.summary.map((s) => new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: s.label, bold: true })] })],
                        shading: { type: ShadingType.SOLID, color: "D6E4F0" },
                        width: { size: 40, type: WidthType.PERCENTAGE },
                    }),
                    new TableCell({ children: [new Paragraph({ text: s.value })] }),
                ],
            })),
        });
        children.push(summaryTable);
        children.push(new Paragraph({ text: "", spacing: { after: 240 } }));
    }
    // Secciones
    for (const section of params.sections) {
        children.push(new Paragraph({ text: section.heading, heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 120 } }));
        children.push(new Paragraph({ children: [new TextRun({ text: section.content, size: 22 })], spacing: { after: 160 } }));
        if (section.table) {
            const tbl = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: section.table.headers.map((h) => new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF" })] })],
                            shading: { type: ShadingType.SOLID, color: "2E75B6" },
                            borders: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "1F4E79" } },
                        })),
                    }),
                    ...section.table.rows.map((row, ri) => new TableRow({
                        children: row.map((cell) => new TableCell({
                            children: [new Paragraph({ text: cell })],
                            shading: { type: ShadingType.SOLID, color: ri % 2 === 0 ? "EBF3FB" : "FFFFFF" },
                        })),
                    })),
                ],
            });
            children.push(tbl);
            children.push(new Paragraph({ text: "", spacing: { after: 200 } }));
        }
    }
    const doc = new Document({
        sections: [{ properties: {}, children }],
        styles: {
            default: {
                document: { run: { font: "Calibri", size: 22 } },
            },
        },
    });
    const filename = `${sanitize(params.title)}_${datestamp()}.docx`;
    const outPath = path.join(REPORTS_DIR, filename);
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buffer);
    return outPath;
}
// ── Generador HTML dashboard ──────────────────────────────────────────────────
function buildHtmlDashboard(params) {
    const COLORS = ["#2E75B6", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5", "#C55A11", "#264478", "#9DC3E6"];
    const kpiCards = params.kpis.map((k) => {
        const trendIcon = k.trend === "up" ? "▲" : k.trend === "down" ? "▼" : "–";
        const trendColor = k.trend === "up" ? "#70AD47" : k.trend === "down" ? "#C00000" : "#888";
        return `
    <div class="kpi-card">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.value}${k.unit ? `<span class="kpi-unit">${k.unit}</span>` : ""}</div>
      <div class="kpi-trend" style="color:${trendColor}">${trendIcon}</div>
    </div>`;
    }).join("");
    const chartScripts = params.charts.map((ch) => {
        const datasets = ch.datasets.map((ds, i) => `{
      label: ${JSON.stringify(ds.label)},
      data: ${JSON.stringify(ds.data)},
      backgroundColor: ${ch.type === "line" ? JSON.stringify((ds.color || COLORS[i % COLORS.length]) + "33") : JSON.stringify(ds.data.map((_, j) => ds.color || COLORS[j % COLORS.length]))},
      borderColor: ${JSON.stringify(ds.color || COLORS[i % COLORS.length])},
      borderWidth: 2,
      borderRadius: ${ch.type === "bar" ? 4 : 0},
      tension: 0.35,
      fill: ${ch.type === "line" ? "true" : "false"}
    }`).join(",");
        return `
    new Chart(document.getElementById(${JSON.stringify(ch.id)}), {
      type: ${JSON.stringify(ch.type)},
      data: {
        labels: ${JSON.stringify(ch.labels)},
        datasets: [${datasets}]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: ${ch.type === "pie" || ch.type === "doughnut" ? "'right'" : "'top'"}, labels: { font: { family: "Segoe UI", size: 12 }, color: "#333" } },
          title: { display: false }
        },
        scales: ${ch.type === "pie" || ch.type === "doughnut" ? "{}" : `{ x: { ticks: { color: "#555" }, grid: { color: "#eee" } }, y: { ticks: { color: "#555" }, grid: { color: "#eee" }, beginAtZero: true } }`}
      }
    });`;
    }).join("\n");
    const chartDivs = params.charts.map((ch) => `
    <div class="chart-card">
      <div class="chart-title">${ch.title}</div>
      <div class="chart-wrap"><canvas id="${ch.id}"></canvas></div>
    </div>`).join("");
    const tableHtml = params.table ? `
  <div class="section-title">${params.table.title}</div>
  <div class="table-wrap">
    <table>
      <thead><tr>${params.table.headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead>
      <tbody>
        ${params.table.rows.map((row, i) => `<tr class="${i % 2 === 0 ? "even" : ""}">${row.map((c) => `<td>${c}</td>`).join("")}</tr>`).join("")}
      </tbody>
    </table>
  </div>` : "";
    return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${params.title}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Segoe UI", sans-serif; background: #F0F4F8; color: #1a1a2e; }
  header { background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%); color: #fff; padding: 24px 40px; }
  header h1 { font-size: 24px; font-weight: 600; }
  header p  { font-size: 12px; opacity: .7; margin-top: 4px; }
  .container { max-width: 1200px; margin: 0 auto; padding: 28px 24px; }
  .kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 16px; margin-bottom: 28px; }
  .kpi-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,.06); border-left: 4px solid #2E75B6; }
  .kpi-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 8px; }
  .kpi-value { font-size: 26px; font-weight: 700; color: #1F3864; }
  .kpi-unit  { font-size: 13px; font-weight: 400; color: #666; margin-left: 4px; }
  .kpi-trend { font-size: 13px; margin-top: 4px; }
  .charts-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(440px, 1fr)); gap: 20px; margin-bottom: 28px; }
  .chart-card  { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
  .chart-title { font-size: 13px; font-weight: 600; color: #333; margin-bottom: 14px; }
  .chart-wrap  { height: 260px; position: relative; }
  .section-title { font-size: 15px; font-weight: 600; color: #1F3864; margin-bottom: 12px; }
  .table-wrap { background: #fff; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  thead tr { background: #2E75B6; color: #fff; }
  th, td { padding: 10px 14px; text-align: left; }
  tbody tr:hover { background: #EBF3FB; }
  tbody tr.even { background: #F5FAFF; }
  footer { text-align: center; font-size: 11px; color: #aaa; padding: 20px; }
</style>
</head>
<body>
<header>
  <h1>${params.title}</h1>
  <p>Reporte generado el ${new Date().toLocaleString("es-BO")}</p>
</header>
<div class="container">
  <div class="kpi-grid">${kpiCards}</div>
  <div class="charts-grid">${chartDivs}</div>
  ${tableHtml}
</div>
<footer>Generado por Excel MCP Server · Analista de Ventas</footer>
<script>${chartScripts}</script>
</body>
</html>`;
}
// ── Helpers internos ──────────────────────────────────────────────────────────
function colLetter(n) {
    let s = "";
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}
function datestamp() {
    return new Date().toISOString().slice(0, 10);
}
function sanitize(str) {
    return str.replace(/[^a-zA-Z0-9_\-\u00C0-\u024F]/g, "_").slice(0, 60);
}
// ── KPIs financieros para tienda de abarrotes ─────────────────────────────────
const CONFIG_PATH = path.join(DOCS_DIR, "kpi_config.json");
function loadKpiConfig() {
    if (fs.existsSync(CONFIG_PATH)) {
        return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8"));
    }
    return {
        margen_minimo: 15,
        stock_minimo_alerta: 20,
        dias_rotacion_max: 30,
        moneda: "BOB",
    };
}
function cleanRows(rows) {
    const seen = new Set();
    return rows
        .filter((row) => {
        const hash = JSON.stringify(row);
        if (seen.has(hash))
            return false;
        seen.add(hash);
        return true;
    })
        .map((row) => {
        const clean = {};
        for (const [k, v] of Object.entries(row)) {
            if (v === null || v === "" || v === undefined)
                continue;
            const str = String(v).trim();
            // Normalizar fechas
            if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(str)) {
                clean[k] = str;
            }
            // Normalizar moneda
            else if (/^[Bb][Ss]\.?\s*[\d,.]/.test(str)) {
                clean[k] = parseFloat(str.replace(/[^0-9.]/g, ""));
            }
            else {
                clean[k] = isNaN(Number(str)) ? str : Number(str);
            }
        }
        return clean;
    })
        .filter((row) => Object.keys(row).length > 0);
}
function detectOutliers(rows, key) {
    const vals = rows.map((r) => Number(r[key])).filter((v) => !isNaN(v));
    if (vals.length < 3)
        return [];
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    const std = Math.sqrt(vals.map((v) => (v - mean) ** 2).reduce((a, b) => a + b, 0) / vals.length);
    return vals.map((v, i) => (Math.abs(v - mean) > 2 * std ? i : -1)).filter((i) => i >= 0);
}
function readAllExcels() {
    const files = findExcelFiles(DOCS_DIR);
    const result = { ventas: [], inventario: [], personal: [], proveedores: [], suministro: [] };
    for (const f of files) {
        const name = path.basename(f).toLowerCase();
        const wb = XLSX.readFile(f);
        const rows = sheetToRows(wb.Sheets[wb.SheetNames[0]]);
        if (name.includes("venta"))
            result.ventas = cleanRows(rows);
        if (name.includes("inventario"))
            result.inventario = cleanRows(rows);
        if (name.includes("personal"))
            result.personal = cleanRows(rows);
        if (name.includes("proveedor"))
            result.proveedores = cleanRows(rows);
        if (name.includes("suministro"))
            result.suministro = cleanRows(rows);
    }
    return result;
}
async function buildFinancialReport(kpis) {
    ensureReportsDir();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("KPIs Financieros");
    ws.mergeCells("A1", "D1");
    const title = ws.getCell("A1");
    title.value = "Reporte Financiero — Tienda de Abarrotes";
    title.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F3864" } };
    title.alignment = { horizontal: "center" };
    ws.getRow(1).height = 36;
    ws.mergeCells("A2", "D2");
    ws.getCell("A2").value = `Generado: ${new Date().toLocaleString("es-BO")}`;
    ws.getCell("A2").alignment = { horizontal: "right" };
    ws.getRow(2).height = 18;
    const headers = ["KPI", "Valor", "Unidad", "Estado"];
    const hRow = ws.getRow(4);
    headers.forEach((h, i) => {
        const cell = hRow.getCell(i + 1);
        cell.value = h;
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E75B6" } };
        cell.alignment = { horizontal: "center" };
    });
    const rows = [
        ["Ventas Totales", kpis.ventas_totales, "BOB", kpis.ventas_totales ? "✅" : "—"],
        ["Margen Bruto", kpis.margen_bruto_pct, "%", Number(kpis.margen_bruto_pct) >= 15 ? "✅" : "⚠️"],
        ["Costo de Ventas", kpis.costo_ventas, "BOB", "—"],
        ["Unidades Vendidas", kpis.unidades_vendidas, "uds", "—"],
        ["Ticket Promedio", kpis.ticket_promedio, "BOB", "—"],
        ["Producto Top", kpis.producto_top, "", "⭐"],
        ["Stock Total", kpis.stock_total, "uds", "—"],
        ["Valor Inventario", kpis.valor_inventario, "BOB", "—"],
        ["Productos Bajo Mínimo", kpis.productos_bajo_minimo, "uds", Number(kpis.productos_bajo_minimo) > 0 ? "⚠️" : "✅"],
        ["Rotación Inventario", kpis.rotacion_inventario, "días", Number(kpis.rotacion_inventario) <= 30 ? "✅" : "⚠️"],
        ["Margen por Producto", kpis.margen_por_producto, "%", "—"],
        ["Total Planilla", kpis.total_planilla, "BOB", "—"],
        ["Nro Empleados", kpis.nro_empleados, "pers", "—"],
        ["Costo Laboral s/Venta", kpis.costo_laboral_pct, "%", Number(kpis.costo_laboral_pct) <= 30 ? "✅" : "⚠️"],
        ["Nro Proveedores", kpis.nro_proveedores, "", "—"],
    ];
    rows.forEach((r, i) => {
        const dataRow = ws.getRow(i + 5);
        r.forEach((val, ci) => {
            const cell = dataRow.getCell(ci + 1);
            cell.value = val;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: i % 2 === 0 ? "FFD6E4F0" : "FFFFFFFF" } };
            cell.font = { name: "Calibri", size: 10 };
        });
        dataRow.height = 20;
    });
    [18, 14, 10, 10].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    const filename = `Reporte_Financiero_${datestamp()}.xlsx`;
    const outPath = path.join(REPORTS_DIR, filename);
    await wb.xlsx.writeFile(outPath);
    return outPath;
}
// ── Servidor MCP ──────────────────────────────────────────────────────────────
const server = new Server({ name: "excel-mcp-server", version: "2.0.0" }, { capabilities: { tools: {} } });
// ── Definición de herramientas ────────────────────────────────────────────────
server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
        // ── LECTURA ──────────────────────────────────────────────────────────────
        {
            name: "list_excel_files",
            description: "Lista todos los archivos Excel en la carpeta Documentos del usuario.",
            inputSchema: {
                type: "object",
                properties: {
                    subdirectory: { type: "string", description: "Subcarpeta opcional dentro de Documentos." },
                },
            },
        },
        {
            name: "read_excel_sheets",
            description: "Lee los nombres de todas las hojas de un archivo Excel.",
            inputSchema: {
                type: "object",
                required: ["file_path"],
                properties: {
                    file_path: { type: "string", description: "Ruta del archivo (relativa a Documentos o con ~)." },
                },
            },
        },
        {
            name: "read_excel_data",
            description: "Lee los datos de una hoja específica de un archivo Excel y los devuelve como JSON.",
            inputSchema: {
                type: "object",
                required: ["file_path"],
                properties: {
                    file_path: { type: "string" },
                    sheet_name: { type: "string", description: "Nombre de la hoja. Si no se indica, se lee la primera." },
                    max_rows: { type: "number", description: "Máximo de filas a devolver (por defecto 200)." },
                },
            },
        },
        {
            name: "search_in_excel",
            description: "Busca un texto en todas las celdas de una hoja Excel.",
            inputSchema: {
                type: "object",
                required: ["file_path", "search_text"],
                properties: {
                    file_path: { type: "string" },
                    search_text: { type: "string" },
                    sheet_name: { type: "string", description: "Hoja donde buscar. Si no se indica, busca en todas." },
                },
            },
        },
        {
            name: "get_excel_summary",
            description: "Devuelve un resumen estadístico de las columnas de una hoja Excel.",
            inputSchema: {
                type: "object",
                required: ["file_path"],
                properties: {
                    file_path: { type: "string" },
                    sheet_name: { type: "string" },
                },
            },
        },
        // ── GENERACIÓN DE REPORTES ────────────────────────────────────────────────
        {
            name: "generate_excel_report",
            description: `Genera un archivo .xlsx formateado con colores corporativos, encabezados, datos con filas alternas y opcionalmente una gráfica (bar, line o pie). 
El archivo se guarda en ~/Documents/Reportes/ y devuelve la ruta. 
Úsalo cuando el usuario pida un reporte, tabla o resumen en Excel.`,
            inputSchema: {
                type: "object",
                required: ["title", "sheet_name", "columns", "rows"],
                properties: {
                    title: { type: "string", description: "Título del reporte." },
                    sheet_name: { type: "string", description: "Nombre de la hoja principal." },
                    columns: {
                        type: "array",
                        description: "Definición de columnas.",
                        items: {
                            type: "object",
                            required: ["key", "label"],
                            properties: {
                                key: { type: "string", description: "Clave del campo en los datos." },
                                label: { type: "string", description: "Encabezado visible." },
                                type: { type: "string", enum: ["text", "number", "currency", "percent"], description: "Formato de celda." },
                            },
                        },
                    },
                    rows: {
                        type: "array",
                        description: "Array de objetos con los datos. Las claves deben coincidir con 'key' de columns.",
                        items: { type: "object" },
                    },
                    chart: {
                        type: "object",
                        description: "Gráfica opcional.",
                        properties: {
                            type: { type: "string", enum: ["bar", "line", "pie"] },
                            label_key: { type: "string", description: "Clave de la columna que sirve de etiqueta (eje X o labels)." },
                            series: {
                                type: "array",
                                items: {
                                    type: "object",
                                    required: ["name", "data_key"],
                                    properties: {
                                        name: { type: "string" },
                                        data_key: { type: "string" },
                                        color: { type: "string", description: "Hex color, ej: #2E75B6" },
                                    },
                                },
                            },
                        },
                    },
                    summary: {
                        type: "array",
                        description: "Filas de resumen opcionales al final de la hoja.",
                        items: {
                            type: "object",
                            properties: {
                                label: { type: "string" },
                                value: { type: "string" },
                            },
                        },
                    },
                },
            },
        },
        {
            name: "generate_word_report",
            description: `Genera un archivo .docx con portada, resumen ejecutivo, secciones de texto y tablas formateadas.
El archivo se guarda en ~/Documents/Reportes/ y devuelve la ruta.
Úsalo cuando el usuario pida un reporte, informe o documento en Word.`,
            inputSchema: {
                type: "object",
                required: ["title", "sections"],
                properties: {
                    title: { type: "string" },
                    subtitle: { type: "string" },
                    sections: {
                        type: "array",
                        items: {
                            type: "object",
                            required: ["heading", "content"],
                            properties: {
                                heading: { type: "string" },
                                content: { type: "string" },
                                table: {
                                    type: "object",
                                    properties: {
                                        headers: { type: "array", items: { type: "string" } },
                                        rows: { type: "array", items: { type: "array", items: { type: "string" } } },
                                    },
                                },
                            },
                        },
                    },
                    summary: {
                        type: "array",
                        items: {
                            type: "object",
                            properties: {
                                label: { type: "string" },
                                value: { type: "string" },
                            },
                        },
                    },
                },
            },
        },
        {
            name: "generate_html_dashboard",
            description: `Genera un dashboard HTML interactivo con KPIs, gráficas (Chart.js) y tabla de datos.
El archivo se guarda en ~/Documents/Reportes/ y devuelve la ruta. Se puede abrir directo en el navegador.
Úsalo cuando el usuario pida un dashboard, visualización o reporte visual.`,
            inputSchema: {
                type: "object",
                required: ["title", "kpis", "charts"],
                properties: {
                    title: { type: "string" },
                    kpis: {
                        type: "array",
                        items: {
                            type: "object",
                            required: ["label", "value"],
                            properties: {
                                label: { type: "string" },
                                value: {},
                                unit: { type: "string" },
                                trend: { type: "string", enum: ["up", "down", "neutral"] },
                            },
                        },
                    },
                    charts: {
                        type: "array",
                        items: {
                            type: "object",
                            required: ["id", "title", "type", "labels", "datasets"],
                            properties: {
                                id: { type: "string" },
                                title: { type: "string" },
                                type: { type: "string", enum: ["bar", "line", "pie", "doughnut"] },
                                labels: { type: "array", items: { type: "string" } },
                                datasets: {
                                    type: "array",
                                    items: {
                                        type: "object",
                                        required: ["label", "data"],
                                        properties: {
                                            label: { type: "string" },
                                            data: { type: "array", items: { type: "number" } },
                                            color: { type: "string" },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    table: {
                        type: "object",
                        properties: {
                            title: { type: "string" },
                            headers: { type: "array", items: { type: "string" } },
                            rows: { type: "array", items: { type: "array" } },
                        },
                    },
                },
            },
        },
        // leer y limpiar datos financieros
        {
            name: "read_financials",
            description: "Lee y normaliza todos los archivos Excel del directorio. Limpia duplicados, normaliza fechas y monedas, detecta outliers. Devuelve los datos limpios listos para análisis.",
            inputSchema: { type: "object", properties: {} },
        },
        {
            name: "analyze_kpis",
            description: "Calcula los KPIs financieros clave para la tienda de abarrotes: ventas totales, margen bruto, rotación de inventario, costo laboral, ticket promedio, productos bajo mínimo de stock y más. Genera reporte Excel automáticamente.",
            inputSchema: { type: "object", properties: {} },
        },
        {
            name: "config_kpis",
            description: "Guarda o lee la configuración de KPIs de la empresa (margen mínimo aceptable, días de rotación máximos, stock mínimo de alerta). Se persiste en un JSON local.",
            inputSchema: {
                type: "object",
                properties: {
                    margen_minimo: { type: "number", description: "Margen bruto mínimo aceptable en %" },
                    stock_minimo_alerta: { type: "number", description: "Unidades mínimas antes de alertar" },
                    dias_rotacion_max: { type: "number", description: "Días máximos de rotación de inventario" },
                    moneda: { type: "string", description: "Moneda (ej: BOB)" },
                },
            },
        },
        {
            name: "clean_data",
            description: "Limpia los datos crudos de un archivo Excel: elimina duplicados por hash, normaliza fechas y monedas, detecta outliers por z-score. No consume tokens de Claude.",
            inputSchema: {
                type: "object",
                required: ["file_path"],
                properties: {
                    file_path: { type: "string", description: "Ruta del archivo a limpiar" },
                    numeric_key: { type: "string", description: "Columna numérica para detectar outliers (opcional)" },
                },
            },
        },
        {
            name: "clean_and_prepare",
            description: `USAR SIEMPRE antes de analyze_kpis o cualquier análisis.
Limpia todos los excels del directorio y los convierte en Markdown optimizado para Claude.
Elimina duplicados por hash MD5, normaliza fechas a YYYY-MM-DD, convierte montos Bs. a float,
detecta outliers con z-score > 2.5, y reduce tokens un ~85%.
Devuelve: markdown_para_claude listo para inyectar en el prompt.`,
            inputSchema: {
                type: "object",
                required: ["file_paths"],
                properties: {
                    file_paths: {
                        type: "array",
                        items: { type: "string" },
                        description: "Rutas de los excels a procesar. Usa list_excel_files primero para obtenerlas.",
                    },
                },
            },
        },
    ],
}));
// ── Implementación ────────────────────────────────────────────────────────────
server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;
    try {
        // ── list_excel_files ──────────────────────────────────────────────────
        if (name === "list_excel_files") {
            const subdir = args?.subdirectory ?? "";
            const searchDir = subdir ? path.join(DOCS_DIR, subdir) : DOCS_DIR;
            if (!fs.existsSync(searchDir))
                return text(`La carpeta no existe: ${searchDir}`);
            const files = findExcelFiles(searchDir);
            if (files.length === 0)
                return text(`No se encontraron archivos Excel en: ${searchDir}`);
            const result = files.map((f) => ({
                ruta_relativa: path.relative(DOCS_DIR, f),
                ruta_absoluta: f,
                tamaño_kb: Math.round(fs.statSync(f).size / 1024),
                modificado: fs.statSync(f).mtime.toISOString().split("T")[0],
            }));
            return json({ total: files.length, archivos: result });
        }
        // ── read_excel_sheets ─────────────────────────────────────────────────
        if (name === "read_excel_sheets") {
            const fp = resolveSafePath(args.file_path);
            assertFileExists(fp);
            const wb = XLSX.readFile(fp);
            return json({ archivo: path.basename(fp), hojas: wb.SheetNames, total_hojas: wb.SheetNames.length });
        }
        // ── read_excel_data ───────────────────────────────────────────────────
        if (name === "read_excel_data") {
            const fp = resolveSafePath(args.file_path);
            assertFileExists(fp);
            const wb = XLSX.readFile(fp);
            const sheetName = args?.sheet_name ?? wb.SheetNames[0];
            if (!wb.SheetNames.includes(sheetName))
                return text(`Hoja "${sheetName}" no encontrada. Disponibles: ${wb.SheetNames.join(", ")}`);
            const maxRows = args?.max_rows ?? 200;
            const rows = sheetToRows(wb.Sheets[sheetName]).slice(0, maxRows);
            return json({
                archivo: path.basename(fp), hoja: sheetName,
                total_filas_leídas: rows.length,
                columnas: rows.length > 0 ? Object.keys(rows[0]) : [],
                datos: rows,
            });
        }
        // ── search_in_excel ───────────────────────────────────────────────────
        if (name === "search_in_excel") {
            const fp = resolveSafePath(args.file_path);
            assertFileExists(fp);
            const searchText = args.search_text.toLowerCase();
            const wb = XLSX.readFile(fp);
            const sheets = args?.sheet_name ? [args.sheet_name] : wb.SheetNames;
            const results = [];
            for (const sn of sheets) {
                if (!wb.SheetNames.includes(sn))
                    continue;
                for (const row of sheetToRows(wb.Sheets[sn])) {
                    if (Object.values(row).some((v) => String(v ?? "").toLowerCase().includes(searchText)))
                        results.push({ hoja: sn, ...row });
                }
            }
            return json({ texto_buscado: args.search_text, total_coincidencias: results.length, resultados: results });
        }
        // ── get_excel_summary ─────────────────────────────────────────────────
        if (name === "get_excel_summary") {
            const fp = resolveSafePath(args.file_path);
            assertFileExists(fp);
            const wb = XLSX.readFile(fp);
            const sheetName = args?.sheet_name ?? wb.SheetNames[0];
            const rows = sheetToRows(wb.Sheets[sheetName]);
            if (rows.length === 0)
                return text("La hoja está vacía.");
            const summary = {};
            for (const col of Object.keys(rows[0])) {
                const nums = rows.map((r) => r[col]).filter((v) => v !== null && v !== "" && !isNaN(Number(v))).map(Number);
                if (nums.length > 0) {
                    summary[col] = { tipo: "numérico", conteo: nums.length, suma: round(nums.reduce((a, b) => a + b, 0)), promedio: round(nums.reduce((a, b) => a + b, 0) / nums.length), mínimo: round(Math.min(...nums)), máximo: round(Math.max(...nums)) };
                }
                else {
                    const uniq = [...new Set(rows.map((r) => String(r[col] ?? "")))];
                    summary[col] = { tipo: "texto", conteo: rows.filter((r) => r[col] !== null && r[col] !== "").length, valores_únicos: uniq.length, muestra: uniq.slice(0, 5) };
                }
            }
            return json({ archivo: path.basename(fp), hoja: sheetName, total_filas: rows.length, resumen_columnas: summary });
        }
        // ── generate_excel_report ─────────────────────────────────────────────
        if (name === "generate_excel_report") {
            const outPath = await buildExcelReport({
                title: args.title,
                sheet_name: args.sheet_name,
                columns: args.columns,
                rows: args.rows,
                chart: args?.chart,
                summary: args?.summary,
            });
            return json({ ok: true, archivo_generado: outPath, mensaje: `Reporte guardado en ${outPath}` });
        }
        // ── generate_word_report ──────────────────────────────────────────────
        if (name === "generate_word_report") {
            const outPath = await buildWordReport({
                title: args.title,
                subtitle: args?.subtitle,
                sections: args.sections,
                summary: args?.summary,
            });
            return json({ ok: true, archivo_generado: outPath, mensaje: `Reporte guardado en ${outPath}` });
        }
        // ── generate_html_dashboard ───────────────────────────────────────────
        if (name === "generate_html_dashboard") {
            ensureReportsDir();
            const html = buildHtmlDashboard({
                title: args.title,
                kpis: args.kpis,
                charts: args.charts,
                table: args?.table,
            });
            const filename = `${sanitize(args.title)}_${datestamp()}.html`;
            const outPath = path.join(REPORTS_DIR, filename);
            fs.writeFileSync(outPath, html, "utf-8");
            return json({ ok: true, archivo_generado: outPath, mensaje: `Dashboard guardado en ${outPath}` });
        }
        // ── read_financials ───────────────────────────────────────────────────
        if (name === "read_financials") {
            const data = readAllExcels();
            const summary = {
                ventas: { filas: data.ventas.length, columnas: data.ventas.length > 0 ? Object.keys(data.ventas[0]).length : 0 },
                inventario: { filas: data.inventario.length, columnas: data.inventario.length > 0 ? Object.keys(data.inventario[0]).length : 0 },
                personal: { filas: data.personal.length, columnas: data.personal.length > 0 ? Object.keys(data.personal[0]).length : 0 },
                proveedores: { filas: data.proveedores.length, columnas: data.proveedores.length > 0 ? Object.keys(data.proveedores[0]).length : 0 },
                suministro: { filas: data.suministro.length, columnas: data.suministro.length > 0 ? Object.keys(data.suministro[0]).length : 0 },
            };
            return json({ ok: true, resumen: summary, muestra_ventas: data.ventas.slice(0, 3), muestra_inventario: data.inventario.slice(0, 3) });
        }
        // ── analyze_kpis ──────────────────────────────────────────────────────
        if (name === "analyze_kpis") {
            const data = readAllExcels();
            const cfg = loadKpiConfig();
            // Ventas
            const ventasTotales = data.ventas.reduce((s, r) => s + (Number(r["Total (BOB)"] ?? r["total"] ?? r["monto"] ?? r["Monto"] ?? 0)), 0);
            const unidadesVendidas = data.ventas.reduce((s, r) => s + (Number(r["Cantidad"] ?? r["cantidad"] ?? r["unidades"] ?? 0)), 0);
            const ticketPromedio = data.ventas.length > 0 ? round(ventasTotales / data.ventas.length) : 0;
            const productoTop = (() => {
                const cnt = {};
                data.ventas.forEach((r) => { const p = String(r["Producto"] ?? r["producto"] ?? ""); cnt[p] = (cnt[p] || 0) + Number(r["Cantidad"] ?? 1); });
                return Object.entries(cnt).sort((a, b) => b[1] - a[1])[0]?.[0] ?? "—";
            })();
            // Inventario
            const stockTotal = data.inventario.reduce((s, r) => s + (Number(r["Stock Actual"] ?? r["stock"] ?? 0)), 0);
            const valorInventario = data.inventario.reduce((s, r) => s + (Number(r["Precio de Compra (BOB)"] ?? r["precio_compra"] ?? 0) * Number(r["Stock Actual"] ?? r["stock"] ?? 0)), 0);
            const bajosMinimo = data.inventario.filter((r) => Number(r["Stock Actual"] ?? 0) < Number(r["Stock Mínimo (Alerta)"] ?? r["stock_minimo"] ?? cfg.stock_minimo_alerta ?? 20)).length;
            const costoVentas = data.inventario.reduce((s, r) => s + (Number(r["Precio de Compra (BOB)"] ?? 0) * Number(r["Stock Actual"] ?? 0)), 0);
            const margenBruto = ventasTotales > 0 ? round(((ventasTotales - costoVentas) / ventasTotales) * 100) : 0;
            const margenPorProducto = data.inventario.length > 0
                ? round(data.inventario.reduce((s, r) => {
                    const pc = Number(r["Precio de Compra (BOB)"] ?? 0);
                    const pv = Number(r["Precio de Venta (BOB)"] ?? 0);
                    return s + (pv > 0 ? ((pv - pc) / pv) * 100 : 0);
                }, 0) / data.inventario.length)
                : 0;
            const rotacionInventario = ventasTotales > 0 && valorInventario > 0
                ? round((valorInventario / (ventasTotales / 30)))
                : 0;
            // Personal
            const totalPlanilla = data.personal.reduce((s, r) => s + (Number(r["Salario (BOB)"] ?? r["salario"] ?? r["sueldo"] ?? 0)), 0);
            const nroEmpleados = data.personal.length;
            const costoLaboralPct = ventasTotales > 0 ? round((totalPlanilla / ventasTotales) * 100) : 0;
            // Proveedores
            const nroProveedores = data.proveedores.length;
            const kpis = {
                ventas_totales: round(ventasTotales),
                unidades_vendidas: unidadesVendidas,
                ticket_promedio: ticketPromedio,
                producto_top: productoTop,
                costo_ventas: round(costoVentas),
                margen_bruto_pct: margenBruto,
                margen_por_producto: margenPorProducto,
                stock_total: stockTotal,
                valor_inventario: round(valorInventario),
                productos_bajo_minimo: bajosMinimo,
                rotacion_inventario: rotacionInventario,
                total_planilla: round(totalPlanilla),
                nro_empleados: nroEmpleados,
                costo_laboral_pct: costoLaboralPct,
                nro_proveedores: nroProveedores,
            };
            const outPath = await buildFinancialReport(kpis);
            return json({ ok: true, kpis, reporte_generado: outPath, alertas: {
                    margen_bajo: margenBruto < Number(cfg.margen_minimo ?? 15),
                    stock_critico: bajosMinimo > 0,
                    costo_laboral_alto: costoLaboralPct > 30,
                    rotacion_lenta: rotacionInventario > Number(cfg.dias_rotacion_max ?? 30),
                } });
        }
        // ── config_kpis ───────────────────────────────────────────────────────
        if (name === "config_kpis") {
            const current = loadKpiConfig();
            const updated = { ...current, ...args };
            fs.writeFileSync(CONFIG_PATH, JSON.stringify(updated, null, 2), "utf-8");
            return json({ ok: true, config_guardada: updated, ruta: CONFIG_PATH });
        }
        // ── clean_data ────────────────────────────────────────────────────────
        if (name === "clean_data") {
            const fp = resolveSafePath(args.file_path);
            assertFileExists(fp);
            const wb = XLSX.readFile(fp);
            const raw = sheetToRows(wb.Sheets[wb.SheetNames[0]]);
            const clean = cleanRows(raw);
            const outliers = args?.numeric_key
                ? detectOutliers(clean, args.numeric_key)
                : [];
            return json({
                ok: true,
                archivo: path.basename(fp),
                filas_originales: raw.length,
                filas_limpias: clean.length,
                duplicados_removidos: raw.length - clean.length,
                outliers_detectados: outliers.length,
                indices_outliers: outliers,
                muestra: clean.slice(0, 5),
            });
        }
        if (name === "clean_and_prepare") {
            const rawPaths = args.file_paths;
            const resolvedPaths = rawPaths.map((p) => resolveSafePath(p));
            const missing = resolvedPaths.filter((p) => !fs.existsSync(p));
            if (missing.length > 0)
                return text(`Archivos no encontrados: ${missing.join(", ")}`);
            const report = await runCleaningAgent(resolvedPaths);
            return json({
                archivos_procesados: report.files.map((f) => path.basename(f)),
                hojas: report.sheets.map((s) => ({
                    nombre: s.name,
                    filas_originales: s.originalRows,
                    filas_limpias: s.cleanedRows,
                    duplicados_eliminados: s.duplicatesRemoved,
                    anomalias_detectadas: s.outliersFound,
                    anomalias: s.outliers,
                })),
                tokens_antes: report.tokenEstimateBefore,
                tokens_despues: report.tokenEstimateAfter,
                reduccion_porcentaje: report.reductionPercent,
                markdown_para_claude: report.globalMarkdown,
            });
        }
        return text(`Herramienta desconocida: ${name}`);
    }
    catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        return text(`Error: ${msg}`);
    }
});
// ── Respuesta helpers ─────────────────────────────────────────────────────────
function text(content) {
    return { content: [{ type: "text", text: content }] };
}
function json(data) {
    return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
}
// ── Arranque ──────────────────────────────────────────────────────────────────
async function main() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("✅ Excel MCP Server v2.0 — listo para Claude");
}
main().catch((err) => {
    console.error("Error fatal:", err);
    process.exit(1);
});
