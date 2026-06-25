import XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";
import * as crypto from "crypto";
// ── Constantes ────────────────────────────────────────────────────────────────
const OUTLIER_THRESHOLD = 2.5; // z-score para marcar outlier
// Patrones para detectar columnas financieras/fecha automáticamente
const DATE_KEYWORDS = /fecha|date|mes|month|periodo|period|dia|day|año|year/i;
const AMOUNT_KEYWORDS = /monto|amount|total|precio|price|venta|sale|ingreso|income|gasto|expense|pago|payment|saldo|balance|valor|value|bs|usd/i;
const ID_KEYWORDS = /^id$|^cod|^codigo|^code|^nro|^num|^numero|^ref/i;
// ── Utilidades ────────────────────────────────────────────────────────────────
function hashRow(row) {
    const normalized = JSON.stringify(Object.fromEntries(Object.entries(row).map(([k, v]) => [
        k.trim().toLowerCase(),
        String(v ?? "").trim().toLowerCase(),
    ])));
    return crypto.createHash("md5").update(normalized).digest("hex");
}
function isNullish(val) {
    if (val === null || val === undefined)
        return true;
    if (typeof val === "string" && val.trim() === "")
        return true;
    return false;
}
function isAllNullish(row) {
    return Object.values(row).every(isNullish);
}
// ── Normalización de fechas ───────────────────────────────────────────────────
function normalizeDate(val) {
    if (isNullish(val))
        return null;
    const s = String(val).trim();
    // Excel serial date number
    const serial = Number(s);
    if (!isNaN(serial) && serial > 1 && serial < 100000) {
        try {
            const d = XLSX.SSF.parse_date_code(serial);
            return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
        }
        catch {
            // not a serial
        }
    }
    // Common formats: DD/MM/YYYY, DD-MM-YYYY, YYYY/MM/DD
    const patterns = [
        { re: /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/, fn: (m) => `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` },
        { re: /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/, fn: (m) => `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}` },
        { re: /^(\d{1,2})[\/\-](\d{4})$/, fn: (m) => `${m[2]}-${m[1].padStart(2, "0")}` },
    ];
    for (const { re, fn } of patterns) {
        const m = s.match(re);
        if (m)
            return fn(m);
    }
    // Try native Date parsing as last resort
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
        return d.toISOString().slice(0, 10);
    }
    return s; // devuelve original si no puede parsear
}
// ── Normalización de montos ───────────────────────────────────────────────────
function normalizeAmount(val) {
    if (isNullish(val))
        return null;
    if (typeof val === "number")
        return val;
    const s = String(val)
        .replace(/Bs\.?\s*/i, "")
        .replace(/USD?\s*/i, "")
        .replace(/\$\s*/, "")
        .replace(/,(\d{3})/g, "$1") // 1,234,567 → 1234567
        .replace(/\./g, "") // 1.234 (europeo) → 1234
        .replace(/,/, ".") // 1234,56 → 1234.56
        .trim();
    const n = Number(s);
    return isNaN(n) ? null : n;
}
// ── Detección de tipo de columna ──────────────────────────────────────────────
function detectColumnType(values) {
    const nonNull = values.filter((v) => !isNullish(v));
    if (nonNull.length === 0)
        return "text";
    let numericCount = 0;
    let dateCount = 0;
    let textCount = 0;
    for (const v of nonNull) {
        if (typeof v === "number") {
            numericCount++;
            continue;
        }
        const s = String(v).trim();
        if (!isNaN(Number(s.replace(/[,\s]/g, "")))) {
            numericCount++;
            continue;
        }
        if (/\d{1,4}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(s)) {
            dateCount++;
            continue;
        }
        textCount++;
    }
    const total = nonNull.length;
    if (numericCount / total > 0.8)
        return "numeric";
    if (dateCount / total > 0.8)
        return "date";
    if (textCount / total > 0.8)
        return "text";
    return "mixed";
}
// ── Z-score outlier detection ─────────────────────────────────────────────────
function detectOutliers(rows, column, values) {
    if (values.length < 4)
        return [];
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const stdDev = Math.sqrt(values.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / values.length);
    if (stdDev === 0)
        return [];
    const flags = [];
    values.forEach((v, i) => {
        const z = Math.abs((v - mean) / stdDev);
        if (z > OUTLIER_THRESHOLD) {
            flags.push({
                row: i + 2, // 1-indexed + header
                column,
                value: v,
                zScore: Math.round(z * 100) / 100,
                mean: Math.round(mean * 100) / 100,
                stdDev: Math.round(stdDev * 100) / 100,
            });
        }
    });
    return flags;
}
// ── Limpieza de una hoja ──────────────────────────────────────────────────────
function cleanSheet(rawRows, sheetName) {
    const originalRows = rawRows.length;
    // 1. Eliminar filas completamente vacías
    let rows = rawRows.filter((r) => !isAllNullish(r));
    // 2. Deduplicación por hash de contenido
    const seen = new Set();
    let duplicatesRemoved = 0;
    rows = rows.filter((row) => {
        const h = hashRow(row);
        if (seen.has(h)) {
            duplicatesRemoved++;
            return false;
        }
        seen.add(h);
        return true;
    });
    if (rows.length === 0) {
        return {
            name: sheetName, originalRows, cleanedRows: 0,
            duplicatesRemoved, outliersFound: 0, outliers: [],
            markdown: `## ${sheetName}\n\n_Hoja vacía o sin datos válidos._\n`,
            summary: { columns: [], totalRows: 0, numericTotals: {} },
        };
    }
    const columns = Object.keys(rows[0]);
    // 3. Detectar tipo de cada columna (por nombre + valores)
    const colTypes = new Map();
    for (const col of columns) {
        const vals = rows.map((r) => r[col]);
        const nameHint = DATE_KEYWORDS.test(col)
            ? "date"
            : AMOUNT_KEYWORDS.test(col)
                ? "numeric"
                : null;
        colTypes.set(col, nameHint ?? detectColumnType(vals));
    }
    // 4. Normalizar valores por tipo
    const normalizedRows = rows.map((row) => {
        const clean = {};
        for (const col of columns) {
            const type = colTypes.get(col);
            if (type === "date") {
                clean[col] = normalizeDate(row[col]) ?? row[col];
            }
            else if (type === "numeric") {
                const n = normalizeAmount(row[col]);
                clean[col] = n !== null ? n : row[col];
            }
            else {
                clean[col] = isNullish(row[col]) ? null : String(row[col]).trim();
            }
        }
        return clean;
    });
    // 5. Detectar outliers en columnas numéricas
    const allOutliers = [];
    for (const col of columns) {
        if (colTypes.get(col) !== "numeric")
            continue;
        const nums = normalizedRows
            .map((r) => r[col])
            .filter((v) => typeof v === "number" && !isNaN(v));
        allOutliers.push(...detectOutliers(normalizedRows, col, nums));
    }
    // 6. Calcular estadísticas por columna
    const columnStats = columns.map((col) => {
        const vals = normalizedRows.map((r) => r[col]);
        const nonNull = vals.filter((v) => !isNullish(v));
        const type = colTypes.get(col);
        if (type === "numeric") {
            const nums = nonNull.filter((v) => typeof v === "number");
            const sum = nums.reduce((a, b) => a + b, 0);
            return {
                name: col, type,
                nonNullCount: nonNull.length,
                min: Math.min(...nums),
                max: Math.max(...nums),
                sum: Math.round(sum * 100) / 100,
                avg: Math.round((sum / nums.length) * 100) / 100,
            };
        }
        const strVals = nonNull.map((v) => String(v));
        return {
            name: col, type,
            nonNullCount: nonNull.length,
            uniqueValues: new Set(strVals).size,
            sample: [...new Set(strVals)].slice(0, 5),
        };
    });
    // 7. Rango de fechas (si hay columna de fecha)
    let dateRange;
    for (const col of columns) {
        if (colTypes.get(col) === "date") {
            const dates = normalizedRows
                .map((r) => r[col])
                .filter((v) => typeof v === "string" && v.length >= 7)
                .sort();
            if (dates.length > 0) {
                dateRange = { from: dates[0], to: dates[dates.length - 1] };
            }
            break;
        }
    }
    // 8. Totales numéricos
    const numericTotals = {};
    for (const stat of columnStats) {
        if (stat.type === "numeric" && stat.sum !== undefined) {
            numericTotals[stat.name] = stat.sum;
        }
    }
    const summary = {
        columns: columnStats,
        totalRows: normalizedRows.length,
        dateRange,
        numericTotals,
    };
    // 9. Construir Markdown optimizado para Claude
    const markdown = buildSheetMarkdown(sheetName, normalizedRows, summary, allOutliers);
    return {
        name: sheetName,
        originalRows,
        cleanedRows: normalizedRows.length,
        duplicatesRemoved,
        outliersFound: allOutliers.length,
        outliers: allOutliers,
        markdown,
        summary,
    };
}
// ── Construcción del Markdown ─────────────────────────────────────────────────
function buildSheetMarkdown(sheetName, rows, summary, outliers) {
    const lines = [];
    lines.push(`## Hoja: ${sheetName}`);
    lines.push(`**Filas:** ${summary.totalRows}`);
    if (summary.dateRange) {
        lines.push(`**Período:** ${summary.dateRange.from} → ${summary.dateRange.to}`);
    }
    // Totales clave
    if (Object.keys(summary.numericTotals).length > 0) {
        lines.push("\n### Totales");
        for (const [col, total] of Object.entries(summary.numericTotals)) {
            lines.push(`- **${col}:** ${total.toLocaleString("es-BO")}`);
        }
    }
    // Estadísticas por columna numérica
    const numericCols = summary.columns.filter((c) => c.type === "numeric");
    if (numericCols.length > 0) {
        lines.push("\n### Estadísticas por columna");
        for (const col of numericCols) {
            lines.push(`- **${col.name}:** min=${col.min?.toLocaleString("es-BO")} / max=${col.max?.toLocaleString("es-BO")} / promedio=${col.avg?.toLocaleString("es-BO")} / total=${col.sum?.toLocaleString("es-BO")}`);
        }
    }
    // Alertas de outliers
    if (outliers.length > 0) {
        lines.push("\n### ⚠ Anomalías detectadas (sin IA)");
        for (const o of outliers) {
            lines.push(`- Fila ${o.row}, columna **${o.column}**: valor \`${o.value.toLocaleString("es-BO")}\` (z-score=${o.zScore}, promedio=${o.mean.toLocaleString("es-BO")})`);
        }
    }
    // Columnas de texto: valores únicos / muestra
    const textCols = summary.columns.filter((c) => c.type === "text" && (c.uniqueValues ?? 0) <= 20);
    if (textCols.length > 0) {
        lines.push("\n### Categorías clave");
        for (const col of textCols) {
            lines.push(`- **${col.name}:** ${col.sample?.join(", ")}`);
        }
    }
    // Tabla de datos (máximo 30 filas para no saturar tokens)
    const tableCols = summary.columns.map((c) => c.name);
    const tableRows = rows.slice(0, 30);
    lines.push("\n### Datos");
    lines.push("| " + tableCols.join(" | ") + " |");
    lines.push("| " + tableCols.map(() => "---").join(" | ") + " |");
    for (const row of tableRows) {
        const cells = tableCols.map((col) => {
            const v = row[col];
            if (isNullish(v))
                return "_null_";
            if (typeof v === "number")
                return v.toLocaleString("es-BO");
            return String(v).replace(/\|/g, "\\|");
        });
        lines.push("| " + cells.join(" | ") + " |");
    }
    if (rows.length > 30) {
        lines.push(`\n_... y ${rows.length - 30} filas más. Usa los totales y estadísticas de arriba para el análisis._`);
    }
    return lines.join("\n");
}
// ── Punto de entrada principal ────────────────────────────────────────────────
export async function runCleaningAgent(filePaths) {
    const sheets = [];
    let tokensBefore = 0;
    let tokensAfter = 0;
    for (const filePath of filePaths) {
        if (!fs.existsSync(filePath))
            continue;
        const ext = path.extname(filePath).toLowerCase();
        if (ext === ".csv") {
            // CSV → tratarlo como una sola hoja
            const content = fs.readFileSync(filePath, "utf-8");
            const wb = XLSX.read(content, { type: "string" });
            const sheetName = path.basename(filePath, ext);
            const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
            tokensBefore += estimateTokens(JSON.stringify(rawRows));
            const cleaned = cleanSheet(rawRows, sheetName);
            tokensAfter += estimateTokens(cleaned.markdown);
            sheets.push(cleaned);
        }
        else {
            // Excel → procesar cada hoja
            const wb = XLSX.readFile(filePath);
            for (const sheetName of wb.SheetNames) {
                const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });
                if (rawRows.length === 0)
                    continue;
                tokensBefore += estimateTokens(JSON.stringify(rawRows));
                const cleaned = cleanSheet(rawRows, sheetName);
                tokensAfter += estimateTokens(cleaned.markdown);
                sheets.push(cleaned);
            }
        }
    }
    // Markdown global que se pasa a Claude
    const globalMarkdown = buildGlobalMarkdown(sheets, filePaths);
    tokensAfter = estimateTokens(globalMarkdown); // usar el global final
    const reductionPercent = tokensBefore > 0
        ? Math.round(((tokensBefore - tokensAfter) / tokensBefore) * 100)
        : 0;
    return {
        processedAt: new Date().toISOString(),
        files: filePaths,
        sheets,
        globalMarkdown,
        tokenEstimateBefore: tokensBefore,
        tokenEstimateAfter: tokensAfter,
        reductionPercent,
    };
}
// ── Markdown global para Claude ───────────────────────────────────────────────
function buildGlobalMarkdown(sheets, files) {
    const lines = [];
    lines.push("# Reporte financiero — datos limpios");
    lines.push(`**Archivos procesados:** ${files.map((f) => path.basename(f)).join(", ")}`);
    lines.push(`**Hojas procesadas:** ${sheets.length}`);
    lines.push(`**Generado:** ${new Date().toLocaleString("es-BO")}`);
    // Resumen ejecutivo de duplicados y anomalías
    const totalDups = sheets.reduce((s, sh) => s + sh.duplicatesRemoved, 0);
    const totalOutliers = sheets.reduce((s, sh) => s + sh.outliersFound, 0);
    const totalOrig = sheets.reduce((s, sh) => s + sh.originalRows, 0);
    const totalClean = sheets.reduce((s, sh) => s + sh.cleanedRows, 0);
    lines.push("\n## Calidad de datos");
    lines.push(`- Filas originales: **${totalOrig}**`);
    lines.push(`- Filas después de limpieza: **${totalClean}**`);
    lines.push(`- Duplicados eliminados: **${totalDups}**`);
    lines.push(`- Anomalías detectadas: **${totalOutliers}**`);
    // Contenido de cada hoja
    for (const sheet of sheets) {
        lines.push("\n---");
        lines.push(sheet.markdown);
    }
    return lines.join("\n");
}
// ── Estimación de tokens ──────────────────────────────────────────────────────
function estimateTokens(text) {
    // Aproximación: 1 token ≈ 4 caracteres en español
    return Math.ceil(text.length / 4);
}
