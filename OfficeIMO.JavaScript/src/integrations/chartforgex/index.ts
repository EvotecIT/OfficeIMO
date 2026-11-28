import { BlobByteSink, checkAbort } from "../../core/index.js";
import type { OutputDestination, ExportResult } from "../../core/index.js";
import { captureFields } from "../../internal/capture.js";
import { captureCommonLimits, captureXlsxOptions, mergeAdapterLimits } from "../options.js";
import { writeXlsxTo } from "../../xlsx/index.js";
import { writePdfDocumentTo } from "../../pdf/index.js";
import type { PdfDocumentBlock, PdfDrawing } from "../../pdf/index.js";
import type { PdfDrawingFont } from "../../pdf/drawing-types.js";
import { capturePdfLimits } from "../../pdf/settings.js";
import { prepareChart } from "./capture.js";
import { svgToPdfDrawing } from "./svg-drawing.js";
import { chartTable } from "./workbook.js";
import { winAnsiText } from "./winansi.js";
import type { ChartForgeXCapture, ChartForgeXDiagnostic, ChartForgeXFormat, ChartForgeXSvgOptions, ChartForgeXWriteOptions } from "./types.js";
export { PdfFont } from "../../pdf/font.js";
export type { ChartForgeXCapture, ChartForgeXSeries, ChartForgeXPoint, ChartForgeXAxis, ChartForgeXFormat, ChartForgeXDiagnostic,
  ChartForgeXSvgOptions, ChartForgeXWriteOptions } from "./types.js";

const slots: readonly PdfDrawingFont[] = ["regular", "bold", "italic", "boldItalic"];
const headerFill = "E7EDF5", headerBorder = Object.freeze({ bottom: Object.freeze({ style: "thin" as const, color: "A6B4C8" }) });

/** Convert resolved chart SVG markup to the PDF writer's portable vector drawing, without a browser. */
export function chartForgeXDrawing(svg: string, options: ChartForgeXSvgOptions = {}): PdfDrawing {
  return svgToPdfDrawing(svg, captureFields(options, ["unsupported", "fonts", "winAnsi", "limits", "onDiagnostic"]));
}

/** Write a ChartForgeX capture as a typed XLSX sheet of its series, or a PDF with its title, description and vector SVG. */
export async function writeChartForgeXTo(capture: ChartForgeXCapture, format: ChartForgeXFormat, destination: OutputDestination,
  options: ChartForgeXWriteOptions = {}): Promise<ExportResult> {
  const selected = captureFields(options, ["layout", "signal", "limits", "onProgress", "onDiagnostic", "unsupported", "xlsx", "pdf"]);
  if (format !== "xlsx" && format !== "pdf") throw new TypeError("ChartForgeX exports are xlsx or pdf.");
  const { signal, onProgress, onDiagnostic, unsupported } = selected, limits = captureCommonLimits(selected.limits);
  if (unsupported !== undefined && unsupported !== "skip" && unsupported !== "reject") throw new TypeError("unsupported must be skip or reject.");
  checkAbort(signal);
  const chart = prepareChart(capture), stream = { ...(signal ? { signal } : {}), ...(onProgress ? { onProgress } : {}) };
  if (format === "xlsx") {
    const input = captureFields(selected.xlsx ?? {}, ["dateMode", "compression", "creator", "created", "modified", "sheetName"]);
    const { sheetName = "Chart data", ...workbook } = input;
    const xlsx = captureXlsxOptions(workbook), xlsxLimits = mergeAdapterLimits(xlsx.limits, limits);
    const table = chartTable(chart, selected.layout, xlsxLimits);
    return writeXlsxTo(table.rows, destination, { dateMode: "utc" as const, ...xlsx, ...stream, columns: table.columns,
      ...(chart.title === undefined ? {} : { title: chart.title }), ...(chart.description === undefined ? {} : { description: chart.description }),
      limits: xlsxLimits,
      sheet: { name: sheetName, headerFill, headerBorder, headerWrap: true, freezeHeader: true, autoFilter: true, boldHeader: true,
        autoSize: { minWidth: 8, maxWidth: 48 }, ...(chart.title ? { title: { text: chart.title } } : {}) } });
  }
  if (chart.svg === undefined) throw new TypeError("A ChartForgeX PDF needs the chart's SVG markup.");
  const pdf = captureFields(selected.pdf ?? {}, ["pageSize", "orientation", "margins", "fonts", "tagged", "compression", "pageNumbers", "pageHeader",
    "pageFooter", "fontSize", "limits"]);
  const pdfLimits = mergeAdapterLimits(pdf.limits === undefined ? undefined : capturePdfLimits(pdf.limits), limits);
  // Keep Unicode labels in the embedded regular face when an emphasis face is missing.
  const faces = pdf.fonts === undefined ? undefined : captureFields(pdf.fonts, slots);
  const available = faces ? slots.filter(slot => slot === "regular" || faces[slot] !== undefined) : slots;
  const reported = new Set<string>();
  const report = (diagnostic: ChartForgeXDiagnostic): void => {
    if (reported.has(diagnostic.code)) return;
    reported.add(diagnostic.code);
    onDiagnostic?.(diagnostic);
  };
  const drawing = svgToPdfDrawing(chart.svg, { fonts: available, winAnsi: !faces?.regular, limits: pdfLimits, ...(unsupported ? { unsupported } : {}),
    onDiagnostic: report }, signal);
  const standard = (text: string | undefined): string | undefined => text === undefined || faces?.regular ? text :
    winAnsiText(text, () => { if (unsupported === "reject") throw new TypeError("Chart text needs an embedded font: supply pdf.fonts.regular."); report({ code: "TEXT_SUBSTITUTED" }); });
  const title = standard(chart.title), description = standard(chart.description);
  // The figure alternative stays exact: it is Unicode metadata, not painted text.
  const blocks: PdfDocumentBlock[] = [...(description ? [{ kind: "paragraph" as const, text: description }] : []),
    { kind: "drawing", drawing, alignment: "center", alternativeText: chart.description || chart.title || "Chart" }];
  return writePdfDocumentTo(blocks, destination, { orientation: drawing.width > drawing.height ? "landscape" : "portrait", pageNumbers: false,
    ...pdf, ...stream, ...(faces ? { fonts: faces } : {}), ...(title ? { title } : {}),
    ...(chart.title ? { metadataTitle: chart.title } : {}), limits: pdfLimits });
}

/** Export a ChartForgeX capture as a Blob. No chart runtime, browser layout or server is involved. */
export async function exportChartForgeX(capture: ChartForgeXCapture, format: ChartForgeXFormat, options: ChartForgeXWriteOptions = {}): Promise<Blob> {
  const sink = new BlobByteSink();
  try {
    await writeChartForgeXTo(capture, format, sink, options);
    return sink.toBlob(format === "xlsx" ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/pdf");
  } catch (error) { sink.discard(); throw error; }
}
