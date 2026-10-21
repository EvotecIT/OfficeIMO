import { ChunkedTextSink } from "../core/sinks.js";
import type { ByteSink } from "../core/sinks.js";
import { escapeXml, xmlDeclaration } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { officeRelationshipsNamespace } from "../opc/index.js";
import { columnName } from "./values.js";
import type { RegisteredChart } from "./charts.js";

export const chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const chartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

/** @internal Graphic frames share a drawing with worksheet images; their nonvisual IDs must be unique. */
export function chartAnchorXml(chart: RegisteredChart, objectId: number, policy: InvalidCharacterPolicy): string {
  const size = ' cx="' + Math.round(chart.width * 9525) + '" cy="' + Math.round(chart.height * 9525) + '"';
  return '<xdr:oneCellAnchor><xdr:from><xdr:col>' + (chart.column - 1) + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + (chart.row - 1) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:ext' + size + '/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + objectId + '" name="Chart ' + chart.id + '" descr="' + escapeXml(chart.description ?? chart.title ?? "", policy) + '"/>' +
    '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="' + chartNamespace + '">' +
    '<c:chart xmlns:c="' + chartNamespace + '" r:id="chart' + chart.id + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>';
}

function titleXml(value: string, policy: InvalidCharacterPolicy): string {
  return '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t xml:space="preserve">' +
    escapeXml(value, policy) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>';
}

function* reference(kind: "str" | "num", formula: string, values: readonly (string | number)[], policy: InvalidCharacterPolicy): Generator<string> {
  yield '<c:' + kind + 'Ref><c:f>' + escapeXml(formula, policy) + '</c:f><c:' + kind + 'Cache>';
  if (kind === "num") yield '<c:formatCode>General</c:formatCode>';
  yield '<c:ptCount val="' + values.length + '"/>';
  for (let i = 0; i < values.length; i++) yield '<c:pt idx="' + i + '"><c:v xml:space="preserve">' + escapeXml(String(values[i]), policy) + '</c:v></c:pt>';
  yield '</c:' + kind + 'Cache></c:' + kind + 'Ref>';
}

function labelsXml(chart: RegisteredChart, policy: InvalidCharacterPolicy): string {
  const labels = chart.dataLabels;
  if (!labels) return "";
  const positions = { center: "ctr", insideEnd: "inEnd", outsideEnd: "outEnd", bestFit: "bestFit", above: "t", below: "b", left: "l", right: "r" };
  return '<c:dLbls>' + (labels.numberFormat === undefined ? "" : '<c:numFmt formatCode="' + escapeXml(labels.numberFormat, policy) + '" sourceLinked="0"/>') +
    (labels.position === undefined ? "" : '<c:dLblPos val="' + positions[labels.position] + '"/>') + '<c:showLegendKey val="0"/><c:showVal val="' + (labels.showValue === false ? 0 : 1) +
    '"/><c:showCatName val="' + (labels.showCategoryName ? 1 : 0) + '"/><c:showSerName val="' + (labels.showSeriesName ? 1 : 0) +
    '"/><c:showPercent val="' + (labels.showPercent ? 1 : 0) + '"/><c:showBubbleSize val="0"/></c:dLbls>';
}

function axesXml(chart: RegisteredChart, policy: InvalidCharacterPolicy): string {
  const horizontal = chart.type === "barClustered", value = chart.valueAxis;
  const scaling = '<c:scaling><c:orientation val="minMax"/>' + (value?.maximum === undefined ? "" : '<c:max val="' + value.maximum + '"/>') +
    (value?.minimum === undefined ? "" : '<c:min val="' + value.minimum + '"/>') + '</c:scaling>';
  const ticks = '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
  return '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (horizontal ? "l" : "b") + '"/>' +
    (chart.categoryAxis?.title ? titleXml(chart.categoryAxis.title, policy) : "") + '<c:numFmt formatCode="General" sourceLinked="1"/>' + ticks +
    '<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
    '<c:valAx><c:axId val="2"/>' + scaling + '<c:delete val="0"/><c:axPos val="' + (horizontal ? "b" : "l") + '"/><c:majorGridlines/>' +
    (value?.title ? titleXml(value.title, policy) : "") + '<c:numFmt formatCode="' + escapeXml(value?.numberFormat ?? "General", policy) + '" sourceLinked="' + (value?.numberFormat === undefined ? 1 : 0) + '"/>' + ticks +
    '<c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
}

function* chartXml(chart: RegisteredChart, policy: InvalidCharacterPolicy): Generator<string> {
  const pie = chart.type === "pie", line = chart.type === "line", tag = pie ? "pieChart" : line ? "lineChart" : "barChart";
  const source = "'" + chart.dataSheetName.replace(/'/g, "''") + "'!", end = chart.data.categories.length + 1;
  yield xmlDeclaration + '<c:chartSpace xmlns:c="' + chartNamespace + '" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + officeRelationshipsNamespace + '"><c:date1904 val="0"/>' +
    (chart.language === undefined ? "" : '<c:lang val="' + escapeXml(chart.language, policy) + '"/>') + '<c:chart>';
  if (chart.title) yield titleXml(chart.title, policy);
  yield '<c:autoTitleDeleted val="' + (chart.title ? 0 : 1) + '"/><c:plotArea><c:layout/><c:' + tag + '>';
  if (!pie) yield line ? '<c:grouping val="standard"/>' : '<c:barDir val="' + (chart.type === "barClustered" ? "bar" : "col") + '"/><c:grouping val="clustered"/>';
  yield '<c:varyColors val="' + (pie ? 1 : 0) + '"/>';
  for (let i = 0; i < chart.data.series.length; i++) {
    const series = chart.data.series[i]!, column = "$" + columnName(i + 2);
    yield '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/><c:tx>';
    const nameFormula = chart.binding ? chart.binding.references.names[i] : source + column + "$1";
    if (nameFormula === undefined) yield '<c:v>' + escapeXml(series.name, policy) + '</c:v>';
    else yield* reference("str", nameFormula, [series.name], policy);
    yield '</c:tx>';
    if (series.color) yield '<c:spPr>' + (line ? '<a:ln w="25400">' : "") + '<a:solidFill><a:srgbClr val="' + series.color + '"/></a:solidFill>' + (line ? '</a:ln>' : "") + '</c:spPr>';
    if (!line && !pie) yield '<c:invertIfNegative val="0"/>';
    yield '<c:cat>';
    yield* reference("str", chart.binding?.references.categories ?? source + '$A$2:$A$' + end, chart.data.categories, policy);
    yield '</c:cat><c:val>';
    yield* reference("num", chart.binding?.references.values[i] ?? source + column + "$2:" + column + "$" + end, series.values, policy);
    yield '</c:val>' + (line ? '<c:smooth val="0"/>' : "") + '</c:ser>';
  }
  yield labelsXml(chart, policy);
  if (!pie && !line) yield '<c:gapWidth val="150"/><c:overlap val="0"/>';
  if (!pie) yield '<c:axId val="1"/><c:axId val="2"/>';
  yield '</c:' + tag + '>';
  if (!pie) yield axesXml(chart, policy);
  yield '</c:plotArea>';
  if (chart.legend !== false) {
    const position = { right: "r", bottom: "b", left: "l", top: "t" }[chart.legend ?? "right"];
    yield '<c:legend><c:legendPos val="' + position + '"/><c:layout/><c:overlay val="0"/></c:legend>';
  }
  // Snapshot data sheets are hidden; bound sources remain visible and respect native visibility.
  yield '<c:plotVisOnly val="' + (chart.binding ? 1 : 0) + '"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart></c:chartSpace>';
}

/** @internal Stream cache points through the existing bounded, cancellation-aware XML sink. */
export async function writeChartXml(chart: RegisteredChart, destination: ByteSink, policy: InvalidCharacterPolicy, signal?: AbortSignal): Promise<void> {
  const text = new ChunkedTextSink(destination, signal);
  for (const fragment of chartXml(chart, policy)) if (text.append(fragment)) await text.flush();
  await text.close();
}
