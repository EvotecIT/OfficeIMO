import { assertGroupedReportSource, groupedReportOperation } from "../core/grouped-report.js";
import type { GroupedReportEvent, GroupedReportSummary } from "../core/grouped-report.js";
import { consumeRows } from "../core/iteration.js";
import { ExportCell } from "../core/presentation.js";
import type { Column, ColumnValueContext, ExportValue } from "../core/index.js";
import type { PdfGroupedTableBlock } from "./document-types.js";
import { captureGroupedPdfTableOptions } from "./grouped-table.js";
import type { PdfSettings } from "./settings.js";
import { settings } from "./settings.js";
import { PdfTableLayout } from "./table-layout.js";
import type { PdfLayoutRow } from "./table-layout.js";
import type { PdfPages } from "./pages.js";
import type { PdfFontResources } from "./font-resources.js";
import { captureCaption, captureParagraph, prepareParagraph } from "./paragraph-layout.js";
import { synchronousText } from "./text.js";

type GroupStart = Extract<GroupedReportEvent, { readonly kind: "groupStart" }>;
interface GroupedPdfContext {
  readonly prepared: PdfSettings;
  readonly pages: PdfPages;
  readonly fonts: PdfFontResources;
  readonly controller: AbortController;
  readonly flushParagraphs: (minimumHeight?: number) => Promise<void>;
  /** Generated rows consume the shared output budget but do not add source rows to progress/results. */
  readonly acceptRow: (source: boolean) => void;
  readonly acceptColumns: (columns: number) => void;
}

/** @internal Render one analytical source in the existing document session. Retain one detail row and
 * at most four closing subtotals. All captions, rows, pages, bytes and fonts use the document owners. */
export async function renderGroupedPdfTable(block: PdfGroupedTableBlock, context: GroupedPdfContext): Promise<void> {
  const { report } = block, { pages, fonts, prepared, controller } = context;
  assertGroupedReportSource(report);
  const semantic = prepared.options.tagged !== undefined;
  const captured = captureGroupedPdfTableOptions(block.options);
  const { labelColumn, subtotalLabel, grandTotalLabel, includeGrandTotal = true, continuationLabel, groupHeadingLevel = 2, ...input } = captured;
  const columns = Object.freeze([Object.freeze({ ...labelColumn }), ...report.columns]);
  let rowIndex = -1;
  const sourceContext = (value: ColumnValueContext): ColumnValueContext => ({ rowIndex,
    columnIndex: value.columnIndex - 1, column: report.columns[value.columnIndex - 1]! });
  const table = settings({ ...pages.settings.options, ...input, columns, includeHeader: input.includeHeader ?? true,
    ...(input.formatValue === undefined ? {} : { formatValue: (value, position) => input.formatValue!(value, sourceContext(position)) }),
    ...(input.formatRuns === undefined ? {} : { formatRuns: (value, position) => input.formatRuns!(value, sourceContext(position)) }) });
  const sharedTable = { ...table, budget: prepared.budget }, layout = new PdfTableLayout(sharedTable, fonts);
  context.acceptColumns(columns.length);
  const operation = groupedReportOperation(report, controller.signal), abort = () => controller.abort(operation.signal?.reason);
  operation.signal?.addEventListener("abort", abort, { once: true });
  if (operation.signal?.aborted) abort();
  try {
    const headings = await layout.headers(), headingHeight = await layout.blockHeight(headings);
    let path: readonly GroupStart[] = [], started = false, firstFragment = false, firstBundleOversized = false, semanticDepth = 0;
    const continuationLabels = new WeakMap<GroupStart, string>();
    let pending: PdfLayoutRow | undefined;
    const closing: PdfLayoutRow[] = [];
    const summaryValues = (label: string, values: readonly (number | null)[]): readonly ExportValue[] => Object.freeze([
      new ExportCell(label, { text: label, presentation: table.options.footerPresentation! }),
      ...values.map(value => new ExportCell(value, { presentation: table.options.footerPresentation! }))
    ]);
    const prepareGroupContext = async () => {
      const caption = await prepareParagraph(captureCaption({ text: path.map(group => group.label).join("\n"),
        presentation: { bold: true }, spaceAfter: 4 }, pages.settings, fonts), pages.settings, fonts);
      const continuation = await prepareParagraph(captureCaption({ text: path.map(group => {
        if (continuationLabel === undefined) return group.label;
        const existing = semantic ? continuationLabels.get(group) : undefined;
        if (existing !== undefined) return existing;
        const label = synchronousText(continuationLabel(group), "continuationLabel");
        if (semantic) continuationLabels.set(group, label);
        return label;
      }).join("\n"), presentation: { bold: true }, spaceAfter: 4 }, pages.settings, fonts), pages.settings, fonts);
      return { caption, continuation };
    };
    const flushLastDetail = async (): Promise<void> => {
      if (pending) {
        if (!started) await startLeaf(pending);
        if (closing.length && !firstBundleOversized) await pages.keptRows([pending, ...closing], firstFragment);
        else if (closing.length) {
          await pages.row(pending, firstFragment);
          for (const row of closing) await pages.row(row);
        }
        else await pages.row(pending, firstFragment);
      } else for (const row of closing) await pages.row(row);
      pending = undefined; closing.length = 0; firstFragment = false; firstBundleOversized = false;
    };
    const finishLeaf = async (): Promise<void> => {
      if (!started && !pending) return;
      await flushLastDetail(); await pages.endTable(); started = false;
    };
    const startLeaf = async (row: PdfLayoutRow): Promise<void> => {
      const { caption, continuation } = await prepareGroupContext();
      const bundleHeight = row.height + closing.reduce((height, summary) => height + summary.height, 0);
      const available = table.page.height - table.margins.top - table.margins.bottom - headingHeight - Math.max(caption?.height ?? 0, continuation.height);
      firstBundleOversized = closing.length > 0 && bundleHeight > available + .001;
      const minimum = closing.length > 0 && !firstBundleOversized ? bundleHeight : layout.minimumRowHeight(row);
      // Defer new group headings until their first real row is known. The complete, at-most-four
      // heading chain stays with a meaningful first table fragment without retaining a group body.
      const groupHeadings = [];
      if (semantic) for (const group of path.slice(semanticDepth)) groupHeadings.push(await prepareParagraph(captureParagraph({
        kind: "paragraph", text: group.label.trim() ? group.label : "", headingLevel: (groupHeadingLevel + group.level) as 1 | 2 | 3 | 4 | 5 | 6,
        presentation: { bold: true }, keepTogether: true, keepWithNext: true, spaceAfter: 4,
        ...(group.label.trim() ? { bookmark: { title: group.label, level: (groupHeadingLevel + group.level) as 1 | 2 | 3 | 4 | 5 | 6 } } : {})
      }, pages.settings, fonts), pages.settings, fonts));
      const groupHeight = groupHeadings.reduce((height, heading) => height + heading.height, 0);
      await context.flushParagraphs((caption?.height ?? 0) + groupHeight + headingHeight + minimum);
      let followingHeight = groupHeight + caption.height + headingHeight + minimum;
      for (const heading of groupHeadings) {
        followingHeight -= heading.height;
        await pages.beginGroup(semanticDepth++, heading, followingHeight);
      }
      await pages.beginTable(layout, headings, minimum, caption, continuation, false, semantic);
      started = true; firstFragment = (caption?.lines.length ?? 0) > 0 || groupHeadings.some(heading => heading.lines.length > 0);
    };
    const semanticSubtotal = async (row: PdfLayoutRow, level: number): Promise<void> => {
      if (level === report.groupBy.length - 1) {
        // A leaf's own subtotal shares its detail table and header identities. Keep the last
        // detail/summary bundle when it fits; the existing flow owns continuation and limits.
        pages.summary(row); closing.push(row); await finishLeaf();
        await pages.endGroup(level); semanticDepth = level; path = path.slice(0, level);
        return;
      }
      await finishLeaf();
      const { caption, continuation } = await prepareGroupContext();
      const minimum = layout.minimumRowHeight(row);
      await context.flushParagraphs(caption.height + headingHeight + minimum);
      // An ancestor subtotal belongs to that ancestor's section, never the final child's table.
      // Its own visible column headings keep the separate logical summary table truthful.
      await pages.beginTable(layout, headings, minimum, caption, continuation, false, true);
      pages.summary(row); await pages.row(row, caption.lines.length > 0 || headings.length > 0); await pages.endTable();
      await pages.endGroup(level); semanticDepth = level; path = path.slice(0, level);
    };
    const complete = async (summary: GroupedReportSummary): Promise<void> => {
      await finishLeaf();
      if (!includeGrandTotal && summary.rows !== 0) return;
      const completeColumns = summary.rows === 0 ? columns : columns.map(column => {
        const { groups: _groups, ...leaf } = column as Column<readonly ExportValue[]>; return leaf;
      });
      const completeTable = settings({ ...table.options, columns: completeColumns,
        includeHeader: input.includeHeader ?? true });
      const completeLayout = summary.rows === 0 ? layout : new PdfTableLayout({ ...completeTable, budget: prepared.budget }, fonts);
      const completeHeadings = summary.rows === 0 ? headings : await completeLayout.headers();
      rowIndex = -1;
      let total: PdfLayoutRow | undefined;
      if (includeGrandTotal) {
        context.acceptRow(false);
        total = await completeLayout.data(summaryValues(synchronousText(grandTotalLabel(summary), "grandTotalLabel"), summary.totals), 0, true);
      }
      if (!completeHeadings.length && !total) return;
      const minimum = total ? completeLayout.minimumRowHeight(total) : 0;
      await context.flushParagraphs(await completeLayout.blockHeight(completeHeadings) + minimum);
      await pages.beginTable(completeLayout, completeHeadings, minimum, undefined, undefined, !semantic && summary.rows !== 0);
      if (total) { if (semantic) pages.summary(total); await pages.row(total); }
      await pages.endTable();
    };
    await consumeRows(operation.events, operation.signal, async event => {
      switch (event.kind) {
        case "groupStart":
          await finishLeaf();
          if (semantic && groupHeadingLevel + event.level > 6)
            throw new RangeError("groupHeadingLevel and the opened group depth exceed heading level 6.");
          path = Object.freeze([...path.slice(0, event.level), event]);
          break;
        case "row": {
          if (pending) await flushLastDetail();
          rowIndex = event.rowIndex; context.acceptRow(true);
          const row = await layout.data([new ExportCell("", { text: "" }), ...event.values], event.rowIndex);
          pending = row; break;
        }
        case "groupEnd": {
          rowIndex = -1; context.acceptRow(false);
          const row = await layout.data(summaryValues(synchronousText(subtotalLabel(event), "subtotalLabel"), event.totals), 0, true);
          if (semantic) await semanticSubtotal(row, event.level); else closing.push(row);
          break;
        }
        case "complete": await complete(event); break;
      }
    });
  } finally {
    operation.signal?.removeEventListener("abort", abort); operation.dispose();
  }
}
