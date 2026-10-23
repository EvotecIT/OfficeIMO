using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    // The RTF model owns format resolution and counter advancement. This plan
    // only associates those resolved markers with destination Word resources.
    // Keep it per conversion: neither the authored model nor a global cache is modified.
    private sealed class RtfWordNumberingPlan {
        private readonly RtfDocument _source;
        private readonly Numbering _numbering;
        private readonly Dictionary<RtfParagraph, int> _bindings = new Dictionary<RtfParagraph, int>();
        private readonly Dictionary<RtfParagraph, string> _fallback = new Dictionary<RtfParagraph, string>();
        private readonly HashSet<RtfNote> _visitedNotes = new HashSet<RtfNote>();
        private int _abstractId;
        private int _instanceId;

        private RtfWordNumberingPlan(RtfDocument source, Numbering numbering) {
            _source = source;
            _numbering = numbering;
            _abstractId = numbering.Elements<AbstractNum>().Select(item => item.AbstractNumberId?.Value ?? 0).DefaultIfEmpty().Max();
            _instanceId = numbering.Elements<NumberingInstance>().Select(item => item.NumberID?.Value ?? 0).DefaultIfEmpty().Max();
        }

        internal static RtfWordNumberingPlan Create(RtfDocument source, WordDocument destination, RtfConversionReport? report) {
            MainDocumentPart main = destination._wordprocessingDocument.MainDocumentPart!;
            var plan = new RtfWordNumberingPlan(source, main.NumberingDefinitionsPart?.Numbering ?? new Numbering());
            plan.VisitBlocks(source.Blocks, new RtfListNumbering(source), new Sequence());
            foreach (RtfHeaderFooter story in source.HeaderFooters.Distinct()) {
                plan.VisitParagraphs(story.Paragraphs, new RtfListNumbering(source), new Sequence());
            }
            if (plan._bindings.Count > 0) {
                NumberingDefinitionsPart part = main.NumberingDefinitionsPart ?? main.AddNewPart<NumberingDefinitionsPart>();
                if (!ReferenceEquals(part.Numbering, plan._numbering)) part.Numbering = plan._numbering;
            }
            if (plan._bindings.Count > 0) report?.Add(RtfConversionSeverity.Information, "RtfWordLegacyNumberingMapped",
                "Legacy RTF paragraph numbering was mapped to Word numbering.", RtfConversionAction.Preserved,
                feature: "legacy-numbering", count: plan._bindings.Count);
            if (plan._fallback.Count > 0) report?.Add(RtfConversionSeverity.Warning, "RtfWordLegacyNumberingFlattened",
                "Legacy numbering outside the supported editable profile was represented by its cached or resolved marker text; numbering semantics or marker formatting may differ.",
                RtfConversionAction.Flattened, feature: "legacy-numbering", count: plan._fallback.Count);
            return plan;
        }

        internal int? GetNumberId(RtfParagraph paragraph) => _bindings.TryGetValue(paragraph, out int id) ? id : (int?)null;

        internal void AppendFallback(WordParagraph destination, RtfParagraph paragraph) {
            if (_fallback.TryGetValue(paragraph, out string? text)) destination.AddText(text);
        }

        private void VisitBlocks(IEnumerable<IRtfBlock> blocks, RtfListNumbering numbering, Sequence sequence) {
            foreach (IRtfBlock block in blocks) {
                if (block is RtfParagraph paragraph) VisitParagraph(paragraph, numbering, sequence);
                else if (block is RtfTable table) {
                    foreach (RtfTableRow row in table.Rows)
                        foreach (RtfTableCell cell in row.Cells) VisitBlocks(cell.Blocks, numbering, sequence);
                }
            }
        }

        private void VisitParagraphs(IEnumerable<RtfParagraph> paragraphs, RtfListNumbering numbering, Sequence sequence) {
            foreach (RtfParagraph paragraph in paragraphs) VisitParagraph(paragraph, numbering, sequence);
        }

        private void VisitParagraph(RtfParagraph paragraph, RtfListNumbering numbering, Sequence sequence) {
            RtfListMarker? marker = numbering.Next(paragraph);
            RtfParagraph effective = _source.GetParagraphFormatting(paragraph);
            if (effective.ListId.HasValue || !effective.LegacyNumbering.HasAnyValue) {
                sequence.Signature = null;
            } else if (marker == null || !TryCreateLevel(effective, marker, out Level? level) ||
                (paragraph.ListText != null && paragraph.ListText.ToPlainText() != marker.Text + marker.Separator)) {
                _fallback[paragraph] = paragraph.ListText?.ToPlainText() ?? (marker == null ? string.Empty : marker.Text + marker.Separator);
                sequence.Signature = null;
            } else {
                // LastValue observes the canonical marker; it does not generate
                // numbers. A cached discontinuity or an anonymous-list reset needs
                // a new Word instance rather than restarting every \pnstart1 item.
                // The resolved current value, not each paragraph's legacy
                // start hint, decides whether this destination instance restarts.
                level!.StartNumberingValue!.Val = 1;
                string signature = marker.Formatting.Identity + "|" + level.OuterXml;
                if (sequence.Signature != signature || marker.Value != sequence.LastValue + 1) {
                    level.StartNumberingValue!.Val = checked((int)marker.Value);
                    int abstractId = checked(++_abstractId);
                    int instanceId = checked(++_instanceId);
                    var definition = new AbstractNum { AbstractNumberId = abstractId };
                    definition.Append(level);
                    _numbering.InsertBefore(definition, _numbering.ChildElements.FirstOrDefault(child =>
                        child is NumberingInstance || child is NumberingIdMacAtCleanup));
                    _numbering.InsertBefore(new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = instanceId },
                        _numbering.GetFirstChild<NumberingIdMacAtCleanup>());
                    sequence.InstanceId = instanceId;
                    sequence.Signature = signature;
                }
                sequence.LastValue = marker.Value;
                _bindings.Add(paragraph, sequence.InstanceId);
            }
            VisitInlineNotes(paragraph.Inlines);
        }

        private void VisitInlineNotes(IEnumerable<IRtfInline> inlines) {
            foreach (IRtfInline inline in inlines) {
                RtfNote? note = inline is RtfRun run ? run.Note : inline is RtfGeneratedText generated ? generated.Note : null;
                if (note != null && _visitedNotes.Add(note))
                    VisitParagraphs(note.Paragraphs, new RtfListNumbering(_source), new Sequence());
                if (inline is RtfField field) VisitInlineNotes(field.Result.Inlines);
            }
        }

        private bool TryCreateLevel(RtfParagraph paragraph, RtfListMarker marker, out Level? result) {
            result = null;
            RtfLegacyNumbering legacy = paragraph.LegacyNumbering;
            RtfListLevel source = marker.Formatting.Level;
            if (marker.Formatting.LevelIndex != 0 || legacy.Level is > 1 || marker.Value < 0 || marker.Value > int.MaxValue ||
                legacy.LevelKind == RtfLegacyNumberingLevelKind.Continue || legacy.IncludePreviousLevels == true ||
                legacy.NumberEachCellOnce == true || legacy.NumberAcrossRows == true || legacy.RestartAfterSection == true ||
                legacy.HangingIndent == true || legacy.SpaceTwips.HasValue || legacy.FontSizeHalfPoints is <= 0 ||
                (source.Kind != RtfListKind.Bullet && source.NumberFormat is not null and not 0)) return false;

            string template = source.Text ?? (source.Kind == RtfListKind.Bullet ? "•" : "%1.");
            if (source.Kind == RtfListKind.Bullet) template = marker.Text;
            else if (marker.Text != template.Replace("%1", marker.Value.ToString(CultureInfo.InvariantCulture))) return false;
            string? font = null;
            if (legacy.FontId.HasValue && !TryGetFontName(_source, legacy.FontId.Value, out font)) return false;
            string? color = legacy.ForegroundColorIndex.HasValue ? GetColorHex(_source, legacy.ForegroundColorIndex.Value) : null;
            if (legacy.ForegroundColorIndex is > 0 && color == null) return false;

            result = CreateWordListLevel(source);
            result.LevelText!.Val = template;
            // Paragraph-level \li/\fi remain separate direct overrides.
            // The list-level values are the canonical resolved fallback.
            if (legacy.Alignment.HasValue) result.LevelJustification = new LevelJustification {
                Val = legacy.Alignment == RtfLegacyNumberingAlignment.Center ? LevelJustificationValues.Center
                    : legacy.Alignment == RtfLegacyNumberingAlignment.Right ? LevelJustificationValues.Right : LevelJustificationValues.Left
            };
            var properties = new NumberingSymbolRunProperties();
            if (font != null) properties.AddChild(new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font }, true);
            if (legacy.Bold.HasValue) properties.AddChild(new Bold { Val = legacy.Bold.Value }, true);
            if (legacy.Italic.HasValue) properties.AddChild(new Italic { Val = legacy.Italic.Value }, true);
            if (legacy.AllCaps.HasValue) properties.AddChild(new Caps { Val = legacy.AllCaps.Value }, true);
            if (legacy.SmallCaps.HasValue) properties.AddChild(new SmallCaps { Val = legacy.SmallCaps.Value }, true);
            if (legacy.Strike.HasValue) properties.AddChild(new Strike { Val = legacy.Strike.Value }, true);
            if (legacy.UnderlineStyle.HasValue) properties.AddChild(new Underline { Val = ToWordUnderlineStyle(legacy.UnderlineStyle.Value) }, true);
            if (legacy.FontSizeHalfPoints.HasValue) properties.AddChild(new FontSize { Val = legacy.FontSizeHalfPoints.Value.ToString(CultureInfo.InvariantCulture) }, true);
            if (legacy.ForegroundColorIndex.HasValue) properties.AddChild(new Color { Val = color ?? "auto" }, true);
            if (properties.HasChildren) result.AddChild(properties, true);
            return true;
        }

        private sealed class Sequence {
            internal string? Signature;
            internal long LastValue;
            internal int InstanceId;
        }
    }
}
