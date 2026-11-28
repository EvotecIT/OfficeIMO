using System.Globalization;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    // The RTF model owns format resolution and counter advancement. This plan
    // only associates those resolved markers with destination Word resources.
    // Keep it per conversion: neither the authored model nor a global cache is modified.
    private sealed class RtfWordNumberingPlan {
        private readonly RtfDocument _source;
        private readonly OfficeXmlNode _numbering;
        private readonly Dictionary<RtfParagraph, int> _bindings = new Dictionary<RtfParagraph, int>();
        private readonly Dictionary<RtfParagraph, string> _fallback = new Dictionary<RtfParagraph, string>();
        private readonly HashSet<RtfNote> _visitedNotes = new HashSet<RtfNote>();
        private int _abstractId;
        private int _instanceId;

        private RtfWordNumberingPlan(RtfDocument source, OfficeXmlNode numbering) {
            _source = source;
            _numbering = numbering;
            _abstractId = numbering.ModelChildren("abstractNum").Select(item => WordXml.ParseInt32(item.WAttr("abstractNumId")) ?? 0).DefaultIfEmpty().Max();
            _instanceId = numbering.ModelChildren("num").Select(item => WordXml.ParseInt32(item.WAttr("numId")) ?? 0).DefaultIfEmpty().Max();
        }

        internal static RtfWordNumberingPlan Create(RtfDocument source, WordDocument destination, RtfConversionReport? report) {
            OfficeXmlNode existing = destination.NumberingNode();
            var plan = new RtfWordNumberingPlan(source, existing.Exists ? existing : WordConverterNodes.CreateDetachedNumbering());
            plan.VisitBlocks(source.Blocks, new RtfListNumbering(source), new Sequence());
            foreach (RtfHeaderFooter story in source.HeaderFooters.Distinct()) {
                plan.VisitParagraphs(story.Paragraphs, new RtfListNumbering(source), new Sequence());
            }
            if (plan._bindings.Count > 0) {
                destination.SetNumberingNode(plan._numbering);
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
            } else if (marker == null || !TryCreateLevel(effective, marker, out OfficeXmlNode level) ||
                (paragraph.ListText != null && paragraph.ListText.ToPlainText() != marker.Text + marker.Separator)) {
                _fallback[paragraph] = paragraph.ListText?.ToPlainText() ?? (marker == null ? string.Empty : marker.Text + marker.Separator);
                sequence.Signature = null;
            } else {
                // LastValue observes the canonical marker; it does not generate
                // numbers. A cached discontinuity or an anonymous-list reset needs
                // a new Word instance rather than restarting every \pnstart1 item.
                // The resolved current value, not each paragraph's legacy
                // start hint, decides whether this destination instance restarts.
                level.ModelChild("start").SetAttribute(WordXml.W, "val", "1");
                string signature = marker.Formatting.Identity + "|" + Encoding.UTF8.GetString(level.OuterXml());
                if (sequence.Signature != signature || marker.Value != sequence.LastValue + 1) {
                    level.ModelChild("start").SetAttribute(WordXml.W, "val", FormatInt(checked((int)marker.Value)));
                    int abstractId = checked(++_abstractId);
                    int instanceId = checked(++_instanceId);
                    OfficeXmlNode definition = _numbering.CreateElement(WordXml.W, "abstractNum");
                    definition.SetAttribute(WordXml.W, "abstractNumId", FormatInt(abstractId));
                    definition.Append(level);
                    _numbering.InsertBefore(definition, _numbering.Elements().FirstOrDefault(child =>
                        child.IsW("num") || child.IsW("numIdMacAtCleanup")));
                    OfficeXmlNode instance = _numbering.CreateElement(WordXml.W, "num");
                    instance.SetAttribute(WordXml.W, "numId", FormatInt(instanceId));
                    OfficeXmlNode abstractReference = instance.CreateElement(WordXml.W, "abstractNumId");
                    abstractReference.SetAttribute(WordXml.W, "val", FormatInt(abstractId));
                    instance.Append(abstractReference);
                    _numbering.InsertBefore(instance, _numbering.ModelChild("numIdMacAtCleanup"));
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

        private bool TryCreateLevel(RtfParagraph paragraph, RtfListMarker marker, out OfficeXmlNode result) {
            result = default;
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

            result = CreateWordListLevel(_numbering.CreateElement(WordXml.W, "abstractNum"), source);
            result.ModelChild("lvlText").SetAttribute(WordXml.W, "val", template);
            // Paragraph-level \li/\fi remain separate direct overrides.
            // The list-level values are the canonical resolved fallback.
            if (legacy.Alignment.HasValue) WordXmlEdits.ReplaceChild(result, "lvlJc", WordXml.LevelOrder,
                legacy.Alignment == RtfLegacyNumberingAlignment.Center ? "center"
                    : legacy.Alignment == RtfLegacyNumberingAlignment.Right ? "right" : "left");
            OfficeXmlNode properties = result.CreateElement(WordXml.W, "rPr");
            if (font != null) {
                OfficeXmlNode fonts = properties.AddOrdered(WordXml.W, "rFonts", WRPrView.Order);
                fonts.SetAttribute(WordXml.W, "ascii", font);
                fonts.SetAttribute(WordXml.W, "hAnsi", font);
                fonts.SetAttribute(WordXml.W, "eastAsia", font);
                fonts.SetAttribute(WordXml.W, "cs", font);
            }
            if (legacy.Bold.HasValue) properties.AddOrdered(WordXml.W, "b", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(legacy.Bold.Value));
            if (legacy.Italic.HasValue) properties.AddOrdered(WordXml.W, "i", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(legacy.Italic.Value));
            if (legacy.AllCaps.HasValue) properties.AddOrdered(WordXml.W, "caps", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(legacy.AllCaps.Value));
            if (legacy.SmallCaps.HasValue) properties.AddOrdered(WordXml.W, "smallCaps", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(legacy.SmallCaps.Value));
            if (legacy.Strikethrough.HasValue) properties.AddOrdered(WordXml.W, "strike", WRPrView.Order).SetAttribute(WordXml.W, "val", FormatOnOff(legacy.Strikethrough.Value));
            if (legacy.UnderlineStyle.HasValue) properties.AddOrdered(WordXml.W, "u", WRPrView.Order).SetAttribute(WordXml.W, "val", ToUnderlineToken(ToWordUnderlineStyle(legacy.UnderlineStyle.Value)));
            if (legacy.FontSizeHalfPoints.HasValue) properties.AddOrdered(WordXml.W, "sz", WRPrView.Order).SetAttribute(WordXml.W, "val", legacy.FontSizeHalfPoints.Value.ToString(CultureInfo.InvariantCulture));
            if (legacy.ForegroundColorIndex.HasValue) properties.AddOrdered(WordXml.W, "color", WRPrView.Order).SetAttribute(WordXml.W, "val", color ?? "auto");
            if (properties.HasChildren()) result.InsertOrdered(properties, WordXml.LevelOrder);
            return true;
        }

        private sealed class Sequence {
            internal string? Signature;
            internal long LastValue;
            internal int InstanceId;
        }
    }
}
