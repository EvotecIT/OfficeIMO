namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Cached source text remains ordinary paragraph text during measurement.
    // Only final paint moves into the widget, so editing or flattening a field
    // can replace its initial appearance without leaving old glyphs underneath.
    private sealed class InlineFormAppearance {
        public double Left { get; set; }
        public double Bottom { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Baseline { get; set; }
        public List<InlineFormTextFragment> Fragments { get; } = new();
    }

    private sealed class InlineFormTextFragment {
        public RichSeg Segment { get; set; } = null!;
        public double X { get; set; }
        public double Baseline { get; set; }
        public PdfColor? DefaultColor { get; set; }
    }

    private static FormFieldAnnotation CreateInlineFormAnnotation(PdfInlineFormField field,
        double left, double bottom, double right, double top, int? parentElementIndex) {
        PdfCanvasFormFieldItem item = field.Field;
        return new FormFieldAnnotation {
            X1 = left, Y1 = bottom, X2 = right, Y2 = top,
            Kind = item.Kind == PdfCanvasFormFieldKind.Text ? FormFieldAnnotationKind.Text
                : item.Kind == PdfCanvasFormFieldKind.CheckBox ? FormFieldAnnotationKind.CheckBox
                : FormFieldAnnotationKind.Choice,
            Name = item.Name, Value = item.Value, AppearanceValue = item.AppearanceValue,
            AppearanceStyle = item.AppearanceStyle, Values = item.Values,
            FontSize = item.FontSize, IsChecked = item.IsSelected,
            CheckedValueName = item.Option, ExportValue = item.ExportValue,
            Options = item.Options, ChoiceOptions = item.ChoiceOptions,
            SelectedIndices = item.SelectedIndices, IsComboBox = item.IsComboBox,
            AllowsMultipleSelection = item.AllowsMultipleSelection,
            Style = item.Style, InlineField = field,
            StructureParentElementIndex = parentElementIndex
        };
    }

    private static void CaptureInlineFormText(LayoutResult.Page page, RichSeg segment,
        double x, double baseline, PdfColor? defaultColor, PdfOptions options,
        int? parentElementIndex, double leadingAdvance) {
        double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
        double rise = TextRiseForBaseline(segment.FontSize, segment.Baseline) + segment.LineAlignmentRise;
        double bottom = baseline + rise - GetDescenderForOptions(segment.Font, segment.NamedFont, size, options);
        double top = baseline + rise + GetAscenderForOptions(segment.Font, segment.NamedFont, size, options);
        double right = x + GetRichSegmentWidth(segment);
        bool whitespaceOnly = right <= x && segment.LeadingSpace && leadingAdvance > 0D;
        if ((right <= x && !whitespaceOnly) || top <= bottom) return;

        // One contiguous source control has one widget on each occupied line.
        // Its shared descriptor, rather than its public field name, identifies
        // continuations; unrelated duplicate names remain invalid.
        FormFieldAnnotation? field = page.FormFields.Count == 0 ? null : page.FormFields[page.FormFields.Count - 1];
        if (field == null || !ReferenceEquals(field.InlineField, segment.FormField) ||
            field.InlineAppearance == null || Math.Abs(field.InlineAppearance.Baseline - baseline) >= .001D)
            field = null;
        // Retain source-owned separators in the replaceable appearance. A
        // separator before the first glyph belongs outside the control unless
        // it arrived as its own whitespace fragment.
        double appearanceLeading = field != null || whitespaceOnly ? leadingAdvance : 0D;
        double fragmentX = x - appearanceLeading;
        if (field == null) {
            field = CreateInlineFormAnnotation(segment.FormField!, fragmentX, bottom, right, top, parentElementIndex);
            field.InlineAppearance = new InlineFormAppearance {
                Left = fragmentX, Bottom = bottom, Right = right, Top = top, Baseline = baseline
            };
            page.FormFields.Add(field);
        } else {
            field.X1 = Math.Min(field.X1, fragmentX);
            field.Y1 = Math.Min(field.Y1, bottom);
            field.X2 = Math.Max(field.X2, right);
            field.Y2 = Math.Max(field.Y2, top);
            field.InlineAppearance!.Left = field.X1;
            field.InlineAppearance.Bottom = field.Y1;
            field.InlineAppearance.Right = field.X2;
            field.InlineAppearance.Top = field.Y2;
        }
        field.InlineAppearance.Fragments.Add(new InlineFormTextFragment {
            Segment = segment.WithLeadingAdvance(appearanceLeading), X = fragmentX, Baseline = baseline,
            DefaultColor = defaultColor
        });
    }

    private static string BuildInlineFormAppearanceContent(FormFieldAnnotation field, PdfOptions options,
        Func<PdfStandardFont, PdfOptions, int> ensureFont,
        Func<PdfNamedFontFace, PdfOptions, int> ensureNamedFont,
        out IReadOnlyList<(string Name, int Id)> fontResources) {
        InlineFormAppearance appearance = field.InlineAppearance!;
        var resources = new Dictionary<string, int>(StringComparer.Ordinal);
        var content = new StringBuilder();
        foreach (InlineFormTextFragment fragment in appearance.Fragments) {
            RichSeg segment = fragment.Segment;
            string resource = GetFontResourceName(segment.Font, segment.NamedFont, ChooseNormal(options.DefaultFont));
            resources[resource] = segment.NamedFont.HasValue
                ? ensureNamedFont(segment.NamedFont.Value, options) : ensureFont(segment.Font, options);
            var paragraph = new RichParagraphBlock(Array.Empty<PdfTextRun>(), PdfAlign.Left, fragment.DefaultColor);
            var lines = new List<List<RichSeg>> { new List<RichSeg> { segment } };
            double leading = Math.Max(segment.FontSize, appearance.Top - appearance.Bottom);
            WriteRichParagraph(content, paragraph, lines, new List<double> { leading }, options,
                fragment.Baseline - appearance.Bottom, segment.FontSize, leading,
                new List<LinkAnnotation>(), xOverride: fragment.X - appearance.Left,
                widthOverride: Math.Max(.001D, GetRichSegmentWidth(segment) + segment.LeadingAdvance), captureInlineForms: false);
        }
        fontResources = resources.Select(entry => (entry.Key, entry.Value)).ToArray();
        return content.ToString();
    }

    private static bool TryBuildInlineCheckBoxAppearance(FormFieldAnnotation field, bool isChecked,
        PdfOptions options, Func<PdfStandardFont, PdfOptions, int> ensureFont,
        Func<PdfNamedFontFace, PdfOptions, int> ensureNamedFont, out string content,
        out IReadOnlyList<(string Name, int Id)> fontResources) {
        PdfTextRun? source = isChecked ? field.InlineField?.CheckedAppearanceRun
            : field.InlineField?.UncheckedAppearanceRun;
        content = string.Empty;
        fontResources = Array.Empty<(string Name, int Id)>();
        if (source == null || !field.InlineWidgetBaseline.HasValue) return false;

        // Both button states use authored glyphs when available. The original
        // baseline belongs to the text line; the annotation box merely clips
        // and carries its replaceable appearance when the field is toggled.
        var appearance = new InlineFormAppearance {
            Left = 0D, Bottom = 0D,
            Right = (field.X2 - field.X1) / field.AppearanceScale,
            Top = (field.Y2 - field.Y1) / field.AppearanceScale,
            Baseline = field.InlineWidgetBaseline.Value
        };
        double x = 0D;
        foreach (PdfTextRun run in NormalizeFallbackRuns(new[] { source }, ChooseNormal(options.DefaultFont), options)) {
            RichSeg segment = CreatePositionedTextSegment(run, options.DefaultFontSize, options);
            appearance.Fragments.Add(new InlineFormTextFragment {
                Segment = segment, X = x, Baseline = appearance.Baseline,
                DefaultColor = field.Style.MarkColor
            });
            x += GetRichSegmentWidth(segment);
        }
        content = BuildInlineFormAppearanceContent(new FormFieldAnnotation { InlineAppearance = appearance },
            options, ensureFont, ensureNamedFont, out fontResources);
        if (isChecked != field.IsChecked && MeasurePositionedTextBox(source, options) is { } measured) {
            // A PDF button has a fixed rectangle after editing. Fit a wider or
            // taller alternate glyph to that rectangle without changing the
            // source state's appearance, surrounding flow, or text baseline.
            double scaleX = measured.Width > appearance.Right ? appearance.Right / measured.Width : 1D;
            double scaleY = 1D;
            if (measured.Ascent > appearance.Top - appearance.Baseline)
                scaleY = Math.Min(scaleY, (appearance.Top - appearance.Baseline) / measured.Ascent);
            if (measured.Descent > appearance.Baseline)
                scaleY = Math.Min(scaleY, appearance.Baseline / measured.Descent);
            if (scaleX < 1D || scaleY < 1D) {
                var fitted = new StringBuilder();
                new ContentStreamBuilder(fitted).SaveState().TransformMatrix(scaleX, 0D, 0D, scaleY,
                    0D, appearance.Baseline * (1D - scaleY), preciseCoordinates: true);
                fitted.Append(content);
                new ContentStreamBuilder(fitted).RestoreState();
                content = fitted.ToString();
            }
        }
        return true;
    }
}
