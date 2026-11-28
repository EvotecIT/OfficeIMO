using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        // Shared immutable snapshots keep nested table projection off the thread stack.
        private sealed record NativeTableStyleDefaults(PdfCore.PdfCellPadding? CellPadding, global::OfficeIMO.OfficeColor? CellFill, PdfCore.PdfCellVerticalAlign? CellVerticalAlignment, (global::OfficeIMO.OfficeColor Color, double Width)? TableBorder, OfficeXmlNode Borders, OfficeXmlNode PreferredWidth, string? Layout, double? LeftIndent, double? CellSpacing, string? Alignment, double? ParagraphLineHeight, double? ParagraphLineSpacingPoints, string? ParagraphLineSpacingRule, double? ParagraphSpacingBefore, double? ParagraphSpacingAfter, string? ParagraphAlignment, double? ParagraphLeftIndent, double? ParagraphRightIndent, double? ParagraphFirstLineIndent, NativeTableRunStyleDefaults RunStyle, NativeTableConditionalStyleDefaults FirstRowStyle, NativeTableConditionalStyleDefaults LastRowStyle, NativeTableConditionalStyleDefaults FirstColumnStyle, NativeTableConditionalStyleDefaults LastColumnStyle, NativeTableConditionalStyleDefaults Band1HorizontalStyle, NativeTableConditionalStyleDefaults Band1VerticalStyle) {
            public NativeLineSpacing LineSpacing { get; init; }
            public NativeParagraphSpacing ParagraphSpacing { get; init; }
            public bool? SeparateCellBorders { get; init; }
            public bool RightToLeft { get; init; }
            // An explicit PDF default supplies missing cell typography while
            // authored run, paragraph and named table styles retain precedence.
            public bool UseConfiguredTypography { get; init; }
            public NativeParagraphPaginationDefaults ParagraphPagination { get; init; }
            public NativeTableConditionalStyleDefaults Band2HorizontalStyle { get; init; }
            public NativeTableConditionalStyleDefaults Band2VerticalStyle { get; init; }
            public int RowBandSize { get; init; }
            public int ColumnBandSize { get; init; }
            public static NativeTableStyleDefaults Empty { get; } = new(null, null, null, null, default, default, null, null, null, null, null, null, null, null, null, null, null, null, null, NativeTableRunStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty, NativeTableConditionalStyleDefaults.Empty);
        }

        private readonly record struct NativeTableRunStyleDefaults(double? FontSize, string? FontFamily, bool? Bold, bool? Italic, OfficeTextDecorationStyle? UnderlineStyle, OfficeTextDecorationStyle? StrikeStyle, bool? AllCaps, string? Baseline, string? ColorHex, string? Highlight, global::OfficeIMO.OfficeColor? Color) {
            public bool? SmallCaps { get; init; }
            public NativeLatinFontFamilies FontFamilies { get; init; }
            public NativeComplexScriptDefaults ComplexScript { get; init; }
            public NativeTextSpacing TextSpacing { get; init; }
            public static NativeTableRunStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null);
        }

        private readonly record struct NativeTableConditionalStyleDefaults(global::OfficeIMO.OfficeColor? CellFill, OfficeXmlNode CellBorders, PdfCore.PdfCellPadding? CellPadding, PdfCore.PdfCellVerticalAlign? CellVerticalAlignment, global::OfficeIMO.OfficeColor? TextColor, double? FontSize, string? FontFamily, bool? Bold, bool? Italic, OfficeTextDecorationStyle? UnderlineStyle, OfficeTextDecorationStyle? StrikeStyle, bool? AllCaps, string? Baseline, string? Highlight, double? ParagraphLineHeight, double? ParagraphLineSpacingPoints, string? ParagraphLineSpacingRule, double? ParagraphSpacingBefore, double? ParagraphSpacingAfter, string? ParagraphAlignment, double? ParagraphLeftIndent, double? ParagraphRightIndent, double? ParagraphFirstLineIndent) {
            public bool? SmallCaps { get; init; }
            public NativeLatinFontFamilies FontFamilies { get; init; }
            public NativeLineSpacing LineSpacing { get; init; }
            public NativeParagraphSpacing ParagraphSpacing { get; init; }
            public NativeComplexScriptDefaults ComplexScript { get; init; }
            public NativeTextSpacing TextSpacing { get; init; }
            public NativeParagraphPaginationDefaults ParagraphPagination { get; init; }
            /// <summary>True when the condition declares a text color; a declared automatic color leaves <see cref="TextColor"/> null.</summary>
            public bool TextColorDeclared { get; init; }
            public static NativeTableConditionalStyleDefaults Empty { get; } = new(null, default, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
        }

        // The w:tblStylePr type tokens of the conditional formats a table style can declare.
        private const string NativeFirstRowStyle = "firstRow", NativeLastRowStyle = "lastRow",
            NativeFirstColumnStyle = "firstCol", NativeLastColumnStyle = "lastCol",
            NativeBand1HorizontalStyle = "band1Horz", NativeBand1VerticalStyle = "band1Vert",
            NativeBand2HorizontalStyle = "band2Horz", NativeBand2VerticalStyle = "band2Vert";

        private static NativeTableStyleDefaults GetNativeTableStyleDefaults(WordTable table, NativeDocumentDefaults nativeDefaults, bool ignoreFallbackTableStyle) {
            string? styleId = GetNativeTableStyleId(table);
            // An unnamed table may inherit an authored custom default. Resolve
            // that chain before applying the configured PDF fallback; built-in
            // Normal Table definitions already produce an empty chain.
            IReadOnlyList<OfficeXmlNode> styleChain = GetNativeTableStyleChain(table.Document, styleId);
            bool rightToLeft = WordConverterNodes.IsTableRightToLeft(table.PropertiesNode());
            if (styleChain.Count == 0) {
                return NativeTableStyleDefaults.Empty with { UseConfiguredTypography = ignoreFallbackTableStyle, RightToLeft = rightToLeft };
            }

            double? marginTop = null;
            double? marginBottom = null;
            double? marginLeft = null;
            double? marginRight = null;
            global::OfficeIMO.OfficeColor? cellFill = null;
            PdfCore.PdfCellVerticalAlign? cellVerticalAlignment = null;
            (global::OfficeIMO.OfficeColor Color, double Width)? tableBorder = null;
            OfficeXmlNode tableBorders = default;
            OfficeXmlNode preferredWidth = default;
            string? layout = null;
            double? leftIndent = null;
            double? cellSpacing = null;
            bool? separateCellBorders = null;
            string? alignment = null;
            double? paragraphLineHeight = null;
            double? paragraphLineSpacingPoints = null;
            string? paragraphLineSpacingRule = null;
            NativeLineSpacing authoredLineSpacing = default;
            NativeParagraphSpacing paragraphSpacing = default;
            NativeParagraphPaginationDefaults paragraphPagination = default;
            double? paragraphSpacingBefore = null;
            double? paragraphSpacingAfter = null;
            string? paragraphAlignment = null;
            double? paragraphLeftIndent = null;
            double? paragraphRightIndent = null;
            double? paragraphFirstLineIndent = null;
            NativeComplexScriptDefaults complexScript = default;
            NativeTextSpacing textSpacing = default;
            double? fontSize = null;
            string? fontFamily = null;
            NativeLatinFontFamilies fontFamilies = default;
            bool? bold = null;
            bool? italic = null;
            OfficeTextDecorationStyle? underlineStyle = null;
            OfficeTextDecorationStyle? strikeStyle = null;
            bool? allCaps = null;
            bool? smallCaps = null;
            string? baseline = null;
            string? colorHex = null;
            string? highlight = null;
            NativeTableConditionalStyleDefaults firstRowStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults lastRowStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults firstColumnStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults lastColumnStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults band1HorizontalStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults band1VerticalStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults band2HorizontalStyle = NativeTableConditionalStyleDefaults.Empty;
            NativeTableConditionalStyleDefaults band2VerticalStyle = NativeTableConditionalStyleDefaults.Empty;
            int rowBandSize = 1, columnBandSize = 1;

            foreach (OfficeXmlNode style in styleChain) {
                OfficeXmlNode runProperties = style.WChild("rPr");
                complexScript = complexScript.Merge(runProperties);
                textSpacing = textSpacing.Merge(runProperties);
                fontSize = GetNativeStyleFontSize(runProperties) ?? fontSize;
                fontFamilies = GetNativeRunFontFamilies(table.Document, runProperties.WChild("rFonts")).Inherit(fontFamilies);
                fontFamily = fontFamilies.Primary;
                bold = runProperties.WChild("b").OnOff() ?? bold;
                italic = runProperties.WChild("i").OnOff() ?? italic;
                underlineStyle = MapNativeUnderlineStyle(runProperties.WChild("u")) ?? underlineStyle;
                strikeStyle = MapNativeStrikeStyle(runProperties) ?? strikeStyle;
                allCaps = runProperties.WChild("caps").OnOff() ?? allCaps;
                smallCaps = runProperties.WChild("smallCaps").OnOff() ?? smallCaps;
                baseline = runProperties.ChildVal("vertAlign") ?? baseline;
                colorHex = WordConverterNodes.ResolveColorHex(runProperties.WChild("color"), table.Document) ?? colorHex;
                highlight = WordConverterNodes.ReadRunHighlight(runProperties) ?? highlight;

                OfficeXmlNode tableProperties = style.WChild("tblPr");
                rowBandSize = WordXml.ParseInt32(tableProperties.ChildVal("tblStyleRowBandSize")) ?? rowBandSize;
                columnBandSize = WordXml.ParseInt32(tableProperties.ChildVal("tblStyleColBandSize")) ?? columnBandSize;
                OfficeXmlNode styleWidth = tableProperties.WChild("tblW");
                if (styleWidth.Exists) preferredWidth = styleWidth;
                layout = tableProperties.WChild("tblLayout").WAttr("type") ?? layout;
                leftIndent = GetNativeTableLeftIndent(tableProperties.WChild("tblInd")) ?? leftIndent;
                OfficeXmlNode spacingDeclaration = tableProperties.WChild("tblCellSpacing");
                if (GetNativeTableCellSpacing(spacingDeclaration) is { } resolvedCellSpacing) {
                    cellSpacing = resolvedCellSpacing;
                    separateCellBorders = spacingDeclaration.WAttr("type") == "dxa";
                }
                alignment = WordConverterNodes.ReadTableAlignment(tableProperties.WChild("jc"), rightToLeft) ?? alignment;

                OfficeXmlNode margins = tableProperties.WChild("tblCellMar");
                if (margins.Exists) {
                    (double? top, double? bottom, double? left, double? right) = ReadNativeTableCellMarginDefault(margins);

                    marginTop = top ?? marginTop;
                    marginBottom = bottom ?? marginBottom;
                    marginLeft = left ?? marginLeft;
                    marginRight = right ?? marginRight;
                }

                OfficeXmlNode tableCellProperties = style.WChild("tcPr");
                OfficeXmlNode shading = tableCellProperties.WChild("shd");
                if (shading.IsNone) shading = tableProperties.WChild("shd");
                if (shading.Exists) {
                    cellFill = ParseNativeColor(WordConverterNodes.ResolveShadingFillColorHex(shading, table.Document));
                }

                cellVerticalAlignment = MapNativeNullableCellVerticalAlign(tableCellProperties.ChildVal("vAlign")) ?? cellVerticalAlignment;

                OfficeXmlNode styleTableBorders = ResolveNativeTableBorderColors(tableProperties.WChild("tblBorders"), table.Document);
                if (styleTableBorders.Exists) {
                    tableBorder = GetNativeUniformTableBorder(styleTableBorders);
                    tableBorders = styleTableBorders;
                }

                OfficeXmlNode paragraphProperties = style.WChild("pPr");
                paragraphSpacing = paragraphSpacing.Merge(paragraphProperties);
                paragraphPagination = paragraphPagination.Merge(paragraphProperties);
                OfficeXmlNode indentation = paragraphProperties.WChild("ind");
                if (indentation.Exists) {
                    paragraphLeftIndent = ConvertNativeTwipsToPoints(indentation.WAttr("left")) ?? paragraphLeftIndent;
                    paragraphRightIndent = ConvertNativeTwipsToPoints(indentation.WAttr("right")) ?? paragraphRightIndent;

                    double? firstLine = ConvertNativeTwipsToPoints(indentation.WAttr("firstLine"));
                    double? hanging = ConvertNativeTwipsToPoints(indentation.WAttr("hanging"));
                    if (hanging.HasValue) {
                        paragraphFirstLineIndent = -hanging.Value;
                    } else if (firstLine.HasValue) {
                        paragraphFirstLineIndent = firstLine.Value;
                    }
                }

                OfficeXmlNode spacing = paragraphProperties.WChild("spacing");
                authoredLineSpacing = ReadNativeLineSpacing(spacing, paragraphProperties.WChild("textAlignment")).Inherit(authoredLineSpacing);
                if (spacing.Exists) {
                    double naturalLineHeight = ResolveNativeWordSingleLineHeight(fontFamily, nativeDefaults.FontFamily);
                    double? styleParagraphLineHeight = GetNativeTableStyleParagraphLineHeight(spacing, fontFamily, nativeDefaults.FontFamily);
                    double? styleParagraphLineSpacingPoints = GetNativeTableStyleParagraphLineSpacingPoints(spacing);
                    if (styleParagraphLineHeight.HasValue || styleParagraphLineSpacingPoints.HasValue) {
                        paragraphLineHeight = styleParagraphLineHeight;
                        paragraphLineSpacingPoints = styleParagraphLineSpacingPoints;
                        paragraphLineSpacingRule = spacing.WAttr("lineRule");
                    }

                    double effectiveFontSize = fontSize ?? nativeDefaults.FontSize;
                    double effectiveLineHeight = styleParagraphLineSpacingPoints.HasValue && effectiveFontSize > 0D
                        ? ResolveNativeLineSpacingHeight(styleParagraphLineSpacingPoints.Value, spacing.WAttr("lineRule"), effectiveFontSize, naturalLineHeight)
                        : styleParagraphLineHeight ?? paragraphLineHeight ?? naturalLineHeight;
                    paragraphSpacingBefore = GetNativeSpacingBeforePoints(spacing, effectiveFontSize, effectiveLineHeight) ?? paragraphSpacingBefore;
                    paragraphSpacingAfter = GetNativeSpacingAfterPoints(spacing, effectiveFontSize, effectiveLineHeight) ?? paragraphSpacingAfter;
                }

                paragraphAlignment = paragraphProperties.ChildVal("jc") ?? paragraphAlignment;
            }

            // Resolve relative conditional spacing after the complete table style chain,
            // so a derived whole-table font also governs inherited conditional spacing.
            var conditionalDefaults = nativeDefaults with {
                FontSize = fontSize ?? nativeDefaults.FontSize,
                FontFamily = fontFamily ?? nativeDefaults.FontFamily,
                FontFamilies = fontFamilies.Primary != null ? fontFamilies : nativeDefaults.FontFamilies
            };
            var conditionalFontDefaults = ResolveNativeConditionalFontDefaults(styleChain, table.Document, conditionalDefaults);
            foreach (OfficeXmlNode style in styleChain) {
                firstRowStyle = GetNativeTableConditionalStyleDefaults(style, NativeFirstRowStyle, firstRowStyle, conditionalFontDefaults[NativeFirstRowStyle], table.Document);
                lastRowStyle = GetNativeTableConditionalStyleDefaults(style, NativeLastRowStyle, lastRowStyle, conditionalFontDefaults[NativeLastRowStyle], table.Document);
                firstColumnStyle = GetNativeTableConditionalStyleDefaults(style, NativeFirstColumnStyle, firstColumnStyle, conditionalFontDefaults[NativeFirstColumnStyle], table.Document);
                lastColumnStyle = GetNativeTableConditionalStyleDefaults(style, NativeLastColumnStyle, lastColumnStyle, conditionalFontDefaults[NativeLastColumnStyle], table.Document);
                band1HorizontalStyle = GetNativeTableConditionalStyleDefaults(style, NativeBand1HorizontalStyle, band1HorizontalStyle, conditionalFontDefaults[NativeBand1HorizontalStyle], table.Document);
                band1VerticalStyle = GetNativeTableConditionalStyleDefaults(style, NativeBand1VerticalStyle, band1VerticalStyle, conditionalFontDefaults[NativeBand1VerticalStyle], table.Document);
                band2HorizontalStyle = GetNativeTableConditionalStyleDefaults(style, NativeBand2HorizontalStyle, band2HorizontalStyle, conditionalFontDefaults[NativeBand2HorizontalStyle], table.Document);
                band2VerticalStyle = GetNativeTableConditionalStyleDefaults(style, NativeBand2VerticalStyle, band2VerticalStyle, conditionalFontDefaults[NativeBand2VerticalStyle], table.Document);
            }

            PdfCore.PdfCellPadding? cellPadding = marginTop.HasValue || marginBottom.HasValue || marginLeft.HasValue || marginRight.HasValue
                ? new PdfCore.PdfCellPadding {
                    Top = (marginTop) is { } points17760 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points17760) : null,
                    Bottom = (marginBottom) is { } points17800 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points17800) : null,
                    Left = (marginLeft) is { } points17841 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points17841) : null,
                    Right = (marginRight) is { } points17881 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points17881) : null
                }
                : null;

            return new NativeTableStyleDefaults(
                cellPadding,
                cellFill,
                cellVerticalAlignment,
                tableBorder,
                tableBorders,
                preferredWidth,
                layout,
                leftIndent,
                cellSpacing,
                alignment,
                paragraphLineHeight,
                paragraphLineSpacingPoints,
                paragraphLineSpacingRule,
                paragraphSpacingBefore,
                paragraphSpacingAfter,
                paragraphAlignment,
                paragraphLeftIndent,
                paragraphRightIndent,
                paragraphFirstLineIndent,
                new NativeTableRunStyleDefaults(
                    fontSize,
                    fontFamily,
                    bold,
                    italic,
                    underlineStyle,
                    strikeStyle,
                    allCaps,
                    baseline,
                    colorHex,
                    highlight,
                    null) { SmallCaps = smallCaps, FontFamilies = fontFamilies, ComplexScript = complexScript, TextSpacing = textSpacing },
                firstRowStyle,
                lastRowStyle,
                firstColumnStyle,
                lastColumnStyle,
                band1HorizontalStyle,
                band1VerticalStyle) { LineSpacing = authoredLineSpacing, ParagraphSpacing = paragraphSpacing,
                    ParagraphPagination = paragraphPagination, SeparateCellBorders = separateCellBorders, RightToLeft = rightToLeft,
                    Band2HorizontalStyle = band2HorizontalStyle, Band2VerticalStyle = band2VerticalStyle,
                    RowBandSize = rowBandSize, ColumnBandSize = columnBandSize };
        }

        private static Dictionary<string, NativeDocumentDefaults> ResolveNativeConditionalFontDefaults(
            IReadOnlyList<OfficeXmlNode> styleChain, WordDocument document, NativeDocumentDefaults tableDefaults) {
            var result = new Dictionary<string, NativeDocumentDefaults>(StringComparer.Ordinal);
            foreach (string type in new[] {
                NativeFirstRowStyle, NativeLastRowStyle,
                NativeFirstColumnStyle, NativeLastColumnStyle,
                NativeBand1HorizontalStyle, NativeBand1VerticalStyle,
                NativeBand2HorizontalStyle, NativeBand2VerticalStyle
            }) {
                NativeDocumentDefaults defaults = tableDefaults;
                foreach (OfficeXmlNode style in styleChain) {
                    foreach (OfficeXmlNode properties in style.WElements("tblStylePr").Where(properties => properties.WAttr("type") == type)) {
                        OfficeXmlNode runs = properties.WChild("rPr");
                        NativeLatinFontFamilies fonts = GetNativeRunFontFamilies(document, runs.WChild("rFonts")).Inherit(defaults.FontFamilies);
                        defaults = defaults with {
                            FontSize = GetNativeRunPropertiesBaseStyleFontSize(runs) ?? defaults.FontSize,
                            FontFamily = fonts.Primary ?? defaults.FontFamily,
                            FontFamilies = fonts
                        };
                    }
                }
                result[type] = defaults;
            }
            return result;
        }

        private static NativeTableConditionalStyleDefaults GetNativeTableConditionalStyleDefaults(OfficeXmlNode style, string type, NativeTableConditionalStyleDefaults inherited, NativeDocumentDefaults nativeDefaults, WordDocument document) {
            NativeTableConditionalStyleDefaults result = inherited;
            foreach (OfficeXmlNode properties in style.WElements("tblStylePr").Where(properties => properties.WAttr("type") == type)) {
                OfficeXmlNode cellProperties = properties.WChild("tcPr");
                global::OfficeIMO.OfficeColor? cellFill = ParseNativeColor(WordConverterNodes.ResolveShadingFillColorHex(cellProperties.WChild("shd"), document));
                OfficeXmlNode cellBorders = ResolveNativeTableBorderColors(cellProperties.WChild("tcBorders"), document);
                PdfCore.PdfCellPadding? cellPadding = CreateNativeConditionalTableCellPadding(cellProperties.WChild("tcMar"));
                PdfCore.PdfCellVerticalAlign? cellVerticalAlignment = MapNativeNullableCellVerticalAlign(cellProperties.ChildVal("vAlign"));

                OfficeXmlNode runProperties = properties.WChild("rPr");
                NativeLatinFontFamilies fontFamilies = GetNativeRunFontFamilies(document, runProperties.WChild("rFonts")).Inherit(result.FontFamilies);
                string? fontFamily = fontFamilies.Primary;
                OfficeXmlNode textColorDeclaration = runProperties.WChild("color");
                global::OfficeIMO.OfficeColor? textColor = ParseNativeColor(WordConverterNodes.ResolveColorHex(textColorDeclaration, document));
                double? fontSize = GetNativeRunPropertiesBaseStyleFontSize(runProperties);
                bool? bold = runProperties.WChild("b").OnOff();
                bool? italic = runProperties.WChild("i").OnOff();
                OfficeTextDecorationStyle? underlineStyle = MapNativeUnderlineStyle(runProperties.WChild("u"));
                OfficeTextDecorationStyle? strikeStyle = MapNativeStrikeStyle(runProperties);
                bool? allCaps = runProperties.WChild("caps").OnOff();
                bool? smallCaps = runProperties.WChild("smallCaps").OnOff();
                string? baseline = runProperties.ChildVal("vertAlign");
                string? highlight = WordConverterNodes.ReadRunHighlight(runProperties);

                OfficeXmlNode paragraphProperties = properties.WChild("pPr");
                OfficeXmlNode spacing = paragraphProperties.WChild("spacing");
                double? paragraphLineHeight = null;
                double? paragraphLineSpacingPoints = null;
                string? paragraphLineSpacingRule = null;
                double? paragraphSpacingBefore = null;
                double? paragraphSpacingAfter = null;
                double? paragraphLeftIndent = null;
                double? paragraphRightIndent = null;
                double? paragraphFirstLineIndent = null;
                OfficeXmlNode indentation = paragraphProperties.WChild("ind");
                if (indentation.Exists) {
                    paragraphLeftIndent = ConvertNativeTwipsToPoints(indentation.WAttr("left"));
                    paragraphRightIndent = ConvertNativeTwipsToPoints(indentation.WAttr("right"));

                    double? firstLine = ConvertNativeTwipsToPoints(indentation.WAttr("firstLine"));
                    double? hanging = ConvertNativeTwipsToPoints(indentation.WAttr("hanging"));
                    if (hanging.HasValue) {
                        paragraphFirstLineIndent = -hanging.Value;
                    } else if (firstLine.HasValue) {
                        paragraphFirstLineIndent = firstLine.Value;
                    }
                }

                if (spacing.Exists) {
                    double naturalLineHeight = ResolveNativeWordSingleLineHeight(nativeDefaults.FontFamily);
                    paragraphLineHeight = GetNativeTableStyleParagraphLineHeight(spacing, nativeDefaults.FontFamily);
                    paragraphLineSpacingPoints = GetNativeTableStyleParagraphLineSpacingPoints(spacing);
                    paragraphLineSpacingRule = paragraphLineHeight.HasValue || paragraphLineSpacingPoints.HasValue
                        ? spacing.WAttr("lineRule")
                        : null;

                    double effectiveFontSize = nativeDefaults.FontSize;
                    double effectiveLineHeight = paragraphLineSpacingPoints.HasValue && effectiveFontSize > 0D
                        ? ResolveNativeLineSpacingHeight(paragraphLineSpacingPoints.Value, spacing.WAttr("lineRule"), effectiveFontSize, naturalLineHeight)
                        : paragraphLineHeight ?? result.ParagraphLineHeight ?? naturalLineHeight;
                    paragraphSpacingBefore = GetNativeSpacingBeforePoints(spacing, effectiveFontSize, effectiveLineHeight);
                    paragraphSpacingAfter = GetNativeSpacingAfterPoints(spacing, effectiveFontSize, effectiveLineHeight);
                }

                string? paragraphAlignment = paragraphProperties.ChildVal("jc");
                result = new NativeTableConditionalStyleDefaults(
                    cellFill ?? result.CellFill,
                    cellBorders.Exists ? cellBorders : result.CellBorders,
                    MergeNativeCellPadding(result.CellPadding, cellPadding),
                    cellVerticalAlignment ?? result.CellVerticalAlignment,
                    textColorDeclaration.Exists ? textColor : result.TextColor,
                    fontSize ?? result.FontSize,
                    fontFamily ?? result.FontFamily,
                    bold ?? result.Bold,
                    italic ?? result.Italic,
                    underlineStyle ?? result.UnderlineStyle,
                    strikeStyle ?? result.StrikeStyle,
                    allCaps ?? result.AllCaps,
                    baseline ?? result.Baseline,
                    highlight ?? result.Highlight,
                    paragraphLineHeight ?? result.ParagraphLineHeight,
                    paragraphLineSpacingPoints ?? result.ParagraphLineSpacingPoints,
                    paragraphLineSpacingRule ?? result.ParagraphLineSpacingRule,
                    paragraphSpacingBefore ?? result.ParagraphSpacingBefore,
                    paragraphSpacingAfter ?? result.ParagraphSpacingAfter,
                    paragraphAlignment ?? result.ParagraphAlignment,
                    paragraphLeftIndent ?? result.ParagraphLeftIndent,
                    paragraphRightIndent ?? result.ParagraphRightIndent,
                paragraphFirstLineIndent ?? result.ParagraphFirstLineIndent) {
                    SmallCaps = smallCaps ?? result.SmallCaps,
                    FontFamilies = fontFamilies,
                    ComplexScript = result.ComplexScript.Merge(runProperties),
                    TextSpacing = result.TextSpacing.Merge(runProperties),
                    LineSpacing = ReadNativeLineSpacing(spacing, paragraphProperties.WChild("textAlignment")).Inherit(result.LineSpacing),
                    ParagraphSpacing = result.ParagraphSpacing.Merge(paragraphProperties),
                    ParagraphPagination = result.ParagraphPagination.Merge(paragraphProperties),
                    TextColorDeclared = textColorDeclaration.Exists || result.TextColorDeclared
                };
            }

            return result;
        }

        private static PdfCore.PdfCellPadding? CreateNativeConditionalTableCellPadding(OfficeXmlNode margins) {
            if (margins.IsNone) {
                return null;
            }

            double? top = ConvertNativeTwipsToPoints(margins.WChild("top").WAttr("w"));
            double? bottom = ConvertNativeTwipsToPoints(margins.WChild("bottom").WAttr("w"));
            double? left = ConvertNativeTwipsToPoints(margins.WChild("left").WAttr("w"));
            double? right = ConvertNativeTwipsToPoints(margins.WChild("right").WAttr("w"));
            if (!top.HasValue && !bottom.HasValue && !left.HasValue && !right.HasValue) {
                return null;
            }

            return new PdfCore.PdfCellPadding {
                Top = (top) is { } points29644 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29644) : null,
                Bottom = (bottom) is { } points29674 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29674) : null,
                Left = (left) is { } points29705 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29705) : null,
                Right = (right) is { } points29735 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29735) : null
            };
        }

        private static double? GetNativeRunPropertiesBaseStyleFontSize(OfficeXmlNode runProperties) {
            string? value = runProperties.WChild("sz").WVal();
            if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double halfPoints) ||
                halfPoints <= 0D ||
                double.IsNaN(halfPoints) ||
                double.IsInfinity(halfPoints)) {
                return null;
            }

            return halfPoints / 2D;
        }

        private static double? GetNativeTableStyleParagraphLineHeight(
            OfficeXmlNode spacing,
            params string?[] fontFamilies) {
            if (ReadNativeTableStyleLineSpacing(spacing)?.MultipleValue is not double multiple || multiple <= 0D) return null;
            return Math.Max(0.01D, ResolveNativeWordSingleLineHeight(fontFamilies) * multiple);
        }

        private static double? GetNativeTableStyleParagraphLineSpacingPoints(OfficeXmlNode spacing) {
            return ReadNativeTableStyleLineSpacing(spacing)?.Height?.ToPoints();
        }

        private static WordLineSpacing? ReadNativeTableStyleLineSpacing(OfficeXmlNode spacing) =>
            spacing.IsNone ? null : WordLineSpacing.Read(spacing.WAttr("line"), spacing.WAttr("lineRule"));

        private static IReadOnlyList<OfficeXmlNode> GetNativeTableStyleChain(WordDocument document, string? styleId) {
            NativeStyleLookupCache? cache = GetNativeStyleLookupCache(document);
            string? resolvedStyleId = cache?.ResolveTableStyleId(styleId);
            if (cache == null || string.IsNullOrWhiteSpace(resolvedStyleId)) {
                return Array.Empty<OfficeXmlNode>();
            }

            if (cache.TableChains.TryGetValue(resolvedStyleId!, out IReadOnlyList<OfficeXmlNode>? cachedChain)) {
                return cachedChain.ToArray();
            }

            var chain = new List<OfficeXmlNode>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? currentStyleId = resolvedStyleId;
            while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!) && cache.TableStyles.TryGetValue(currentStyleId!, out OfficeXmlNode style)) {
                // Word ignores the built-in Normal Table definition's child
                // properties, including when it is a named style's base.
                if (IsNativeNormalTableStyleId(currentStyleId)) {
                    break;
                }
                cache.RecordStyleChainReference(chain.Count + 1);
                chain.Add(style);
                currentStyleId = style.ChildVal("basedOn");
            }

            chain.Reverse();
            IReadOnlyList<OfficeXmlNode> result = chain.ToArray();
            cache.TableChains[resolvedStyleId!] = result;
            return result;
        }

        private static string? GetNativeTableStyleId(WordTable table) =>
            table.PropertiesNode().ChildVal("tblStyle");

        private static bool IsNativeFallbackTableStyleId(string? styleId) =>
            string.IsNullOrWhiteSpace(styleId) ||
            IsNativeNormalTableStyleId(styleId);

        private static bool IsNativeNormalTableStyleId(string? styleId) =>
            string.Equals(styleId, "TableNormal", StringComparison.Ordinal) ||
            string.Equals(styleId, "NormalTable", StringComparison.Ordinal);
    }
}
