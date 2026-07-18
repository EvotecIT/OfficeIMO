using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class TextContentParser {
    private readonly struct TextGraphicsState {
        public Matrix2D Ctm { get; }
        public string Font { get; }
        public double Size { get; }
        public double Leading { get; }
        public double CharSpacing { get; }
        public double WordSpacing { get; }
        public double HScale { get; }
        public double TextRise { get; }
        public OfficeColor FillColor { get; }
        public PdfPageColorSpace FillColorSpace { get; }
        public OfficeColor StrokeColor { get; }
        public PdfPageColorSpace StrokeColorSpace { get; }
        public double? FillOpacity { get; }
        public double? StrokeOpacity { get; }
        public int TextRenderingMode { get; }
        public PdfPageClipPath? ClipPath { get; }

        public TextGraphicsState(Matrix2D ctm, string font, double size, double leading, double charSpacing, double wordSpacing, double hScale, double textRise, OfficeColor fillColor, PdfPageColorSpace fillColorSpace, OfficeColor strokeColor, PdfPageColorSpace strokeColorSpace, double? fillOpacity, double? strokeOpacity, int textRenderingMode, PdfPageClipPath? clipPath) {
            Ctm = ctm;
            Font = font;
            Size = size;
            Leading = leading;
            CharSpacing = charSpacing;
            WordSpacing = wordSpacing;
            HScale = hScale;
            TextRise = textRise;
            FillColor = fillColor;
            FillColorSpace = fillColorSpace;
            StrokeColor = strokeColor;
            StrokeColorSpace = strokeColorSpace;
            FillOpacity = fillOpacity;
            StrokeOpacity = strokeOpacity;
            TextRenderingMode = textRenderingMode;
            ClipPath = clipPath;
        }
    }

    private sealed class InlineDictionary {
        public Dictionary<string, object> Items { get; } = new(StringComparer.Ordinal);
        public PdfInlineOptionalContentReferences? OptionalContentReferences { get; set; }
    }

    private readonly struct ActualTextValue {
        public string Text { get; }

        public ActualTextValue(string text) {
            Text = text;
        }
    }

    private sealed class MarkedContentState {
        public string ActualText { get; }
        public bool HasActualText { get; }
        public bool IsArtifact { get; }
        public bool IsHidden { get; }
        public bool ActualTextEmitted { get; set; }

        public MarkedContentState(ActualTextValue? actualText, bool isArtifact, bool isHidden) {
            ActualText = actualText?.Text ?? string.Empty;
            HasActualText = actualText.HasValue;
            IsArtifact = isArtifact;
            IsHidden = isHidden;
        }
    }

    internal readonly struct FormInvocation {
        public string Name { get; }
        public Matrix2D Transform { get; }
        public double PaintOrder { get; }
        public OfficeColor FillColor { get; }
        public PdfPageColorSpace FillColorSpace { get; }
        public OfficeColor StrokeColor { get; }
        public PdfPageColorSpace StrokeColorSpace { get; }
        public double? FillOpacity { get; }
        public double? StrokeOpacity { get; }
        public int TextRenderingMode { get; }
        public PdfPageClipPath? ClipPath { get; }

        public FormInvocation(
            string name,
            Matrix2D transform,
            double paintOrder = 0D,
            OfficeColor? fillColor = null,
            PdfPageColorSpace fillColorSpace = default,
            OfficeColor? strokeColor = null,
            PdfPageColorSpace strokeColorSpace = default,
            double? fillOpacity = null,
            double? strokeOpacity = null,
            int textRenderingMode = 0,
            PdfPageClipPath? clipPath = null) {
            Name = name;
            Transform = transform;
            PaintOrder = paintOrder;
            FillColor = fillColor ?? OfficeColor.Black;
            FillColorSpace = fillColorSpace;
            StrokeColor = strokeColor ?? OfficeColor.Black;
            StrokeColorSpace = strokeColorSpace;
            FillOpacity = fillOpacity;
            StrokeOpacity = strokeOpacity;
            TextRenderingMode = textRenderingMode;
            ClipPath = clipPath;
        }
    }

    public static List<PdfTextSpan> Parse(
        string content,
        System.Func<string, byte[], string> decodeWithFont,
        System.Func<string, byte[], double> sumWidth1000ForFont,
        bool adjustKerningFromTJ = true,
        System.Func<string, string?>? actualTextForProperty = null,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates = null,
        IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces = null,
        System.Func<string, string?>? baseFontForResource = null,
        System.Func<string, string?>? drawingFontFamilyForResource = null,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null,
        double pageHeight = 0D,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        double? initialFillOpacity = null,
        double? initialStrokeOpacity = null,
        int initialTextRenderingMode = 0,
        PdfPageClipPath? initialClipPath = null,
        bool useLogicalTextFilters = true,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations) {
        var spans = new List<PdfTextSpan>();
        // Text state
        bool inText = false;
        string font = "F1"; double size = 12; double leading = size * 1.2; double charSpacing = 0, wordSpacing = 0; double hScale = 1.0; double textRise = 0;
        OfficeColor fillColor = initialFillColor ?? OfficeColor.Black;
        PdfPageColorSpace fillColorSpace = initialFillColorSpace;
        OfficeColor strokeColor = initialStrokeColor ?? OfficeColor.Black;
        PdfPageColorSpace strokeColorSpace = initialStrokeColorSpace;
        double? fillOpacity = initialFillOpacity;
        double? strokeOpacity = initialStrokeOpacity;
        int textRenderingMode = ReadTextRenderingMode(initialTextRenderingMode);
        PdfPageClipPath? clipPath = initialClipPath;
        var clipPathBuilder = new PdfPageClipPathBuilder(pageHeight);
        Matrix2D textMatrix = Matrix2D.Identity;
        Matrix2D lineMatrix = Matrix2D.Identity;
        // Graphics state (CTM) and stack
        Matrix2D ctm = Matrix2D.Identity; var gstack = new System.Collections.Generic.Stack<TextGraphicsState>();
        // Operand buffer (tokens collected since last operator)
        var args = new List<object>(8);
        int i = 0; int n = content.Length;
        int operationCount = 0;
        // Kerning state between text runs in TJ arrays (points) and rolling output buffer for gap checks
        double pendingGapPt = 0;
        int pendingLineBreaks = 0;
        bool emittedTextInTextObject = false;
        var sbOutGlobal = new StringBuilder();
        var markedContentStack = new Stack<MarkedContentState>();
        while (i < n) {
            SkipWs(); if (i >= n) break;
            char c = content[i];
            if (c == '%') { // comment till end of line
                while (i < n && content[i] != '\n' && content[i] != '\r') i++;
                continue;
            }
            if (c == '/') { args.Add(ReadName()); continue; }
            if (c == '(') { args.Add(ReadLiteralStringBytes()); continue; }
            if (c == '<') {
                if (i + 1 < n && content[i + 1] == '<') { args.Add(ReadInlineDictionary()); continue; }
                args.Add(ReadHexStringBytes()); continue;
            }
            if (c == '[') { args.Add(ReadArray()); continue; }
            if (c == ']' || c == '>') { i++; continue; }
            if (IsNumberStart(c)) { args.Add(ReadNumber()); continue; }
            // operator (BT, ET, Tf, Tm, Td, TD, T*, TL, Tc, Tw, Tz, Ts, cm, q, Q, Tj, TJ, ', ")
            double paintOrder = GetPaintOrder(i);
            string op = ReadOperator();
            if (op.Length == 0) { i++; continue; }
            if (++operationCount > maxOperations) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperations, maxOperations, operationCount);
            }

            switch (op) {
                case "BT": inText = true; textMatrix = Matrix2D.Identity; lineMatrix = Matrix2D.Identity; pendingGapPt = 0; pendingLineBreaks = 0; emittedTextInTextObject = false; args.Clear(); break;
                case "ET": inText = false; pendingGapPt = 0; pendingLineBreaks = 0; emittedTextInTextObject = false; args.Clear(); break;
                case "Tf": if (args.Count >= 2) { size = ToDouble(args[args.Count - 1]); font = ToName(args[args.Count - 2]); args.Clear(); } break;
                case "Tm": if (args.Count >= 6) { SetTextMatrix(args); args.Clear(); } break;
                case "Td": if (args.Count >= 2) { MoveTextLine(ToDouble(args[args.Count - 2]), ToDouble(args[args.Count - 1])); args.Clear(); } break;
                case "TD": if (args.Count >= 2) { double tx = ToDouble(args[args.Count - 2]); double ty = ToDouble(args[args.Count - 1]); leading = -ty; MoveTextLine(tx, ty); args.Clear(); } break;
                case "TL": if (args.Count >= 1) { leading = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "T*": MoveToNextTextLine(); args.Clear(); break;
                case "Tc": if (args.Count >= 1) { charSpacing = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "Tw": if (args.Count >= 1) { wordSpacing = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "Tz": if (args.Count >= 1) { hScale = ToDouble(args[args.Count - 1]) / 100.0; args.Clear(); } break;
                case "Ts": if (args.Count >= 1) { textRise = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "Tr": if (args.Count >= 1) { textRenderingMode = ReadTextRenderingMode(ToDouble(args[args.Count - 1])); args.Clear(); } break;
                case "q":
                    gstack.Push(new TextGraphicsState(ctm, font, size, leading, charSpacing, wordSpacing, hScale, textRise, fillColor, fillColorSpace, strokeColor, strokeColorSpace, fillOpacity, strokeOpacity, textRenderingMode, clipPath));
                    args.Clear();
                    break;
                case "Q":
                    if (gstack.Count > 0) {
                        var state = gstack.Pop();
                        ctm = state.Ctm;
                        font = state.Font;
                        size = state.Size;
                        leading = state.Leading;
                        charSpacing = state.CharSpacing;
                        wordSpacing = state.WordSpacing;
                        hScale = state.HScale;
                        textRise = state.TextRise;
                        fillColor = state.FillColor;
                        fillColorSpace = state.FillColorSpace;
                        strokeColor = state.StrokeColor;
                        strokeColorSpace = state.StrokeColorSpace;
                        fillOpacity = state.FillOpacity;
                        strokeOpacity = state.StrokeOpacity;
                        textRenderingMode = state.TextRenderingMode;
                        clipPath = state.ClipPath;
                    } else {
                        ctm = Matrix2D.Identity;
                        fillColor = OfficeColor.Black;
                        fillColorSpace = PdfPageColorSpaceKind.DeviceGray;
                        strokeColor = OfficeColor.Black;
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceGray;
                        fillOpacity = null;
                        strokeOpacity = null;
                        textRenderingMode = 0;
                        clipPath = null;
                    }
                    args.Clear();
                    break;
                case "cm": if (args.Count >= 6) { var m2 = new Matrix2D(ToDouble(args[args.Count - 6]), ToDouble(args[args.Count - 5]), ToDouble(args[args.Count - 4]), ToDouble(args[args.Count - 3]), ToDouble(args[args.Count - 2]), ToDouble(args[args.Count - 1])); ctm = Matrix2D.Multiply(ctm, m2); args.Clear(); } break;
                case "re":
                    if (args.Count >= 4) {
                        clipPathBuilder.AddRectanglePath(
                            ctm,
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "m":
                    if (args.Count >= 2) {
                        clipPathBuilder.MoveTo(ctm, ToDouble(args[args.Count - 2]), ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "l":
                    if (args.Count >= 2) {
                        clipPathBuilder.LineTo(ctm, ToDouble(args[args.Count - 2]), ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "c":
                    if (args.Count >= 6) {
                        clipPathBuilder.CubicTo(
                            ctm,
                            ToDouble(args[args.Count - 6]),
                            ToDouble(args[args.Count - 5]),
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "v":
                    if (args.Count >= 4) {
                        clipPathBuilder.CubicToWithCurrentFirstControl(
                            ctm,
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "y":
                    if (args.Count >= 4) {
                        clipPathBuilder.CubicToWithEndSecondControl(
                            ctm,
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "h":
                    clipPathBuilder.ClosePath();
                    args.Clear();
                    break;
                case "W":
                case "W*":
                    if (clipPathBuilder.TryCreateClipPath(op == "W*" ? OfficeFillRule.EvenOdd : OfficeFillRule.NonZero, out PdfPageClipPath parsedClipPath)) {
                        clipPath = PdfPageClipPath.ResolveActiveClip(clipPath, parsedClipPath);
                    }

                    args.Clear();
                    break;
                case "n":
                    clipPathBuilder.Clear();
                    args.Clear();
                    break;
                case "f":
                case "F":
                case "f*":
                case "S":
                case "B":
                case "B*":
                    clipPathBuilder.Clear();
                    args.Clear();
                    break;
                case "s":
                case "b":
                case "b*":
                    clipPathBuilder.ClosePath();
                    clipPathBuilder.Clear();
                    args.Clear();
                    break;
                case "gs":
                    if (args.Count >= 1) {
                        ApplyGraphicsStateResource(ToName(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "cs":
                    if (args.Count >= 1 && TryReadColorSpace(ToName(args[args.Count - 1]), out PdfPageColorSpace parsedColorSpace)) {
                        fillColorSpace = parsedColorSpace;
                    }

                    args.Clear();
                    break;
                case "CS":
                    if (args.Count >= 1 && TryReadColorSpace(ToName(args[args.Count - 1]), out PdfPageColorSpace parsedStrokeColorSpace)) {
                        strokeColorSpace = parsedStrokeColorSpace;
                    }

                    args.Clear();
                    break;
                case "rg":
                    if (args.Count >= 3) {
                        fillColor = ReadRgb(args.Count - 3);
                        fillColorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    }

                    args.Clear();
                    break;
                case "RG":
                    if (args.Count >= 3) {
                        strokeColor = ReadRgb(args.Count - 3);
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    }

                    args.Clear();
                    break;
                case "g":
                    if (args.Count >= 1) {
                        fillColor = ReadGray(args.Count - 1);
                        fillColorSpace = PdfPageColorSpaceKind.DeviceGray;
                    }

                    args.Clear();
                    break;
                case "G":
                    if (args.Count >= 1) {
                        strokeColor = ReadGray(args.Count - 1);
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceGray;
                    }

                    args.Clear();
                    break;
                case "k":
                    if (args.Count >= 4) {
                        fillColor = ReadCmyk(args.Count - 4);
                        fillColorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    }

                    args.Clear();
                    break;
                case "K":
                    if (args.Count >= 4) {
                        strokeColor = ReadCmyk(args.Count - 4);
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    }

                    args.Clear();
                    break;
                case "sc":
                case "scn":
                    if (TryReadColor(fillColorSpace, out OfficeColor parsedFillColor)) {
                        fillColor = parsedFillColor;
                    }

                    args.Clear();
                    break;
                case "SC":
                case "SCN":
                    if (TryReadColor(strokeColorSpace, out OfficeColor parsedStrokeColor)) {
                        strokeColor = parsedStrokeColor;
                    }

                    args.Clear();
                    break;
                case "BI":
                    SkipInlineImageBody();
                    args.Clear();
                    break;
                case "'": // move to next line and show text
                    if (args.Count >= 1) { MoveToNextTextLine(); ShowTextRun(ToBytes(args[args.Count - 1]), paintOrder); pendingGapPt = 0; }
                    args.Clear();
                    break;
                case "\"": // set spacing and show text
                    if (args.Count >= 3) { wordSpacing = ToDouble(args[args.Count - 3]); charSpacing = ToDouble(args[args.Count - 2]); MoveToNextTextLine(); ShowTextRun(ToBytes(args[args.Count - 1]), paintOrder); pendingGapPt = 0; }
                    args.Clear();
                    break;
                case "Tj": if (args.Count >= 1) { ShowTextRun(ToBytes(args[args.Count - 1]), paintOrder); pendingGapPt = 0; args.Clear(); } break;
                case "TJ": if (args.Count >= 1) { ShowTextArray(args[args.Count - 1], paintOrder); args.Clear(); } break;
                case "BDC":
                    markedContentStack.Push(new MarkedContentState(
                        GetActualText(args.Count > 0 ? args[args.Count - 1] : null),
                        IsArtifactTag(args.Count > 1 ? args[args.Count - 2] : null),
                        IsHiddenOptionalContent(args.Count > 1 ? args[args.Count - 2] : null, args.Count > 0 ? args[args.Count - 1] : null)));
                    args.Clear();
                    break;
                case "BMC":
                    markedContentStack.Push(new MarkedContentState(null, IsArtifactTag(args.Count > 0 ? args[args.Count - 1] : null), false));
                    args.Clear();
                    break;
                case "EMC":
                    if (markedContentStack.Count > 0) {
                        markedContentStack.Pop();
                    }

                    args.Clear();
                    break;
                default: args.Clear(); break;
            }
        }
        return spans;

        // Helpers
        void SetTextMatrix(List<object> operands) {
            lineMatrix = new Matrix2D(
                ToDouble(operands[operands.Count - 6]),
                ToDouble(operands[operands.Count - 5]),
                ToDouble(operands[operands.Count - 4]),
                ToDouble(operands[operands.Count - 3]),
                ToDouble(operands[operands.Count - 2]),
                ToDouble(operands[operands.Count - 1]));
            textMatrix = lineMatrix;
            pendingGapPt = 0;
            pendingLineBreaks = 0;
        }

        void MoveTextLine(double tx, double ty) {
            lineMatrix = Matrix2D.Multiply(lineMatrix, Matrix2D.Translation(tx, ty));
            textMatrix = lineMatrix;
            pendingGapPt = 0;
            if (emittedTextInTextObject && Math.Abs(ty) > 0.000001D) {
                pendingLineBreaks++;
            }
        }

        void MoveToNextTextLine() {
            lineMatrix = Matrix2D.Multiply(lineMatrix, Matrix2D.Translation(0, -leading));
            textMatrix = lineMatrix;
            pendingGapPt = 0;
            if (emittedTextInTextObject) {
                pendingLineBreaks++;
            }
        }

        double GetPaintOrder(int operatorIndex) => paintOrderBase + ((operatorIndex + paintOrderOffset) * paintOrderScale);

        void MaybeInsertSpaceBeforeRun() {
            // Insert a space depending on kerning gap accumulated from TJ array numbers
            if (pendingGapPt <= 0) return;
            double prevAvg = Math.Max(1.0, size * 0.5); // fallback if we can't infer
            double emThreshold = size * 0.24; // about quarter em
            double glyphThreshold = prevAvg * 0.6;
            double threshold = Math.Max(emThreshold, glyphThreshold);
            // Tighten when previous char is wordish
            bool prevWord = sbOutGlobal.Length > 0 && (char.IsLetterOrDigit(sbOutGlobal[sbOutGlobal.Length - 1]) || sbOutGlobal[sbOutGlobal.Length - 1] == '\'' || sbOutGlobal[sbOutGlobal.Length - 1] == '-' || sbOutGlobal[sbOutGlobal.Length - 1] == '/');
            if (prevWord) threshold = Math.Min(threshold, 2.0);
            if (pendingGapPt >= threshold) sbOutGlobal.Append(' ');
            pendingGapPt = 0;
        }
        void ShowTextRun(byte[] bytes, double paintOrder) {
            if (!inText || bytes == null || bytes.Length == 0) return;
            MaybeInsertSpaceBeforeRun();
            // Detect 2-byte CIDs (Identity-H) vs single-byte
            bool twoByte = false;
            if (bytes.Length >= 2) {
                string one = decodeWithFont(font, new byte[] { bytes[0] });
                string two = decodeWithFont(font, new byte[] { bytes[0], bytes[1] });
                double firstByteWidth = sumWidth1000ForFont(font, new byte[] { bytes[0] });
                double secondByteWidth = sumWidth1000ForFont(font, new byte[] { bytes[1] });
                double pairWidth = sumWidth1000ForFont(font, new byte[] { bytes[0], bytes[1] });
                twoByte = (IsNullOrEmptyDecodedGlyph(one) && !IsNullOrEmptyDecodedGlyph(two)) ||
                    (firstByteWidth <= 0 && secondByteWidth <= 0 && pairWidth > 0);
            }
            var sbOut = new StringBuilder(bytes.Length);
            double advTotal = 0;
            char prevChar = '\0';
            string wholeDecoded = NormalizeDecodedGlyphText(decodeWithFont(font, bytes) ?? string.Empty);
            for (int idx = 0; idx < bytes.Length;) {
                int step = twoByte ? (idx + 1 < bytes.Length ? 2 : 1) : 1;
                byte[] g = step == 1 ? new byte[] { bytes[idx] } : new byte[] { bytes[idx], bytes[idx + 1] };
                string t = NormalizeDecodedGlyphText(decodeWithFont(font, g) ?? string.Empty);
                char ch = (t.Length > 0) ? t[0] : '\0';
                double w1000 = sumWidth1000ForFont(font, g);
                double advGlyph = ((w1000 / 1000.0) * size + charSpacing + (ch == ' ' ? wordSpacing : 0)) * hScale;
                // Drop thin spaces between letters/digits (visual join) but still advance
                double thinSpacePt = Math.Max(1.0, size * 0.12);
                bool dropSpace = false;
                if (ch == ' ') {
                    // Keep explicit space glyphs; rely on higher-level normalization to fix accidental splits
                } else if (advGlyph <= thinSpacePt && prevChar != '\0') {
                    // Drop non-space thin separators
                    dropSpace = true;
                }
                if (dropSpace) {
                    // do not append, but keep advance
                } else if (ch != '\0') {
                    sbOut.Append(t);
                    prevChar = t[t.Length - 1];
                }
                advTotal += advGlyph;
                idx += step;
            }
            if (ShouldUseWholeDecodedText(sbOut.ToString(), wholeDecoded)) {
                sbOut.Clear();
                sbOut.Append(wholeDecoded);
            }
            var actualTextState = useLogicalTextFilters ? GetActiveActualTextState() : null;
            bool isArtifact = useLogicalTextFilters && HasActiveArtifact();
            bool isHidden = HasActiveHiddenContent();
            bool isVisibleText = IsTextRenderingModeVisible(textRenderingMode);
            if (sbOut.Length == 0 && actualTextState is null && !isArtifact && !isHidden) return;
            string textOut = sbOut.ToString();
            var textOrigin = textMatrix.Transform(0, textRise);
            var (dx, dy) = ctm.Transform(textOrigin.X, textOrigin.Y);
            var textEnd = textMatrix.Transform(advTotal, textRise);
            var (endX, endY) = ctm.Transform(textEnd.X, textEnd.Y);
            double transformedAdvance = Math.Sqrt(((endX - dx) * (endX - dx)) + ((endY - dy) * (endY - dy)));
            double rotationDegrees = CalculateRotationDegrees(endX - dx, endY - dy);
            OfficeColor paintColor = ResolveTextPaintColor(textRenderingMode, fillColor, strokeColor);
            OfficeColor visibleColor = ApplyTextOpacity(paintColor, textRenderingMode);
            PdfPageClipPath? spanClipPath = clipPath;
            if (isHidden) {
                // Hidden optional-content still advances text state but should not emit visible/logical spans.
            } else if (isArtifact) {
                // Artifact content is visual decoration, not logical page text.
            } else if (actualTextState is not null && !actualTextState.ActualTextEmitted) {
                textOut = actualTextState.ActualText;
                actualTextState.ActualTextEmitted = true;
                if (textOut.Length > 0) {
                    AddTextSpan(textOut);
                }
            } else if (actualTextState is null && textOut.Length > 0) {
                AddTextSpan(textOut);
            }

            if (!isHidden) {
                ApplyTextClippingPath(advTotal);
            }

            textMatrix = Matrix2D.Multiply(textMatrix, Matrix2D.Translation(advTotal, 0));

            void AddTextSpan(string rawText) {
                bool logicalLeadingSpace = char.IsWhiteSpace(rawText[0]);
                bool logicalTrailingSpace = char.IsWhiteSpace(rawText[rawText.Length - 1]);
                string normalizedText = NormalizeShatteredSpan(rawText);
                if (normalizedText.Length == 0) {
                    return;
                }

                spans.Add(new PdfTextSpan(
                    normalizedText,
                    font,
                    size,
                    dx,
                    dy,
                    transformedAdvance,
                    visibleColor,
                    isVisibleText,
                    rotationDegrees,
                    baseFontForResource?.Invoke(font),
                    spanClipPath,
                    paintOrder,
                    drawingFontFamilyForResource?.Invoke(font),
                    pendingLineBreaks,
                    logicalLeadingSpace,
                    logicalTrailingSpace));
                sbOutGlobal.Append(normalizedText);
                emittedTextInTextObject = true;
                pendingLineBreaks = 0;
            }
        }

        void ApplyTextClippingPath(double advance) {
            if (!AddsTextToClippingPath(textRenderingMode) || size <= 0D || Math.Abs(advance) <= 0.000001D) {
                return;
            }

            double left = advance < 0D ? advance : 0D;
            double width = Math.Abs(advance);
            double descent = Math.Max(0.001D, size * 0.25D);
            double height = Math.Max(0.001D, size + descent);
            Matrix2D textToPage = Matrix2D.Multiply(ctm, textMatrix);
            var textClipBuilder = new PdfPageClipPathBuilder(pageHeight);
            textClipBuilder.AddRectanglePath(textToPage, left, textRise - descent, width, height);
            if (textClipBuilder.TryCreateClipPath(OfficeFillRule.NonZero, out PdfPageClipPath textClipPath)) {
                clipPath = PdfPageClipPath.ResolveActiveClip(clipPath, textClipPath);
            }
        }

        MarkedContentState? GetActiveActualTextState() {
            foreach (var state in markedContentStack) {
                if (state.HasActualText) {
                    return state;
                }
            }

            return null;
        }

        bool HasActiveArtifact() {
            foreach (var state in markedContentStack) {
                if (state.IsArtifact) {
                    return true;
                }
            }

            return false;
        }

        bool HasActiveHiddenContent() {
            foreach (var state in markedContentStack) {
                if (state.IsHidden) {
                    return true;
                }
            }

            return false;
        }

        ActualTextValue? GetActualText(object? propertyObject) {
            if (propertyObject is string propertyName) {
                string? text = actualTextForProperty?.Invoke(propertyName);
                return text is null ? (ActualTextValue?)null : new ActualTextValue(text);
            }

            if (propertyObject is InlineDictionary dictionary &&
                dictionary.Items.TryGetValue("ActualText", out var value) &&
                value is byte[] bytes) {
                return new ActualTextValue(PdfTextString.Decode(bytes));
            }

            return null;
        }

        static bool IsArtifactTag(object? tag) =>
            tag is string name && string.Equals(name, "Artifact", StringComparison.Ordinal);

        bool IsHiddenOptionalContent(object? tag, object? property) =>
            tag is string tagName &&
            string.Equals(tagName, "OC", StringComparison.Ordinal) &&
            ((property is string propertyName &&
                optionalContentVisibility?.IsHidden(propertyName) == true) ||
             (property is InlineDictionary dictionary &&
                dictionary.OptionalContentReferences != null &&
                optionalContentVisibility?.IsHidden(dictionary.OptionalContentReferences) == true));

        void ShowTextArray(object arrObj, double paintOrder) {
            if (!inText || arrObj == null) return;
            var list = arrObj as List<object>;
            if (list == null) return;
            for (int j = 0; j < list.Count; j++) {
                var it = list[j];
                if (it is byte[] b) { ShowTextRun(b, paintOrder); }
                else if (adjustKerningFromTJ && it is double num) {
                    double delta = -num / 1000.0 * size * hScale;
                    textMatrix = Matrix2D.Multiply(textMatrix, Matrix2D.Translation(delta, 0));
                    // Only positive visual gap should suggest a space
                    if (delta > 0) pendingGapPt += delta; else pendingGapPt = 0;
                }
            }
        }

        void SkipWs() { while (i < n && char.IsWhiteSpace(content[i])) i++; }
        static bool IsDigit(char ch) => ch >= '0' && ch <= '9';
        bool IsNumberStart(char ch) => ch == '-' || ch == '+' || ch == '.' || IsDigit(ch);

        double ReadNumber() {
            int start = i; i++;
            while (i < n) { char ch = content[i]; if (!(IsDigit(ch) || ch == '.' || ch == 'E' || ch == 'e' || ch == '-' || ch == '+')) break; i++; }
            var s = content.Substring(start, i - start);
            if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) v = 0;
            return v;
        }

        string ReadName() {
            i++; int start = i;
            while (i < n) { char ch = content[i]; if (char.IsWhiteSpace(ch) || ch == '%' || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>') break; i++; }
            return PdfSyntax.DecodeName(content.Substring(start, i - start));
        }

        byte[] ReadLiteralStringBytes() {
            int start = ++i; int depth = 1; bool esc = false; var sb = new StringBuilder();
            while (i < n && depth > 0) {
                char ch = content[i++];
                if (esc) { sb.Append('\\'); sb.Append(ch); esc = false; }
                else if (ch == '\\') esc = true;
                else if (ch == '(') { depth++; sb.Append(ch); }
                else if (ch == ')') { depth--; if (depth > 0) sb.Append(ch); }
                else sb.Append(ch);
            }
            return PdfStringParser.ParseLiteralToBytes(sb.ToString());
        }

        byte[] ReadHexStringBytes() {
            i++; int start = i; while (i < n && content[i] != '>') i++; int end = i; if (i < n && content[i] == '>') i++;
            string hex = content.Substring(start, end - start);
            var sb = new StringBuilder(hex.Length);
            for (int k = 0; k < hex.Length; k++) { char ch = hex[k]; if (!char.IsWhiteSpace(ch)) sb.Append(ch); }
            hex = sb.ToString();
            if (hex.Length % 2 == 1) hex += "0";
            var bytes = new byte[hex.Length / 2];
            for (int k = 0; k < bytes.Length; k++) {
                var hi = HexNibble(hex[k * 2]);
                var lo = HexNibble(hex[k * 2 + 1]);
                bytes[k] = (byte)((hi << 4) | lo);
            }
            return bytes;

            static int HexNibble(char c) {
                if (c >= '0' && c <= '9') return c - '0';
                if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
                if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
                throw new FormatException($"Invalid hex character '{c}'.");
            }
        }

        InlineDictionary ReadInlineDictionary() {
            int start = i;
            i += 2;
            var dictionary = new InlineDictionary();
            while (i < n) {
                SkipWs();
                if (i + 1 < n && content[i] == '>' && content[i + 1] == '>') {
                    i += 2;
                    break;
                }

                if (i >= n) {
                    break;
                }

                if (content[i] != '/') {
                    SkipInlineDictionaryValue();
                    continue;
                }

                string key = ReadName();
                SkipWs();
                if (i >= n) {
                    break;
                }

                if (content[i] == '(') {
                    dictionary.Items[key] = ReadLiteralStringBytes();
                } else if (content[i] == '<' && i + 1 < n && content[i + 1] != '<') {
                    dictionary.Items[key] = ReadHexStringBytes();
                } else {
                    SkipInlineDictionaryValue();
                }
            }

            dictionary.OptionalContentReferences = PdfInlineOptionalContentReferenceParser.Parse(content, start, Math.Max(0, i - start));
            return dictionary;
        }

        void SkipInlineDictionaryValue() {
            if (i >= n) {
                return;
            }

            char ch = content[i];
            if (ch == '(') {
                ReadLiteralStringBytes();
                return;
            }

            if (ch == '<') {
                if (i + 1 < n && content[i + 1] == '<') {
                    ReadInlineDictionary();
                } else {
                    ReadHexStringBytes();
                }

                return;
            }

            if (ch == '[') {
                ReadArray();
                return;
            }

            if (ch == '/') {
                ReadName();
                return;
            }

            if (IsNumberStart(ch)) {
                ReadNumber();
                return;
            }

            ReadOperator();
        }

        List<object> ReadArray() {
            var list = new List<object>();
            i++; // skip [
            while (i < n) {
                SkipWs(); if (i >= n) break; char ch = content[i]; if (ch == ']') { i++; break; }
                if (ch == '(') { list.Add(ReadLiteralStringBytes()); continue; }
                if (ch == '<') { if (i + 1 < n && content[i + 1] == '<') { i += 2; continue; } list.Add(ReadHexStringBytes()); continue; }
                if (IsNumberStart(ch)) { list.Add(ReadNumber()); continue; }
                if (ch == '/') { list.Add(ReadName()); continue; }
                if (ch == '[') { i++; continue; } // ignore nested
                // unknown token inside array -> treat as operator and skip
                ReadOperator();
            }
            return list;
        }

        string ReadOperator() {
            int start = i; char ch = content[i++];
            if (ch == '\'' || ch == '"') return ch.ToString();
            while (i < n) {
                char c2 = content[i];
                if (char.IsWhiteSpace(c2) || c2 == '%' || c2 == '(' || c2 == '[' || c2 == '/' || c2 == '<' || c2 == '>') break;
                i++;
            }
            return content.Substring(start, i - start);
        }

        void SkipInlineImageBody() {
            while (i < n) {
                SkipWs();
                if (IsOperatorAt("ID")) {
                    i += 2;
                    break;
                }

                if (i >= n) {
                    return;
                }

                if (content[i] == '(') {
                    ReadLiteralStringBytes();
                } else if (content[i] == '<') {
                    if (i + 1 < n && content[i + 1] == '<') {
                        ReadInlineDictionary();
                    } else {
                        ReadHexStringBytes();
                    }
                } else if (content[i] == '[') {
                    ReadArray();
                } else if (content[i] == '/') {
                    ReadName();
                } else if (IsNumberStart(content[i])) {
                    ReadNumber();
                } else {
                    ReadOperator();
                }
            }

            if (i < n && char.IsWhiteSpace(content[i])) {
                i++;
            }

            while (i + 1 < n) {
                if (PdfInlineImageDataScanner.IsTerminatorAt(content, i)) {
                    i += 2;
                    return;
                }

                i++;
            }

            i = n;
        }

        bool IsOperatorAt(string value) {
            if (i + value.Length > n) {
                return false;
            }

            for (int j = 0; j < value.Length; j++) {
                if (content[i + j] != value[j]) {
                    return false;
                }
            }

            return (i == 0 || IsInlineImageBoundary(content[i - 1])) &&
                   (i + value.Length >= n || IsInlineImageBoundary(content[i + value.Length]));
        }

        static bool IsInlineImageBoundary(char ch) =>
            char.IsWhiteSpace(ch) || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>' || ch == '%';

        static double ToDouble(object o) { return o is double d ? d : 0.0; }
        static string ToName(object o) { return o as string ?? string.Empty; }
        static byte[] ToBytes(object o) { return o as byte[] ?? Array.Empty<byte>(); }
        double NumberAt(int index) => args[index] is double value ? value : 0D;
        void ApplyGraphicsStateResource(string name) {
            if (graphicsStates != null && graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource resource)) {
                fillOpacity = resource.FillOpacity ?? fillOpacity;
                strokeOpacity = resource.StrokeOpacity ?? strokeOpacity;
            }
        }
        OfficeColor ApplyTextOpacity(OfficeColor color, int renderingMode) {
            double? opacity = UsesStrokeTextPaint(renderingMode) ? strokeOpacity : fillOpacity;
            if (!opacity.HasValue) {
                return color;
            }

            return OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * Clamp01(opacity.Value)));
        }
        OfficeColor ReadRgb(int startIndex) =>
            OfficeColor.FromRgb(ToByte(NumberAt(startIndex)), ToByte(NumberAt(startIndex + 1)), ToByte(NumberAt(startIndex + 2)));
        OfficeColor ReadGray(int index) {
            byte value = ToByte(NumberAt(index));
            return OfficeColor.FromRgb(value, value, value);
        }
        OfficeColor ReadCmyk(int startIndex) {
            return OfficeColorSpaceConverter.FromCmyk(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2), NumberAt(startIndex + 3));
        }
        bool TryReadColor(PdfPageColorSpace colorSpace, out OfficeColor color) {
            color = OfficeColor.Black;
            int componentCount = GetColorComponentCount(colorSpace);
            int endIndex = args.Count;
            while (endIndex > 0 && args[endIndex - 1] is not double) {
                endIndex--;
            }

            if (endIndex < componentCount) {
                return false;
            }

            int startIndex = endIndex - componentCount;
            switch (colorSpace.Kind) {
                case PdfPageColorSpaceKind.DeviceRgb:
                    color = ReadRgb(startIndex);
                    return true;
                case PdfPageColorSpaceKind.DeviceCmyk:
                    color = ReadCmyk(startIndex);
                    return true;
                case PdfPageColorSpaceKind.CalGray:
                    color = PdfPageColorConverter.FromCalGray(NumberAt(startIndex));
                    return true;
                case PdfPageColorSpaceKind.CalRgb:
                    color = PdfPageColorConverter.FromCalRgb(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2), colorSpace);
                    return true;
                case PdfPageColorSpaceKind.Lab:
                    color = PdfPageColorConverter.FromLab(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2));
                    return true;
                default:
                    color = ReadGray(startIndex);
                    return true;
            }
        }
        bool TryReadColorSpace(string name, out PdfPageColorSpace colorSpace) {
            switch (name) {
                case "DeviceRGB":
                case "RGB":
                    colorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    return true;
                case "DeviceCMYK":
                case "CMYK":
                    colorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    return true;
                case "DeviceGray":
                case "G":
                    colorSpace = PdfPageColorSpaceKind.DeviceGray;
                    return true;
                case "CalGray":
                    colorSpace = PdfPageColorSpaceKind.CalGray;
                    return true;
                case "CalRGB":
                    colorSpace = PdfPageColorSpaceKind.CalRgb;
                    return true;
                case "Lab":
                    colorSpace = PdfPageColorSpaceKind.Lab;
                    return true;
                default:
                    if (colorSpaces != null && colorSpaces.TryGetValue(name, out colorSpace)) {
                        return true;
                    }

                    colorSpace = PdfPageColorSpaceKind.DeviceGray;
                    return false;
            }
        }
        static int GetColorComponentCount(PdfPageColorSpace colorSpace) {
            switch (colorSpace.Kind) {
                case PdfPageColorSpaceKind.DeviceRgb:
                case PdfPageColorSpaceKind.CalRgb:
                case PdfPageColorSpaceKind.Lab:
                    return 3;
                case PdfPageColorSpaceKind.DeviceCmyk:
                    return 4;
                default:
                    return 1;
            }
        }
        static byte ToByte(double value) => (byte)Math.Round(Clamp01(value) * 255D);
        static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
        static int ReadTextRenderingMode(double value) {
            int mode = (int)Math.Round(value);
            return mode < 0 || mode > 7 ? 0 : mode;
        }
        static OfficeColor ResolveTextPaintColor(int renderingMode, OfficeColor fill, OfficeColor stroke) =>
            UsesStrokeTextPaint(renderingMode) ? stroke : fill;

        static bool UsesStrokeTextPaint(int renderingMode) =>
            renderingMode == 1 || renderingMode == 5;

        static bool IsTextRenderingModeVisible(int renderingMode) =>
            renderingMode != 3 && renderingMode != 7;

        static bool AddsTextToClippingPath(int renderingMode) =>
            renderingMode >= 4 && renderingMode <= 7;

        static double CalculateRotationDegrees(double x, double y) {
            if (Math.Abs(x) <= 0.000001D && Math.Abs(y) <= 0.000001D) {
                return 0D;
            }

            double angle = Math.Atan2(y, x) * 180D / Math.PI;
            return Math.Abs(angle) <= 0.000001D ? 0D : angle;
        }

        static string NormalizeDecodedGlyphText(string text) =>
            text.Length == 0
                ? text
                : text
                    .Replace("\uFB00", "ff")
                    .Replace("\uFB01", "fi")
                    .Replace("\uFB02", "fl")
                    .Replace("\uFB03", "ffi")
                    .Replace("\uFB04", "ffl");

        static bool ShouldUseWholeDecodedText(string chunkedText, string wholeDecoded) {
            if (string.IsNullOrEmpty(wholeDecoded)) {
                return false;
            }

            if (string.IsNullOrEmpty(chunkedText)) {
                return true;
            }

            return ContainsNonTextControl(chunkedText) && !ContainsNonTextControl(wholeDecoded);
        }

        static bool ContainsNonTextControl(string text) {
            for (int index = 0; index < text.Length; index++) {
                char ch = text[index];
                if (char.IsControl(ch) && ch != '\t' && ch != '\n' && ch != '\r') {
                    return true;
                }
            }

            return false;
        }

        // Helpers (left empty for future metrics)
        // NormalizeThinSpaces removed in favor of per-glyph join logic

        static string NormalizeShatteredSpan(string s) {
            if (string.IsNullOrEmpty(s)) return s;
            string normalized = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ");
            string trimmed = normalized.Trim();
            return trimmed.Length == 0 && normalized.Length > 0 ? " " : trimmed;
        }
    }

    private static bool IsNullOrEmptyDecodedGlyph(string? value) =>
        string.IsNullOrEmpty(value) || value.All(static character => character == '\0');

    public static List<FormInvocation> ExtractFormInvocations(
        string content,
        PdfPageOptionalContentVisibility? optionalContentVisibility = null,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates = null,
        IReadOnlyDictionary<string, PdfPageColorSpace>? colorSpaces = null,
        double pageHeight = 0D,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        double? initialFillOpacity = null,
        double? initialStrokeOpacity = null,
        int initialTextRenderingMode = 0,
        PdfPageClipPath? initialClipPath = null,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations) {
        var invocations = new List<FormInvocation>();
        Matrix2D ctm = Matrix2D.Identity;
        OfficeColor fillColor = initialFillColor ?? OfficeColor.Black;
        PdfPageColorSpace fillColorSpace = initialFillColorSpace;
        OfficeColor strokeColor = initialStrokeColor ?? OfficeColor.Black;
        PdfPageColorSpace strokeColorSpace = initialStrokeColorSpace;
        double? fillOpacity = initialFillOpacity;
        double? strokeOpacity = initialStrokeOpacity;
        int textRenderingMode = ReadTextRenderingMode(initialTextRenderingMode);
        PdfPageClipPath? clipPath = initialClipPath;
        var clipPathBuilder = new PdfPageClipPathBuilder(pageHeight);
        var gstack = new Stack<TextGraphicsState>();
        var hiddenContentStack = new Stack<bool>();
        var args = new List<object>(8);
        int i = 0;
        int n = content.Length;
        int operationCount = 0;

        while (i < n) {
            SkipWs();
            if (i >= n) break;

            char c = content[i];
            if (c == '%') {
                while (i < n && content[i] != '\n' && content[i] != '\r') i++;
                continue;
            }

            if (c == '/') { args.Add(ReadName()); continue; }
            if (c == '(') { ReadLiteralStringBytes(); continue; }
            if (c == '<') {
                if (i + 1 < n && content[i + 1] == '<') {
                    args.Add(PdfInlineOptionalContentReferenceParser.Read(content, ref i));
                    continue;
                }

                ReadHexStringBytes();
                continue;
            }
            if (c == '[') { ReadArray(); continue; }
            if (c == ']' || c == '>') { i++; continue; }
            if (IsNumberStart(c)) { args.Add(ReadNumber()); continue; }

            double paintOrder = GetPaintOrder(i);
            string op = ReadOperator();
            if (op.Length == 0) { i++; continue; }
            if (++operationCount > maxOperations) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperations, maxOperations, operationCount);
            }

            switch (op) {
                case "q":
                    gstack.Push(new TextGraphicsState(ctm, string.Empty, 0D, 0D, 0D, 0D, 1D, 0D, fillColor, fillColorSpace, strokeColor, strokeColorSpace, fillOpacity, strokeOpacity, textRenderingMode, clipPath));
                    args.Clear();
                    break;
                case "Q":
                    if (gstack.Count > 0) {
                        TextGraphicsState state = gstack.Pop();
                        ctm = state.Ctm;
                        fillColor = state.FillColor;
                        fillColorSpace = state.FillColorSpace;
                        strokeColor = state.StrokeColor;
                        strokeColorSpace = state.StrokeColorSpace;
                        fillOpacity = state.FillOpacity;
                        strokeOpacity = state.StrokeOpacity;
                        textRenderingMode = state.TextRenderingMode;
                        clipPath = state.ClipPath;
                    } else {
                        ctm = Matrix2D.Identity;
                        fillColor = initialFillColor ?? OfficeColor.Black;
                        fillColorSpace = initialFillColorSpace;
                        strokeColor = initialStrokeColor ?? OfficeColor.Black;
                        strokeColorSpace = initialStrokeColorSpace;
                        fillOpacity = initialFillOpacity;
                        strokeOpacity = initialStrokeOpacity;
                        textRenderingMode = ReadTextRenderingMode(initialTextRenderingMode);
                        clipPath = initialClipPath;
                    }

                    args.Clear();
                    break;
                case "cm":
                    if (args.Count >= 6) {
                        var m2 = new Matrix2D(
                            ToDouble(args[args.Count - 6]),
                            ToDouble(args[args.Count - 5]),
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                        ctm = Matrix2D.Multiply(ctm, m2);
                    }
                    args.Clear();
                    break;
                case "re":
                    if (args.Count >= 4) {
                        clipPathBuilder.AddRectanglePath(
                            ctm,
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "m":
                    if (args.Count >= 2) {
                        clipPathBuilder.MoveTo(ctm, ToDouble(args[args.Count - 2]), ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "l":
                    if (args.Count >= 2) {
                        clipPathBuilder.LineTo(ctm, ToDouble(args[args.Count - 2]), ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "c":
                    if (args.Count >= 6) {
                        clipPathBuilder.CubicTo(
                            ctm,
                            ToDouble(args[args.Count - 6]),
                            ToDouble(args[args.Count - 5]),
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "v":
                    if (args.Count >= 4) {
                        clipPathBuilder.CubicToWithCurrentFirstControl(
                            ctm,
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "y":
                    if (args.Count >= 4) {
                        clipPathBuilder.CubicToWithEndSecondControl(
                            ctm,
                            ToDouble(args[args.Count - 4]),
                            ToDouble(args[args.Count - 3]),
                            ToDouble(args[args.Count - 2]),
                            ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "h":
                    clipPathBuilder.ClosePath();
                    args.Clear();
                    break;
                case "W":
                case "W*":
                    if (clipPathBuilder.TryCreateClipPath(op == "W*" ? OfficeFillRule.EvenOdd : OfficeFillRule.NonZero, out PdfPageClipPath parsedClipPath)) {
                        clipPath = PdfPageClipPath.ResolveActiveClip(clipPath, parsedClipPath);
                    }

                    args.Clear();
                    break;
                case "n":
                    clipPathBuilder.Clear();
                    args.Clear();
                    break;
                case "f":
                case "F":
                case "f*":
                case "S":
                case "B":
                case "B*":
                    clipPathBuilder.Clear();
                    args.Clear();
                    break;
                case "s":
                case "b":
                case "b*":
                    clipPathBuilder.ClosePath();
                    clipPathBuilder.Clear();
                    args.Clear();
                    break;
                case "gs":
                    if (args.Count >= 1) {
                        ApplyGraphicsStateResource(ToName(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "cs":
                    if (args.Count >= 1 && TryReadColorSpace(ToName(args[args.Count - 1]), out PdfPageColorSpace parsedColorSpace)) {
                        fillColorSpace = parsedColorSpace;
                    }

                    args.Clear();
                    break;
                case "CS":
                    if (args.Count >= 1 && TryReadColorSpace(ToName(args[args.Count - 1]), out PdfPageColorSpace parsedStrokeColorSpace)) {
                        strokeColorSpace = parsedStrokeColorSpace;
                    }

                    args.Clear();
                    break;
                case "rg":
                    if (args.Count >= 3) {
                        fillColor = ReadRgb(args.Count - 3);
                        fillColorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    }

                    args.Clear();
                    break;
                case "RG":
                    if (args.Count >= 3) {
                        strokeColor = ReadRgb(args.Count - 3);
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    }

                    args.Clear();
                    break;
                case "g":
                    if (args.Count >= 1) {
                        fillColor = ReadGray(args.Count - 1);
                        fillColorSpace = PdfPageColorSpaceKind.DeviceGray;
                    }

                    args.Clear();
                    break;
                case "G":
                    if (args.Count >= 1) {
                        strokeColor = ReadGray(args.Count - 1);
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceGray;
                    }

                    args.Clear();
                    break;
                case "k":
                    if (args.Count >= 4) {
                        fillColor = ReadCmyk(args.Count - 4);
                        fillColorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    }

                    args.Clear();
                    break;
                case "K":
                    if (args.Count >= 4) {
                        strokeColor = ReadCmyk(args.Count - 4);
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    }

                    args.Clear();
                    break;
                case "sc":
                case "scn":
                    if (TryReadColor(fillColorSpace, out OfficeColor parsedFillColor)) {
                        fillColor = parsedFillColor;
                    }

                    args.Clear();
                    break;
                case "SC":
                case "SCN":
                    if (TryReadColor(strokeColorSpace, out OfficeColor parsedStrokeColor)) {
                        strokeColor = parsedStrokeColor;
                    }

                    args.Clear();
                    break;
                case "Tr":
                    if (args.Count >= 1) {
                        textRenderingMode = ReadTextRenderingMode(ToDouble(args[args.Count - 1]));
                    }

                    args.Clear();
                    break;
                case "Do":
                    if (!HasHiddenContent() && args.Count >= 1) {
                        string name = ToName(args[args.Count - 1]);
                        if (!string.IsNullOrEmpty(name)) {
                            invocations.Add(new FormInvocation(name, ctm, paintOrder, fillColor, fillColorSpace, strokeColor, strokeColorSpace, fillOpacity, strokeOpacity, textRenderingMode, clipPath));
                        }
                    }
                    args.Clear();
                    break;
                case "BDC":
                    hiddenContentStack.Push(IsHiddenOptionalContent(args.Count > 1 ? args[args.Count - 2] : null, args.Count > 0 ? args[args.Count - 1] : null));
                    args.Clear();
                    break;
                case "BMC":
                    hiddenContentStack.Push(false);
                    args.Clear();
                    break;
                case "EMC":
                    if (hiddenContentStack.Count > 0) {
                        hiddenContentStack.Pop();
                    }

                    args.Clear();
                    break;
                default:
                    args.Clear();
                    break;
            }
        }

        return invocations;

        double GetPaintOrder(int operatorIndex) => paintOrderBase + ((operatorIndex + paintOrderOffset) * paintOrderScale);

        bool HasHiddenContent() {
            foreach (bool hidden in hiddenContentStack) {
                if (hidden) {
                    return true;
                }
            }

            return false;
        }

        bool IsHiddenOptionalContent(object? tag, object? property) =>
            tag is string tagName &&
            string.Equals(tagName, "OC", StringComparison.Ordinal) &&
            ((property is string propertyName &&
                optionalContentVisibility?.IsHidden(propertyName) == true) ||
             (property is PdfInlineOptionalContentReferences references &&
                optionalContentVisibility?.IsHidden(references) == true));

        void SkipWs() { while (i < n && char.IsWhiteSpace(content[i])) i++; }
        static bool IsDigit(char ch) => ch >= '0' && ch <= '9';
        bool IsNumberStart(char ch) => ch == '-' || ch == '+' || ch == '.' || IsDigit(ch);

        double ReadNumber() {
            int start = i;
            i++;
            while (i < n) {
                char ch = content[i];
                if (!(IsDigit(ch) || ch == '.' || ch == 'E' || ch == 'e' || ch == '-' || ch == '+')) break;
                i++;
            }
            var s = content.Substring(start, i - start);
            if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) v = 0;
            return v;
        }

        string ReadName() {
            i++;
            int start = i;
            while (i < n) {
                char ch = content[i];
                if (char.IsWhiteSpace(ch) || ch == '%' || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>') break;
                i++;
            }
            return PdfSyntax.DecodeName(content.Substring(start, i - start));
        }

        void ReadLiteralStringBytes() {
            i++;
            int depth = 1;
            bool esc = false;
            while (i < n && depth > 0) {
                char ch = content[i++];
                if (esc) esc = false;
                else if (ch == '\\') esc = true;
                else if (ch == '(') depth++;
                else if (ch == ')') depth--;
            }
        }

        void ReadHexStringBytes() {
            i++;
            while (i < n && content[i] != '>') i++;
            if (i < n && content[i] == '>') i++;
        }

        void ReadArray() {
            i++;
            while (i < n) {
                SkipWs();
                if (i >= n) break;
                char ch = content[i];
                if (ch == ']') { i++; break; }
                if (ch == '(') { ReadLiteralStringBytes(); continue; }
                if (ch == '<') {
                    if (i + 1 < n && content[i + 1] == '<') { i += 2; continue; }
                    ReadHexStringBytes();
                    continue;
                }
                if (IsNumberStart(ch)) { ReadNumber(); continue; }
                if (ch == '/') { ReadName(); continue; }
                if (ch == '[') { i++; continue; }
                ReadOperator();
            }
        }

        string ReadOperator() {
            int start = i;
            char ch = content[i++];
            if (ch == '\'' || ch == '"') return ch.ToString();
            while (i < n) {
                char c2 = content[i];
                if (char.IsWhiteSpace(c2) || c2 == '%' || c2 == '(' || c2 == '[' || c2 == '/' || c2 == '<' || c2 == '>') break;
                i++;
            }
            return content.Substring(start, i - start);
        }

        double NumberAt(int index) => args[index] is double value ? value : 0D;
        void ApplyGraphicsStateResource(string name) {
            if (graphicsStates == null || !graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource resource)) {
                return;
            }

            fillOpacity = resource.FillOpacity ?? fillOpacity;
            strokeOpacity = resource.StrokeOpacity ?? strokeOpacity;
        }

        OfficeColor ReadRgb(int startIndex) =>
            OfficeColor.FromRgb(ToByte(NumberAt(startIndex)), ToByte(NumberAt(startIndex + 1)), ToByte(NumberAt(startIndex + 2)));
        OfficeColor ReadGray(int index) {
            byte value = ToByte(NumberAt(index));
            return OfficeColor.FromRgb(value, value, value);
        }

        OfficeColor ReadCmyk(int startIndex) {
            return OfficeColorSpaceConverter.FromCmyk(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2), NumberAt(startIndex + 3));
        }

        bool TryReadColor(PdfPageColorSpace colorSpace, out OfficeColor color) {
            color = OfficeColor.Black;
            int componentCount = GetColorComponentCount(colorSpace);
            int endIndex = args.Count;
            while (endIndex > 0 && args[endIndex - 1] is not double) {
                endIndex--;
            }

            if (endIndex < componentCount) {
                return false;
            }

            int startIndex = endIndex - componentCount;
            switch (colorSpace.Kind) {
                case PdfPageColorSpaceKind.DeviceRgb:
                    color = ReadRgb(startIndex);
                    return true;
                case PdfPageColorSpaceKind.DeviceCmyk:
                    color = ReadCmyk(startIndex);
                    return true;
                case PdfPageColorSpaceKind.CalGray:
                    color = PdfPageColorConverter.FromCalGray(NumberAt(startIndex));
                    return true;
                case PdfPageColorSpaceKind.CalRgb:
                    color = PdfPageColorConverter.FromCalRgb(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2), colorSpace);
                    return true;
                case PdfPageColorSpaceKind.Lab:
                    color = PdfPageColorConverter.FromLab(NumberAt(startIndex), NumberAt(startIndex + 1), NumberAt(startIndex + 2));
                    return true;
                default:
                    color = ReadGray(startIndex);
                    return true;
            }
        }

        bool TryReadColorSpace(string name, out PdfPageColorSpace colorSpace) {
            switch (name) {
                case "DeviceRGB":
                case "RGB":
                    colorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    return true;
                case "DeviceCMYK":
                case "CMYK":
                    colorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    return true;
                case "DeviceGray":
                case "G":
                    colorSpace = PdfPageColorSpaceKind.DeviceGray;
                    return true;
                case "CalGray":
                    colorSpace = PdfPageColorSpaceKind.CalGray;
                    return true;
                case "CalRGB":
                    colorSpace = PdfPageColorSpaceKind.CalRgb;
                    return true;
                case "Lab":
                    colorSpace = PdfPageColorSpaceKind.Lab;
                    return true;
                default:
                    if (colorSpaces != null && colorSpaces.TryGetValue(name, out colorSpace)) {
                        return true;
                    }

                    colorSpace = PdfPageColorSpaceKind.DeviceGray;
                    return false;
            }
        }

        static int GetColorComponentCount(PdfPageColorSpace colorSpace) {
            switch (colorSpace.Kind) {
                case PdfPageColorSpaceKind.DeviceRgb:
                case PdfPageColorSpaceKind.CalRgb:
                case PdfPageColorSpaceKind.Lab:
                    return 3;
                case PdfPageColorSpaceKind.DeviceCmyk:
                    return 4;
                default:
                    return 1;
            }
        }

        static byte ToByte(double value) => (byte)Math.Round(Clamp01(value) * 255D);
        static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
        static int ReadTextRenderingMode(double value) {
            int mode = (int)Math.Round(value);
            return mode < 0 || mode > 7 ? 0 : mode;
        }

        static double ToDouble(object o) => o is double d ? d : 0.0;
        static string ToName(object o) => o as string ?? string.Empty;
    }
}
