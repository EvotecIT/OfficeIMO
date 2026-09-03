using OfficeIMO.Drawing;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

internal static class TextContentParser {
    [Flags]
    private enum PersistentGraphicsStateFlags {
        None = 0,
        FillColor = 1,
        StrokeColor = 2,
        Transform = 4,
        LineWidth = 8,
        LineCap = 16,
        LineJoin = 32,
        MiterLimit = 64,
        Dash = 128,
        RenderingIntent = 256,
        Flatness = 512,
        ExtendedState = 1024,
        TextFont = 2048,
        TextCharacterSpacing = 4096,
        TextWordSpacing = 8192,
        TextHorizontalScale = 16384,
        TextLeading = 32768,
        TextRenderingMode = 65536,
        TextRise = 131072
    }

    private sealed class PendingPersistentTextState {
        internal PendingPersistentTextState(int firstSpanIndex, int spanCount, int graphicsDepth, PersistentGraphicsStateFlags flags) {
            FirstSpanIndex = firstSpanIndex;
            SpanCount = spanCount;
            GraphicsDepth = graphicsDepth;
            Flags = flags;
        }
        internal int FirstSpanIndex { get; }
        internal int SpanCount { get; }
        internal int GraphicsDepth { get; }
        internal PersistentGraphicsStateFlags Flags { get; set; }
    }

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
        public OfficeBlendMode BlendMode { get; }
        public bool HasSoftMask { get; }
        public bool HasUnsupportedEffect { get; }
        public bool FillColorResolved { get; }
        public OfficeIccRenderingIntent RenderingIntent { get; }
        public PdfPaintColorSelection? FillColorSelection { get; }
        public PdfPaintColorSelection? StrokeColorSelection { get; }

        public TextGraphicsState(Matrix2D ctm, string font, double size, double leading, double charSpacing, double wordSpacing, double hScale, double textRise, OfficeColor fillColor, PdfPageColorSpace fillColorSpace, OfficeColor strokeColor, PdfPageColorSpace strokeColorSpace, double? fillOpacity, double? strokeOpacity, int textRenderingMode, PdfPageClipPath? clipPath, OfficeBlendMode blendMode = OfficeBlendMode.Normal, bool hasSoftMask = false, bool hasUnsupportedEffect = false, bool fillColorResolved = true, OfficeIccRenderingIntent renderingIntent = OfficeIccRenderingIntent.RelativeColorimetric, PdfPaintColorSelection? fillColorSelection = null, PdfPaintColorSelection? strokeColorSelection = null) {
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
            BlendMode = blendMode;
            HasSoftMask = hasSoftMask;
            HasUnsupportedEffect = hasUnsupportedEffect;
            FillColorResolved = fillColorResolved;
            RenderingIntent = renderingIntent;
            FillColorSelection = fillColorSelection;
            StrokeColorSelection = strokeColorSelection;
        }
    }

    private readonly struct ActualTextValue {
        private readonly string? _text;
        private readonly byte[]? _bytes;

        public ActualTextValue(string text) {
            _text = text;
            _bytes = null;
        }

        public ActualTextValue(byte[] bytes) {
            _text = null;
            _bytes = bytes;
        }

        public string Decode(TextOutputBudget budget) {
            if (_bytes != null) {
                budget.EnsureActualTextMayFit(PdfTextString.GetDecodedCharacterCount(_bytes));
                return PdfTextString.Decode(_bytes);
            }

            string text = _text ?? string.Empty;
            budget.EnsureActualTextMayFit(text.Length);
            return text;
        }
    }

    private sealed class MarkedContentState {
        private readonly ActualTextValue? _actualText;
        public bool HasActualText { get; }
        public bool IsArtifact { get; }
        public bool IsHidden { get; }
        public int? Mcid { get; }
        public bool HasMcid => Mcid.HasValue;
        public bool IsOptionalContent { get; }
        public bool ActualTextEmitted { get; set; }

        public MarkedContentState(ActualTextValue? actualText, bool isArtifact, bool isHidden, int? mcid = null, bool isOptionalContent = false) {
            _actualText = actualText;
            HasActualText = actualText.HasValue;
            IsArtifact = isArtifact;
            IsHidden = isHidden;
            Mcid = mcid;
            IsOptionalContent = isOptionalContent;
        }

        public string DecodeActualText(TextOutputBudget budget) =>
            _actualText?.Decode(budget) ?? string.Empty;
    }

    internal sealed class TextOutputBudget {
        private readonly int _maxActualTextCharacters;
        private readonly int _maxDecodedTextCharacters;
        private long _actualTextCharacters;
        private long _decodedTextCharacters;

        internal TextOutputBudget(int maxActualTextCharacters, int maxDecodedTextCharacters) {
#if NET8_0_OR_GREATER
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxActualTextCharacters);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxDecodedTextCharacters);
#else
            if (maxActualTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxActualTextCharacters));
            if (maxDecodedTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxDecodedTextCharacters));
#endif
            _maxActualTextCharacters = maxActualTextCharacters;
            _maxDecodedTextCharacters = maxDecodedTextCharacters;
        }

        internal void ChargeActualText(int characters) {
            long next = _actualTextCharacters + characters;
            if (next > _maxActualTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ActualTextCharacters, _maxActualTextCharacters, next);
            }

            _actualTextCharacters = next;
        }

        internal void EnsureActualTextMayFit(int characters) {
            long next = _actualTextCharacters + characters;
            if (next > _maxActualTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ActualTextCharacters, _maxActualTextCharacters, next);
            }
        }

        internal void ChargeDecodedText(int characters) {
            long next = _decodedTextCharacters + characters;
            if (next > _maxDecodedTextCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, _maxDecodedTextCharacters, next);
            }

            _decodedTextCharacters = next;
        }

        internal void ThrowDecodedTextLimitExceeded() =>
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.DecodedTextCharacters,
                _maxDecodedTextCharacters,
                (long)_maxDecodedTextCharacters + 1L);

        internal int GetDecodedTextBufferCapacity(int requestedCapacity) {
            return Math.Min(GetRemainingDecodedTextCharacters(), Math.Max(0, requestedCapacity));
        }

        internal int GetRemainingDecodedTextCharacters() {
            return (int)Math.Max(0L, _maxDecodedTextCharacters - _decodedTextCharacters);
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
        public int SourceOperatorIndex { get; }
        public bool HasUnsupportedEffect { get; }
        public bool FillColorResolved { get; }
        public OfficeIccRenderingIntent RenderingIntent { get; }
        public PdfPaintColorSelection? FillColorSelection { get; }
        public PdfPaintColorSelection? StrokeColorSelection { get; }
        public PdfTextStateSnapshot TextState { get; }

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
            PdfPageClipPath? clipPath = null,
            int sourceOperatorIndex = 0,
            bool hasUnsupportedEffect = false,
            bool fillColorResolved = true,
            OfficeIccRenderingIntent renderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
            PdfPaintColorSelection? fillColorSelection = null,
            PdfPaintColorSelection? strokeColorSelection = null,
            PdfTextStateSnapshot? textState = null) {
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
            SourceOperatorIndex = sourceOperatorIndex;
            HasUnsupportedEffect = hasUnsupportedEffect;
            FillColorResolved = fillColorResolved;
            RenderingIntent = renderingIntent;
            FillColorSelection = fillColorSelection;
            StrokeColorSelection = strokeColorSelection;
            TextState = textState ?? PdfTextStateSnapshot.Default.WithTextRenderingMode(textRenderingMode);
        }
    }

    public static List<PdfTextSpan> Parse(
        string content,
        System.Func<string, byte[], string> decodeWithFont,
        System.Func<string, byte[], double> sumWidth1000ForFont,
        bool adjustKerningFromTJ = true,
        System.Func<string, byte[]?>? actualTextForProperty = null,
        System.Func<string, int?>? mcidForProperty = null,
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
        bool includeArtifactText = false,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands,
        int maxActualTextCharacters = PdfReadLimits.DefaultMaxActualTextCharacters,
        int maxDecodedTextCharacters = PdfReadLimits.DefaultMaxDecodedTextCharacters,
        TextOutputBudget? textOutputBudget = null,
        PdfTextClippingBudget? textClippingBudget = null,
        System.Func<string, byte[], int, string>? decodeWithFontWithinLimit = null,
        PdfContentOrderKey? contentOrderPrefix = null,
        int contentOrderOffset = 0,
        bool initialUnsupportedEffect = false,
        OfficeIccRenderingIntent initialRenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        PdfPaintColorSelection? initialFillColorSelection = null,
        PdfPaintColorSelection? initialStrokeColorSelection = null,
        PdfOutputIntentColorTransform? outputIntentColorTransform = null,
        Func<string, int>? inlineImageComponentCount = null,
        Func<PdfArray, int>? inlineImageArrayComponentCount = null,
        int? contentStreamObjectNumber = null,
        Func<int, int?>? contentStreamObjectNumberAtOffset = null,
        Action? cancellationCheck = null,
        PdfTextStateSnapshot? initialTextState = null) {
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxActualTextCharacters);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxDecodedTextCharacters);
#else
        if (maxActualTextCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxActualTextCharacters));
        }
        if (maxDecodedTextCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxDecodedTextCharacters));
        }
#endif

        textOutputBudget ??= new TextOutputBudget(maxActualTextCharacters, maxDecodedTextCharacters);
        textClippingBudget ??= new PdfTextClippingBudget();

        var spans = new List<PdfTextSpan>();
        PdfTextStateSnapshot startingTextState = initialTextState ?? PdfTextStateSnapshot.Default;
        // Text state
        bool inText = false;
        string font = startingTextState.FontResource;
        double size = startingTextState.FontSize;
        double leading = startingTextState.Leading;
        double charSpacing = startingTextState.CharacterSpacing;
        double wordSpacing = startingTextState.WordSpacing;
        double hScale = startingTextState.HorizontalScaling;
        double textRise = startingTextState.TextRise;
        OfficeColor fillColor = initialFillColor ?? OfficeColor.Black;
        PdfPageColorSpace fillColorSpace = initialFillColorSpace;
        OfficeColor strokeColor = initialStrokeColor ?? OfficeColor.Black;
        PdfPageColorSpace strokeColorSpace = initialStrokeColorSpace;
        double? fillOpacity = initialFillOpacity;
        double? strokeOpacity = initialStrokeOpacity;
        int textRenderingMode = ReadTextRenderingMode(initialTextState.HasValue
            ? startingTextState.TextRenderingMode
            : initialTextRenderingMode);
        PdfPageClipPath? clipPath = initialClipPath;
        OfficeBlendMode blendMode = OfficeBlendMode.Normal;
        bool hasSoftMask = false;
        bool hasUnsupportedEffect = initialUnsupportedEffect;
        bool fillColorResolved = initialFillColorSpace.Kind != PdfPageColorSpaceKind.Pattern;
        OfficeIccRenderingIntent renderingIntent = initialRenderingIntent;
        PdfPaintColorSelection? fillColorSelection = initialFillColorSelection;
        PdfPaintColorSelection? strokeColorSelection = initialStrokeColorSelection;
        if (fillColorSelection != null && fillColorSelection.TryConvert(renderingIntent, out OfficeColor selectedFillColor)) {
            fillColor = selectedFillColor;
            fillColorSpace = fillColorSelection.ColorSpace;
        } else if (!initialFillColor.HasValue && outputIntentColorTransform != null &&
            PdfPaintColorSelection.TryCreateDefaultBlack(renderingIntent, outputIntentColorTransform, out fillColorSelection, out OfficeColor defaultFillColor)) {
            fillColor = defaultFillColor;
            fillColorSpace = PdfPageColorSpaceKind.DeviceGray;
        }
        if (strokeColorSelection != null && strokeColorSelection.TryConvert(renderingIntent, out OfficeColor selectedStrokeColor)) {
            strokeColor = selectedStrokeColor;
            strokeColorSpace = strokeColorSelection.ColorSpace;
        } else if (!initialStrokeColor.HasValue && outputIntentColorTransform != null &&
            PdfPaintColorSelection.TryCreateDefaultBlack(renderingIntent, outputIntentColorTransform, out strokeColorSelection, out OfficeColor defaultStrokeColor)) {
            strokeColor = defaultStrokeColor;
            strokeColorSpace = PdfPageColorSpaceKind.DeviceGray;
        }
        OfficeColor effectiveInitialFillColor = fillColor;
        PdfPageColorSpace effectiveInitialFillColorSpace = fillColorSpace;
        OfficeColor effectiveInitialStrokeColor = strokeColor;
        PdfPageColorSpace effectiveInitialStrokeColorSpace = strokeColorSpace;
        PdfPaintColorSelection? effectiveInitialFillColorSelection = fillColorSelection;
        PdfPaintColorSelection? effectiveInitialStrokeColorSelection = strokeColorSelection;
        var clipPathBuilder = new PdfPageClipPathBuilder(pageHeight);
        var pendingTextClipPaths = new List<PdfPageClipPath>();
        Matrix2D textMatrix = Matrix2D.Identity;
        Matrix2D lineMatrix = Matrix2D.Identity;
        // Graphics state (CTM) and stack
        Matrix2D ctm = Matrix2D.Identity; var gstack = new System.Collections.Generic.Stack<TextGraphicsState>();
        // Operand buffer for the current operator.
        var args = new List<object>(8);
        // Kerning state between text runs in TJ arrays (points) and rolling output buffer for gap checks
        double pendingGapPt = 0;
        int pendingLineBreaks = 0;
        bool emittedTextInTextObject = false;
        PdfContentOrderKey? currentContentOrderKey = null;
        PdfContentOrderKey? currentTextObjectOrderKey = null;
        int textObjectFirstSpanIndex = 0;
        PersistentGraphicsStateFlags textObjectPersistentState = PersistentGraphicsStateFlags.None;
        bool textObjectHasCollateralVisual = false;
        var pendingPersistentTextStates = new List<PendingPersistentTextState>();
        var sbOutGlobal = new StringBuilder();
        var markedContentStack = new Stack<MarkedContentState>();
        int? currentContentStreamObjectNumber = contentStreamObjectNumber;
        PdfContentStreamInterpreter.Interpret(content, maxOperations, operation => {
            cancellationCheck?.Invoke();
            currentContentStreamObjectNumber = contentStreamObjectNumberAtOffset?.Invoke(operation.OperatorOffset)
                ?? contentStreamObjectNumber;
            args.Clear();
            args.AddRange(operation.Operands);
            double paintOrder = GetPaintOrder(operation.OperatorOffset);
            currentContentOrderKey = contentOrderPrefix?.Append(operation.OperatorOffset + contentOrderOffset);
            string op = operation.Name;
            if (string.Equals(op, "BT", StringComparison.Ordinal)) {
                currentTextObjectOrderKey = currentContentOrderKey;
                textObjectFirstSpanIndex = spans.Count;
                textObjectPersistentState = PersistentGraphicsStateFlags.None;
                textObjectHasCollateralVisual = false;
            } else {
                PersistentGraphicsStateFlags changedState = GetPersistentGraphicsStateFlag(op);
                if (changedState != PersistentGraphicsStateFlags.None) {
                    OverridePendingPersistentState(changedState);
                    if (inText) textObjectPersistentState |= changedState;
                }
                if (string.Equals(op, "Q", StringComparison.Ordinal)) DiscardRestoredPendingState();
                if (IsTextShowOperator(op)) MarkPendingPersistentStateAsUnsafe();
                if (inText && IsNonTextVisualConsumer(op)) textObjectHasCollateralVisual = true;
                else if (!inText && IsNonTextVisualConsumer(op)) MarkPendingPersistentStateAsUnsafe();
            }
            switch (op) {
                case "BT": ApplyPendingTextClippingPath(); inText = true; textMatrix = Matrix2D.Identity; lineMatrix = Matrix2D.Identity; pendingGapPt = 0; pendingLineBreaks = 0; emittedTextInTextObject = false; args.Clear(); break;
                case "ET": ApplyPendingTextClippingPath(); QueueCurrentTextObjectPersistentState(); inText = false; pendingGapPt = 0; pendingLineBreaks = 0; emittedTextInTextObject = false; args.Clear(); break;
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
                    gstack.Push(new TextGraphicsState(ctm, font, size, leading, charSpacing, wordSpacing, hScale, textRise, fillColor, fillColorSpace, strokeColor, strokeColorSpace, fillOpacity, strokeOpacity, textRenderingMode, clipPath, blendMode, hasSoftMask, hasUnsupportedEffect, fillColorResolved, renderingIntent, fillColorSelection, strokeColorSelection));
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
                        blendMode = state.BlendMode;
                        hasSoftMask = state.HasSoftMask;
                        hasUnsupportedEffect = state.HasUnsupportedEffect;
                        fillColorResolved = state.FillColorResolved;
                        renderingIntent = state.RenderingIntent;
                        fillColorSelection = state.FillColorSelection;
                        strokeColorSelection = state.StrokeColorSelection;
                    } else {
                        ctm = Matrix2D.Identity;
                        fillColor = effectiveInitialFillColor;
                        fillColorSpace = effectiveInitialFillColorSpace;
                        strokeColor = effectiveInitialStrokeColor;
                        strokeColorSpace = effectiveInitialStrokeColorSpace;
                        fillOpacity = null;
                        strokeOpacity = null;
                        textRenderingMode = 0;
                        clipPath = null;
                        blendMode = OfficeBlendMode.Normal;
                        hasSoftMask = false;
                        hasUnsupportedEffect = initialUnsupportedEffect;
                        fillColorResolved = effectiveInitialFillColorSpace.Kind != PdfPageColorSpaceKind.Pattern;
                        renderingIntent = initialRenderingIntent;
                        fillColorSelection = effectiveInitialFillColorSelection;
                        strokeColorSelection = effectiveInitialStrokeColorSelection;
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
                        clipPath = textClippingBudget.ResolveActiveClip(clipPath, parsedClipPath);
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
                        fillColorSelection = null;
                        fillColorSpace = parsedColorSpace;
                        fillColorResolved = parsedColorSpace.Kind != PdfPageColorSpaceKind.Pattern;
                    } else {
                        fillColorSpace = PdfPageColorSpaceKind.Pattern;
                        fillColorResolved = false;
                    }

                    args.Clear();
                    break;
                case "CS":
                    if (args.Count >= 1 && TryReadColorSpace(ToName(args[args.Count - 1]), out PdfPageColorSpace parsedStrokeColorSpace)) {
                        strokeColorSelection = null;
                        strokeColorSpace = parsedStrokeColorSpace;
                    }

                    args.Clear();
                    break;
                case "rg":
                    if (args.Count >= 3 && TryApplyFillColor(PdfPageColorSpaceKind.DeviceRgb, out OfficeColor rgbFill)) {
                        fillColor = rgbFill;
                        fillColorSpace = PdfPageColorSpaceKind.DeviceRgb;
                        fillColorResolved = true;
                    }

                    args.Clear();
                    break;
                case "RG":
                    if (args.Count >= 3 && TryApplyStrokeColor(PdfPageColorSpaceKind.DeviceRgb, out OfficeColor rgbStroke)) {
                        strokeColor = rgbStroke;
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceRgb;
                    }

                    args.Clear();
                    break;
                case "g":
                    if (args.Count >= 1 && TryApplyFillColor(PdfPageColorSpaceKind.DeviceGray, out OfficeColor grayFill)) {
                        fillColor = grayFill;
                        fillColorSpace = PdfPageColorSpaceKind.DeviceGray;
                        fillColorResolved = true;
                    }

                    args.Clear();
                    break;
                case "G":
                    if (args.Count >= 1 && TryApplyStrokeColor(PdfPageColorSpaceKind.DeviceGray, out OfficeColor grayStroke)) {
                        strokeColor = grayStroke;
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceGray;
                    }

                    args.Clear();
                    break;
                case "k":
                    if (args.Count >= 4 && TryApplyFillColor(PdfPageColorSpaceKind.DeviceCmyk, out OfficeColor cmykFill)) {
                        fillColor = cmykFill;
                        fillColorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                        fillColorResolved = true;
                    }

                    args.Clear();
                    break;
                case "K":
                    if (args.Count >= 4 && TryApplyStrokeColor(PdfPageColorSpaceKind.DeviceCmyk, out OfficeColor cmykStroke)) {
                        strokeColor = cmykStroke;
                        strokeColorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                    }

                    args.Clear();
                    break;
                case "sc":
                case "scn":
                    fillColorResolved = TryApplyFillColor(fillColorSpace, out OfficeColor parsedFillColor);
                    if (fillColorResolved) {
                        fillColor = parsedFillColor;
                    }

                    args.Clear();
                    break;
                case "SC":
                case "SCN":
                    if (TryApplyStrokeColor(strokeColorSpace, out OfficeColor parsedStrokeColor)) {
                        strokeColor = parsedStrokeColor;
                    }

                    args.Clear();
                    break;
                case "ri":
                    if (args.Count == 1 && args[0] is string renderingIntentName) {
                        ApplyRenderingIntent(PdfRenderingIntentResolver.FromName(renderingIntentName));
                    } else {
                        hasUnsupportedEffect = true;
                        MarkPendingPersistentStateAsUnsafe();
                        if (inText) textObjectHasCollateralVisual = true;
                    }
                    args.Clear();
                    break;
                case "BI":
                    args.Clear();
                    break;
                case "'": // move to next line and show text
                    if (args.Count >= 1) { MoveToNextTextLine(); ShowTextRun(ToBytes(args[args.Count - 1]), paintOrder, forceCannotRestamp: false); pendingGapPt = 0; }
                    args.Clear();
                    break;
                case "\"": // set spacing and show text
                    if (args.Count >= 3) { wordSpacing = ToDouble(args[args.Count - 3]); charSpacing = ToDouble(args[args.Count - 2]); MoveToNextTextLine(); ShowTextRun(ToBytes(args[args.Count - 1]), paintOrder, forceCannotRestamp: false); pendingGapPt = 0; }
                    args.Clear();
                    break;
                case "Tj": if (args.Count >= 1) { ShowTextRun(ToBytes(args[args.Count - 1]), paintOrder, forceCannotRestamp: false); pendingGapPt = 0; args.Clear(); } break;
                case "TJ": if (args.Count >= 1) { ShowTextArray(args[args.Count - 1], paintOrder); args.Clear(); } break;
                case "BDC":
                    markedContentStack.Push(new MarkedContentState(
                        GetActualText(args.Count > 0 ? args[args.Count - 1] : null),
                        IsArtifactTag(args.Count > 1 ? args[args.Count - 2] : null),
                        operation.HasInvalidOperands ||
                        IsHiddenOptionalContent(args.Count > 1 ? args[args.Count - 2] : null, args.Count > 0 ? args[args.Count - 1] : null),
                        GetMcid(args.Count > 0 ? args[args.Count - 1] : null),
                        IsOptionalContentTag(args.Count > 1 ? args[args.Count - 2] : null)));
                    args.Clear();
                    break;
                case "BMC":
                    markedContentStack.Push(new MarkedContentState(
                        null,
                        IsArtifactTag(args.Count > 0 ? args[args.Count - 1] : null),
                        operation.HasInvalidOperands));
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
        }, inlineImageComponentCount: inlineImageComponentCount, maxNestingDepth: maxNestingDepth, maxOperands: maxOperands, inlineImageArrayComponentCount: inlineImageArrayComponentCount);
        ApplyPendingTextClippingPath();
        return spans;

        // Helpers
        void QueueCurrentTextObjectPersistentState() {
            int spanCount = spans.Count - textObjectFirstSpanIndex;
            if (spanCount <= 0) return;
            if (textObjectHasCollateralVisual) {
                MarkSpansCannotRestamp(textObjectFirstSpanIndex, spanCount);
                return;
            }
            if (textObjectPersistentState != PersistentGraphicsStateFlags.None) {
                pendingPersistentTextStates.Add(new PendingPersistentTextState(textObjectFirstSpanIndex, spanCount, gstack.Count, textObjectPersistentState));
            }
        }

        void OverridePendingPersistentState(PersistentGraphicsStateFlags changedState) {
            PersistentGraphicsStateFlags overrideable = changedState & ~PersistentGraphicsStateFlags.Transform & ~PersistentGraphicsStateFlags.ExtendedState;
            if (overrideable == PersistentGraphicsStateFlags.None) return;
            for (int index = pendingPersistentTextStates.Count - 1; index >= 0; index--) {
                PendingPersistentTextState pending = pendingPersistentTextStates[index];
                pending.Flags &= ~overrideable;
                if (pending.Flags == PersistentGraphicsStateFlags.None) pendingPersistentTextStates.RemoveAt(index);
            }
        }

        void DiscardRestoredPendingState() {
            int restoredDepth = gstack.Count;
            for (int index = pendingPersistentTextStates.Count - 1; index >= 0; index--) {
                if (pendingPersistentTextStates[index].GraphicsDepth >= restoredDepth) pendingPersistentTextStates.RemoveAt(index);
            }
        }

        void MarkPendingPersistentStateAsUnsafe() {
            for (int index = 0; index < pendingPersistentTextStates.Count; index++) {
                PendingPersistentTextState pending = pendingPersistentTextStates[index];
                MarkSpansCannotRestamp(pending.FirstSpanIndex, pending.SpanCount);
            }
            pendingPersistentTextStates.Clear();
        }

        void MarkSpansCannotRestamp(int firstSpanIndex, int spanCount) {
            int end = Math.Min(spans.Count, firstSpanIndex + spanCount);
            for (int index = firstSpanIndex; index < end; index++) spans[index] = spans[index].WithCanRestamp(false);
        }

        static PersistentGraphicsStateFlags GetPersistentGraphicsStateFlag(string name) => name switch {
            "g" or "rg" or "k" or "cs" or "sc" or "scn" => PersistentGraphicsStateFlags.FillColor,
            "G" or "RG" or "K" or "CS" or "SC" or "SCN" => PersistentGraphicsStateFlags.StrokeColor,
            "cm" => PersistentGraphicsStateFlags.Transform,
            "w" => PersistentGraphicsStateFlags.LineWidth,
            "J" => PersistentGraphicsStateFlags.LineCap,
            "j" => PersistentGraphicsStateFlags.LineJoin,
            "M" => PersistentGraphicsStateFlags.MiterLimit,
            "d" => PersistentGraphicsStateFlags.Dash,
            "ri" => PersistentGraphicsStateFlags.RenderingIntent,
            "i" => PersistentGraphicsStateFlags.Flatness,
            "gs" => PersistentGraphicsStateFlags.ExtendedState,
            "Tf" => PersistentGraphicsStateFlags.TextFont,
            "Tc" => PersistentGraphicsStateFlags.TextCharacterSpacing,
            "Tw" => PersistentGraphicsStateFlags.TextWordSpacing,
            "Tz" => PersistentGraphicsStateFlags.TextHorizontalScale,
            "TL" or "TD" => PersistentGraphicsStateFlags.TextLeading,
            "Tr" => PersistentGraphicsStateFlags.TextRenderingMode,
            "Ts" => PersistentGraphicsStateFlags.TextRise,
            "\"" => PersistentGraphicsStateFlags.TextCharacterSpacing | PersistentGraphicsStateFlags.TextWordSpacing,
            _ => PersistentGraphicsStateFlags.None
        };

        static bool IsTextShowOperator(string name) => name is "Tj" or "TJ" or "'" or "\"";

        static bool IsNonTextVisualConsumer(string name) => name switch {
            "S" or "s" or "f" or "F" or "f*" or "B" or "B*" or "b" or "b*" or "Do" or "sh" or "BI" => true,
            _ => false
        };

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
        void ShowTextRun(byte[] bytes, double paintOrder, bool forceCannotRestamp) {
            if (!inText || bytes == null || bytes.Length == 0) return;
            MaybeInsertSpaceBeforeRun();
            string DecodeRun(byte[] value, int? maximumCharacters = null) {
                int remaining = maximumCharacters ?? textOutputBudget.GetRemainingDecodedTextCharacters();
                if (remaining == 0) {
                    textOutputBudget.ThrowDecodedTextLimitExceeded();
                }
                return decodeWithFontWithinLimit != null
                    ? decodeWithFontWithinLimit(font, value, remaining)
                    : decodeWithFont(font, value);
            }
            // Detect 2-byte CIDs (Identity-H) vs single-byte
            bool twoByte = false;
            if (bytes.Length >= 2) {
                string one = DecodeRun(new byte[] { bytes[0] });
                string two = DecodeRun(new byte[] { bytes[0], bytes[1] });
                double firstByteWidth = sumWidth1000ForFont(font, new byte[] { bytes[0] });
                double secondByteWidth = sumWidth1000ForFont(font, new byte[] { bytes[1] });
                double pairWidth = sumWidth1000ForFont(font, new byte[] { bytes[0], bytes[1] });
                twoByte = (IsNullOrEmptyDecodedGlyph(one) && !IsNullOrEmptyDecodedGlyph(two)) ||
                    (firstByteWidth <= 0 && secondByteWidth <= 0 && pairWidth > 0);
            }
            var sbOut = new StringBuilder(textOutputBudget.GetDecodedTextBufferCapacity(bytes.Length));
            var decodedAdvances = new List<double>();
            var decodedGlyphCharacterLengths = new List<int>();
            var decodedGlyphBytes = new List<byte[]>();
            double advTotal = 0;
            string wholeDecoded = NormalizeDecodedGlyphText(DecodeRun(bytes) ?? string.Empty);
            int decodedGlyphCharacters = 0;
            for (int idx = 0; idx < bytes.Length;) {
                int step = twoByte ? (idx + 1 < bytes.Length ? 2 : 1) : 1;
                byte[] g = step == 1 ? new byte[] { bytes[idx] } : new byte[] { bytes[idx], bytes[idx + 1] };
                int remainingGlyphCharacters = textOutputBudget.GetRemainingDecodedTextCharacters() - decodedGlyphCharacters;
                if (remainingGlyphCharacters <= 0) {
                    textOutputBudget.ThrowDecodedTextLimitExceeded();
                }
                string t = NormalizeDecodedGlyphText(DecodeRun(g, remainingGlyphCharacters) ?? string.Empty);
                if (t.Length > remainingGlyphCharacters) {
                    textOutputBudget.ThrowDecodedTextLimitExceeded();
                }
                decodedGlyphCharacters += t.Length;
                char ch = (t.Length > 0) ? t[0] : '\0';
                double w1000 = sumWidth1000ForFont(font, g);
                double advGlyph = ((w1000 / 1000.0) * size + charSpacing + (step == 1 && bytes[idx] == 0x20 ? wordSpacing : 0)) * hScale;
                if (ch != '\0') {
                    sbOut.Append(t);
                    decodedGlyphCharacterLengths.Add(t.Length);
                    decodedGlyphBytes.Add(g);
                    double perCharacterAdvance = advGlyph / Math.Max(1, t.Length);
                    for (int characterIndex = 0; characterIndex < t.Length; characterIndex++) decodedAdvances.Add(perCharacterAdvance);
                }
                advTotal += advGlyph;
                idx += step;
            }
            textOutputBudget.ChargeDecodedText(Math.Max(wholeDecoded.Length, decodedGlyphCharacters));
            if (ShouldUseWholeDecodedText(sbOut.ToString(), wholeDecoded)) {
                sbOut.Clear();
                sbOut.Append(wholeDecoded);
                decodedAdvances.Clear();
                decodedGlyphCharacterLengths.Clear();
                decodedGlyphBytes.Clear();
            }
            var actualTextState = useLogicalTextFilters ? GetActiveActualTextState() : null;
            bool hasActiveArtifact = HasActiveArtifact();
            bool isArtifact = useLogicalTextFilters && !includeArtifactText && hasActiveArtifact;
            bool isHidden = HasActiveHiddenContent();
            bool usesVisibleFill = UsesFillTextPaint(textRenderingMode) && !fillColorSpace.SuppressesPaint;
            bool usesVisibleStroke = UsesStrokeTextPaint(textRenderingMode) && !strokeColorSpace.SuppressesPaint;
            bool isVisibleText = usesVisibleFill || usesVisibleStroke;
            if (sbOut.Length == 0 && actualTextState is null && !isArtifact && !isHidden) return;
            string textOut = sbOut.ToString();
            var textOrigin = textMatrix.Transform(0, textRise);
            var (dx, dy) = ctm.Transform(textOrigin.X, textOrigin.Y);
            var textEnd = textMatrix.Transform(advTotal, textRise);
            var (endX, endY) = ctm.Transform(textEnd.X, textEnd.Y);
            double transformedAdvance = Math.Sqrt(((endX - dx) * (endX - dx)) + ((endY - dy) * (endY - dy)));
            double rotationDegrees = CalculateRotationDegrees(endX - dx, endY - dy);
            var textUnitX = textMatrix.Transform(1D, textRise);
            var textUnitY = textMatrix.Transform(0D, textRise + 1D);
            var (unitXPageX, unitXPageY) = ctm.Transform(textUnitX.X, textUnitX.Y);
            var (unitYPageX, unitYPageY) = ctm.Transform(textUnitY.X, textUnitY.Y);
            double unitXLength = Math.Sqrt(((unitXPageX - dx) * (unitXPageX - dx)) + ((unitXPageY - dy) * (unitXPageY - dy)));
            double unitYLength = Math.Sqrt(((unitYPageX - dx) * (unitYPageX - dx)) + ((unitYPageY - dy) * (unitYPageY - dy)));
            double unitDot = ((unitXPageX - dx) * (unitYPageX - dx)) + ((unitXPageY - dy) * (unitYPageY - dy));
            double unitDeterminant = ((unitXPageX - dx) * (unitYPageY - dy)) - ((unitXPageY - dy) * (unitYPageX - dx));
            bool canRestamp = unitXLength > 0.000001D &&
                unitYLength > 0.000001D &&
                unitDeterminant > 0D &&
                Math.Abs(unitXLength - unitYLength) <= Math.Max(unitXLength, unitYLength) * 0.0001D &&
                Math.Abs(unitDot) <= unitXLength * unitYLength * 0.0001D &&
                Math.Abs(hScale - 1D) <= 0.0001D &&
                Math.Abs(charSpacing) <= 0.000001D &&
                Math.Abs(wordSpacing) <= 0.000001D &&
                blendMode == OfficeBlendMode.Normal &&
                !hasSoftMask &&
                !hasUnsupportedEffect &&
                isVisibleText &&
                fillColorResolved &&
                !HasActiveMcid() &&
                !HasActiveOptionalContent() &&
                !hasActiveArtifact &&
                outputIntentColorTransform == null &&
                !forceCannotRestamp;
            double restampFontSize = size * unitYLength;
            IReadOnlyList<double>? transformedCharacterAdvances = decodedAdvances.Count == textOut.Length
                ? decodedAdvances.Select(advance => advance * unitXLength).ToArray()
                : null;
            bool useStrokePaint = !usesVisibleFill && usesVisibleStroke;
            OfficeColor paintColor = useStrokePaint ? strokeColor : fillColor;
            OfficeColor visibleColor = ApplyTextOpacity(paintColor, useStrokePaint);
            PdfPageClipPath? spanClipPath = clipPath;
            if (isHidden) {
                // Hidden optional-content still advances text state but should not emit visible/logical spans.
            } else if (isArtifact) {
                // Artifact content is visual decoration, not logical page text.
            } else if (actualTextState is not null && !actualTextState.ActualTextEmitted) {
                textOut = actualTextState.DecodeActualText(textOutputBudget);
                textOutputBudget.ChargeActualText(textOut.Length);
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
                string paintedText = sbOut.ToString();
                bool visibleGlyphsMatchLogicalText = string.Equals(
                    NormalizeShatteredSpan(paintedText),
                    normalizedText,
                    StringComparison.Ordinal);

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
                    logicalTrailingSpace,
                    currentContentOrderKey,
                    string.Equals(normalizedText, sbOut.ToString(), StringComparison.Ordinal) ? transformedCharacterAdvances : null,
                    textRenderingMode,
                    canRestamp && visibleGlyphsMatchLogicalText,
                    restampFontSize,
                    paintedText,
                    Math.Abs(charSpacing) <= 0.000001D && Math.Abs(wordSpacing) <= 0.000001D,
                    GetActiveMcid(),
                    currentContentStreamObjectNumber,
                    currentTextObjectOrderKey,
                    Matrix2D.Multiply(ctm, textMatrix),
                    string.Join(",", new[] {
                        fillColor.R.ToString(CultureInfo.InvariantCulture), fillColor.G.ToString(CultureInfo.InvariantCulture), fillColor.B.ToString(CultureInfo.InvariantCulture), fillColor.A.ToString(CultureInfo.InvariantCulture),
                        strokeColor.R.ToString(CultureInfo.InvariantCulture), strokeColor.G.ToString(CultureInfo.InvariantCulture), strokeColor.B.ToString(CultureInfo.InvariantCulture), strokeColor.A.ToString(CultureInfo.InvariantCulture),
                        fillOpacity?.ToString("R", CultureInfo.InvariantCulture) ?? "null", strokeOpacity?.ToString("R", CultureInfo.InvariantCulture) ?? "null",
                        ((int)blendMode).ToString(CultureInfo.InvariantCulture), hasSoftMask ? "1" : "0", hasUnsupportedEffect ? "1" : "0"
                    }),
                    decodedGlyphCharacterLengths.Sum() == paintedText.Length
                        ? decodedGlyphCharacterLengths
                        : null,
                    decodedGlyphCharacterLengths.Sum() == paintedText.Length &&
                        decodedGlyphBytes.Count == decodedGlyphCharacterLengths.Count
                        ? decodedGlyphBytes
                        : null));
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
                textClippingBudget.ChargePath();
                pendingTextClipPaths.Add(textClipPath);
            }
        }

        void ApplyPendingTextClippingPath() {
            if (PdfPageClipPath.TryCombineTextClippingPaths(pendingTextClipPaths, out PdfPageClipPath textClipPath)) {
                clipPath = textClippingBudget.ResolveActiveClip(clipPath, textClipPath);
            }
            pendingTextClipPaths.Clear();
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

        bool HasActiveMcid() {
            foreach (var state in markedContentStack) if (state.HasMcid) return true;
            return false;
        }

        bool HasActiveOptionalContent() {
            foreach (var state in markedContentStack) if (state.IsOptionalContent) return true;
            return false;
        }

        int? GetActiveMcid() {
            foreach (var state in markedContentStack) if (state.Mcid.HasValue) return state.Mcid;
            return null;
        }

        int? GetMcid(object? propertyObject) {
            if (propertyObject is string propertyName) return mcidForProperty?.Invoke(propertyName);
            if (propertyObject is not PdfContentDictionary dictionary || !dictionary.Items.TryGetValue("MCID", out object? value)) return null;
            return TryGetMcid(value);
        }

        static bool IsOptionalContentTag(object? tag) =>
            tag is string name && string.Equals(name, "OC", StringComparison.Ordinal);

        ActualTextValue? GetActualText(object? propertyObject) {
            if (propertyObject is string propertyName) {
                byte[]? propertyBytes = actualTextForProperty?.Invoke(propertyName);
                return propertyBytes is null ? (ActualTextValue?)null : new ActualTextValue(propertyBytes);
            }

            if (propertyObject is PdfContentDictionary dictionary &&
                dictionary.Items.TryGetValue("ActualText", out var value) &&
                value is byte[] bytes) {
                return new ActualTextValue(bytes);
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
             (property is PdfContentDictionary dictionary &&
                dictionary.OptionalContentReferences != null &&
                optionalContentVisibility?.IsHidden(dictionary.OptionalContentReferences) == true));

        void ShowTextArray(object arrObj, double paintOrder) {
            if (!inText || arrObj == null) return;
            var list = arrObj as List<object>;
            if (list == null) return;
            bool hasPositioningAdjustment = list.Any(static item => item is double value && Math.Abs(value) > 0.000001D);
            for (int j = 0; j < list.Count; j++) {
                var it = list[j];
                if (it is byte[] b) { ShowTextRun(b, paintOrder, hasPositioningAdjustment); }
                else if (adjustKerningFromTJ && it is double num) {
                    double delta = -num / 1000.0 * size * hScale;
                    textMatrix = Matrix2D.Multiply(textMatrix, Matrix2D.Translation(delta, 0));
                    // Only positive visual gap should suggest a space
                    if (delta > 0) pendingGapPt += delta; else pendingGapPt = 0;
                }
            }
        }

        static double ToDouble(object o) { return o is double d ? d : 0.0; }
        static string ToName(object o) { return o as string ?? string.Empty; }
        static byte[] ToBytes(object o) { return o as byte[] ?? Array.Empty<byte>(); }
        void ApplyGraphicsStateResource(string name) {
            if (graphicsStates != null && graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource resource)) {
                fillOpacity = resource.FillOpacity ?? fillOpacity;
                strokeOpacity = resource.StrokeOpacity ?? strokeOpacity;
                blendMode = resource.BlendMode ?? blendMode;
                if (resource.SoftMaskEnabled.HasValue) {
                    hasSoftMask = resource.SoftMaskEnabled == true && resource.SoftMask != null;
                }
                hasUnsupportedEffect = hasUnsupportedEffect ||
                    resource.HasUnsupportedBlendMode ||
                    resource.HasUnsupportedSoftMask ||
                    resource.HasUnsupportedTextRestampEffect;
                if (resource.RenderingIntent.HasValue) ApplyRenderingIntent(resource.RenderingIntent.Value);
            }
        }
        void ApplyRenderingIntent(OfficeIccRenderingIntent intent) {
            renderingIntent = intent;
            if (fillColorSelection != null && fillColorSelection.TryConvert(intent, out OfficeColor selectedFill)) {
                fillColor = selectedFill;
                fillColorSpace = fillColorSelection.ColorSpace;
                fillColorResolved = true;
            }
            if (strokeColorSelection != null && strokeColorSelection.TryConvert(intent, out OfficeColor selectedStroke)) {
                strokeColor = selectedStroke;
                strokeColorSpace = strokeColorSelection.ColorSpace;
            }
        }
        OfficeColor ApplyTextOpacity(OfficeColor color, bool useStrokePaint) {
            double? opacity = useStrokePaint ? strokeOpacity : fillOpacity;
            if (!opacity.HasValue) {
                return color;
            }

            return OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * Clamp01(opacity.Value)));
        }
        bool TryApplyFillColor(PdfPageColorSpace colorSpace, out OfficeColor color) =>
            PdfPaintColorSelection.TryCreate(args, colorSpace, renderingIntent, out fillColorSelection, out color, outputIntentColorTransform);
        bool TryApplyStrokeColor(PdfPageColorSpace colorSpace, out OfficeColor color) =>
            PdfPaintColorSelection.TryCreate(args, colorSpace, renderingIntent, out strokeColorSelection, out color, outputIntentColorTransform);
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
        static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
        static int ReadTextRenderingMode(double value) {
            int mode = (int)Math.Round(value);
            return mode < 0 || mode > 7 ? 0 : mode;
        }
        static bool UsesStrokeTextPaint(int renderingMode) =>
            renderingMode is 1 or 2 or 5 or 6;

        static bool UsesFillTextPaint(int renderingMode) =>
            renderingMode is 0 or 2 or 4 or 6;

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
        bool initialUnsupportedEffect = false,
        System.Func<string, int?>? mcidForProperty = null,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands,
        PdfTextClippingBudget? textClippingBudget = null,
        OfficeIccRenderingIntent initialRenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        PdfPaintColorSelection? initialFillColorSelection = null,
        PdfPaintColorSelection? initialStrokeColorSelection = null,
        PdfOutputIntentColorTransform? outputIntentColorTransform = null,
        Func<string, int>? inlineImageComponentCount = null,
        Func<PdfArray, int>? inlineImageArrayComponentCount = null,
        Action? cancellationCheck = null,
        PdfTextStateSnapshot? initialTextState = null) {
        textClippingBudget ??= new PdfTextClippingBudget();
        var invocations = new List<FormInvocation>();
        Matrix2D ctm = Matrix2D.Identity;
        OfficeColor fillColor = initialFillColor ?? OfficeColor.Black;
        PdfPageColorSpace fillColorSpace = initialFillColorSpace;
        OfficeColor strokeColor = initialStrokeColor ?? OfficeColor.Black;
        PdfPageColorSpace strokeColorSpace = initialStrokeColorSpace;
        double? fillOpacity = initialFillOpacity;
        double? strokeOpacity = initialStrokeOpacity;
        PdfTextStateSnapshot startingTextState = initialTextState ??
            PdfTextStateSnapshot.Default.WithTextRenderingMode(initialTextRenderingMode);
        PdfTextStateSnapshot textState = startingTextState;
        int textRenderingMode = ReadTextRenderingMode(startingTextState.TextRenderingMode);
        PdfPageClipPath? clipPath = initialClipPath;
        bool hasUnsupportedEffect = initialUnsupportedEffect;
        bool fillColorResolved = initialFillColorSpace.Kind != PdfPageColorSpaceKind.Pattern;
        OfficeIccRenderingIntent renderingIntent = initialRenderingIntent;
        PdfPaintColorSelection? fillColorSelection = initialFillColorSelection;
        PdfPaintColorSelection? strokeColorSelection = initialStrokeColorSelection;
        if (fillColorSelection != null && fillColorSelection.TryConvert(renderingIntent, out OfficeColor selectedFillColor)) {
            fillColor = selectedFillColor;
            fillColorSpace = fillColorSelection.ColorSpace;
        } else if (!initialFillColor.HasValue && outputIntentColorTransform != null &&
            PdfPaintColorSelection.TryCreateDefaultBlack(renderingIntent, outputIntentColorTransform, out fillColorSelection, out OfficeColor defaultFillColor)) {
            fillColor = defaultFillColor;
            fillColorSpace = PdfPageColorSpaceKind.DeviceGray;
        }
        if (strokeColorSelection != null && strokeColorSelection.TryConvert(renderingIntent, out OfficeColor selectedStrokeColor)) {
            strokeColor = selectedStrokeColor;
            strokeColorSpace = strokeColorSelection.ColorSpace;
        } else if (!initialStrokeColor.HasValue && outputIntentColorTransform != null &&
            PdfPaintColorSelection.TryCreateDefaultBlack(renderingIntent, outputIntentColorTransform, out strokeColorSelection, out OfficeColor defaultStrokeColor)) {
            strokeColor = defaultStrokeColor;
            strokeColorSpace = PdfPageColorSpaceKind.DeviceGray;
        }
        OfficeColor effectiveInitialFillColor = fillColor;
        PdfPageColorSpace effectiveInitialFillColorSpace = fillColorSpace;
        OfficeColor effectiveInitialStrokeColor = strokeColor;
        PdfPageColorSpace effectiveInitialStrokeColorSpace = strokeColorSpace;
        PdfPaintColorSelection? effectiveInitialFillColorSelection = fillColorSelection;
        PdfPaintColorSelection? effectiveInitialStrokeColorSelection = strokeColorSelection;
        var clipPathBuilder = new PdfPageClipPathBuilder(pageHeight);
        var gstack = new Stack<TextGraphicsState>();
        var textStateStack = new Stack<PdfTextStateSnapshot>();
        var hiddenContentStack = new Stack<bool>();
        var unsupportedRestampContentStack = new Stack<bool>();
        var args = new List<object>(8);

        PdfContentStreamInterpreter.Interpret(content, maxOperations, operation => {
            cancellationCheck?.Invoke();
            args.Clear();
            args.AddRange(operation.Operands);
            double paintOrder = GetPaintOrder(operation.OperatorOffset);
            string op = operation.Name;
            switch (op) {
                case "q":
                    gstack.Push(new TextGraphicsState(ctm, string.Empty, 0D, 0D, 0D, 0D, 1D, 0D, fillColor, fillColorSpace, strokeColor, strokeColorSpace, fillOpacity, strokeOpacity, textRenderingMode, clipPath, hasUnsupportedEffect: hasUnsupportedEffect, fillColorResolved: fillColorResolved, renderingIntent: renderingIntent, fillColorSelection: fillColorSelection, strokeColorSelection: strokeColorSelection));
                    textStateStack.Push(textState);
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
                        hasUnsupportedEffect = state.HasUnsupportedEffect;
                        fillColorResolved = state.FillColorResolved;
                        renderingIntent = state.RenderingIntent;
                        fillColorSelection = state.FillColorSelection;
                        strokeColorSelection = state.StrokeColorSelection;
                        textState = textStateStack.Count > 0 ? textStateStack.Pop() : startingTextState;
                    } else {
                        ctm = Matrix2D.Identity;
                        fillColor = effectiveInitialFillColor;
                        fillColorSpace = effectiveInitialFillColorSpace;
                        strokeColor = effectiveInitialStrokeColor;
                        strokeColorSpace = effectiveInitialStrokeColorSpace;
                        fillOpacity = initialFillOpacity;
                        strokeOpacity = initialStrokeOpacity;
                        textState = startingTextState;
                        textRenderingMode = ReadTextRenderingMode(startingTextState.TextRenderingMode);
                        clipPath = initialClipPath;
                        hasUnsupportedEffect = initialUnsupportedEffect;
                        fillColorResolved = effectiveInitialFillColorSpace.Kind != PdfPageColorSpaceKind.Pattern;
                        renderingIntent = initialRenderingIntent;
                        fillColorSelection = effectiveInitialFillColorSelection;
                        strokeColorSelection = effectiveInitialStrokeColorSelection;
                    }

                    args.Clear();
                    break;
                case "Tf" when args.Count >= 2:
                    textState = textState.WithFont(
                        ToName(args[args.Count - 2]),
                        ToDouble(args[args.Count - 1]));
                    args.Clear();
                    break;
                case "Tc" when args.Count >= 1:
                    textState = textState.WithCharacterSpacing(ToDouble(args[args.Count - 1]));
                    args.Clear();
                    break;
                case "Tw" when args.Count >= 1:
                    textState = textState.WithWordSpacing(ToDouble(args[args.Count - 1]));
                    args.Clear();
                    break;
                case "Tz" when args.Count >= 1:
                    textState = textState.WithHorizontalScaling(ToDouble(args[args.Count - 1]) / 100D);
                    args.Clear();
                    break;
                case "TL" when args.Count >= 1:
                    textState = textState.WithLeading(ToDouble(args[args.Count - 1]));
                    args.Clear();
                    break;
                case "TD" when args.Count >= 2:
                    textState = textState.WithLeading(-ToDouble(args[args.Count - 1]));
                    args.Clear();
                    break;
                case "Ts" when args.Count >= 1:
                    textState = textState.WithTextRise(ToDouble(args[args.Count - 1]));
                    args.Clear();
                    break;
                case "\"" when args.Count >= 3:
                    textState = textState
                        .WithWordSpacing(ToDouble(args[args.Count - 3]))
                        .WithCharacterSpacing(ToDouble(args[args.Count - 2]));
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
                        clipPath = textClippingBudget.ResolveActiveClip(clipPath, parsedClipPath);
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
                case "ri":
                    if (args.Count == 1 && args[0] is string renderingIntentName) {
                        ApplyRenderingIntent(PdfRenderingIntentResolver.FromName(renderingIntentName));
                    } else {
                        hasUnsupportedEffect = true;
                    }
                    args.Clear();
                    break;
                case "cs":
                    if (args.Count >= 1 && TryReadColorSpace(ToName(args[args.Count - 1]), out PdfPageColorSpace parsedColorSpace)) {
                        fillColorSpace = parsedColorSpace;
                        fillColorResolved = parsedColorSpace.Kind != PdfPageColorSpaceKind.Pattern;
                        fillColorSelection = null;
                    } else {
                        fillColorSpace = PdfPageColorSpaceKind.Pattern;
                        fillColorResolved = false;
                    }

                    args.Clear();
                    break;
                case "CS":
                    if (args.Count >= 1 && TryReadColorSpace(ToName(args[args.Count - 1]), out PdfPageColorSpace parsedStrokeColorSpace)) {
                        strokeColorSpace = parsedStrokeColorSpace;
                        strokeColorSelection = null;
                    }

                    args.Clear();
                    break;
                case "rg":
                    if (args.Count >= 3) {
                        SetDirectFillColor(PdfPageColorSpaceKind.DeviceRgb);
                        fillColorResolved = true;
                    }

                    args.Clear();
                    break;
                case "RG":
                    if (args.Count >= 3) {
                        SetDirectStrokeColor(PdfPageColorSpaceKind.DeviceRgb);
                    }

                    args.Clear();
                    break;
                case "g":
                    if (args.Count >= 1) {
                        SetDirectFillColor(PdfPageColorSpaceKind.DeviceGray);
                        fillColorResolved = true;
                    }

                    args.Clear();
                    break;
                case "G":
                    if (args.Count >= 1) {
                        SetDirectStrokeColor(PdfPageColorSpaceKind.DeviceGray);
                    }

                    args.Clear();
                    break;
                case "k":
                    if (args.Count >= 4) {
                        SetDirectFillColor(PdfPageColorSpaceKind.DeviceCmyk);
                        fillColorResolved = true;
                    }

                    args.Clear();
                    break;
                case "K":
                    if (args.Count >= 4) {
                        SetDirectStrokeColor(PdfPageColorSpaceKind.DeviceCmyk);
                    }

                    args.Clear();
                    break;
                case "sc":
                case "scn":
                    fillColorResolved = PdfPaintColorSelection.TryCreate(args, fillColorSpace, renderingIntent, out fillColorSelection, out OfficeColor parsedFillColor, outputIntentColorTransform);
                    if (fillColorResolved) fillColor = parsedFillColor;

                    args.Clear();
                    break;
                case "SC":
                case "SCN":
                    if (PdfPaintColorSelection.TryCreate(args, strokeColorSpace, renderingIntent, out strokeColorSelection, out OfficeColor parsedStrokeColor, outputIntentColorTransform)) strokeColor = parsedStrokeColor;

                    args.Clear();
                    break;
                case "Tr":
                    if (args.Count >= 1) {
                        textRenderingMode = ReadTextRenderingMode(ToDouble(args[args.Count - 1]));
                        textState = textState.WithTextRenderingMode(textRenderingMode);
                    }

                    args.Clear();
                    break;
                case "Do":
                    if (!HasHiddenContent() && args.Count >= 1) {
                        string name = ToName(args[args.Count - 1]);
                        if (!string.IsNullOrEmpty(name)) {
                            invocations.Add(new FormInvocation(
                                name,
                                ctm,
                                paintOrder,
                                fillColor,
                                fillColorSpace,
                                strokeColor,
                                strokeColorSpace,
                                fillOpacity,
                                strokeOpacity,
                                textRenderingMode,
                                clipPath,
                                operation.OperatorOffset,
                                hasUnsupportedEffect || HasUnsupportedRestampContent(),
                                fillColorResolved,
                                renderingIntent,
                                fillColorSelection,
                                strokeColorSelection,
                                textState));
                        }
                    }
                    args.Clear();
                    break;
                case "BDC":
                    hiddenContentStack.Push(
                        operation.HasInvalidOperands ||
                        IsHiddenOptionalContent(
                            args.Count > 1 ? args[args.Count - 2] : null,
                            args.Count > 0 ? args[args.Count - 1] : null));
                    unsupportedRestampContentStack.Push(
                        operation.HasInvalidOperands ||
                        IsOptionalContentTag(args.Count > 1 ? args[args.Count - 2] : null) ||
                        GetMcid(args.Count > 0 ? args[args.Count - 1] : null).HasValue);
                    args.Clear();
                    break;
                case "BMC":
                    hiddenContentStack.Push(operation.HasInvalidOperands);
                    unsupportedRestampContentStack.Push(operation.HasInvalidOperands);
                    args.Clear();
                    break;
                case "EMC":
                    if (hiddenContentStack.Count > 0) {
                        hiddenContentStack.Pop();
                    }
                    if (unsupportedRestampContentStack.Count > 0) unsupportedRestampContentStack.Pop();

                    args.Clear();
                    break;
                default:
                    args.Clear();
                    break;
            }
        }, inlineImageComponentCount: inlineImageComponentCount, maxNestingDepth: maxNestingDepth, maxOperands: maxOperands, inlineImageArrayComponentCount: inlineImageArrayComponentCount);

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

        bool HasUnsupportedRestampContent() {
            foreach (bool unsupported in unsupportedRestampContentStack) if (unsupported) return true;
            return false;
        }

        int? GetMcid(object? propertyObject) {
            if (propertyObject is string propertyName) return mcidForProperty?.Invoke(propertyName);
            if (propertyObject is not PdfContentDictionary dictionary || !dictionary.Items.TryGetValue("MCID", out object? value)) return null;
            return TryGetMcid(value);
        }

        static bool IsOptionalContentTag(object? tag) =>
            tag is string name && string.Equals(name, "OC", StringComparison.Ordinal);

        bool IsHiddenOptionalContent(object? tag, object? property) =>
            tag is string tagName &&
            string.Equals(tagName, "OC", StringComparison.Ordinal) &&
            ((property is string propertyName &&
                optionalContentVisibility?.IsHidden(propertyName) == true) ||
             (property is PdfInlineOptionalContentReferences references &&
                optionalContentVisibility?.IsHidden(references) == true) ||
             (property is PdfContentDictionary dictionary &&
                dictionary.OptionalContentReferences is not null &&
                optionalContentVisibility?.IsHidden(dictionary.OptionalContentReferences) == true));

        void ApplyGraphicsStateResource(string name) {
            if (graphicsStates == null || !graphicsStates.TryGetValue(name, out PdfPageGraphicsStateResource resource)) {
                return;
            }

            fillOpacity = resource.FillOpacity ?? fillOpacity;
            strokeOpacity = resource.StrokeOpacity ?? strokeOpacity;
            hasUnsupportedEffect = hasUnsupportedEffect ||
                resource.HasUnsupportedBlendMode ||
                resource.HasUnsupportedSoftMask ||
                resource.HasUnsupportedTextRestampEffect ||
                resource.BlendMode is OfficeBlendMode mode && mode != OfficeBlendMode.Normal ||
                resource.SoftMask != null;
            if (resource.RenderingIntent.HasValue) ApplyRenderingIntent(resource.RenderingIntent.Value);
        }

        void ApplyRenderingIntent(OfficeIccRenderingIntent intent) {
            renderingIntent = intent;
            if (fillColorSelection != null && fillColorSelection.TryConvert(intent, out OfficeColor convertedFill)) fillColor = convertedFill;
            if (strokeColorSelection != null && strokeColorSelection.TryConvert(intent, out OfficeColor convertedStroke)) strokeColor = convertedStroke;
        }

        void SetDirectFillColor(PdfPageColorSpace colorSpace) {
            fillColorSelection = null;
            if (PdfPaintColorSelection.TryCreate(args, colorSpace, renderingIntent, out PdfPaintColorSelection? selection, out OfficeColor color, outputIntentColorTransform)) {
                fillColorSelection = selection;
                fillColor = color;
                fillColorSpace = colorSpace;
            }
        }

        void SetDirectStrokeColor(PdfPageColorSpace colorSpace) {
            strokeColorSelection = null;
            if (PdfPaintColorSelection.TryCreate(args, colorSpace, renderingIntent, out PdfPaintColorSelection? selection, out OfficeColor color, outputIntentColorTransform)) {
                strokeColorSelection = selection;
                strokeColor = color;
                strokeColorSpace = colorSpace;
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

        static int ReadTextRenderingMode(double value) {
            int mode = (int)Math.Round(value);
            return mode < 0 || mode > 7 ? 0 : mode;
        }

        static double ToDouble(object o) => o is double d ? d : 0.0;
        static string ToName(object o) => o as string ?? string.Empty;
    }

    private static int? TryGetMcid(object? value) {
        switch (value) {
            case int integer when integer >= 0:
                return integer;
            case long integer when integer >= 0 && integer <= int.MaxValue:
                return (int)integer;
            case double number when number >= 0D && number <= int.MaxValue && Math.Truncate(number) == number:
                return (int)number;
            default:
                return null;
        }
    }
}
