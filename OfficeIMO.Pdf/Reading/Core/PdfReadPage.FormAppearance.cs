namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal WidgetAppearanceScanBudget CreateWidgetAppearanceScanBudget() => new WidgetAppearanceScanBudget(this);

    internal bool DoesWidgetNormalAppearancePresentAllText(
        int widgetObjectNumber,
        IReadOnlyList<string> expectedValues,
        WidgetAppearanceScanBudget budget,
        double maximumTinyFontSizePoints) {
        if (!_objects.TryGetValue(widgetObjectNumber, out PdfIndirectObject? widgetObject) ||
            widgetObject.Value is not PdfDictionary widget ||
            !TryGetNormalAppearanceStream(widget, out PdfStream appearance)) {
            return false;
        }

        try {
            var spans = new List<PdfTextSpan>();
            if (!TryReadBox(
                    appearance.Dictionary.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null,
                    out (double X1, double Y1, double X2, double Y2) bbox) ||
                bbox.X2 <= bbox.X1 ||
                bbox.Y2 <= bbox.Y1) {
                return false;
            }
            PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
            PdfDictionary? appearanceResources = ResolveDictionary(
                appearance.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject)
                    ? resourcesObject
                    : null) ?? pageResources;
            PdfFontResourceSet fontResources = _fontResourceCache.GetOrCreate(appearanceResources, _objects);
            string content = WrapFormContentWithBoundingBoxClip(
                PdfEncoding.Latin1GetString(budget._pageContentBudget.Decode(appearance)),
                appearance.Dictionary);
            if (content.Length == 0) return false;

            CollectTextAndForms(
                content,
                appearanceResources,
                fontResources.Decoders,
                fontResources.WidthProviders,
                fontResources.Fonts,
                spans,
                new HashSet<PdfStream>(),
                GetPageSize().Height,
                includeArtifactText: true,
                textOutputBudget: budget._textOutputBudget,
                textClippingBudget: budget._textClippingBudget,
                pageContentBudget: budget._pageContentBudget);

            double pageHeight = GetPageSize().Height;
            string presentedText = string.Concat(spans
                .Where(span => span.IsVisible &&
                    span.Color?.A > 3 &&
                    Math.Abs(span.Advance) > 0.01D &&
                    span.FontSize > maximumTinyFontSizePoints &&
                    span.CanProjectCompleteText(pageHeight))
                .Select(static span => span.Text));
            for (int i = 0; i < expectedValues.Count; i++) {
                string expected = expectedValues[i];
                if (!string.IsNullOrEmpty(expected) &&
                    presentedText.IndexOf(expected, StringComparison.Ordinal) < 0) {
                    return false;
                }
            }
            return true;
        } catch (System.IO.InvalidDataException) {
            return false;
        }
    }

    internal bool DoesWidgetNormalAppearancePresentButtonState(
        int widgetObjectNumber,
        WidgetAppearanceScanBudget budget) {
        if (!_objects.TryGetValue(widgetObjectNumber, out PdfIndirectObject? widgetObject) ||
            widgetObject.Value is not PdfDictionary widget ||
            !TryGetNormalAppearanceStream(widget, out PdfStream appearance)) {
            return false;
        }

        try {
            string content = PdfEncoding.Latin1GetString(budget._pageContentBudget.Decode(appearance));
            bool hasPaint = false;
            int textRenderingMode = 0;
            var textRenderingModeStack = new Stack<int>();
            PdfContentStreamInterpreter.Interpret(
                content,
                _limits.MaxContentOperations,
                operation => {
                    if (hasPaint) return;
                    switch (operation.Name) {
                        case "q":
                            textRenderingModeStack.Push(textRenderingMode);
                            break;
                        case "Q":
                            textRenderingMode = textRenderingModeStack.Count > 0 ? textRenderingModeStack.Pop() : 0;
                            break;
                        case "Tr" when operation.Operands.Count > 0:
                            textRenderingMode = (int)Convert.ToDouble(
                                operation.Operands[operation.Operands.Count - 1],
                                System.Globalization.CultureInfo.InvariantCulture);
                            break;
                        case "Tj":
                        case "TJ":
                        case "'":
                        case "\"":
                            hasPaint = textRenderingMode != 3 && ContainsTextBytes(operation.Operands);
                            break;
                        case "S":
                        case "s":
                        case "f":
                        case "F":
                        case "f*":
                        case "B":
                        case "B*":
                        case "b":
                        case "b*":
                        case "Do":
                        case "sh":
                            hasPaint = true;
                            break;
                        case "BI" when operation.InlineImage is not null:
                            hasPaint = operation.InlineImage.Data.Length > 0;
                            break;
                    }
                },
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands);
            return hasPaint;
        } catch (System.IO.InvalidDataException) {
            return false;
        } catch (FormatException) {
            return false;
        } catch (InvalidCastException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool ContainsTextBytes(IEnumerable<object> operands) {
        foreach (object operand in operands) {
            if (operand is byte[] bytes && bytes.Length > 0) return true;
            if (operand is string text && text.Length > 0) return true;
            if (operand is IEnumerable<object> nested && ContainsTextBytes(nested)) return true;
        }
        return false;
    }

    internal sealed class WidgetAppearanceScanBudget {
        internal readonly PageContentBudget _pageContentBudget;
        internal readonly TextContentParser.TextOutputBudget _textOutputBudget;
        internal readonly PdfTextClippingBudget _textClippingBudget;

        internal WidgetAppearanceScanBudget(PdfReadPage page) {
            _pageContentBudget = new PageContentBudget(page);
            _textOutputBudget = new TextContentParser.TextOutputBudget(
                page._limits.MaxActualTextCharacters,
                page._limits.MaxDecodedTextCharacters);
            _textClippingBudget = new PdfTextClippingBudget();
        }
    }
}
