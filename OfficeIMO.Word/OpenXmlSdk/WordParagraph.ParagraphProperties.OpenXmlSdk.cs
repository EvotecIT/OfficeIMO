using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // The SDK side of paragraph formatting: what the members in WordParagraph.ParagraphProperties.cs do once the Open XML
    // SDK holds the document, after OpenXmlDocument or for packages the owned layer doesn't open.
    public partial class WordParagraph {
        private string? SdkStyleId => _paragraphProperties?.GetFirstChild<ParagraphStyleId>()?.Val;

        private WordParagraphAlignment? SdkParagraphAlignment {
            get => _paragraphProperties?.Justification?.Val?.Value.ToOfficeEnum();
            set {
                if (value == null) {
                    _paragraphProperties?.RemoveAllChildren<Justification>();
                    return;
                }
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.Justification = new Justification {
                    Val = value.Value.ToOpenXml()
                };
            }
        }

        private WordVerticalCharacterAlignment? SdkVerticalCharacterAlignment {
            get => _paragraphProperties?.TextAlignment?.Val?.Value.ToOfficeEnum();
            set {
                if (value == null) {
                    _paragraphProperties?.RemoveAllChildren<TextAlignment>();
                    return;
                }
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.TextAlignment = new TextAlignment { Val = value.Value.ToOpenXml() };
            }
        }

        private global::OfficeIMO.OfficeLength? GetSdkParagraphTwips(string element, string attribute) {
            string? text = element == "ind"
                ? attribute switch {
                    "left" => _paragraphProperties?.Indentation?.Left,
                    "right" => _paragraphProperties?.Indentation?.Right,
                    "firstLine" => _paragraphProperties?.Indentation?.FirstLine,
                    _ => _paragraphProperties?.Indentation?.Hanging
                }
                : attribute == "before" ? _paragraphProperties?.SpacingBetweenLines?.Before : _paragraphProperties?.SpacingBetweenLines?.After;
            return !string.IsNullOrEmpty(text) && int.TryParse(text, out int twips) ? global::OfficeIMO.OfficeLength.Points(twips / 20D) : null;
        }

        private void SetSdkParagraphTwips(string element, string attribute, global::OfficeIMO.OfficeLength? value) {
            var wordNativeValue = checked((int?)(value?.ToTwips()));
            var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
            if (element == "ind") {
                var indentation = props.Indentation ?? new Indentation();
                StringValue? text = wordNativeValue.HasValue ? new StringValue(wordNativeValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)) : null;
                switch (attribute) {
                    case "left": indentation.Left = text; break;
                    case "right": indentation.Right = text; break;
                    case "firstLine": indentation.FirstLine = text; break;
                    default: indentation.Hanging = text; break;
                }
                props.Indentation = indentation;
                return;
            }
            // Spacing goes through the string conversion, which writes an empty value when cleared.
            var spacing = props.SpacingBetweenLines ?? new SpacingBetweenLines();
            if (attribute == "before") spacing.Before = wordNativeValue?.ToString();
            else spacing.After = wordNativeValue?.ToString();
            props.SpacingBetweenLines = spacing;
        }

        private bool GetSdkParagraphOnOff(string element) {
            OnOffType? value = element switch {
                "pageBreakBefore" => _paragraphProperties?.PageBreakBefore,
                "keepNext" => _paragraphProperties?.KeepNext,
                "keepLines" => _paragraphProperties?.KeepLines,
                "widowControl" => _paragraphProperties?.WidowControl,
                _ => _paragraphProperties?.BiDi
            };
            return ReadParagraphOnOff(value);
        }

        private void SetSdkParagraphOnOff(string element, bool value) {
            var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
            switch (element) {
                case "pageBreakBefore":
                    props.PageBreakBefore = value ? new PageBreakBefore() : null;
                    break;
                case "keepNext":
                    if (value) {
                        props.KeepNext ??= new KeepNext();
                        props.KeepNext.Val = null;
                    } else {
                        props.KeepNext?.Remove();
                    }
                    break;
                case "keepLines":
                    if (value) {
                        props.KeepLines ??= new KeepLines();
                        props.KeepLines.Val = null;
                    } else {
                        props.KeepLines?.Remove();
                    }
                    break;
                case "widowControl":
                    if (value) {
                        props.WidowControl ??= new WidowControl();
                        props.WidowControl.Val = null;
                    } else {
                        props.WidowControl = new WidowControl { Val = false };
                    }
                    break;
                default:
                    if (value) {
                        props.BiDi ??= new BiDi();
                    } else {
                        props.BiDi?.Remove();
                    }
                    break;
            }
        }

        private WordTextDirection? SdkTextDirection {
            get => _paragraphProperties?.TextDirection?.Val?.Value.ToOfficeEnum();
            set {
                var textDirection = new TextDirection { Val = value?.ToOpenXml() };
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.TextDirection = textDirection;
            }
        }

        private void GetSdkLineAttributes(out string? line, out string? rule) {
            SpacingBetweenLines? spacing = _paragraph.GetFirstChild<ParagraphProperties>()?.SpacingBetweenLines;
            line = spacing?.Line?.Value;
            // The rule is read as written: EnumValue.Value throws on a token the SDK doesn't know.
            rule = spacing?.LineRule?.InnerText;
        }

        /// <summary>
        /// Sets or clears <c>w:spacing</c> line attributes. Clearing both leaves a paragraph without spacing as it is;
        /// anything else adds the paragraph properties and the spacing element when missing.
        /// </summary>
        private void SetSdkLineAttributes(bool setLine, string? line, bool setRule, string? rule) {
            if (setLine && setRule && line == null && rule == null) {
                var existing = _paragraph.GetFirstChild<ParagraphProperties>()?.SpacingBetweenLines;
                if (existing != null) { existing.Line = null; existing.LineRule = null; }
                return;
            }
            var props = _paragraph.ParagraphProperties ??= new ParagraphProperties();
            var spacing = props.SpacingBetweenLines ??= new SpacingBetweenLines();
            if (setLine) spacing.Line = line;
            if (setRule) spacing.LineRule = rule == null ? null : new EnumValue<LineSpacingRuleValues>(new LineSpacingRuleValues(rule));
        }

        private WordVerticalTextPosition? SdkVerticalTextAlignment {
            get => ScopedRunProperties?.VerticalTextAlignment?.Val?.Value.ToOfficeEnum();
            set {
                RunProperties? runProperties = ScopedRunProperties;
                if (value == null) {
                    if (runProperties?.VerticalTextAlignment == null) {
                        return;
                    }
                    runProperties.VerticalTextAlignment = null;
                } else {
                    if (IsHyperLink && _stdRun == null) {
                        WordHyperLink hyperlink = Hyperlink!;
                        runProperties = VerifyRunProperties(
                            hyperlink._hyperlink!,
                            hyperlink._run!,
                            hyperlink._runProperties);
                    } else {
                        runProperties = VerifyRunProperties();
                    }
                    runProperties.VerticalTextAlignment = new VerticalTextAlignment { Val = value.Value.ToOpenXml() };
                }
            }
        }

        private static bool ReadParagraphOnOff(OnOffType? value) =>
            value?.Val?.Value ?? value is not null;

        internal static bool HasPageBreakBefore(Paragraph paragraph) =>
            ReadParagraphOnOff(paragraph.GetFirstChild<ParagraphProperties>()?.PageBreakBefore);
    }
}
