using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private bool TryReadNote(IElement token) {
            string? kindValue = GetAttribute(token, "data-officeimo-rtf-note");
            if (string.IsNullOrWhiteSpace(kindValue) || !TryParseNoteKind(kindValue!, out RtfNoteKind kind) || _lastRun == null) {
                return false;
            }

            var note = new RtfNote(kind) {
                Id = GetAttribute(token, "data-officeimo-rtf-note-id"),
                Author = GetAttribute(token, "data-officeimo-rtf-note-author")
            };

            string? created = GetAttribute(token, "data-officeimo-rtf-note-created");
            if (!string.IsNullOrWhiteSpace(created) &&
                DateTime.TryParse(created, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime timestamp)) {
                note.Created = timestamp;
            }

            AddNoteContent(note, GetAttribute(token, "data-officeimo-rtf-note-content"));
            _lastRun.Note = note;
            return true;
        }

        private static bool TryParseNoteKind(string value, out RtfNoteKind kind) {
            switch (value.Trim().ToLowerInvariant()) {
                case "annotation":
                    kind = RtfNoteKind.Annotation;
                    return true;
                case "endnote":
                    kind = RtfNoteKind.Endnote;
                    return true;
                case "footnote":
                    kind = RtfNoteKind.Footnote;
                    return true;
                default:
                    kind = RtfNoteKind.Footnote;
                    return false;
            }
        }

        private void AddNoteContent(RtfNote note, string? encodedContent) {
            string? html = DecodeNoteContent(encodedContent);
            if (string.IsNullOrEmpty(html)) {
                note.AddParagraph();
                return;
            }

            RtfDocument noteDocument = html!.LoadFromHtml();
            foreach (RtfParagraph paragraph in noteDocument.Paragraphs) {
                RtfParagraph noteParagraph = note.AddParagraph();
                CopyParagraphInlines(paragraph, noteParagraph, noteDocument);
            }

            if (note.Paragraphs.Count == 0) {
                note.AddParagraph();
            }
        }

        private static string? DecodeNoteContent(string? encodedContent) {
            if (string.IsNullOrWhiteSpace(encodedContent)) {
                return null;
            }

            try {
                return Encoding.UTF8.GetString(Convert.FromBase64String(encodedContent!));
            } catch (FormatException) {
                return null;
            }
        }

        private void CopyParagraphInlines(RtfParagraph source, RtfParagraph target, RtfDocument sourceDocument) {
            foreach (IRtfInline inline in source.Inlines) {
                CopyInline(inline, target, sourceDocument);
            }
        }

        private void CopyInline(IRtfInline inline, RtfParagraph target, RtfDocument sourceDocument) {
            switch (inline) {
                case RtfRun run:
                    CopyRun(run, target.AddText(run.Text), sourceDocument);
                    break;
                case RtfBreak rtfBreak:
                    target.AddBreak(rtfBreak.Kind);
                    break;
                case RtfBookmarkMarker marker when marker.Kind == RtfBookmarkMarkerKind.Start:
                    target.AddBookmarkStart(marker.Name);
                    break;
                case RtfBookmarkMarker marker:
                    target.AddBookmarkEnd(marker.Name);
                    break;
                case RtfField field:
                    RtfField copiedField = target.AddField(field.Instruction);
                    copiedField.FormFieldData = field.FormFieldData;
                    CopyParagraphInlines(field.Result, copiedField.Result, sourceDocument);
                    break;
                case RtfImage image:
                    RtfImage copiedImage = target.AddImage(image.Format, image.Data);
                    copiedImage.Description = image.Description;
                    copiedImage.SourceWidth = image.SourceWidth;
                    copiedImage.SourceHeight = image.SourceHeight;
                    copiedImage.DesiredWidthTwips = image.DesiredWidthTwips;
                    copiedImage.DesiredHeightTwips = image.DesiredHeightTwips;
                    break;
                case RtfObject rtfObject:
                    RtfObject copiedObject = target.AddObject(rtfObject.Kind, rtfObject.Data);
                    copiedObject.ClassName = rtfObject.ClassName;
                    copiedObject.Name = rtfObject.Name;
                    copiedObject.Width = rtfObject.Width;
                    copiedObject.Height = rtfObject.Height;
                    copiedObject.ScaleX = rtfObject.ScaleX;
                    copiedObject.ScaleY = rtfObject.ScaleY;
                    copiedObject.ResultImage = rtfObject.ResultImage;
                    CopyParagraphInlines(rtfObject.Result, copiedObject.Result, sourceDocument);
                    break;
                case RtfShape shape:
                    RtfShape copiedShape = target.AddShape();
                    foreach (RtfShapeInstruction instruction in shape.Instructions) {
                        copiedShape.AddInstruction(instruction.Name, instruction.Parameter, instruction.HasParameter);
                    }

                    foreach (RtfShapeProperty property in shape.Properties) {
                        copiedShape.AddProperty(property.Name, property.Value);
                    }

                    foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) {
                        RtfParagraph copiedParagraph = copiedShape.AddTextBoxParagraph();
                        CopyParagraphInlines(paragraph, copiedParagraph, sourceDocument);
                    }

                    break;
            }
        }

        private void CopyRun(RtfRun source, RtfRun target, RtfDocument sourceDocument) {
            target.Bold = source.Bold;
            target.Italic = source.Italic;
            target.UnderlineStyle = source.UnderlineStyle;
            target.Strike = source.Strike;
            target.DoubleStrike = source.DoubleStrike;
            target.Hidden = source.Hidden;
            target.Outline = source.Outline;
            target.Shadow = source.Shadow;
            target.Emboss = source.Emboss;
            target.Imprint = source.Imprint;
            target.CapsStyle = source.CapsStyle;
            target.VerticalPosition = source.VerticalPosition;
            target.FontSize = source.FontSize;
            target.FontId = RemapFontId(source.FontId, sourceDocument);
            target.ForegroundColorIndex = RemapColorIndex(source.ForegroundColorIndex, sourceDocument);
            target.HighlightColorIndex = RemapColorIndex(source.HighlightColorIndex, sourceDocument);
            target.CharacterBackgroundColorIndex = RemapColorIndex(source.CharacterBackgroundColorIndex, sourceDocument);
            target.CharacterShadingForegroundColorIndex = RemapColorIndex(source.CharacterShadingForegroundColorIndex, sourceDocument);
            target.CharacterShadingPatternPercent = source.CharacterShadingPatternPercent;
            target.CharacterShadingPattern = source.CharacterShadingPattern;
            target.CharacterBorder.Style = source.CharacterBorder.Style;
            target.CharacterBorder.Width = source.CharacterBorder.Width;
            target.CharacterBorder.ColorIndex = RemapColorIndex(source.CharacterBorder.ColorIndex, sourceDocument);
            target.UnderlineColorIndex = RemapColorIndex(source.UnderlineColorIndex, sourceDocument);
            target.CharacterSpacingTwips = source.CharacterSpacingTwips;
            target.CharacterScalePercent = source.CharacterScalePercent;
            target.KerningHalfPoints = source.KerningHalfPoints;
            target.CharacterOffsetHalfPoints = source.CharacterOffsetHalfPoints;
            target.StyleId = source.StyleId;
            target.Direction = source.Direction;
            target.LanguageId = source.LanguageId;
            target.Hyperlink = source.Hyperlink;
            target.Note = source.Note;
            target.RevisionKind = source.RevisionKind;
            target.RevisionAuthorIndex = source.RevisionAuthorIndex;
            target.RevisionTimestampValue = source.RevisionTimestampValue;
            target.CharacterRevisionSaveId = source.CharacterRevisionSaveId;
            target.InsertionRevisionSaveId = source.InsertionRevisionSaveId;
            target.DeletionRevisionSaveId = source.DeletionRevisionSaveId;
        }

        private int? RemapFontId(int? sourceFontId, RtfDocument sourceDocument) {
            if (!sourceFontId.HasValue) {
                return null;
            }

            RtfFont? sourceFont = sourceDocument.Fonts.FirstOrDefault(font => font.Id == sourceFontId.Value);
            return sourceFont == null ? sourceFontId : _document.AddFont(sourceFont.Name);
        }

        private int? RemapColorIndex(int? sourceColorIndex, RtfDocument sourceDocument) {
            if (!sourceColorIndex.HasValue || sourceColorIndex.Value <= 0 || sourceColorIndex.Value > sourceDocument.Colors.Count) {
                return sourceColorIndex;
            }

            RtfColor color = sourceDocument.Colors[sourceColorIndex.Value - 1];
            return GetOrAddColorIndex(color);
        }
    }
}
