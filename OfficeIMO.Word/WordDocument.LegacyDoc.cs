using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Loads a legacy binary `.doc` document and projects supported content into a normal OfficeIMO Word document.
        /// The resulting document saves through the normal Open XML path.
        /// </summary>
        public static WordDocument LoadLegacyDoc(string path, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(path, options);
            return ProjectLoadedLegacyDocDocument(document, path);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` document and returns both the projected OfficeIMO document and import report.
        /// </summary>
        public static LegacyDocLoadResult LoadLegacyDocWithReport(string path, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(path, options);
            return CreateLegacyDocLoadResult(document, path);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` stream and projects supported content into a normal OfficeIMO Word document.
        /// The resulting document saves through the normal Open XML path.
        /// </summary>
        public static WordDocument LoadLegacyDoc(Stream stream, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(stream, options);
            return ProjectLoadedLegacyDocDocument(document, sourcePath: null);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` stream and returns both the projected OfficeIMO document and import report.
        /// </summary>
        public static LegacyDocLoadResult LoadLegacyDocWithReport(Stream stream, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(stream, options);
            return CreateLegacyDocLoadResult(document, sourcePath: null);
        }

        private static WordDocument LoadLegacyDocFromNormalFlow(byte[] bytes, string? sourcePath, bool autoSave) {
            if (autoSave) {
                throw new NotSupportedException("Auto-save is not supported when loading legacy binary .doc files. Load the document, then save explicitly to a .docx path.");
            }

            LegacyDocDocument document = LegacyDocDocument.Load(bytes, new LegacyDocImportOptions());
            LegacyDocImportDiagnostic[] errors = document.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy DOC import failed: " + FormatLegacyDocDiagnostics(errors));
            }

            return ProjectLoadedLegacyDocDocument(document, sourcePath);
        }

        private static LegacyDocLoadResult CreateLegacyDocLoadResult(LegacyDocDocument legacyDocument, string? sourcePath) {
            try {
                return new LegacyDocLoadResult(ProjectLoadedLegacyDocDocument(legacyDocument, sourcePath), legacyDocument);
            } catch (InvalidDataException exception) {
                return new LegacyDocLoadResult(document: null, legacyDocument, exception);
            }
        }

        private static WordDocument ProjectLoadedLegacyDocDocument(LegacyDocDocument legacyDocument, string? sourcePath) {
            LegacyDocImportDiagnostic[] errors = legacyDocument.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy DOC import failed: " + FormatLegacyDocDiagnostics(errors));
            }

            WordDocument document = CreateInternal(filePath: null, stream: null, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, autoSave: false);
            ApplyLegacyDocProperties(document, legacyDocument.DocumentProperties);
            WordSection section = document.Sections.Count > 0
                ? document.Sections[0]
                : new WordSection(document, null!, null!);

            if (legacyDocument.BodyBlocks.Count == 0) {
                section.AddParagraph();
            } else {
                foreach (LegacyDocBodyBlock block in legacyDocument.BodyBlocks) {
                    if (block is LegacyDocParagraphBlock paragraphBlock) {
                        AddLegacyDocParagraph(section, paragraphBlock.Runs, paragraphBlock.Format);
                    } else if (block is LegacyDocTableBlock tableBlock) {
                        AddLegacyDocTable(section, tableBlock);
                    }
                }
            }

            document.MarkLoadedFromLegacyDoc(sourcePath, legacyDocument);
            return document;
        }

        private static void AddLegacyDocTable(WordSection section, LegacyDocTableBlock tableBlock) {
            int rowCount = tableBlock.Rows.Count;
            int columnCount = tableBlock.Rows.Count == 0
                ? 0
                : tableBlock.Rows.Max(row => row.Cells.Count);
            if (rowCount == 0 || columnCount == 0) {
                return;
            }

            WordTable table = section.AddTable(rowCount, columnCount, WordTableStyle.TableGrid);
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < sourceRow.Cells.Count && columnIndex < columnCount; columnIndex++) {
                    table.Rows[rowIndex].Cells[columnIndex].AddParagraph(sourceRow.Cells[columnIndex].Text, removeExistingParagraphs: true);
                }
            }
        }

        private static void AddLegacyDocParagraph(WordSection section, IReadOnlyList<LegacyDocTextRun> paragraphRuns, LegacyDocParagraphFormat paragraphFormat) {
            if (paragraphRuns.Count == 0) {
                ApplyLegacyDocParagraphFormatting(section.AddParagraph(), paragraphFormat);
                return;
            }

            WordParagraph paragraph = section.AddParagraph(string.Empty);
            ApplyLegacyDocParagraphFormatting(paragraph, paragraphFormat);
            foreach (LegacyDocTextRun legacyRun in paragraphRuns) {
                WordParagraph run = paragraph.AddText(legacyRun.Text);
                ApplyLegacyDocRunFormatting(run, legacyRun);
            }
        }

        private static void ApplyLegacyDocParagraphFormatting(WordParagraph paragraph, LegacyDocParagraphFormat paragraphFormat) {
            if (paragraphFormat.StyleIndex != null && TryMapBuiltInParagraphStyle(paragraphFormat.StyleIndex.Value, out WordParagraphStyles style)) {
                paragraph.SetStyle(style);
            }

            if (paragraphFormat.Alignment != null && TryMapParagraphAlignment(paragraphFormat.Alignment.Value, out JustificationValues alignment)) {
                paragraph.ParagraphAlignment = alignment;
            }

            if (paragraphFormat.SpacingBeforeTwips != null) {
                paragraph.LineSpacingBefore = paragraphFormat.SpacingBeforeTwips;
            }

            if (paragraphFormat.SpacingAfterTwips != null) {
                paragraph.LineSpacingAfter = paragraphFormat.SpacingAfterTwips;
            }

            if (paragraphFormat.LineSpacingTwips != null) {
                paragraph.LineSpacing = paragraphFormat.LineSpacingTwips;
            }

            if (paragraphFormat.LeftIndentTwips != null) {
                paragraph.IndentationBefore = paragraphFormat.LeftIndentTwips;
            }

            if (paragraphFormat.RightIndentTwips != null) {
                paragraph.IndentationAfter = paragraphFormat.RightIndentTwips;
            }

            if (paragraphFormat.FirstLineIndentTwips != null) {
                if (paragraphFormat.FirstLineIndentTwips.Value < 0) {
                    paragraph.IndentationHanging = -paragraphFormat.FirstLineIndentTwips.Value;
                } else {
                    paragraph.IndentationFirstLine = paragraphFormat.FirstLineIndentTwips;
                }
            }
        }

        private static bool TryMapBuiltInParagraphStyle(ushort styleIndex, out WordParagraphStyles style) {
            switch (styleIndex) {
                case 0:
                    style = WordParagraphStyles.Normal;
                    return true;
                case 1:
                    style = WordParagraphStyles.Heading1;
                    return true;
                case 2:
                    style = WordParagraphStyles.Heading2;
                    return true;
                case 3:
                    style = WordParagraphStyles.Heading3;
                    return true;
                case 4:
                    style = WordParagraphStyles.Heading4;
                    return true;
                case 5:
                    style = WordParagraphStyles.Heading5;
                    return true;
                case 6:
                    style = WordParagraphStyles.Heading6;
                    return true;
                case 7:
                    style = WordParagraphStyles.Heading7;
                    return true;
                case 8:
                    style = WordParagraphStyles.Heading8;
                    return true;
                case 9:
                    style = WordParagraphStyles.Heading9;
                    return true;
                default:
                    style = default;
                    return false;
            }
        }

        private static void ApplyLegacyDocRunFormatting(WordParagraph run, LegacyDocTextRun legacyRun) {
            if (legacyRun.Bold) {
                run.SetBold();
            }

            if (legacyRun.Italic) {
                run.SetItalic();
            }

            if (legacyRun.Underline != null && TryMapUnderline(legacyRun.Underline.Value, out UnderlineValues underline)) {
                run.Underline = underline;
            }

            if (legacyRun.FontSizeHalfPoints != null) {
                RunProperties runProperties = run._runProperties ?? new RunProperties();
                run._runProperties = runProperties;
                runProperties.FontSize = new FontSize {
                    Val = legacyRun.FontSizeHalfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                };
            }

            if (!string.IsNullOrEmpty(legacyRun.ColorHex)) {
                run.ColorHex = legacyRun.ColorHex!;
            }

            if (!string.IsNullOrEmpty(legacyRun.FontFamily)) {
                run.SetFontFamily(legacyRun.FontFamily!);
            }
        }

        private static bool TryMapUnderline(LegacyDocUnderlineKind underline, out UnderlineValues value) {
            switch (underline) {
                case LegacyDocUnderlineKind.Single:
                    value = UnderlineValues.Single;
                    return true;
                case LegacyDocUnderlineKind.Words:
                    value = UnderlineValues.Words;
                    return true;
                case LegacyDocUnderlineKind.Double:
                    value = UnderlineValues.Double;
                    return true;
                case LegacyDocUnderlineKind.Dotted:
                    value = UnderlineValues.Dotted;
                    return true;
                case LegacyDocUnderlineKind.Thick:
                    value = UnderlineValues.Thick;
                    return true;
                case LegacyDocUnderlineKind.Dash:
                    value = UnderlineValues.Dash;
                    return true;
                case LegacyDocUnderlineKind.DotDash:
                    value = UnderlineValues.DotDash;
                    return true;
                case LegacyDocUnderlineKind.DotDotDash:
                    value = UnderlineValues.DotDotDash;
                    return true;
                case LegacyDocUnderlineKind.Wave:
                    value = UnderlineValues.Wave;
                    return true;
                case LegacyDocUnderlineKind.DottedHeavy:
                    value = UnderlineValues.DottedHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashedHeavy:
                    value = UnderlineValues.DashedHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashDotHeavy:
                    value = UnderlineValues.DashDotHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashDotDotHeavy:
                    value = UnderlineValues.DashDotDotHeavy;
                    return true;
                case LegacyDocUnderlineKind.WavyHeavy:
                    value = UnderlineValues.WavyHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashLong:
                    value = UnderlineValues.DashLong;
                    return true;
                case LegacyDocUnderlineKind.WavyDouble:
                    value = UnderlineValues.WavyDouble;
                    return true;
                case LegacyDocUnderlineKind.DashLongHeavy:
                    value = UnderlineValues.DashLongHeavy;
                    return true;
                default:
                    value = default;
                    return false;
            }
        }

        private static bool TryMapParagraphAlignment(LegacyDocParagraphAlignment alignment, out JustificationValues value) {
            switch (alignment) {
                case LegacyDocParagraphAlignment.Left:
                    value = JustificationValues.Left;
                    return true;
                case LegacyDocParagraphAlignment.Center:
                    value = JustificationValues.Center;
                    return true;
                case LegacyDocParagraphAlignment.Right:
                    value = JustificationValues.Right;
                    return true;
                case LegacyDocParagraphAlignment.Justify:
                    value = JustificationValues.Both;
                    return true;
                default:
                    value = default;
                    return false;
            }
        }

        private static void ApplyLegacyDocProperties(WordDocument document, LegacyDocDocumentProperties properties) {
            if (!properties.HasAnyProperties) {
                return;
            }

            document.BuiltinDocumentProperties.Title = properties.Title;
            document.BuiltinDocumentProperties.Subject = properties.Subject;
            document.BuiltinDocumentProperties.Creator = properties.Creator;
            document.BuiltinDocumentProperties.Keywords = properties.Keywords;
            document.BuiltinDocumentProperties.Description = properties.Description;
            document.BuiltinDocumentProperties.Category = properties.Category;
            document.BuiltinDocumentProperties.LastModifiedBy = properties.LastModifiedBy;
            document.BuiltinDocumentProperties.Revision = properties.Revision;
            document.BuiltinDocumentProperties.Created = properties.Created;
            document.BuiltinDocumentProperties.Modified = properties.Modified;
            document.BuiltinDocumentProperties.LastPrinted = properties.LastPrinted;

            if (properties.Company != null) {
                document.ApplicationProperties.Company = properties.Company;
            }

            if (properties.Manager != null) {
                document.ApplicationProperties.Manager = new Manager { Text = properties.Manager };
            }

            foreach (KeyValuePair<string, LegacyDocDocumentPropertyValue> property in properties.CustomProperties) {
                if (TryCreateWordCustomProperty(property.Value, out WordCustomProperty? wordProperty)) {
                    document.CustomDocumentProperties[property.Key] = wordProperty!;
                }
            }
        }

        private static bool TryCreateWordCustomProperty(LegacyDocDocumentPropertyValue property, out WordCustomProperty? wordProperty) {
            switch (property.Kind) {
                case LegacyDocDocumentPropertyValueKind.Text:
                    wordProperty = new WordCustomProperty(Convert.ToString(property.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                    return true;
                case LegacyDocDocumentPropertyValueKind.Boolean:
                    wordProperty = new WordCustomProperty(Convert.ToBoolean(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.DateTime:
                    wordProperty = new WordCustomProperty(Convert.ToDateTime(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.Integer:
                    wordProperty = new WordCustomProperty(Convert.ToInt32(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.Number:
                    wordProperty = new WordCustomProperty(Convert.ToDouble(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                default:
                    wordProperty = null;
                    return false;
            }
        }

        private static string FormatLegacyDocDiagnostics(IEnumerable<LegacyDocImportDiagnostic> diagnostics) {
            const int maxDiagnostics = 6;
            LegacyDocImportDiagnostic[] selected = diagnostics.Take(maxDiagnostics + 1).ToArray();
            string message = string.Join("; ", selected.Take(maxDiagnostics).Select(diagnostic => diagnostic.ToString()));
            if (selected.Length > maxDiagnostics) {
                message += $"; and {selected.Length - maxDiagnostics} more";
            }

            return message;
        }
    }
}
