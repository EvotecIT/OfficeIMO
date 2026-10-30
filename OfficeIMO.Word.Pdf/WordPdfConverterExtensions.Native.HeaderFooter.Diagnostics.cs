using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void ConfigureNativePageNumbering(PdfCore.PdfPageBuilder page, WordSection section) {
            W.PageNumberType? pageNumberType = section._sectionProperties.GetFirstChild<W.PageNumberType>();
            if (section.GetEffectivePageNumberStart() is int start) {
                page.PageNumberStart(start);
            }

            PdfCore.PdfPageNumberStyle? style = MapNativePageNumberStyle(pageNumberType?.Format?.Value);
            if (style.HasValue) {
                page.PageNumberStyle(style.Value);
            }
        }

        private static PdfCore.PdfPageNumberStyle? MapNativePageNumberStyle(W.NumberFormatValues? format) {
            if (format == W.NumberFormatValues.LowerRoman) {
                return PdfCore.PdfPageNumberStyle.LowerRoman;
            }

            if (format == W.NumberFormatValues.UpperRoman) {
                return PdfCore.PdfPageNumberStyle.UpperRoman;
            }

            if (format == W.NumberFormatValues.LowerLetter) {
                return PdfCore.PdfPageNumberStyle.LowerLetter;
            }

            if (format == W.NumberFormatValues.UpperLetter) {
                return PdfCore.PdfPageNumberStyle.UpperLetter;
            }

            if (format == W.NumberFormatValues.Decimal || format == W.NumberFormatValues.DecimalZero) {
                return PdfCore.PdfPageNumberStyle.Arabic;
            }

            return null;
        }

        private static void ConfigureNativeHeaderFooter(PdfCore.PdfPageBuilder page, WordSection section, WordPdfExportOptions? options, double headerMarginExpansion, double footerMarginExpansion, NativeFontMap nativeFontMap, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            bool hasEvenPageStory = section.DocumentOddEvenSettingEnabled;
            RecordNativeHeaderFooterDiagnostics(section.ResolveDefaultHeader(), options, "default header");
            if (section.DifferentFirstPage)
                RecordNativeHeaderFooterDiagnostics(section.ResolveFirstHeader(), options, "first header");
            if (hasEvenPageStory)
                RecordNativeHeaderFooterDiagnostics(section.ResolveEvenHeader(), options, "even header");
            RecordNativeHeaderFooterDiagnostics(section.ResolveDefaultFooter(), options, "default footer");
            if (section.DifferentFirstPage)
                RecordNativeHeaderFooterDiagnostics(section.ResolveFirstFooter(), options, "first footer");
            if (hasEvenPageStory)
                RecordNativeHeaderFooterDiagnostics(section.ResolveEvenFooter(), options, "even footer");
            ApplyNativeSectionWatermark(page, section, options);
            page.Options.UsePageNumberParityForHeaderFooter = true;

            // Flat projections still supply field-numbering metadata. Their
            // approximation diagnostics do not describe admitted running flow.
            WordPdfExportOptions? headerTextDiagnostics = UsesNativeRunningHeader(section, options) ? null : options;
            WordPdfExportOptions? footerTextDiagnostics = UsesNativeRunningFooter(section, options) ? null : options;
            RecordNativeRunningBarTabDiagnostics(section, options,
                headerTextDiagnostics == null, footerTextDiagnostics == null);
            RecordNativeRunningImageWrapDiagnostics(section, options,
                headerTextDiagnostics == null, footerTextDiagnostics == null);
            NativeHeaderFooterText? defaultHeader = GetNativeHeaderFooterText(section.ResolveDefaultHeader(), listMarkers, nativeFontMap, headerTextDiagnostics);
            NativeHeaderFooterText? firstHeader = section.DifferentFirstPage ? GetNativeHeaderFooterText(section.ResolveFirstHeader(), listMarkers, nativeFontMap, headerTextDiagnostics) : null;
            NativeHeaderFooterText? evenHeader = hasEvenPageStory ? GetNativeHeaderFooterText(section.ResolveEvenHeader(), listMarkers, nativeFontMap, headerTextDiagnostics) : null;
            NativeHeaderFooterText? defaultFooter = GetNativeHeaderFooterText(section.ResolveDefaultFooter(), listMarkers, nativeFontMap, footerTextDiagnostics);
            NativeHeaderFooterText? firstFooter = section.DifferentFirstPage ? GetNativeHeaderFooterText(section.ResolveFirstFooter(), listMarkers, nativeFontMap, footerTextDiagnostics) : null;
            NativeHeaderFooterText? evenFooter = hasEvenPageStory ? GetNativeHeaderFooterText(section.ResolveEvenFooter(), listMarkers, nativeFontMap, footerTextDiagnostics) : null;
            IReadOnlyList<NativeHeaderFooterImage> defaultHeaderImages = GetNativeHeaderFooterImages(section.ResolveDefaultHeader(), options, "default header image");
            IReadOnlyList<NativeHeaderFooterImage> firstHeaderImages = section.DifferentFirstPage ? GetNativeHeaderFooterImages(section.ResolveFirstHeader(), options, "first header image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterImage> evenHeaderImages = hasEvenPageStory ? GetNativeHeaderFooterImages(section.ResolveEvenHeader(), options, "even header image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterImage> defaultFooterImages = GetNativeHeaderFooterImages(section.ResolveDefaultFooter(), options, "default footer image");
            IReadOnlyList<NativeHeaderFooterImage> firstFooterImages = section.DifferentFirstPage ? GetNativeHeaderFooterImages(section.ResolveFirstFooter(), options, "first footer image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterImage> evenFooterImages = hasEvenPageStory ? GetNativeHeaderFooterImages(section.ResolveEvenFooter(), options, "even footer image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterShape> defaultHeaderShapes = GetNativeHeaderFooterShapes(section.ResolveDefaultHeader());
            IReadOnlyList<NativeHeaderFooterShape> firstHeaderShapes = section.DifferentFirstPage ? GetNativeHeaderFooterShapes(section.ResolveFirstHeader()) : Array.Empty<NativeHeaderFooterShape>();
            IReadOnlyList<NativeHeaderFooterShape> evenHeaderShapes = hasEvenPageStory ? GetNativeHeaderFooterShapes(section.ResolveEvenHeader()) : Array.Empty<NativeHeaderFooterShape>();
            IReadOnlyList<NativeHeaderFooterShape> defaultFooterShapes = GetNativeHeaderFooterShapes(section.ResolveDefaultFooter());
            IReadOnlyList<NativeHeaderFooterShape> firstFooterShapes = section.DifferentFirstPage ? GetNativeHeaderFooterShapes(section.ResolveFirstFooter()) : Array.Empty<NativeHeaderFooterShape>();
            IReadOnlyList<NativeHeaderFooterShape> evenFooterShapes = hasEvenPageStory ? GetNativeHeaderFooterShapes(section.ResolveEvenFooter()) : Array.Empty<NativeHeaderFooterShape>();
            PdfCore.PdfStandardFont? headerFont = ResolveNativeHeaderFooterFont(
                ResolveNativeHeaderFooterBaseFont(section._document, options, isHeader: true),
                nativeFontMap,
                section.ResolveDefaultHeader(),
                section.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            PdfCore.PdfStandardFont? footerFont = ResolveNativeHeaderFooterFont(
                ResolveNativeHeaderFooterBaseFont(section._document, options, isHeader: false),
                nativeFontMap,
                section.ResolveDefaultFooter(),
                section.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            string? headerFontFamily = ResolveNativeHeaderFooterFontFamily(
                nativeFontMap,
                section.ResolveDefaultHeader(),
                section.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            string? footerFontFamily = ResolveNativeHeaderFooterFontFamily(
                nativeFontMap,
                section.ResolveDefaultFooter(),
                section.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            global::OfficeIMO.OfficeColor? headerColor = ResolveNativeHeaderFooterColor(
                section.ResolveDefaultHeader(),
                section.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            global::OfficeIMO.OfficeColor? footerColor = ResolveNativeHeaderFooterColor(
                section.ResolveDefaultFooter(),
                section.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            double? headerFontSize = ResolveNativeHeaderFooterFontSize(
                section.ResolveDefaultHeader(),
                section.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            double? footerFontSize = ResolveNativeHeaderFooterFontSize(
                section.ResolveDefaultFooter(),
                section.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            ApplyNativeHeaderFooterPageNumberStyle(page, defaultHeader, firstHeader, evenHeader, defaultFooter, firstFooter, evenFooter);
            bool hasFirstHeaderVariant = section.DifferentFirstPage;
            bool hasEvenHeaderVariant = hasEvenPageStory;
            bool hasFirstFooterVariant = section.DifferentFirstPage;
            bool hasEvenFooterVariant = hasEvenPageStory;
            if (defaultHeader != null || hasFirstHeaderVariant || hasEvenHeaderVariant ||
                defaultHeaderImages.Count > 0 || firstHeaderImages.Count > 0 || evenHeaderImages.Count > 0 ||
                defaultHeaderShapes.Count > 0 || firstHeaderShapes.Count > 0 || evenHeaderShapes.Count > 0) {
                page.Header(header => {
                    if (headerMarginExpansion > 0D) {
                        header.Offset(GetNativeHeaderOffset(options, headerMarginExpansion));
                    }

                    if (headerFont.HasValue) {
                        header.Font(headerFont.Value);
                    }

                    if (!string.IsNullOrWhiteSpace(headerFontFamily)) {
                        header.FontFamily(headerFontFamily!);
                    }

                    if (headerColor.HasValue) {
                        header.Color(headerColor.Value);
                    }

                    if (headerFontSize.HasValue) {
                        header.FontSize(headerFontSize.Value);
                    }

                    if (defaultHeader?.HasStyledZones == true) {
                        header.StyledZones(defaultHeader.CreateHeaderZone(NativeHeaderFooterZone.Left), defaultHeader.CreateHeaderZone(NativeHeaderFooterZone.Center), defaultHeader.CreateHeaderZone(NativeHeaderFooterZone.Right));
                    } else if (defaultHeader != null) {
                        header.Zones(defaultHeader.Left, defaultHeader.Center, defaultHeader.Right);
                    }

                    AddNativeHeaderImages(header, defaultHeaderImages, W.HeaderFooterValues.Default);
                    AddNativeHeaderShapes(header, defaultHeaderShapes, W.HeaderFooterValues.Default);

                    if (firstHeader?.HasStyledZones == true) {
                        header.FirstPageStyledZones(firstHeader.CreateHeaderZone(NativeHeaderFooterZone.Left), firstHeader.CreateHeaderZone(NativeHeaderFooterZone.Center), firstHeader.CreateHeaderZone(NativeHeaderFooterZone.Right));
                    } else if (firstHeader != null) {
                        header.FirstPageZones(firstHeader.Left, firstHeader.Center, firstHeader.Right);
                    } else if (hasFirstHeaderVariant) {
                        header.FirstPageText(string.Empty);
                    }

                    AddNativeHeaderImages(header, firstHeaderImages, W.HeaderFooterValues.First);
                    AddNativeHeaderShapes(header, firstHeaderShapes, W.HeaderFooterValues.First);

                    if (evenHeader?.HasStyledZones == true) {
                        header.EvenPagesStyledZones(evenHeader.CreateHeaderZone(NativeHeaderFooterZone.Left), evenHeader.CreateHeaderZone(NativeHeaderFooterZone.Center), evenHeader.CreateHeaderZone(NativeHeaderFooterZone.Right));
                    } else if (evenHeader != null) {
                        header.EvenPagesZones(evenHeader.Left, evenHeader.Center, evenHeader.Right);
                    } else if (hasEvenHeaderVariant) {
                        header.EvenPagesText(string.Empty);
                    }

                    AddNativeHeaderImages(header, evenHeaderImages, W.HeaderFooterValues.Even);
                    AddNativeHeaderShapes(header, evenHeaderShapes, W.HeaderFooterValues.Even);
                });
            }

            bool includePageNumbers = options?.IncludePageNumbers ?? false;
            if (!includePageNumbers && defaultFooter == null && !hasFirstFooterVariant && !hasEvenFooterVariant &&
                defaultFooterImages.Count == 0 && firstFooterImages.Count == 0 && evenFooterImages.Count == 0 &&
                defaultFooterShapes.Count == 0 && firstFooterShapes.Count == 0 && evenFooterShapes.Count == 0) {
                return;
            }

            string pageNumberFormat = GetNativePageNumberFormat(options);
            page.Footer(footer => {
                footer.Offset(GetNativeFooterOffset(options));
                if (footerFont.HasValue) {
                    footer.Font(footerFont.Value);
                }

                if (!string.IsNullOrWhiteSpace(footerFontFamily)) {
                    footer.FontFamily(footerFontFamily!);
                }

                if (footerColor.HasValue) {
                    footer.Color(footerColor.Value);
                }

                if (footerFontSize.HasValue) {
                    footer.FontSize(footerFontSize.Value);
                }

                NativeHeaderFooterText? resolvedDefaultFooter = WithNativeFooterPageNumber(defaultFooter, includePageNumbers, pageNumberFormat);
                if (resolvedDefaultFooter?.HasStyledZones == true) {
                    footer.StyledZones(resolvedDefaultFooter.CreateFooterZone(NativeHeaderFooterZone.Left), resolvedDefaultFooter.CreateFooterZone(NativeHeaderFooterZone.Center), resolvedDefaultFooter.CreateFooterZone(NativeHeaderFooterZone.Right));
                } else if (resolvedDefaultFooter != null) {
                    footer.Zones(resolvedDefaultFooter.Left, resolvedDefaultFooter.Center, resolvedDefaultFooter.Right);
                }

                AddNativeFooterImages(footer, defaultFooterImages, W.HeaderFooterValues.Default);
                AddNativeFooterShapes(footer, defaultFooterShapes, W.HeaderFooterValues.Default);

                NativeHeaderFooterText? resolvedFirstFooter = WithNativeFooterPageNumber(firstFooter, includePageNumbers && firstFooter != null, pageNumberFormat);
                if (resolvedFirstFooter?.HasStyledZones == true) {
                    footer.FirstPageStyledZones(resolvedFirstFooter.CreateFooterZone(NativeHeaderFooterZone.Left), resolvedFirstFooter.CreateFooterZone(NativeHeaderFooterZone.Center), resolvedFirstFooter.CreateFooterZone(NativeHeaderFooterZone.Right));
                } else if (resolvedFirstFooter != null) {
                    footer.FirstPageZones(resolvedFirstFooter.Left, resolvedFirstFooter.Center, resolvedFirstFooter.Right);
                } else if (hasFirstFooterVariant) {
                    footer.FirstPageText(string.Empty);
                }

                AddNativeFooterImages(footer, firstFooterImages, W.HeaderFooterValues.First);
                AddNativeFooterShapes(footer, firstFooterShapes, W.HeaderFooterValues.First);

                NativeHeaderFooterText? resolvedEvenFooter = WithNativeFooterPageNumber(evenFooter, includePageNumbers && evenFooter != null, pageNumberFormat);
                if (resolvedEvenFooter?.HasStyledZones == true) {
                    footer.EvenPagesStyledZones(resolvedEvenFooter.CreateFooterZone(NativeHeaderFooterZone.Left), resolvedEvenFooter.CreateFooterZone(NativeHeaderFooterZone.Center), resolvedEvenFooter.CreateFooterZone(NativeHeaderFooterZone.Right));
                } else if (resolvedEvenFooter != null) {
                    footer.EvenPagesZones(resolvedEvenFooter.Left, resolvedEvenFooter.Center, resolvedEvenFooter.Right);
                } else if (hasEvenFooterVariant) {
                    footer.EvenPagesText(string.Empty);
                }

                AddNativeFooterImages(footer, evenFooterImages, W.HeaderFooterValues.Even);
                AddNativeFooterShapes(footer, evenFooterShapes, W.HeaderFooterValues.Even);
            });
        }

        private static double GetNativeHeaderOffset(WordPdfExportOptions? options, double headerMarginExpansion) {
            double configuredOffset = options?.PdfOptions?.HeaderOffsetY.ToPoints() ?? NativeHeaderFooterDefaultOffset;
            return configuredOffset + headerMarginExpansion;
        }

        private static double GetNativeFooterOffset(WordPdfExportOptions? options) {
            return options?.PdfOptions?.FooterOffsetY.ToPoints() ?? NativeFooterDefaultOffset;
        }

        private static void AddNativeHeaderImages(PdfCore.PdfHeaderBuilder header, IReadOnlyList<NativeHeaderFooterImage> images, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterImage image in images) {
                if (variant == W.HeaderFooterValues.First) {
                    header.FirstPageImage(image.ToPdfImage());
                } else if (variant == W.HeaderFooterValues.Even) {
                    header.EvenPagesImage(image.ToPdfImage());
                } else {
                    header.Image(image.ToPdfImage());
                }
            }
        }

        private static void AddNativeHeaderShapes(PdfCore.PdfHeaderBuilder header, IReadOnlyList<NativeHeaderFooterShape> shapes, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterShape shape in shapes) {
                if (variant == W.HeaderFooterValues.First) {
                    header.FirstPageShape(shape.Shape, shape.Align);
                } else if (variant == W.HeaderFooterValues.Even) {
                    header.EvenPagesShape(shape.Shape, shape.Align);
                } else {
                    header.Shape(shape.Shape, shape.Align);
                }
            }
        }

        private static void AddNativeFooterImages(PdfCore.PdfFooterBuilder footer, IReadOnlyList<NativeHeaderFooterImage> images, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterImage image in images) {
                if (variant == W.HeaderFooterValues.First) {
                    footer.FirstPageImage(image.ToPdfImage());
                } else if (variant == W.HeaderFooterValues.Even) {
                    footer.EvenPagesImage(image.ToPdfImage());
                } else {
                    footer.Image(image.ToPdfImage());
                }
            }
        }

        private static void AddNativeFooterShapes(PdfCore.PdfFooterBuilder footer, IReadOnlyList<NativeHeaderFooterShape> shapes, W.HeaderFooterValues variant) {
            foreach (NativeHeaderFooterShape shape in shapes) {
                if (variant == W.HeaderFooterValues.First) {
                    footer.FirstPageShape(shape.Shape, shape.Align);
                } else if (variant == W.HeaderFooterValues.Even) {
                    footer.EvenPagesShape(shape.Shape, shape.Align);
                } else {
                    footer.Shape(shape.Shape, shape.Align);
                }
            }
        }

        private static void RecordNativeHeaderFooterDiagnostics(WordHeaderFooter? headerFooter, WordPdfExportOptions? options, string source) {
            if (headerFooter == null || options == null) {
                return;
            }

            var diagnosedParagraphs = new HashSet<W.Paragraph>();
            foreach (WordElement element in headerFooter.Elements) {
                RecordNativeHeaderFooterElementDiagnostics(element, options, source, diagnosedParagraphs);
            }
        }

        private static void RecordNativeHeaderFooterElementDiagnostics(WordElement element, WordPdfExportOptions options, string source,
            HashSet<W.Paragraph> diagnosedParagraphs) {
            switch (element) {
                case WordParagraph paragraph:
                    RecordNativeHeaderFooterParagraphDiagnostics(paragraph, options, source, diagnosedParagraphs);
                    break;
                case WordTable table:
                    RecordNativeHeaderFooterTableDiagnostics(table, options, source + " table", diagnosedParagraphs);
                    break;
                case WordEmbeddedDocument:
                    AddNativeExportWarning(
                        options,
                        "native-header-footer-embedded-document-unsupported",
                        source,
                        "Embedded documents in Word headers and footers are not mapped by the OfficeIMO PDF engine yet.");
                    break;
            }
        }

        private static void RecordNativeHeaderFooterTableDiagnostics(WordTable table, WordPdfExportOptions options, string source,
            HashSet<W.Paragraph> diagnosedParagraphs) {
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    List<WordElement> cellElements = EnumerateNativeTableCellElements(cell).ToList();
                    if (cellElements.Any(static element => element is WordTable)) {
                        AddNativeNestedTableLayoutWarning(options, source);
                    }

                    foreach (WordElement element in cellElements) {
                        RecordNativeHeaderFooterElementDiagnostics(element, options, source, diagnosedParagraphs);
                    }
                }
            }
        }

        private static void RecordNativeHeaderFooterParagraphDiagnostics(WordParagraph paragraph, WordPdfExportOptions options, string source,
            HashSet<W.Paragraph> diagnosedParagraphs) {
            if (paragraph.Shape != null && CreateNativeShape(paragraph.Shape) == null) {
                AddNativeExportWarning(
                    options,
                    "native-header-footer-shape-unsupported",
                    source,
                    "Word header and footer shapes without supported geometry are not mapped by the OfficeIMO PDF engine yet.");
            }

            if (paragraph.IsEquation && string.IsNullOrWhiteSpace(GetNativeEquationText(paragraph))) {
                AddNativeExportWarning(
                    options,
                    "native-header-footer-equation-unsupported",
                    source,
                    "Equations in Word headers and footers are not mapped by the OfficeIMO PDF engine yet.");
            }

            // Elements expose separate run wrappers over the same source
            // paragraph. Keep the run-specific checks above, but inspect its
            // shared XML content only once, including paragraphs in tables.
            if (paragraph._paragraph != null && !diagnosedParagraphs.Add(paragraph._paragraph)) return;

            RecordNativeFieldResultDiagnostics(paragraph, options, source, runningStory: true);
            RecordNativeParagraphFrameDiagnostics(paragraph, options, source);

            if (HasNativeUnsupportedHeaderFooterTextBox(paragraph)) {
                AddNativeExportWarning(
                    options,
                    "native-header-footer-text-box-unsupported",
                    source,
                    "Word header and footer text boxes without extractable text are not mapped by the OfficeIMO PDF engine yet.");
            }

            // Running-story callbacks paint diagrams after the conversion report is captured.
            // Qualify their semantic projection now so omissions and approximations reach that report.
            RecordNativeHeaderFooterSmartArtDiagnostics(paragraph, options, source);

            if (HasNativeUnsupportedHeaderFooterContentControl(paragraph)) {
                AddNativeExportWarning(
                    options,
                    "native-header-footer-content-control-unsupported",
                    source,
                    "Content controls in Word headers and footers are not mapped by the OfficeIMO PDF engine yet.");
            }
        }

        private static bool HasNativeUnsupportedHeaderFooterTextBox(WordParagraph paragraph) =>
            FindNativeParagraphTextBox(paragraph) != null &&
            string.IsNullOrWhiteSpace(GetNativeParagraphTextBoxPlainText(paragraph));

        private static void RecordNativeBodyTableDiagnostics(WordTable table, WordPdfExportOptions? options, string source) {
            if (options == null) {
                return;
            }

            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    List<WordElement> cellElements = EnumerateNativeTableCellElements(cell).ToList();
                    foreach (WordElement element in cellElements) {
                        if (element is WordParagraph paragraph) {
                            RecordNativeBodyParagraphDiagnostics(paragraph, options, source, mapsCheckBoxes: true, mapsFormFields: true, mapsPictureControls: true, mapsRepeatingSections: true);
                        } else {
                            RecordNativeBodyElementDiagnostics(element, options, source);
                        }
                    }
                }
            }
        }

        private static void AddNativeNestedTableLayoutWarning(WordPdfExportOptions options, string source) {
            AddNativeExportWarning(
                options,
                "native-nested-table-layout-approximated",
                source,
                "This nested table profile retains its text through flattened cell flow; its inner grid and mixed-content placement are approximated.");
        }

        private static void RecordNativeBodyElementDiagnostics(WordElement element, WordPdfExportOptions options, string source) {
            switch (element) {
                case WordParagraph paragraph:
                    RecordNativeBodyParagraphDiagnostics(paragraph, options, source, mapsCheckBoxes: false, mapsFormFields: false, mapsPictureControls: false, mapsRepeatingSections: false);
                    break;
                case WordTable table:
                    RecordNativeBodyTableDiagnostics(table, options, source + " table");
                    break;
                case WordStructuredDocumentTag structuredDocumentTag:
                    foreach (WordElement structuredElement in EnumerateNativeStructuredContentElements(
                        new[] { structuredDocumentTag },
                        structuredDocumentTagDepth: 0)) {
                        RecordNativeBodyElementDiagnostics(structuredElement, options, source);
                    }
                    break;
                case WordEmbeddedDocument:
                    AddNativeExportWarning(
                        options,
                        "native-body-embedded-document-unsupported",
                        source,
                        "Embedded documents in Word body content are not mapped by the OfficeIMO PDF engine yet.");
                    break;
            }
        }

        private static void RecordNativeBodyParagraphDiagnostics(WordParagraph paragraph, WordPdfExportOptions? options, string source, bool mapsCheckBoxes, bool mapsFormFields, bool mapsPictureControls, bool mapsRepeatingSections, bool mapsSmartArt = false) {
            if (options == null) {
                return;
            }

            RecordNativeFieldResultDiagnostics(paragraph, options, source, runningStory: false);
            RecordNativeParagraphFrameDiagnostics(paragraph, options, source);
            if (paragraph._paragraph.Ancestors<W.TableCell>().Any())
                RecordNativeCellBarTabDiagnostics(paragraph, options, source);

            if (!mapsSmartArt && GetNativeRuns(paragraph).Any(run => run.IsSmartArt)) {
                AddNativeExportWarning(
                    options,
                    "native-body-smart-art-unsupported",
                    source,
                    "SmartArt in Word body content is not mapped by the OfficeIMO PDF engine yet.");
            }

            if (paragraph.IsEquation && string.IsNullOrWhiteSpace(GetNativeEquationText(paragraph))) {
                AddNativeExportWarning(
                    options,
                    "native-body-equation-unsupported",
                    source,
                    "Equations in Word body content are not mapped by the OfficeIMO PDF engine yet.");
            }

            if (HasNativeUnsupportedBodyContentControl(paragraph, mapsCheckBoxes, mapsFormFields, mapsPictureControls, mapsRepeatingSections)) {
                AddNativeExportWarning(
                    options,
                    "native-body-content-control-unsupported",
                    source,
                    "Content controls in Word body content are not mapped by the OfficeIMO PDF engine yet.");
            }
        }

        private static bool HasNativeUnsupportedHeaderFooterContentControl(WordParagraph paragraph) =>
            (paragraph.IsCheckBox && GetNativeCheckBoxControls(paragraph).Count == 0) ||
            ((paragraph.IsDatePicker || paragraph.IsDropDownList || paragraph.IsComboBox) && GetNativeFormFieldControls(paragraph).Count == 0) ||
            (paragraph.IsPictureControl && paragraph.PictureControl?.Image == null) ||
            (paragraph.IsRepeatingSection && paragraph.RepeatingSection?.TextItems.Count == 0) ||
            paragraph._paragraph?.Descendants<W.SdtRun>().Any(sdtRun =>
                !IsNativeSimpleTextContentControl(sdtRun) &&
                !IsNativeCheckBoxControl(sdtRun) &&
                !IsNativeSupportedFormFieldContentControl(sdtRun) &&
                !IsNativePictureControlWithImage(paragraph, sdtRun) &&
                !IsNativeRepeatingSectionWithText(sdtRun) &&
                !IsNativeRepeatingSectionChildControl(sdtRun)) == true ||
            paragraph._paragraph?.Descendants<W.SdtBlock>().Any(sdtBlock => !IsNativeSimpleBlockTextContentControl(sdtBlock)) == true ||
            paragraph._paragraph?.Descendants<W.SdtCell>().Any() == true;

        private static bool HasNativeUnsupportedBodyContentControl(WordParagraph paragraph, bool mapsCheckBoxes, bool mapsFormFields, bool mapsPictureControls, bool mapsRepeatingSections) =>
            (!mapsCheckBoxes && paragraph.IsCheckBox) ||
            (!mapsFormFields && (paragraph.IsDatePicker || paragraph.IsDropDownList || paragraph.IsComboBox)) ||
            (paragraph.IsPictureControl && (!mapsPictureControls || paragraph.PictureControl?.Image == null)) ||
            (paragraph.IsRepeatingSection && (!mapsRepeatingSections || paragraph.RepeatingSection?.TextItems.Count == 0)) ||
            paragraph._paragraph?.Descendants<W.SdtRun>().Any(sdtRun =>
                (!mapsCheckBoxes || !IsNativeCheckBoxControl(sdtRun)) &&
                (!mapsFormFields || !IsNativeSupportedFormFieldContentControl(sdtRun)) &&
                (!mapsPictureControls || !IsNativePictureControl(sdtRun)) &&
                (!mapsRepeatingSections || !IsNativeRepeatingSectionControl(sdtRun) && !IsNativeRepeatingSectionChildControl(sdtRun)) &&
                !IsNativeSimpleTextContentControl(sdtRun)) == true ||
            paragraph._paragraph?.Descendants<W.SdtBlock>().Any(sdtBlock => !IsNativeSimpleBlockTextContentControl(sdtBlock)) == true ||
            paragraph._paragraph?.Descendants<W.SdtCell>().Any() == true;

        private static IReadOnlyList<W.SdtRun> GetNativeCheckBoxControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }

            return paragraph._paragraph.Descendants<W.SdtRun>()
                .Where(IsNativeCheckBoxControl)
                .ToList();
        }

        private static bool IsNativeCheckBoxControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().Any() == true;

        private static bool IsNativePictureControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W.SdtContentPicture>().Any() == true;

        private static bool IsNativePictureControlWithImage(WordParagraph paragraph, W.SdtRun sdtRun) {
            if (!IsNativePictureControl(sdtRun) || paragraph._paragraph == null) {
                return false;
            }

            var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph, sdtRun);
            return pictureParagraph.PictureControl?.Image != null;
        }

        private static IReadOnlyList<W.SdtRun> GetNativePictureControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }
            var controls = new List<W.SdtRun>();
            var visibility = WordComplexFieldRunVisibility.ForParagraph(paragraph._paragraph);
            CollectNativeVisiblePictureControls(paragraph._paragraph, visibility, controls);
            return controls;
        }

        private static void CollectNativeVisiblePictureControls(
            DocumentFormat.OpenXml.OpenXmlCompositeElement container,
            WordComplexFieldRunVisibility visibility,
            List<W.SdtRun> controls) {
            foreach (DocumentFormat.OpenXml.OpenXmlElement child in container.ChildElements) {
                if (child is W.Run run) {
                    visibility.GetVisibleRun(run);
                } else if (child is W.SdtRun control) {
                    if (visibility.IsVisible && IsNativePictureControl(control)) controls.Add(control);
                    if (control.SdtContentRun != null)
                        CollectNativeVisiblePictureControls(control.SdtContentRun, visibility, controls);
                } else if (child is W.SimpleField field) {
                    if (visibility.IsVisible)
                        CollectNativeVisiblePictureControls(field, visibility, controls);
                    else
                        visibility.ObserveDescendantRuns(field);
                } else if (child is W.Hyperlink or W.CustomXmlRun) {
                    CollectNativeVisiblePictureControls((DocumentFormat.OpenXml.OpenXmlCompositeElement)child, visibility, controls);
                }
            }
        }

        private static bool IsNativeRepeatingSectionControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W15.SdtRepeatedSection>().Any() == true;

        private static bool IsNativeRepeatingSectionChildControl(W.SdtRun sdtRun) =>
            sdtRun.Ancestors<W.SdtRun>().Any(IsNativeRepeatingSectionControl);

        private static bool IsNativeRepeatingSectionWithText(W.SdtRun sdtRun) =>
            IsNativeRepeatingSectionControl(sdtRun) &&
            GetNativeRepeatingSectionItems(sdtRun).Count > 0;

        private static IReadOnlyList<W.SdtRun> GetNativeFormFieldControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }

            return paragraph._paragraph.Descendants<W.SdtRun>()
                .Where(IsNativeSupportedFormFieldContentControl)
                .ToList();
        }

        private static IReadOnlyList<W.SdtRun> GetNativeRepeatingSectionControls(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return Array.Empty<W.SdtRun>();
            }

            return paragraph._paragraph.Descendants<W.SdtRun>()
                .Where(IsNativeRepeatingSectionControl)
                .ToList();
        }

        private static bool IsNativeSupportedFormFieldContentControl(W.SdtRun sdtRun) =>
            IsNativeDatePickerControl(sdtRun) ||
            GetNativeChoiceFieldOptions(sdtRun).Count > 0;

        private static bool IsNativeDatePickerControl(W.SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<W.SdtContentDate>().Any() == true;

        private static IReadOnlyList<string> GetNativeChoiceFieldOptions(W.SdtRun sdtRun) {
            W.SdtProperties? properties = sdtRun.SdtProperties;
            if (properties == null) {
                return Array.Empty<string>();
            }

            IEnumerable<W.ListItem> items = properties.Elements<W.SdtContentDropDownList>().FirstOrDefault()?.Elements<W.ListItem>() ??
                properties.Elements<W.SdtContentComboBox>().FirstOrDefault()?.Elements<W.ListItem>() ??
                Enumerable.Empty<W.ListItem>();

            var options = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (W.ListItem item in items) {
                string? option = item.DisplayText?.Value ?? item.Value?.Value;
                if (string.IsNullOrWhiteSpace(option) || !seen.Add(option!)) {
                    continue;
                }

                options.Add(option!);
            }

            return options;
        }

        private static string? GetNativeChoiceFieldValue(W.SdtRun sdtRun, IReadOnlyList<string> options) {
            W.SdtContentComboBox? comboBox = sdtRun.SdtProperties?.Elements<W.SdtContentComboBox>().FirstOrDefault();
            string? lastValue = comboBox?.LastValue?.Value;
            if (!string.IsNullOrWhiteSpace(lastValue)) {
                string? displayValue = GetNativeChoiceDisplayValue(sdtRun, lastValue!);
                if (!string.IsNullOrWhiteSpace(displayValue) && options.Contains(displayValue!, StringComparer.Ordinal)) {
                    return displayValue;
                }

                if (options.Contains(lastValue!, StringComparer.Ordinal)) {
                    return lastValue;
                }
            }

            string? contentText = GetNativeSdtText(sdtRun);
            if (!string.IsNullOrWhiteSpace(contentText) && options.Contains(contentText!, StringComparer.Ordinal)) {
                return contentText;
            }

            return options.Count > 0 ? options[0] : null;
        }

        private static string? GetNativeChoiceDisplayValue(W.SdtRun sdtRun, string value) {
            IEnumerable<W.ListItem> items = sdtRun.SdtProperties?.Elements<W.SdtContentDropDownList>().FirstOrDefault()?.Elements<W.ListItem>() ??
                sdtRun.SdtProperties?.Elements<W.SdtContentComboBox>().FirstOrDefault()?.Elements<W.ListItem>() ??
                Enumerable.Empty<W.ListItem>();

            W.ListItem? match = items.FirstOrDefault(item => string.Equals(item.Value?.Value, value, StringComparison.Ordinal));
            return match?.DisplayText?.Value ?? match?.Value?.Value;
        }

        private static string GetNativeDatePickerValue(W.SdtRun sdtRun) {
            W.SdtContentDate? datePicker = sdtRun.SdtProperties?.Elements<W.SdtContentDate>().FirstOrDefault();
            if (datePicker?.FullDate?.Value is DateTime value) {
                return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return GetNativeSdtText(sdtRun) ?? string.Empty;
        }

        private static string? GetNativeSdtText(W.SdtRun sdtRun) {
            if (sdtRun.SdtContentRun == null) {
                return null;
            }

            string text = string.Concat(sdtRun.SdtContentRun.Descendants<W.Text>().Select(WordParagraph.ReadVisibleText));
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private static bool IsNativeSimpleTextContentControl(W.SdtRun sdtRun) {
            W.SdtProperties? properties = sdtRun.SdtProperties;
            if (properties == null) {
                return false;
            }

            if (IsNativeBuiltInPropertyContentControl(properties)) {
                return true;
            }

            if (properties.Elements<W14.SdtContentCheckBox>().Any() ||
                properties.Elements<W.SdtContentDate>().Any() ||
                properties.Elements<W.SdtContentDropDownList>().Any() ||
                properties.Elements<W.SdtContentComboBox>().Any() ||
                properties.Elements<W.SdtContentPicture>().Any() ||
                properties.Elements<W15.SdtRepeatedSection>().Any()) {
                return false;
            }

            return sdtRun.SdtContentRun?.Descendants<W.Text>().Any() == true;
        }

        private static bool IsNativeSimpleBlockTextContentControl(W.SdtBlock sdtBlock) {
            W.SdtProperties? properties = sdtBlock.SdtProperties;
            if (properties == null) {
                return false;
            }

            if (IsNativeBuiltInPropertyContentControl(properties) ||
                IsNativeSupportedDocPartBlockContentControl(sdtBlock)) {
                return true;
            }

            if (properties.Elements<W.SdtContentDocPartObject>().Any() ||
                properties.Elements<W.SdtContentDate>().Any() ||
                properties.Elements<W.SdtContentPicture>().Any()) {
                return false;
            }

            return sdtBlock.SdtContentBlock?.Descendants<W.Text>().Any() == true;
        }

        private static bool IsNativeSupportedDocPartBlockContentControl(W.SdtBlock sdtBlock) {
            string? gallery = sdtBlock.SdtProperties?
                .Elements<W.SdtContentDocPartObject>()
                .FirstOrDefault()?
                .Elements<W.DocPartGallery>()
                .FirstOrDefault()?
                .Val?
                .Value;
            return gallery != null &&
                   (gallery.Equals("Cover Pages", StringComparison.OrdinalIgnoreCase) ||
                    gallery.Equals("Table of Contents", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsNativeBuiltInPropertyContentControl(W.SdtProperties properties) {
            string key = (properties.Elements<W.SdtAlias>().FirstOrDefault()?.Val?.Value ?? string.Empty).Trim();
            string? xPath = properties.Elements<W.DataBinding>().FirstOrDefault()?.XPath?.Value;
            if (string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(xPath)) {
                key = xPath!;
            }

            return key.IndexOf("Company", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   key.Equals("Title", StringComparison.OrdinalIgnoreCase) ||
                   key.IndexOf("Title[", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   key.Equals("Subtitle", StringComparison.OrdinalIgnoreCase) ||
                   key.Equals("Subject", StringComparison.OrdinalIgnoreCase) ||
                   key.Equals("Author", StringComparison.OrdinalIgnoreCase) ||
                   key.Equals("Creator", StringComparison.OrdinalIgnoreCase) ||
                   key.Equals("Date", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsNativeCheckBoxChecked(W.SdtRun sdtRun) {
            W14.SdtContentCheckBox? checkBox = sdtRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().FirstOrDefault();
            W14.Checked? checkedState = checkBox?.Elements<W14.Checked>().FirstOrDefault();
            return checkedState?.Val?.Value == W14.OnOffValues.One;
        }

        private static string GetNativeCheckBoxFieldName(W.SdtRun sdtRun, int index, string fallbackPrefix = "WordCheckBox") {
            return GetNativeContentControlFieldName(sdtRun, index, fallbackPrefix);
        }

        private static string GetNativeContentControlFieldName(W.SdtRun sdtRun, int index, string fallbackPrefix) {
            string? tag = sdtRun.SdtProperties?.Elements<W.Tag>().FirstOrDefault()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(tag)) {
                return tag!;
            }

            string? alias = sdtRun.SdtProperties?.Elements<W.SdtAlias>().FirstOrDefault()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(alias)) {
                return alias!;
            }

            int? sdtId = sdtRun.SdtProperties?.Elements<W.SdtId>().FirstOrDefault()?.Val?.Value;
            return sdtId.HasValue
                ? fallbackPrefix + "." + sdtId.Value.ToString(CultureInfo.InvariantCulture)
                : fallbackPrefix + "." + (index + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static void AddNativeExportWarning(WordPdfExportOptions options, string code, string source, string message) {
            var warning = new PdfExportWarning(code, source, message);
            options.Warnings.Add(warning);
            options.Report.Add(warning.ToConversionWarning());
        }

    }
}
