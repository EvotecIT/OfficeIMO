using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using System.Xml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static void ConfigureNativePageNumbering(PdfCore.PdfPageBuilder page, WordSection section) {
            string? format = section.Node().WChild("pgNumType").WAttr("fmt");
            if (section.GetEffectivePageNumberStart() is int start) {
                page.PageNumberStart(start);
            }

            PdfCore.PdfPageNumberStyle? style = MapNativePageNumberStyle(format);
            page.Options.LayoutPageNumberStyle = style ?? PdfCore.PdfPageNumberStyle.Arabic;
            page.Options.LayoutPageNumberStyleUnsupported = format != null
                && (!style.HasValue || format == "decimalZero");
            if (style.HasValue) {
                page.PageNumberStyle(style.Value);
            }
        }

        private static PdfCore.PdfPageNumberStyle? MapNativePageNumberStyle(string? format) {
            switch (format) {
                case "lowerRoman":
                    return PdfCore.PdfPageNumberStyle.LowerRoman;
                case "upperRoman":
                    return PdfCore.PdfPageNumberStyle.UpperRoman;
                case "lowerLetter":
                    return PdfCore.PdfPageNumberStyle.LowerLetter;
                case "upperLetter":
                    return PdfCore.PdfPageNumberStyle.UpperLetter;
                case "decimal":
                case "decimalZero":
                    return PdfCore.PdfPageNumberStyle.Arabic;
                default:
                    return null;
            }
        }

        private static void ConfigureNativeHeaderFooter(PdfCore.PdfPageBuilder page, WordSection section, WordPdfExportOptions? options, double headerMarginExpansion, double footerMarginExpansion, NativeFontMap nativeFontMap, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            bool hasEvenPageStory = section.DocumentOddEvenSettingEnabled;
            RecordNativeHeaderFooterDiagnostics(section.ResolveDefaultHeader(), options, "default header");
            if (section.PageSetup.DifferentFirstPage)
                RecordNativeHeaderFooterDiagnostics(section.ResolveFirstHeader(), options, "first header");
            if (hasEvenPageStory)
                RecordNativeHeaderFooterDiagnostics(section.ResolveEvenHeader(), options, "even header");
            RecordNativeHeaderFooterDiagnostics(section.ResolveDefaultFooter(), options, "default footer");
            if (section.PageSetup.DifferentFirstPage)
                RecordNativeHeaderFooterDiagnostics(section.ResolveFirstFooter(), options, "first footer");
            if (hasEvenPageStory)
                RecordNativeHeaderFooterDiagnostics(section.ResolveEvenFooter(), options, "even footer");
            ApplyNativeSectionWatermark(page, section, options);
            page.Options.PageSetup.UseVisiblePageNumberParity = true;

            // Flat projections still supply field-numbering metadata. Their
            // approximation diagnostics do not describe admitted running flow.
            WordPdfExportOptions? headerTextDiagnostics = UsesNativeRunningHeader(section, options) ? null : options;
            WordPdfExportOptions? footerTextDiagnostics = UsesNativeRunningFooter(section, options) ? null : options;
            RecordNativeRunningBarTabDiagnostics(section, options,
                headerTextDiagnostics == null, footerTextDiagnostics == null);
            RecordNativeRunningImageWrapDiagnostics(section, options,
                headerTextDiagnostics == null, footerTextDiagnostics == null);
            RecordNativeRunningShapeGroupDiagnostics(section, options,
                headerTextDiagnostics == null, footerTextDiagnostics == null);
            NativeHeaderFooterText? defaultHeader = GetNativeHeaderFooterText(section.ResolveDefaultHeader(), listMarkers, nativeFontMap, headerTextDiagnostics);
            NativeHeaderFooterText? firstHeader = section.PageSetup.DifferentFirstPage ? GetNativeHeaderFooterText(section.ResolveFirstHeader(), listMarkers, nativeFontMap, headerTextDiagnostics) : null;
            NativeHeaderFooterText? evenHeader = hasEvenPageStory ? GetNativeHeaderFooterText(section.ResolveEvenHeader(), listMarkers, nativeFontMap, headerTextDiagnostics) : null;
            NativeHeaderFooterText? defaultFooter = GetNativeHeaderFooterText(section.ResolveDefaultFooter(), listMarkers, nativeFontMap, footerTextDiagnostics);
            NativeHeaderFooterText? firstFooter = section.PageSetup.DifferentFirstPage ? GetNativeHeaderFooterText(section.ResolveFirstFooter(), listMarkers, nativeFontMap, footerTextDiagnostics) : null;
            NativeHeaderFooterText? evenFooter = hasEvenPageStory ? GetNativeHeaderFooterText(section.ResolveEvenFooter(), listMarkers, nativeFontMap, footerTextDiagnostics) : null;
            IReadOnlyList<NativeHeaderFooterImage> defaultHeaderImages = GetNativeHeaderFooterImages(section.ResolveDefaultHeader(), options, "default header image");
            IReadOnlyList<NativeHeaderFooterImage> firstHeaderImages = section.PageSetup.DifferentFirstPage ? GetNativeHeaderFooterImages(section.ResolveFirstHeader(), options, "first header image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterImage> evenHeaderImages = hasEvenPageStory ? GetNativeHeaderFooterImages(section.ResolveEvenHeader(), options, "even header image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterImage> defaultFooterImages = GetNativeHeaderFooterImages(section.ResolveDefaultFooter(), options, "default footer image");
            IReadOnlyList<NativeHeaderFooterImage> firstFooterImages = section.PageSetup.DifferentFirstPage ? GetNativeHeaderFooterImages(section.ResolveFirstFooter(), options, "first footer image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterImage> evenFooterImages = hasEvenPageStory ? GetNativeHeaderFooterImages(section.ResolveEvenFooter(), options, "even footer image") : Array.Empty<NativeHeaderFooterImage>();
            IReadOnlyList<NativeHeaderFooterShape> defaultHeaderShapes = GetNativeHeaderFooterShapes(section.ResolveDefaultHeader());
            IReadOnlyList<NativeHeaderFooterShape> firstHeaderShapes = section.PageSetup.DifferentFirstPage ? GetNativeHeaderFooterShapes(section.ResolveFirstHeader()) : Array.Empty<NativeHeaderFooterShape>();
            IReadOnlyList<NativeHeaderFooterShape> evenHeaderShapes = hasEvenPageStory ? GetNativeHeaderFooterShapes(section.ResolveEvenHeader()) : Array.Empty<NativeHeaderFooterShape>();
            IReadOnlyList<NativeHeaderFooterShape> defaultFooterShapes = GetNativeHeaderFooterShapes(section.ResolveDefaultFooter());
            IReadOnlyList<NativeHeaderFooterShape> firstFooterShapes = section.PageSetup.DifferentFirstPage ? GetNativeHeaderFooterShapes(section.ResolveFirstFooter()) : Array.Empty<NativeHeaderFooterShape>();
            IReadOnlyList<NativeHeaderFooterShape> evenFooterShapes = hasEvenPageStory ? GetNativeHeaderFooterShapes(section.ResolveEvenFooter()) : Array.Empty<NativeHeaderFooterShape>();
            PdfCore.PdfStandardFont? headerFont = ResolveNativeHeaderFooterFont(
                ResolveNativeHeaderFooterBaseFont(section._document, options, isHeader: true),
                nativeFontMap,
                section.ResolveDefaultHeader(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            PdfCore.PdfStandardFont? footerFont = ResolveNativeHeaderFooterFont(
                ResolveNativeHeaderFooterBaseFont(section._document, options, isHeader: false),
                nativeFontMap,
                section.ResolveDefaultFooter(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            string? headerFontFamily = ResolveNativeHeaderFooterFontFamily(
                nativeFontMap,
                section.ResolveDefaultHeader(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            string? footerFontFamily = ResolveNativeHeaderFooterFontFamily(
                nativeFontMap,
                section.ResolveDefaultFooter(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            global::OfficeIMO.OfficeColor? headerColor = ResolveNativeHeaderFooterColor(
                section.ResolveDefaultHeader(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            global::OfficeIMO.OfficeColor? footerColor = ResolveNativeHeaderFooterColor(
                section.ResolveDefaultFooter(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            double? headerFontSize = ResolveNativeHeaderFooterFontSize(
                section.ResolveDefaultHeader(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                hasEvenPageStory ? section.ResolveEvenHeader() : null);
            double? footerFontSize = ResolveNativeHeaderFooterFontSize(
                section.ResolveDefaultFooter(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                hasEvenPageStory ? section.ResolveEvenFooter() : null);
            ApplyNativeHeaderFooterPageNumberStyle(page, defaultHeader, firstHeader, evenHeader, defaultFooter, firstFooter, evenFooter);
            bool hasFirstHeaderVariant = section.PageSetup.DifferentFirstPage;
            bool hasEvenHeaderVariant = hasEvenPageStory;
            bool hasFirstFooterVariant = section.PageSetup.DifferentFirstPage;
            bool hasEvenFooterVariant = hasEvenPageStory;
            if (defaultHeader != null || hasFirstHeaderVariant || hasEvenHeaderVariant ||
                defaultHeaderImages.Count > 0 || firstHeaderImages.Count > 0 || evenHeaderImages.Count > 0 ||
                defaultHeaderShapes.Count > 0 || firstHeaderShapes.Count > 0 || evenHeaderShapes.Count > 0) {
                page.Header(header => {
                    if (headerMarginExpansion > 0D) {
                        header.Offset(global::OfficeIMO.OfficeLength.Points(GetNativeHeaderOffset(options, headerMarginExpansion)));
                    }

                    if (headerFont.HasValue) {
                        header.StandardFont(headerFont.Value);
                    }

                    if (!string.IsNullOrWhiteSpace(headerFontFamily)) {
                        header.FontFamily(headerFontFamily!);
                    }

                    if (headerColor.HasValue) {
                        header.Color(headerColor.Value);
                    }

                    if (headerFontSize.HasValue) {
                        header.FontSize(global::OfficeIMO.OfficeLength.Points(headerFontSize.Value));
                    }

                    if (defaultHeader?.HasStyledZones == true) {
                        header.StyledZones(defaultHeader.CreateHeaderZone(NativeHeaderFooterZone.Left), defaultHeader.CreateHeaderZone(NativeHeaderFooterZone.Center), defaultHeader.CreateHeaderZone(NativeHeaderFooterZone.Right));
                    } else if (defaultHeader != null) {
                        header.Zones(defaultHeader.Left, defaultHeader.Center, defaultHeader.Right);
                    }

                    AddNativeHeaderImages(header, defaultHeaderImages, "default");
                    AddNativeHeaderShapes(header, defaultHeaderShapes, "default");

                    if (firstHeader?.HasStyledZones == true) {
                        header.FirstPageStyledZones(firstHeader.CreateHeaderZone(NativeHeaderFooterZone.Left), firstHeader.CreateHeaderZone(NativeHeaderFooterZone.Center), firstHeader.CreateHeaderZone(NativeHeaderFooterZone.Right));
                    } else if (firstHeader != null) {
                        header.FirstPageZones(firstHeader.Left, firstHeader.Center, firstHeader.Right);
                    } else if (hasFirstHeaderVariant) {
                        header.FirstPageText(string.Empty);
                    }

                    AddNativeHeaderImages(header, firstHeaderImages, "first");
                    AddNativeHeaderShapes(header, firstHeaderShapes, "first");

                    if (evenHeader?.HasStyledZones == true) {
                        header.EvenPagesStyledZones(evenHeader.CreateHeaderZone(NativeHeaderFooterZone.Left), evenHeader.CreateHeaderZone(NativeHeaderFooterZone.Center), evenHeader.CreateHeaderZone(NativeHeaderFooterZone.Right));
                    } else if (evenHeader != null) {
                        header.EvenPagesZones(evenHeader.Left, evenHeader.Center, evenHeader.Right);
                    } else if (hasEvenHeaderVariant) {
                        header.EvenPagesText(string.Empty);
                    }

                    AddNativeHeaderImages(header, evenHeaderImages, "even");
                    AddNativeHeaderShapes(header, evenHeaderShapes, "even");
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
                footer.Offset(global::OfficeIMO.OfficeLength.Points(GetNativeFooterOffset(options)));
                if (footerFont.HasValue) {
                    footer.StandardFont(footerFont.Value);
                }

                if (!string.IsNullOrWhiteSpace(footerFontFamily)) {
                    footer.FontFamily(footerFontFamily!);
                }

                if (footerColor.HasValue) {
                    footer.Color(footerColor.Value);
                }

                if (footerFontSize.HasValue) {
                    footer.FontSize(global::OfficeIMO.OfficeLength.Points(footerFontSize.Value));
                }

                NativeHeaderFooterText? resolvedDefaultFooter = WithNativeFooterPageNumber(defaultFooter, includePageNumbers, pageNumberFormat);
                if (resolvedDefaultFooter?.HasStyledZones == true) {
                    footer.StyledZones(resolvedDefaultFooter.CreateFooterZone(NativeHeaderFooterZone.Left), resolvedDefaultFooter.CreateFooterZone(NativeHeaderFooterZone.Center), resolvedDefaultFooter.CreateFooterZone(NativeHeaderFooterZone.Right));
                } else if (resolvedDefaultFooter != null) {
                    footer.Zones(resolvedDefaultFooter.Left, resolvedDefaultFooter.Center, resolvedDefaultFooter.Right);
                }

                AddNativeFooterImages(footer, defaultFooterImages, "default");
                AddNativeFooterShapes(footer, defaultFooterShapes, "default");

                NativeHeaderFooterText? resolvedFirstFooter = WithNativeFooterPageNumber(firstFooter, includePageNumbers && firstFooter != null, pageNumberFormat);
                if (resolvedFirstFooter?.HasStyledZones == true) {
                    footer.FirstPageStyledZones(resolvedFirstFooter.CreateFooterZone(NativeHeaderFooterZone.Left), resolvedFirstFooter.CreateFooterZone(NativeHeaderFooterZone.Center), resolvedFirstFooter.CreateFooterZone(NativeHeaderFooterZone.Right));
                } else if (resolvedFirstFooter != null) {
                    footer.FirstPageZones(resolvedFirstFooter.Left, resolvedFirstFooter.Center, resolvedFirstFooter.Right);
                } else if (hasFirstFooterVariant) {
                    footer.FirstPageText(string.Empty);
                }

                AddNativeFooterImages(footer, firstFooterImages, "first");
                AddNativeFooterShapes(footer, firstFooterShapes, "first");

                NativeHeaderFooterText? resolvedEvenFooter = WithNativeFooterPageNumber(evenFooter, includePageNumbers && evenFooter != null, pageNumberFormat);
                if (resolvedEvenFooter?.HasStyledZones == true) {
                    footer.EvenPagesStyledZones(resolvedEvenFooter.CreateFooterZone(NativeHeaderFooterZone.Left), resolvedEvenFooter.CreateFooterZone(NativeHeaderFooterZone.Center), resolvedEvenFooter.CreateFooterZone(NativeHeaderFooterZone.Right));
                } else if (resolvedEvenFooter != null) {
                    footer.EvenPagesZones(resolvedEvenFooter.Left, resolvedEvenFooter.Center, resolvedEvenFooter.Right);
                } else if (hasEvenFooterVariant) {
                    footer.EvenPagesText(string.Empty);
                }

                AddNativeFooterImages(footer, evenFooterImages, "even");
                AddNativeFooterShapes(footer, evenFooterShapes, "even");
            });
        }

        private static double GetNativeHeaderOffset(WordPdfExportOptions? options, double headerMarginExpansion) {
            double configuredOffset = options?.PdfOptions?.PageSetup.Header.Offset.ToPoints() ?? NativeHeaderFooterDefaultOffset;
            return configuredOffset + headerMarginExpansion;
        }

        private static double GetNativeFooterOffset(WordPdfExportOptions? options) {
            return options?.PdfOptions?.PageSetup.Footer.Offset.ToPoints() ?? NativeFooterDefaultOffset;
        }

        private static void AddNativeHeaderImages(PdfCore.PdfHeaderBuilder header, IReadOnlyList<NativeHeaderFooterImage> images, string variant) {
            foreach (NativeHeaderFooterImage image in images) {
                if (string.Equals(variant, "first", StringComparison.Ordinal)) {
                    header.FirstPageImage(image.ToPdfImage());
                } else if (string.Equals(variant, "even", StringComparison.Ordinal)) {
                    header.EvenPagesImage(image.ToPdfImage());
                } else {
                    header.Image(image.ToPdfImage());
                }
            }
        }

        private static void AddNativeHeaderShapes(PdfCore.PdfHeaderBuilder header, IReadOnlyList<NativeHeaderFooterShape> shapes, string variant) {
            foreach (NativeHeaderFooterShape shape in shapes) {
                if (string.Equals(variant, "first", StringComparison.Ordinal)) {
                    header.FirstPageShape(shape.Shape, shape.Align);
                } else if (string.Equals(variant, "even", StringComparison.Ordinal)) {
                    header.EvenPagesShape(shape.Shape, shape.Align);
                } else {
                    header.Shape(shape.Shape, shape.Align);
                }
            }
        }

        private static void AddNativeFooterImages(PdfCore.PdfFooterBuilder footer, IReadOnlyList<NativeHeaderFooterImage> images, string variant) {
            foreach (NativeHeaderFooterImage image in images) {
                if (string.Equals(variant, "first", StringComparison.Ordinal)) {
                    footer.FirstPageImage(image.ToPdfImage());
                } else if (string.Equals(variant, "even", StringComparison.Ordinal)) {
                    footer.EvenPagesImage(image.ToPdfImage());
                } else {
                    footer.Image(image.ToPdfImage());
                }
            }
        }

        private static void AddNativeFooterShapes(PdfCore.PdfFooterBuilder footer, IReadOnlyList<NativeHeaderFooterShape> shapes, string variant) {
            foreach (NativeHeaderFooterShape shape in shapes) {
                if (string.Equals(variant, "first", StringComparison.Ordinal)) {
                    footer.FirstPageShape(shape.Shape, shape.Align);
                } else if (string.Equals(variant, "even", StringComparison.Ordinal)) {
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

            var diagnosedParagraphs = new HashSet<OfficeXmlNode>();
            foreach (WordElement element in headerFooter.Elements) {
                RecordNativeHeaderFooterElementDiagnostics(element, options, source, diagnosedParagraphs);
            }
        }

        private static void RecordNativeHeaderFooterElementDiagnostics(WordElement element, WordPdfExportOptions options, string source,
            HashSet<OfficeXmlNode> diagnosedParagraphs) {
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
            HashSet<OfficeXmlNode> diagnosedParagraphs) {
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
            HashSet<OfficeXmlNode> diagnosedParagraphs) {
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
            OfficeXmlNode sourceParagraph = paragraph.Node();
            if (sourceParagraph.Exists && !diagnosedParagraphs.Add(sourceParagraph)) return;

            RecordNativeFieldResultDiagnostics(paragraph, options, source, runningStory: true);
            RecordNativeParagraphFrameDiagnostics(paragraph, options, source);
            RecordNativeCellTextBoxDiagnostic(paragraph, options, source);

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
            if (paragraph.Node().Ancestor(WordXml.W, "tc").Exists) {
                RecordNativeCellBarTabDiagnostics(paragraph, options, source);
                RecordNativeCellTextBoxDiagnostic(paragraph, options, source);
            }

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
            EnumerateNativeContentControls(paragraph.Node(), WordContentControlLevel.Run).Any(sdtRun =>
                !IsNativeSimpleTextContentControl(sdtRun) &&
                !IsNativeCheckBoxControl(sdtRun) &&
                !IsNativeSupportedFormFieldContentControl(sdtRun) &&
                !IsNativePictureControlWithImage(paragraph, sdtRun) &&
                !IsNativeRepeatingSectionWithText(sdtRun) &&
                !IsNativeRepeatingSectionChildControl(sdtRun)) ||
            EnumerateNativeContentControls(paragraph.Node(), WordContentControlLevel.Block).Any(sdtBlock => !IsNativeSimpleBlockTextContentControl(sdtBlock)) ||
            EnumerateNativeContentControls(paragraph.Node(), WordContentControlLevel.Cell).Any();

        private static bool HasNativeUnsupportedBodyContentControl(WordParagraph paragraph, bool mapsCheckBoxes, bool mapsFormFields, bool mapsPictureControls, bool mapsRepeatingSections) =>
            (!mapsCheckBoxes && paragraph.IsCheckBox) ||
            (!mapsFormFields && (paragraph.IsDatePicker || paragraph.IsDropDownList || paragraph.IsComboBox)) ||
            (paragraph.IsPictureControl && (!mapsPictureControls || paragraph.PictureControl?.Image == null)) ||
            (paragraph.IsRepeatingSection && (!mapsRepeatingSections || paragraph.RepeatingSection?.TextItems.Count == 0)) ||
            EnumerateNativeContentControls(paragraph.Node(), WordContentControlLevel.Run).Any(sdtRun =>
                (!mapsCheckBoxes || !IsNativeCheckBoxControl(sdtRun)) &&
                (!mapsFormFields || !IsNativeSupportedFormFieldContentControl(sdtRun)) &&
                (!mapsPictureControls || !IsNativePictureControl(sdtRun)) &&
                (!mapsRepeatingSections || !IsNativeRepeatingSectionControl(sdtRun) && !IsNativeRepeatingSectionChildControl(sdtRun)) &&
                !IsNativeSimpleTextContentControl(sdtRun)) ||
            EnumerateNativeContentControls(paragraph.Node(), WordContentControlLevel.Block).Any(sdtBlock => !IsNativeSimpleBlockTextContentControl(sdtBlock)) ||
            EnumerateNativeContentControls(paragraph.Node(), WordContentControlLevel.Cell).Any();

        /// <summary>The <c>w:sdt</c> elements of one level below <paramref name="container"/>, in document order.</summary>
        private static IEnumerable<OfficeXmlNode> EnumerateNativeContentControls(OfficeXmlNode container, WordContentControlLevel level) =>
            container.ModelDescendants("sdt").Where(sdt => sdt.ContentControlLevel() == level);

        private static IReadOnlyList<OfficeXmlNode> GetNativeCheckBoxControls(WordParagraph paragraph) {
            OfficeXmlNode source = paragraph.Node();
            if (source.IsNone) {
                return Array.Empty<OfficeXmlNode>();
            }

            return EnumerateNativeContentControls(source, WordContentControlLevel.Run)
                .Where(IsNativeCheckBoxControl)
                .ToList();
        }

        private static bool IsNativeCheckBoxControl(OfficeXmlNode sdtRun) =>
            sdtRun.WChild("sdtPr").Child(WordXml.W14, "checkbox").Exists;

        private static bool IsNativePictureControl(OfficeXmlNode sdtRun) =>
            sdtRun.WChild("sdtPr").WChild("picture").Exists;

        private static bool IsNativePictureControlWithImage(WordParagraph paragraph, OfficeXmlNode sdtRun) {
            OfficeXmlNode source = paragraph.Node();
            if (!IsNativePictureControl(sdtRun) || source.IsNone) {
                return false;
            }

            WordParagraph pictureParagraph = paragraph._document.SdtRunView(source, sdtRun);
            return pictureParagraph.PictureControl?.Image != null;
        }

        private static IReadOnlyList<OfficeXmlNode> GetNativePictureControls(WordParagraph paragraph) {
            OfficeXmlNode source = paragraph.Node();
            if (source.IsNone) {
                return Array.Empty<OfficeXmlNode>();
            }
            var controls = new List<OfficeXmlNode>();
            var visibility = WordConverterNodes.FieldVisibilityForParagraph(source);
            CollectNativeVisiblePictureControls(source, visibility, controls);
            return controls;
        }

        private static void CollectNativeVisiblePictureControls(
            OfficeXmlNode container,
            WordComplexFieldRunVisibility visibility,
            List<OfficeXmlNode> controls) {
            foreach (OfficeXmlNode child in container.Elements()) {
                if (child.IsW("r")) {
                    visibility.ObserveRun(child);
                } else if (child.IsW("sdt")) {
                    if (visibility.IsVisible && IsNativePictureControl(child)) controls.Add(child);
                    OfficeXmlNode content = child.WChild("sdtContent");
                    if (content.Exists)
                        CollectNativeVisiblePictureControls(content, visibility, controls);
                } else if (child.IsW("fldSimple")) {
                    if (visibility.IsVisible)
                        CollectNativeVisiblePictureControls(child, visibility, controls);
                    else
                        visibility.ObserveDescendantRuns(child);
                } else if (child.IsW("hyperlink") || child.IsW("customXml")) {
                    CollectNativeVisiblePictureControls(child, visibility, controls);
                }
            }
        }

        private static bool IsNativeRepeatingSectionControl(OfficeXmlNode sdtRun) =>
            sdtRun.WChild("sdtPr").Child(WordXml.W15, "repeatingSection").Exists;

        private static bool IsNativeRepeatingSectionChildControl(OfficeXmlNode sdtRun) =>
            sdtRun.Ancestors().Any(ancestor => ancestor.ContentControlLevel() == WordContentControlLevel.Run &&
                IsNativeRepeatingSectionControl(ancestor));

        private static bool IsNativeRepeatingSectionWithText(OfficeXmlNode sdtRun) =>
            IsNativeRepeatingSectionControl(sdtRun) &&
            GetNativeRepeatingSectionItems(sdtRun).Count > 0;

        private static IReadOnlyList<OfficeXmlNode> GetNativeFormFieldControls(WordParagraph paragraph) {
            OfficeXmlNode source = paragraph.Node();
            if (source.IsNone) {
                return Array.Empty<OfficeXmlNode>();
            }

            return EnumerateNativeContentControls(source, WordContentControlLevel.Run)
                .Where(IsNativeSupportedFormFieldContentControl)
                .ToList();
        }

        private static IReadOnlyList<OfficeXmlNode> GetNativeRepeatingSectionControls(WordParagraph paragraph) {
            OfficeXmlNode source = paragraph.Node();
            if (source.IsNone) {
                return Array.Empty<OfficeXmlNode>();
            }

            return EnumerateNativeContentControls(source, WordContentControlLevel.Run)
                .Where(IsNativeRepeatingSectionControl)
                .ToList();
        }

        private static bool IsNativeSupportedFormFieldContentControl(OfficeXmlNode sdtRun) =>
            IsNativeDatePickerControl(sdtRun) ||
            GetNativeChoiceFieldOptions(sdtRun).Count > 0;

        private static bool IsNativeDatePickerControl(OfficeXmlNode sdtRun) =>
            sdtRun.WChild("sdtPr").WChild("date").Exists;

        // A drop-down list's items, else a combo box's: a present drop-down list wins even when empty.
        private static IEnumerable<OfficeXmlNode> GetNativeChoiceFieldItems(OfficeXmlNode properties) {
            OfficeXmlNode list = properties.WChild("dropDownList");
            if (list.IsNone) list = properties.WChild("comboBox");
            return list.WElements("listItem");
        }

        private static IReadOnlyList<string> GetNativeChoiceFieldOptions(OfficeXmlNode sdtRun) {
            OfficeXmlNode properties = sdtRun.WChild("sdtPr");
            if (properties.IsNone) {
                return Array.Empty<string>();
            }

            var options = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode item in GetNativeChoiceFieldItems(properties)) {
                string? option = item.WAttr("displayText") ?? item.WAttr("value");
                if (string.IsNullOrWhiteSpace(option) || !seen.Add(option!)) {
                    continue;
                }

                options.Add(option!);
            }

            return options;
        }

        private static string? GetNativeChoiceFieldValue(OfficeXmlNode sdtRun, IReadOnlyList<string> options) {
            string? lastValue = sdtRun.WChild("sdtPr").WChild("comboBox").WAttr("lastValue");
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

        private static string? GetNativeChoiceDisplayValue(OfficeXmlNode sdtRun, string value) {
            OfficeXmlNode match = GetNativeChoiceFieldItems(sdtRun.WChild("sdtPr"))
                .FirstOrDefault(item => string.Equals(item.WAttr("value"), value, StringComparison.Ordinal));
            return match.WAttr("displayText") ?? match.WAttr("value");
        }

        private static string GetNativeDatePickerValue(OfficeXmlNode sdtRun) {
            string? fullDate = sdtRun.WChild("sdtPr").WChild("date").WAttr("fullDate");
            if (TryParseNativeXmlDateTime(fullDate, out DateTime value)) {
                return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return GetNativeSdtText(sdtRun) ?? string.Empty;
        }

        // An xsd:dateTime value, read as the SDK's DateTimeValue reads it.
        private static bool TryParseNativeXmlDateTime(string? text, out DateTime value) {
            value = default;
            if (text == null) return false;
            try {
                value = XmlConvert.ToDateTime(text, XmlDateTimeSerializationMode.RoundtripKind);
                return true;
            } catch (FormatException) {
                return false;
            }
        }

        private static string? GetNativeSdtText(OfficeXmlNode sdtRun) {
            OfficeXmlNode content = sdtRun.WChild("sdtContent");
            if (content.IsNone) {
                return null;
            }

            string text = string.Concat(content.Descendants(WordXml.W, "t").Select(WordConverterNodes.VisibleText));
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private static bool IsNativeSimpleTextContentControl(OfficeXmlNode sdtRun) {
            OfficeXmlNode properties = sdtRun.WChild("sdtPr");
            if (properties.IsNone) {
                return false;
            }

            if (IsNativeBuiltInPropertyContentControl(properties)) {
                return true;
            }

            if (properties.Child(WordXml.W14, "checkbox").Exists ||
                properties.WChild("date").Exists ||
                properties.WChild("dropDownList").Exists ||
                properties.WChild("comboBox").Exists ||
                properties.WChild("picture").Exists ||
                properties.Child(WordXml.W15, "repeatingSection").Exists) {
                return false;
            }

            return sdtRun.WChild("sdtContent").HasDescendant(WordXml.W, "t");
        }

        private static bool IsNativeSimpleBlockTextContentControl(OfficeXmlNode sdtBlock) {
            OfficeXmlNode properties = sdtBlock.WChild("sdtPr");
            if (properties.IsNone) {
                return false;
            }

            if (IsNativeBuiltInPropertyContentControl(properties) ||
                IsNativeSupportedDocPartBlockContentControl(sdtBlock)) {
                return true;
            }

            if (properties.WChild("docPartObj").Exists ||
                properties.WChild("date").Exists ||
                properties.WChild("picture").Exists) {
                return false;
            }

            return sdtBlock.WChild("sdtContent").HasDescendant(WordXml.W, "t");
        }

        private static bool IsNativeSupportedDocPartBlockContentControl(OfficeXmlNode sdtBlock) {
            string? gallery = sdtBlock.WChild("sdtPr").WChild("docPartObj").WChild("docPartGallery").WVal();
            return gallery != null &&
                   (gallery.Equals("Cover Pages", StringComparison.OrdinalIgnoreCase) ||
                    gallery.Equals("Table of Contents", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsNativeBuiltInPropertyContentControl(OfficeXmlNode properties) {
            string key = (properties.WChild("alias").WVal() ?? string.Empty).Trim();
            string? xPath = properties.WChild("dataBinding").WAttr("xpath");
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

        private static bool IsNativeCheckBoxChecked(OfficeXmlNode sdtRun) =>
            string.Equals(sdtRun.WChild("sdtPr").Child(WordXml.W14, "checkbox").Child(WordXml.W14, "checked")
                .Attr(WordXml.W14, "val"), "1", StringComparison.Ordinal);

        private static string GetNativeCheckBoxFieldName(OfficeXmlNode sdtRun, int index, string fallbackPrefix = "WordCheckBox") {
            return GetNativeContentControlFieldName(sdtRun, index, fallbackPrefix);
        }

        private static string GetNativeContentControlFieldName(OfficeXmlNode sdtRun, int index, string fallbackPrefix) {
            OfficeXmlNode properties = sdtRun.WChild("sdtPr");
            string? tag = properties.WChild("tag").WVal();
            if (!string.IsNullOrWhiteSpace(tag)) {
                return tag!;
            }

            string? alias = properties.WChild("alias").WVal();
            if (!string.IsNullOrWhiteSpace(alias)) {
                return alias!;
            }

            int? sdtId = WordXml.ParseInt32(properties.WChild("id").WVal());
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
