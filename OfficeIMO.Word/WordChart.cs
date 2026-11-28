using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides functionality for creating and manipulating charts within a
    /// <see cref="WordDocument"/>.
    /// </summary>
    public partial class WordChart : WordElement {
        private static int _axisIdSeed = 148921728;
        private bool _ownsChartDefaults;

        /// <summary>
        /// Initializes a <see cref="WordChart"/> instance from an existing drawing.
        /// </summary>
        /// <param name="document">Parent document that owns the chart.</param>
        /// <param name="paragraph">Paragraph containing the chart.</param>
        /// <param name="drawing">Existing <c>w:drawing</c> element with chart data.</param>
        internal WordChart(WordDocument document, WordParagraph paragraph, OfficeXmlNode drawing) {
            _document = document;
            _drawing = drawing;
            _paragraph = paragraph;
            _chart = ChartSpaceNode.IsNone ? default : ChartSpaceNode.Child(NsC, "chart");
        }

        /// <summary>
        /// Creates a new chart and inserts it into the specified paragraph.
        /// </summary>
        /// <param name="document">Parent document that will contain the chart.</param>
        /// <param name="paragraph">Paragraph where the chart will be placed.</param>
        /// <param name="title">Optional chart title.</param>
        /// <param name="roundedCorners">Specifies whether the chart frame should have rounded corners.</param>
        /// <param name="width">Initial chart width; defaults to 600 pixels.</param>
        /// <param name="height">Initial chart height; defaults to 600 pixels.</param>
        public WordChart(WordDocument document, WordParagraph paragraph, string title = "", bool roundedCorners = false, OfficeLength? width = null, OfficeLength? height = null) {
            _document = document;
            _paragraph = paragraph;
            SetTitle(title);
            var frameWidth = width ?? OfficeLength.Pixels(600);
            var frameHeight = height ?? OfficeLength.Pixels(600);
            if (frameWidth.ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            if (frameHeight.ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(height));
            InsertChart(paragraph, roundedCorners, frameWidth, frameHeight);
        }

        /// <summary>The next series index: one past the highest <c>c:idx</c> in the plot area.</summary>
        private uint _index {
            get {
                OfficeXmlNode plotArea = PlotAreaOf(_chart);
                if (plotArea.IsNone) return 0U;
                bool any = false;
                uint max = 0U;
                foreach (OfficeXmlNode index in PlotIndexes(plotArea)) {
                    uint value = index.UIntVal() ?? 0U;
                    if (!any || value > max) max = value;
                    any = true;
                }
                return any ? max + 1U : 0U;
            }
        }

        /// <summary>
        /// The <c>c:idx</c> elements of a plot area where the chart schema puts them: in the series of its chart groups,
        /// their data points and data labels, the groups' own data labels and surface band formats. Markup elsewhere,
        /// such as in vendor elements, is not an index.
        /// </summary>
        private static IEnumerable<OfficeXmlNode> PlotIndexes(OfficeXmlNode plotArea) {
            for (OfficeXmlNode group = plotArea.FirstChild; !group.IsNone; group = group.NextSibling) {
                if (group.NamespaceId != NsC || !IsChartGroup(group)) continue;
                for (OfficeXmlNode child = group.FirstChild; !child.IsNone; child = child.NextSibling) {
                    if (child.Is(NsC, "ser")) {
                        for (OfficeXmlNode item = child.FirstChild; !item.IsNone; item = item.NextSibling) {
                            if (item.Is(NsC, "idx")) yield return item;
                            else if (item.Is(NsC, "dPt") && !item.Child(NsC, "idx").IsNone) yield return item.Child(NsC, "idx");
                            else if (item.Is(NsC, "dLbls")) foreach (OfficeXmlNode index in LabelIndexes(item)) yield return index;
                        }
                    } else if (child.Is(NsC, "dLbls")) {
                        foreach (OfficeXmlNode index in LabelIndexes(child)) yield return index;
                    } else if (child.Is(NsC, "bandFmts")) {
                        foreach (OfficeXmlNode band in child.Children(NsC, "bandFmt")) {
                            if (!band.Child(NsC, "idx").IsNone) yield return band.Child(NsC, "idx");
                        }
                    }
                }
            }
        }

        private static IEnumerable<OfficeXmlNode> LabelIndexes(OfficeXmlNode labels) {
            foreach (OfficeXmlNode label in labels.Children(NsC, "dLbl")) {
                if (!label.Child(NsC, "idx").IsNone) yield return label.Child(NsC, "idx");
            }
        }

        private OfficeXmlNode AddCategoryAxisInternal(OfficeXmlNode plotArea, uint axisId, uint crossingAxis, string position) {
            OfficeXmlNode categoryAxis = NewC(plotArea, "catAx");
            categoryAxis.DeclareNamespace("c", ChartNamespaceUri);
            AddC(categoryAxis, "axId", axisId);
            AddC(AddC(categoryAxis, "scaling"), "orientation", "minMax");
            AddC(categoryAxis, "delete", false);
            AddC(categoryAxis, "axPos", position);
            if (!string.IsNullOrEmpty(_xAxisTitle)) {
                categoryAxis.Append(AddAxisTitle(categoryAxis, _xAxisTitle!));
            }
            AddC(categoryAxis, "majorTickMark", "out");
            AddC(categoryAxis, "minorTickMark", "none");
            AddC(categoryAxis, "tickLblPos", "nextTo");
            AddC(categoryAxis, "crossAx", crossingAxis);
            AddC(categoryAxis, "crosses", "autoZero");
            AddC(categoryAxis, "auto", true);
            AddC(categoryAxis, "lblAlgn", "ctr");
            AddC(categoryAxis, "lblOffset", "100");
            AddC(categoryAxis, "noMultiLvlLbl", false);
            return categoryAxis;
        }

        private OfficeXmlNode AddValueAxisInternal(OfficeXmlNode plotArea, uint axisId, uint crossingAxis, string position) {
            OfficeXmlNode valueAxis = NewC(plotArea, "valAx");
            valueAxis.DeclareNamespace("c", ChartNamespaceUri);
            AddC(valueAxis, "axId", axisId);
            AddC(AddC(valueAxis, "scaling"), "orientation", "minMax");
            AddC(valueAxis, "delete", false);
            AddC(valueAxis, "axPos", position);
            AddC(valueAxis, "majorGridlines");
            if (!string.IsNullOrEmpty(_yAxisTitle)) {
                valueAxis.Append(AddAxisTitle(valueAxis, _yAxisTitle!));
            }
            OfficeXmlNode numberFormat = AddC(valueAxis, "numFmt");
            numberFormat.SetAttribute("formatCode", "General");
            numberFormat.SetAttribute("sourceLinked", OfficeChartXml.Text(false));
            AddC(valueAxis, "majorTickMark", "out");
            AddC(valueAxis, "minorTickMark", "none");
            AddC(valueAxis, "tickLblPos", "nextTo");
            AddC(valueAxis, "crossAx", crossingAxis);
            AddC(valueAxis, "crosses", "autoZero");
            AddC(valueAxis, "crossBetween", "between");
            return valueAxis;
        }

        /// <summary>A detached <c>c:dLbls</c> for a chart group: values and percentages shown, the rest hidden.</summary>
        private static OfficeXmlNode AddDataLabel(OfficeXmlNode group) {
            OfficeXmlNode dataLabels = NewC(group, "dLbls");
            dataLabels.DeclareNamespace("c", ChartNamespaceUri);
            AddC(dataLabels, "showLegendKey", false);
            // Show numeric values by default to make charts executive‑friendly
            AddC(dataLabels, "showVal", true);
            AddC(dataLabels, "showCatName", false);
            AddC(dataLabels, "showSerName", false);
            // For pie charts, percent labels are helpful; harmless on others
            AddC(dataLabels, "showPercent", true);
            AddC(dataLabels, "showBubbleSize", false);
            AddC(dataLabels, "showLeaderLines", false);
            return dataLabels;
        }

        private void InsertChart(WordParagraph paragraph, bool roundedCorners, OfficeLength width, OfficeLength height) {
            _ownsChartDefaults = true;
            (string id, OfficeXmlNode space) = AddChartPart(paragraph);
            AddC(space, "roundedCorners", roundedCorners);

            OfficeXmlNode paragraphNode = ParagraphNode(paragraph);
            OfficeXmlNode run = paragraphNode.CreateElement(OfficeXmlNamespaces.WordprocessingML, "r");
            OfficeXmlNode drawing = CreateChartDrawing(run, id, width, height);
            run.Append(drawing);
            run = paragraphNode.Append(run);
            _drawing = run.Child(OfficeXmlNamespaces.WordprocessingML, "drawing");
        }

        // New charts have no chart-area frame or fill. Without c:spPr, Word and Excel draw a
        // border around the chart; the document's own layout already separates it.
        private void EnsureFramelessChartArea() {
            OfficeXmlNode space = ChartSpaceNode;
            OfficeXmlNode chart = space.IsNone ? default : space.Child(NsC, "chart");
            if (chart.IsNone) return;
            if (_ownsChartDefaults) {
                OfficeOpenXmlChartThemeWriter.Apply(SharedChartPart!, OfficeTheme.Default);
                _ownsChartDefaults = false;
            }
            if (!space.Child(NsC, "spPr").IsNone) return;
            OfficeXmlNode properties = NewC(space, "spPr");
            AddA(properties, "noFill");
            AddA(AddA(properties, "ln"), "noFill");
            chart.InsertAfterSelf(properties);
        }

        /// <summary>A detached <c>c:chart</c> with an empty plot area and the chart-level defaults.</summary>
        private OfficeXmlNode GenerateChart(string title = "") {
            OfficeXmlNode chart = NewC(ChartSpaceNode, "chart");

            // Add elements in the correct OpenXML schema order
            // 1. Title (optional, but must be first if present)
            if (!string.IsNullOrEmpty(title)) {
                chart.Append(AddTitle(chart, title));
            }

            // 2. AutoTitleDeleted (must come after Title if Title exists)
            AddC(chart, "autoTitleDeleted", string.IsNullOrEmpty(title));

            // 8. PlotArea (comes after View3D, Floor, SideWall, BackWall if they were present)
            AddC(AddC(chart, "plotArea"), "layout");

            // 10. PlotVisibleOnly
            AddC(chart, "plotVisOnly", true);

            // 11. DisplayBlanksAs
            AddC(chart, "dispBlanksAs", "gap");

            // 12. ShowDataLabelsOverMaximum
            AddC(chart, "showDLblsOverMax", false);

            return chart;
        }

        /// <summary>Appends a newly built chart to the chart space and makes it this instance's chart.</summary>
        private void AttachNewChart(OfficeXmlNode chart) {
            OfficeXmlNode space = ChartSpaceNode;
            _chart = space.IsNone ? chart : space.Append(chart);
        }

        /// <summary>
        /// Set the title of the chart
        /// </summary>
        /// <param name="title"></param>
        /// <returns></returns>
        public WordChart SetTitle(string title) {
            PrivateTitle = title;
            UpdateTitle();
            return this;
        }

        /// <summary>
        /// Update the title of the chart
        /// </summary>
        internal void UpdateTitle() {
            if (string.IsNullOrEmpty(PrivateTitle) || _chart.IsNone) return;
            OfficeXmlNode title = _chart.Child(NsC, "title");
            if (title.IsNone) {
                // Insert title at the beginning (before AutoTitleDeleted) to maintain schema order
                OfficeXmlNode newTitle = AddTitle(_chart, PrivateTitle);
                OfficeXmlNode autoTitleDeleted = _chart.Child(NsC, "autoTitleDeleted");
                if (!autoTitleDeleted.IsNone) {
                    autoTitleDeleted.SetAttribute("val", OfficeChartXml.Text(false));
                    _chart.InsertBefore(newTitle, autoTitleDeleted);
                } else {
                    // If no AutoTitleDeleted found, insert at the beginning
                    _chart.InsertBefore(newTitle, _chart.FirstChild);
                }
            } else {
                ReplaceTitleText(title, PrivateTitle);
            }
        }

        /// <summary>Replaces a saved title's text, keeping the formatting of its first run.</summary>
        /// <remarks>A title without literal text (a cell reference, its cache or an automatic title) gets a literal rich-text body.</remarks>
        private void ReplaceTitleText(OfficeXmlNode title, string text) {
            OfficeXmlNode chartText = title.Child(NsC, "tx");
            OfficeXmlNode rich = chartText.IsNone ? default : chartText.Child(NsC, "rich");
            OfficeXmlNode firstRun = rich.IsNone ? default : FirstTextRun(rich);
            if (!firstRun.IsNone) {
                OfficeXmlNode paragraph = firstRun.Parent;
                foreach (OfficeXmlNode other in rich.Children(NsA, "p").ToList()) {
                    if (other != paragraph) other.Remove();
                }
                foreach (OfficeXmlNode other in paragraph.Children().ToList()) {
                    if (other != firstRun && (other.Is(NsA, "r") || other.Is(NsA, "fld") || other.Is(NsA, "br"))) other.Remove();
                }
                OfficeXmlNode runText = firstRun.Child(NsA, "t");
                if (runText.IsNone) runText = AddA(firstRun, "t");
                runText.SetInnerText(text);
                return;
            }

            OfficeXmlNode literal = AddTitle(title.Parent.IsNone ? title : title.Parent, text).Child(NsC, "tx");
            literal.Remove();
            if (!chartText.IsNone) {
                chartText.InsertAfterSelf(literal);
                chartText.Remove();
            } else {
                title.InsertBefore(literal, title.FirstChild);
            }
        }

        /// <summary>
        /// The first run of a text body's paragraphs. Runs sit directly in <c>a:p</c>; an <c>a:r</c> anywhere else, such
        /// as in vendor markup, is not a text run.
        /// </summary>
        private static OfficeXmlNode FirstTextRun(OfficeXmlNode textBody) {
            foreach (OfficeXmlNode paragraph in textBody.Children(NsA, "p")) {
                OfficeXmlNode run = paragraph.Child(NsA, "r");
                if (!run.IsNone) return run;
            }
            return default;
        }

        /// <summary>A detached <c>c:title</c> holding <paramref name="title"/>, made for <paramref name="context"/>.</summary>
        private static OfficeXmlNode AddTitle(OfficeXmlNode context, string title) {
            OfficeXmlNode titleElement = NewC(context, "title");
            OfficeXmlNode rich = AddC(AddC(titleElement, "tx"), "rich");
            AddA(rich, "bodyPr");
            AddA(rich, "lstStyle");
            OfficeXmlNode paragraph = AddA(rich, "p");
            AddA(AddA(paragraph, "pPr"), "defRPr");
            AddText(AddA(paragraph, "r"), NsA, "t", title);
            AddC(titleElement, "overlay", false);
            OfficeOpenXmlChartThemeWriter.ApplyTitle(titleElement, OfficeTheme.Default);
            return titleElement;
        }

        /// <summary>A detached axis <c>c:title</c> in the axis-title font, size and color, made for <paramref name="axis"/>.</summary>
        private OfficeXmlNode AddAxisTitle(OfficeXmlNode axis, string titleText) {
            OfficeXmlNode title = NewC(axis, "title");
            OfficeXmlNode rich = AddC(AddC(title, "tx"), "rich");
            AddA(rich, "bodyPr");
            AddA(rich, "lstStyle");
            OfficeXmlNode paragraph = AddA(rich, "p");
            OfficeXmlNode runProperties = AddA(AddA(paragraph, "pPr"), "defRPr");
            runProperties.SetAttribute("sz", OfficeChartXml.Text(_axisTitleFontSize));
            runProperties.Append(CreateChartFill(runProperties, _axisTitleColor));
            AddA(runProperties, "latin").SetAttribute("typeface", _axisTitleFontName);
            AddText(AddA(paragraph, "r"), NsA, "t", titleText);
            AddC(title, "overlay", false);
            return title;
        }

        /// <summary>Applies the axis titles set on this instance without discarding titles it has not set.</summary>
        /// <param name="restyleExisting">Applies the current axis-title font, size and color to saved titles in place.</param>
        internal void UpdateAxisTitles(bool restyleExisting = false) {
            OfficeXmlNode plot = PlotAreaOf(_chart);
            if (plot.IsNone) return;
            OfficeXmlNode categoryAxis = plot.Child(NsC, "catAx");
            if (!categoryAxis.IsNone) UpdateAxisTitle(categoryAxis, _xAxisTitle, restyleExisting, CCatAxView.Order);
            OfficeXmlNode valueAxis = plot.Child(NsC, "valAx");
            if (!valueAxis.IsNone) UpdateAxisTitle(valueAxis, _yAxisTitle, restyleExisting, CValAxView.Order);
        }

        private void UpdateAxisTitle(OfficeXmlNode axis, string? assignedTitle, bool restyleExisting, OfficeXmlChildOrder order) {
            OfficeXmlNode existing = axis.Child(NsC, "title");
            if (assignedTitle == null) {
                // A title this instance has not assigned belongs to the saved chart: keep its text, paragraphs,
                // layout and run formatting, and change only the font, size and color when asked to.
                if (restyleExisting && !existing.IsNone) RestyleAxisTitle(existing);
                return;
            }
            if (!existing.IsNone) existing.Remove();
            if (assignedTitle.Length == 0) return;
            // The title goes in schema order, after any major and minor gridlines.
            OfficeChartXml.SetOrdered(axis, AddAxisTitle(axis, assignedTitle), order);
        }

        private void RestyleAxisTitle(OfficeXmlNode title) {
            // Rich text carries its own paragraphs. A title that shows a cell reference or its cache is
            // formatted through c:txPr, which is created when the saved title has none.
            OfficeXmlNode chartText = title.Child(NsC, "tx");
            OfficeXmlNode textBody = chartText.IsNone ? default : chartText.Child(NsC, "rich");
            if (textBody.IsNone) textBody = GetOrCreateTitleTextProperties(title);
            if (textBody.Child(NsA, "p").IsNone) {
                AddA(AddA(textBody, "p"), "endParaRPr");
            }
            foreach (OfficeXmlNode paragraph in textBody.Children(NsA, "p").ToList()) {
                OfficeXmlNode paragraphProperties = paragraph.Child(NsA, "pPr");
                if (paragraphProperties.IsNone) paragraphProperties = paragraph.InsertBefore(paragraph.CreateElement(NsA, "pPr"), paragraph.FirstChild);
                OfficeXmlNode defaults = paragraphProperties.Child(NsA, "defRPr");
                if (defaults.IsNone) defaults = AddA(paragraphProperties, "defRPr");
                ApplyAxisTitleFormat(defaults);
                foreach (OfficeXmlNode run in paragraph.Children(NsA, "r").ToList()) {
                    ApplyAxisTitleFormat(GetOrPrependRunProperties(run));
                }
                foreach (OfficeXmlNode field in paragraph.Children(NsA, "fld").ToList()) {
                    ApplyAxisTitleFormat(GetOrPrependRunProperties(field));
                }
            }
        }

        /// <summary>The run's <c>a:rPr</c>, created as its first child when it has none.</summary>
        private static OfficeXmlNode GetOrPrependRunProperties(OfficeXmlNode run) {
            OfficeXmlNode properties = run.Child(NsA, "rPr");
            return properties.IsNone ? run.InsertBefore(run.CreateElement(NsA, "rPr"), run.FirstChild) : properties;
        }

        private static OfficeXmlNode GetOrCreateTitleTextProperties(OfficeXmlNode title) {
            OfficeXmlNode properties = title.Child(NsC, "txPr");
            if (!properties.IsNone) return properties;
            properties = NewC(title, "txPr");
            AddA(properties, "bodyPr");
            AddA(properties, "lstStyle");
            return OfficeChartXml.SetOrdered(title, properties, CTitleView.Order);
        }

        private void ApplyAxisTitleFormat(OfficeXmlNode properties) {
            properties.SetAttribute("sz", OfficeChartXml.Text(_axisTitleFontSize));
            foreach (string fill in FillNames) properties.RemoveChildren(NsA, fill);
            OfficeXmlNode existingLatin = properties.Child(NsA, "latin");
            if (!existingLatin.IsNone) existingLatin.Remove();
            OfficeXmlChildOrder order = DmlTextCharacterPropertiesView.Order;
            OfficeChartXml.SetOrdered(properties, CreateChartFill(properties, _axisTitleColor), order);
            OfficeXmlNode latin = properties.CreateElement(NsA, "latin");
            latin.SetAttribute("typeface", _axisTitleFontName);
            OfficeChartXml.SetOrdered(properties, latin, order);
        }

        /// <summary>A detached <c>c:val</c> of literal numbers, made for <paramref name="series"/>.</summary>
        internal OfficeXmlNode AddValuesAxisData<T>(OfficeXmlNode series, List<T> dataList) {
            OfficeXmlNode values = NewC(series, "val");
            OfficeXmlNode literal = AddC(values, "numLit");
            AddText(literal, NsC, "formatCode", "General");
            AddC(literal, "ptCount", (uint)dataList.Count);
            var index = 0;
            foreach (var data in dataList) {
                // Ensure decimal values use invariant culture (period as decimal separator) for OpenXML compatibility
                string valueText = data is double number ? OfficeIMO.Core.Internal.OfficeRoundTripText.Format(number, System.Globalization.CultureInfo.InvariantCulture)
                    : data is float single ? OfficeIMO.Core.Internal.OfficeRoundTripText.Format(single, System.Globalization.CultureInfo.InvariantCulture)
                    : data is System.IFormattable numeric
                    ? numeric.ToString(null, System.Globalization.CultureInfo.InvariantCulture)
                    : data?.ToString() ?? "0";
                AddPoint(literal, Convert.ToUInt32(index), valueText);
                index++;
            }
            return values;
        }

        /// <summary>A detached <c>c:cat</c> of literal strings, made for <paramref name="series"/>.</summary>
        internal OfficeXmlNode AddCategoryAxisData(OfficeXmlNode series, List<string> categories) {
            OfficeXmlNode categoryAxis = NewC(series, "cat");
            OfficeXmlNode stringLiteral = AddC(categoryAxis, "strLit");
            AddC(stringLiteral, "ptCount", (uint)categories.Count);
            int index = 0;
            foreach (string category in categories) {
                AddPoint(stringLiteral, Convert.ToUInt32(index), category);
                index++;
            }
            return categoryAxis;
        }

        /// <summary>Appends a <c>c:pt</c> with its index and <c>c:v</c> text to a literal or cache.</summary>
        private static OfficeXmlNode AddPoint(OfficeXmlNode literal, uint index, string text) =>
            literal.Append(NewPoint(literal, index, text));

        /// <summary>A detached <c>c:pt</c> with its index and <c>c:v</c> text, made for <paramref name="literal"/>.</summary>
        private static OfficeXmlNode NewPoint(OfficeXmlNode literal, uint index, string text) {
            OfficeXmlNode point = NewC(literal, "pt");
            point.SetAttribute("idx", OfficeChartXml.Text(index));
            AddText(point, NsC, "v", text);
            return point;
        }

        // Literal series names use c:v. A c:strRef needs a formula into an embedded
        // workbook; Word charts built here have none, and other consumers reject the whole
        // document when the reference formula is empty.
        private static void AddSeriesText(OfficeXmlNode series, string name) =>
            AddText(AddC(series, "tx"), NsC, "v", name);

        /// <summary>A detached inline <c>w:drawing</c> showing the chart <paramref name="id"/>, made for <paramref name="run"/>.</summary>
        internal OfficeXmlNode CreateChartDrawing(OfficeXmlNode run, string id, OfficeLength width, OfficeLength height) {
            OfficeXmlNode drawing = run.CreateElement(OfficeXmlNamespaces.WordprocessingML, "drawing");

            OfficeXmlNode inline = drawing.AppendNew(NsWp, "inline");
            inline.DeclareNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            OfficeXmlNode extent = inline.AppendNew(NsWp, "extent");
            extent.SetAttribute("cx", OfficeChartXml.Text(width.ToEmus()));
            extent.SetAttribute("cy", OfficeChartXml.Text(height.ToEmus()));
            OfficeXmlNode effectExtent = inline.AppendNew(NsWp, "effectExtent");
            effectExtent.SetAttribute("l", "0");
            effectExtent.SetAttribute("t", "0");
            effectExtent.SetAttribute("r", "19050");
            effectExtent.SetAttribute("b", "19050");
            OfficeXmlNode docProperties = inline.AppendNew(NsWp, "docPr");
            docProperties.SetAttribute("id", OfficeChartXml.Text(WordDrawingIdAllocator.Reserve(_document)));
            docProperties.SetAttribute("name", "chart");

            OfficeXmlNode graphic = AddA(inline, "graphic");
            graphic.DeclareNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            OfficeXmlNode graphicData = AddA(graphic, "graphicData");
            graphicData.SetAttribute("uri", ChartNamespaceUri);
            OfficeXmlNode chartReference = AddC(graphicData, "chart");
            chartReference.SetAttribute(NsR, "id", id);
            chartReference.DeclareNamespace("p6", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartReference.DeclareNamespace("c", ChartNamespaceUri);
            return drawing;
        }
    }
}
