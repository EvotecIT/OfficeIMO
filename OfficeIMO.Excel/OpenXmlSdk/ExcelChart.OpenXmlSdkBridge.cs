using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    /// <summary>
    /// What chart code still needs from the SDK while the SDK holds the chart part: element classes the SDK picks by
    /// context, which earlier output depends on, and the workbook theme colour resolver, which takes SDK elements. These
    /// go when charts are written to OfficeIMO's own package.
    /// </summary>
    public sealed partial class ExcelChart {
        /// <summary>
        /// New <c>c:dPt</c> and <c>c:dLbls</c> as the SDK's typed classes, also under series whose schema has no such child
        /// (surface series), where an element made by name would be an unknown element and send the package through the
        /// SDK's save.
        /// </summary>
        private static OfficeXmlNode NewDataPoint(OfficeXmlNode series) =>
            series.IsSdk ? OfficeXmlNode.FromSdk(new C.DataPoint()) : series.CreateElement(ExcelChartXml.C, "dPt");

        private static OfficeXmlNode NewDataLabels(OfficeXmlNode parent) =>
            parent.IsSdk ? OfficeXmlNode.FromSdk(new C.DataLabels()) : parent.CreateElement(ExcelChartXml.C, "dLbls");

        /// <summary>An ARGB literal as a colour, through <see cref="ExcelColor"/>.</summary>
        private static OfficeColor? ColorFromArgb(string argb) => ExcelColor.FromSpreadsheetColor(null, argb);

        /// <summary>The workbook theme's value for a <c>a:solidFill</c>, as worksheet colours resolve.</summary>
        private static string? ResolveImageExportSolidFill(OfficeXmlNode solidFill, ExcelPartRef workbookPart) =>
            ExcelThemeColorResolver.Resolve(solidFill.SdkElement as A.SolidFill, workbookPart.SdkPart as WorkbookPart);

        // The SDK reads an axis's c:spPr as ChartShapeProperties; only a c:spPr OfficeIMO added to an axis in memory
        // (SetAxisLine) is a ShapeProperties. The image export reads the axis line from the latter alone, so a loaded
        // chart's axis line style is not drawn.
        private static OfficeXmlNode ImageExportAxisShapeProperties(OfficeXmlNode axis) {
            foreach (OfficeXmlNode properties in axis.Children(ExcelChartXml.C, "spPr")) {
                if (properties.SdkElement is C.ShapeProperties) return properties;
            }
            return OfficeXmlNode.None;
        }

        /// <summary>
        /// The axis code looks up <c>c:spPr</c> and <c>c:extLst</c> by SDK element class: the SDK reads an axis' <c>c:spPr</c>
        /// as <c>ChartShapeProperties</c> and its <c>c:extLst</c> as an axis-specific extension list. So an axis line goes
        /// into a separate <c>ShapeProperties</c> element unless an earlier axis line call made one, and the image export
        /// reads only that element.
        /// </summary>
        private static class AxisShapePropertiesQuirk {
            /// <summary>The axis' <c>c:spPr</c> made by an axis line call (SDK <c>ShapeProperties</c>), or none.</summary>
            internal static OfficeXmlNode AuthoredShapeProperties(OfficeXmlNode axis) =>
                OfficeXmlNode.FromSdk(((OpenXmlElement)axis.SdkElement!).GetFirstChild<C.ShapeProperties>());

            /// <summary>The axis' <c>c:spPr</c> as the SDK reads it (SDK <c>ChartShapeProperties</c>), or none.</summary>
            internal static OfficeXmlNode ChartShapeProperties(OfficeXmlNode axis) =>
                OfficeXmlNode.FromSdk(((OpenXmlElement)axis.SdkElement!).GetFirstChild<C.ChartShapeProperties>());

            /// <summary>A generic chart <c>c:extLst</c> under the axis, which the SDK never reads there, or none.</summary>
            internal static OfficeXmlNode GenericExtensionList(OfficeXmlNode axis) =>
                OfficeXmlNode.FromSdk(((OpenXmlElement)axis.SdkElement!).GetFirstChild<C.ExtensionList>());

            /// <summary>A new, detached SDK <c>ShapeProperties</c> for an axis line.</summary>
            internal static OfficeXmlNode NewAuthoredShapeProperties() => OfficeXmlNode.FromSdk(new C.ShapeProperties());

            /// <summary>Inserts the axis line's <c>ShapeProperties</c> keeping its SDK class, before the reference or last.</summary>
            internal static void InsertAuthoredShapeProperties(OfficeXmlNode axis, OfficeXmlNode properties, OfficeXmlNode insertBefore) {
                var target = (OpenXmlElement)axis.SdkElement!;
                var element = (OpenXmlElement)properties.SdkElement!;
                if (insertBefore.Exists) target.InsertBefore(element, (OpenXmlElement)insertBefore.SdkElement!);
                else target.AppendChild(element);
            }
        }
    }

    public partial class ExcelDocument {
        // Formula references and sheets for the chart data sheet, from the SDK workbook model.
        private IEnumerable<MutationFormulaContext> EnumerateMutationFormulaContexts(
            IReadOnlyList<OfficeXmlNode> sheets, int editedSheetIndex, string? excludedDefinedName) {
            var sdkSheets = new List<DocumentFormat.OpenXml.Spreadsheet.Sheet>(sheets.Count);
            foreach (OfficeXmlNode sheet in sheets) sdkSheets.Add((DocumentFormat.OpenXml.Spreadsheet.Sheet)sheet.SdkElement!);
            return EnumerateMutationFormulaContexts(sdkSheets, editedSheetIndex, excludedDefinedName);
        }

        private static IEnumerable<string> EnumerateMutationFormulaTexts(OfficeXmlNode root) =>
            EnumerateMutationFormulaTexts(root.SdkElement as OpenXmlPartRootElement);

        private ExcelSheet WrapChartDataSheet(OfficeXmlNode sheet) =>
            new ExcelSheet(this, _spreadSheetDocument, (DocumentFormat.OpenXml.Spreadsheet.Sheet)sheet.SdkElement!);

        private ExcelSheet CreateChartDataWorksheet(string name) =>
            new ExcelSheet(this, _workBookPart, _spreadSheetDocument, name);

        // The text a plain number would be stored as in its cell; false when the value does not store as a number.
        private bool TryCoerceNumberCellText(object value, out string? text) {
            var (cellValue, cellType) = CoerceValueHelper.Coerce(value, static shared => new DocumentFormat.OpenXml.Spreadsheet.CellValue(shared), dateSystem: DateSystem);
            text = cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number ? cellValue.Text : null;
            return cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
        }
    }

    public partial class ExcelSheet {
        // The ChartEx frame reference and the SVG blip are the SDK's typed elements, as the code built them before, so code
        // that reads them by SDK class finds them before the package is saved and read again.
        private static OfficeXmlNode NewChartExReference(OfficeXmlNode context, string relationshipId) {
            if (context.IsSdk) return OfficeXmlNode.FromSdk(new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId { Id = relationshipId });
            OfficeXmlNode reference = context.CreateElement(ExcelChartXml.ChartExUri, "chart", "cx");
            reference.SetAttribute(ExcelChartXml.R, "id", relationshipId);
            return reference;
        }

        private static OfficeXmlNode NewSvgBlip(OfficeXmlNode context, string? relationshipId) {
            if (context.IsSdk) return OfficeXmlNode.FromSdk(new DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip { Embed = relationshipId });
            OfficeXmlNode blip = context.CreateElement("http://schemas.microsoft.com/office/drawing/2016/SVG/main", "svgBlip", "asvg");
            blip.SetAttribute(ExcelChartXml.R, "embed", relationshipId);
            return blip;
        }

        /// <summary>The SDK image part type for an image content type, for header/footer and authored-source images.</summary>
        private static PartTypeInfo ToImagePartType(string? contentType) {
            var (partContentType, extension) = ToImagePartContentType(contentType);
            return new PartTypeInfo(partContentType, extension);
        }
    }
}
