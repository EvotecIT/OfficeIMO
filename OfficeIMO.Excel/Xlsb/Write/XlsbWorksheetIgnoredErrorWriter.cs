using OfficeIMO.Core.Internal.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Encodes ignored-error ranges as an IGNOREECS collection (MS-XLSB 2.4.16, 2.4.322, 2.4.373).</summary>
    internal static class XlsbWorksheetIgnoredErrorWriter {
        internal const int BrtBeginCellIgnoreECs = 648;
        internal const int BrtCellIgnoreEC = 649;
        internal const int BrtEndCellIgnoreECs = 650;
        private const int MaximumRanges = 8191;

        // BrtCellIgnoreEC flag bits from the least significant bit, paired with their CT_IgnoredError attributes.
        internal static readonly (string Attribute, uint Flag, Action<IgnoredError> Set)[] Flags = {
            ("evalError", 0x001U, item => item.EvalError = true),
            ("emptyCellReference", 0x002U, item => item.EmptyCellReference = true),
            ("numberStoredAsText", 0x004U, item => item.NumberStoredAsText = true),
            ("formulaRange", 0x008U, item => item.FormulaRange = true),
            ("formula", 0x010U, item => item.Formula = true),
            ("twoDigitTextYear", 0x020U, item => item.TwoDigitTextYear = true),
            ("unlockedFormula", 0x040U, item => item.UnlockedFormula = true),
            ("listDataValidation", 0x080U, item => item.ListDataValidation = true),
            ("calculatedColumn", 0x100U, item => item.CalculatedColumn = true),
        };

        private static readonly string[] CollectionAttributes = { "count" };
        private static readonly string[] EntryAttributes = Flags.Select(flag => flag.Attribute).Concat(new[] { "sqref" }).ToArray();

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) =>
            records.IgnoredErrors.AddRange(Create(context));

        internal static IReadOnlyList<XlsbGeneratedRecord> Create(XlsbWorksheetFeatureContext context) {
            OfficeXmlNode collection = XlsbWorksheetMarkup.Child(context.WorksheetMarkup, "ignoredErrors");
            if (collection.IsNone) return Array.Empty<XlsbGeneratedRecord>();
            string sheetName = context.Sheet.Name;
            List<OfficeXmlNode> items = collection.Children(OfficeXmlNamespaces.SpreadsheetML, "ignoredError").ToList();
            if (collection.Children().Count != items.Count
                || XlsbWorkbookMarkup.UnsupportedAttribute(collection, CollectionAttributes) != null) {
                throw new NotSupportedException($"Native XLSB generation does not support ignored-error extensions on worksheet '{sheetName}'.");
            }
            if (items.Count == 0) return Array.Empty<XlsbGeneratedRecord>();

            var result = new List<XlsbGeneratedRecord>(items.Count + 2) {
                new XlsbGeneratedRecord(BrtBeginCellIgnoreECs, Array.Empty<byte>())
            };
            foreach (OfficeXmlNode item in items) {
                if (item.FirstChild.Exists || XlsbWorkbookMarkup.UnsupportedAttribute(item, EntryAttributes) != null) {
                    throw new NotSupportedException($"Native XLSB generation found an unsupported ignored-error entry on worksheet '{sheetName}'.");
                }
                uint flags = 0;
                foreach (var entry in Flags) {
                    if (XlsbWorkbookMarkup.ReadBoolean(item, entry.Attribute) == true) flags |= entry.Flag;
                }
                using var payload = new MemoryStream();
                XlsbPayloadWriter.WriteUInt32(payload, flags);
                XlsbPayloadWriter.WriteSqRfX(payload, item.Attribute("sqref"), sheetName, "ignored-error", MaximumRanges);
                result.Add(new XlsbGeneratedRecord(BrtCellIgnoreEC, payload.ToArray()));
            }
            result.Add(new XlsbGeneratedRecord(BrtEndCellIgnoreECs, Array.Empty<byte>()));
            return result;
        }
    }
}
