using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Encodes ignored-error ranges as an IGNOREECS collection (MS-XLSB 2.4.16, 2.4.322, 2.4.373).</summary>
    internal static class XlsbWorksheetIgnoredErrorWriter {
        internal const int BrtBeginCellIgnoreECs = 648;
        internal const int BrtCellIgnoreEC = 649;
        internal const int BrtEndCellIgnoreECs = 650;
        private const int MaximumRanges = 8191;

        // BrtCellIgnoreEC flag bits from the least significant bit, paired with their CT_IgnoredError attributes.
        internal static readonly (string Attribute, uint Flag, Func<IgnoredError, BooleanValue?> Get, Action<IgnoredError> Set)[] Flags = {
            ("evalError", 0x001U, item => item.EvalError, item => item.EvalError = true),
            ("emptyCellReference", 0x002U, item => item.EmptyCellReference, item => item.EmptyCellReference = true),
            ("numberStoredAsText", 0x004U, item => item.NumberStoredAsText, item => item.NumberStoredAsText = true),
            ("formulaRange", 0x008U, item => item.FormulaRange, item => item.FormulaRange = true),
            ("formula", 0x010U, item => item.Formula, item => item.Formula = true),
            ("twoDigitTextYear", 0x020U, item => item.TwoDigitTextYear, item => item.TwoDigitTextYear = true),
            ("unlockedFormula", 0x040U, item => item.UnlockedFormula, item => item.UnlockedFormula = true),
            ("listDataValidation", 0x080U, item => item.ListDataValidation, item => item.ListDataValidation = true),
            ("calculatedColumn", 0x100U, item => item.CalculatedColumn, item => item.CalculatedColumn = true),
        };

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) =>
            records.IgnoredErrors.AddRange(Create(context));

        internal static IReadOnlyList<XlsbGeneratedRecord> Create(XlsbWorksheetFeatureContext context) {
            IgnoredErrors? collection = context.Worksheet.GetFirstChild<IgnoredErrors>();
            if (collection == null) return Array.Empty<XlsbGeneratedRecord>();
            string sheetName = context.Sheet.Name;
            if (collection.ChildElements.Any(child => child is not IgnoredError)
                || XlsbOpenXmlAttributeValidator.FindUnsupported(collection, new[] { "count" }) is OpenXmlAttribute) {
                throw new NotSupportedException($"Native XLSB generation does not support ignored-error extensions on worksheet '{sheetName}'.");
            }
            IgnoredError[] items = collection.Elements<IgnoredError>().ToArray();
            if (items.Length == 0) return Array.Empty<XlsbGeneratedRecord>();

            var result = new List<XlsbGeneratedRecord>(items.Length + 2) {
                new XlsbGeneratedRecord(BrtBeginCellIgnoreECs, Array.Empty<byte>())
            };
            string[] allowed = Flags.Select(flag => flag.Attribute).Concat(new[] { "sqref" }).ToArray();
            foreach (IgnoredError item in items) {
                if (item.HasChildren || XlsbOpenXmlAttributeValidator.FindUnsupported(item, allowed) is OpenXmlAttribute) {
                    throw new NotSupportedException($"Native XLSB generation found an unsupported ignored-error entry on worksheet '{sheetName}'.");
                }
                uint flags = 0;
                foreach (var entry in Flags) {
                    if (entry.Get(item)?.Value == true) flags |= entry.Flag;
                }
                using var payload = new MemoryStream();
                XlsbPayloadWriter.WriteUInt32(payload, flags);
                XlsbPayloadWriter.WriteSqRfX(payload, item.SequenceOfReferences?.InnerText, sheetName, "ignored-error", MaximumRanges);
                result.Add(new XlsbGeneratedRecord(BrtCellIgnoreEC, payload.ToArray()));
            }
            result.Add(new XlsbGeneratedRecord(BrtEndCellIgnoreECs, Array.Empty<byte>()));
            return result;
        }
    }
}
