using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelNativeFormulaOperandTypeTests {
        [Theory]
        [InlineData("excel-functions.xls", false, ExcelFileFormat.Xls)]
        [InlineData("excel-functions.encrypted.xls", true, ExcelFileFormat.Xls)]
        [InlineData("excel-functions.xlsb", false, ExcelFileFormat.Xlsb)]
        [InlineData("excel-functions.encrypted.xlsb", true, ExcelFileFormat.Xlsb)]
        public void Native_Excel_formula_results_survive_a_preserving_encrypted_save(string file, bool encrypted, ExcelFileFormat format) {
            string corpus = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelFormulaOperandCorpus");
            using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(corpus, "manifest.json")));
            using var input = File.OpenRead(Path.Combine(corpus, file));
            using ExcelDocument document = encrypted ? ExcelDocument.LoadEncrypted(input, "test-password") : ExcelDocument.Load(input);
            Check(document);
            document.GetSheet("Data").CellValue(10, 4, "Preserving edit");
            using var output = new MemoryStream(); document.SaveEncrypted(output, "rotated-pass", format);
            using ExcelDocument copy = ExcelDocument.LoadEncrypted(output, "rotated-pass");
            Check(copy);
            Assert.Equal("Preserving edit", copy.GetSheet("Data").CellAt(10, 4).GetValue<string>());

            void Check(ExcelDocument workbook) {
                ExcelSheet sheet = workbook.GetSheet("Formulas");
                foreach (JsonElement item in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
                    double expected = item.GetProperty("expected").GetDouble();
                    double actual = sheet.CellAt(item.GetProperty("row").GetInt32(), item.GetProperty("column").GetInt32()).GetValue<double>();
                    Assert.InRange(Math.Abs(actual - expected), 0, 1e-7);
                }
                for (int row = 1; row <= 4; row++) Assert.Equal(row * 10D, sheet.CellAt(row, 6).GetValue<double>());
                string extension = format == ExcelFileFormat.Xls ? "xls" : "xlsb";
                foreach (JsonElement item in manifest.RootElement.GetProperty("additionalArrayCases").EnumerateArray()) {
                    if (!item.GetProperty("formats").EnumerateArray().Any(value => value.GetString() == extension)) continue;
                    Assert.Equal(item.GetProperty("expected").GetDouble(), sheet.CellAt(item.GetProperty("row").GetInt32(), item.GetProperty("column").GetInt32()).GetValue<double>());
                }
            }
        }

        [Theory]
        [InlineData("SUM(ABS(Data!A1:A4))", true)]
        [InlineData("SUM(Data!A1:A4*Data!B1:B4)", true)]
        [InlineData("ROWS(ABS(Data!A1:A4))", false)]
        [InlineData("COLUMNS(ABS(Data!A1:D1))", false)]
        [InlineData("MODE(ABS(Data!B1:B4-3))", false)]
        public void Computed_arrays_retain_array_evaluation_across_function_arguments(string expression, bool arrayFormula) {
            foreach (ExcelFileFormat format in new[] { ExcelFileFormat.Xls, ExcelFileFormat.Xlsb }) {
                using ExcelDocument seed = ExcelDocument.Create(); seed.AddWorksheet("Data"); seed.AddWorksheet("Formulas");
                using ExcelDocument document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(format)));
                if (arrayFormula) document.Sheets[1].SetArrayFormula("D1:D1", expression);
                else document.Sheets[1].CellFormula(1, 4, expression);
                byte[] output = document.ToBytes(format);
                List<byte> operands;
                if (format == ExcelFileFormat.Xls) {
                    operands = LegacyXlsWorkbook.Load(output).FormulaTokenRecords
                        .Where(token => token.Token >= 0x20 && (token.Token & 0x1f) == 0x1b).Select(token => token.Token).ToList();
                } else {
                    var sheet = XlsbWorkbookReader.Load(output).Sheets[1];
                    byte[] parsed = arrayFormula ? Assert.Single(sheet.FormulaGroups).Materialize(1, 4) : Assert.Single(sheet.Cells).FormulaBytes!;
                    operands = new List<byte>();
                    int count = arrayFormula ? checked((int)(uint)(parsed[2] | parsed[3] << 8 | parsed[4] << 16 | parsed[5] << 24)) : parsed.Length;
                    XlsbFormulaTokenScanner.Visit(parsed, arrayFormula ? 6 : 0, count, (_, token, _) => {
                        if (token >= 0x20 && (token & 0x1f) == 0x1b) operands.Add(token);
                    });
                }
                Assert.NotEmpty(operands);
                Assert.All(operands, token => Assert.Equal((byte)0x7b, token));
            }
        }

        [Fact]
        public void Legacy_CSE_conditionals_keep_array_operands_under_an_aggregate() {
            using ExcelDocument document = ExcelDocument.Create(); document.AddWorksheet("Data");
            document.AddWorksheet("Formulas").SetArrayFormula("D1:D1", "SUM(IF(Data!A1:A4>0,Data!A1:A4,0))");
            var native = LegacyXlsWorkbook.Load(document.ToBytes(ExcelFileFormat.Xls));
            var areas = native.FormulaTokenRecords.Where(token => token.TokenName == "PtgArea3d").ToArray();
            Assert.Equal(2, areas.Length);
            Assert.All(areas, token => Assert.Equal((byte)0x7b, token.Token));
        }

        [Theory]
        [InlineData("ABSREF(1,A1)")]
        [InlineData("RELREF(A1,B1)")]
        [InlineData("DEREF(A1)")]
        [InlineData("REFTEXT(A1)")]
        [InlineData("ARGUMENT(1,2,A1)")]
        [InlineData("GETPIVOTDATA(\"Total\",A1)")]
        [InlineData("PHONETIC(A1)")]
        public void Legacy_reference_arguments_are_encoded_as_references(string expression) {
            Assert.True(LegacyXlsFormulaEncoder.TryEncode(expression, out byte[] tokens, out string? reason), reason);
            var scanned = new List<LegacyXlsFormulaTokenRecord>();
            BiffFormulaTokenScanner.ScanTokens(tokens, "CellFormula", null, null, 0, 6, scanned);
            var references = scanned.Where(token => token.Token >= 0x20 && (token.Token & 0x1f) == 4).ToArray();
            Assert.NotEmpty(references);
            Assert.All(references, token => Assert.Equal((byte)0x24, token.Token));
        }

        [Theory]
        [InlineData("CHOOSE(A1,B1,C1)>0")]
        [InlineData("CHOOSE(A1,IF(B1>0,C1,D1),E1)>0")]
        public void Choose_formula_references_rebase_for_validation_and_conditional_formatting(string expression) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Rules");
            sheet.ValidationCustomFormula("D4:D5", expression);
            sheet.AddConditionalFormulaRule("D4:D5", expression);
            var native = LegacyXlsWorkbook.Load(document.ToBytes(ExcelFileFormat.Xls));
            Assert.Equal(expression, Assert.Single(native.Worksheets[0].DataValidations).Formula1);
            Assert.Equal(expression, Assert.Single(native.Worksheets[0].ConditionalFormattings).Formula1);
        }

        [Theory]
        [InlineData("IF(1,10,20)")]
        [InlineData("IF(0,99)")]
        [InlineData("IF(1,CHOOSE(2,10,20,30),0)")]
        [InlineData("SUM(IF(1,{1,2},0))")]
        public void Native_control_flow_offsets_land_on_token_boundaries(string expression) {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Formulas").CellFormula(1, 1, expression);
            var native = LegacyXlsWorkbook.Load(document.ToBytes(ExcelFileFormat.Xls));
            var formula = BiffRecordReader.ReadRecords(native.SourceCompoundFile!.Streams["Workbook"], new()).Single(record => record.Type == 6);
            byte[] bytes = formula.Payload.Skip(22).Take(BiffRecordReader.ReadUInt16(formula.Payload, 20)).ToArray();
            var scanned = new List<LegacyXlsFormulaTokenRecord>();
            BiffFormulaTokenScanner.ScanTokens(bytes, "CellFormula", null, null, 0, 6, scanned);
            var boundaries = new HashSet<int>(scanned.Select(token => token.TokenOffset)) { bytes.Length };
            foreach (var token in scanned.Where(token => token.Token == 0x19)) {
                int start = token.TokenOffset, operand = start + 1;
                if (bytes[operand] == 0x02) Assert.Contains(start + 4 + BiffRecordReader.ReadUInt16(bytes, operand + 1), boundaries);
                else if (bytes[operand] == 0x08) Assert.Contains(start + 5 + BiffRecordReader.ReadUInt16(bytes, operand + 1), boundaries);
                else if (bytes[operand] == 0x04) {
                    int count = BiffRecordReader.ReadUInt16(bytes, operand + 1) + 1;
                    for (int index = 0; index < count; index++)
                        Assert.Contains(start + 4 + BiffRecordReader.ReadUInt16(bytes, start + 4 + index * 2), boundaries);
                }
            }
        }

        [Theory]
        [InlineData("SUM(Data!A1:A3)", 0x3b)]
        [InlineData("SUM((Data!A1:A3))", 0x3b)]
        [InlineData("ABS(Data!A1)", 0x5a)]
        [InlineData("SUMPRODUCT(ABS(Data!A1:A3))", 0x7b)]
        [InlineData("VLOOKUP(A1,Data!A1:B3,2,FALSE)", 0x7b)]
        [InlineData("IRR(Data!A1:A3)", 0x7b)]
        [InlineData("IF(A1>0,SUM(Data!A1:A3),0)", 0x3b)]
        public void Native_reference_classes_preserve_function_argument_semantics(string expression, byte expectedReference) {
            using ExcelDocument seed = ExcelDocument.Create();
            ExcelSheet data = seed.AddWorksheet("Data");
            data.CellValue(1, 1, 10D); data.CellValue(2, 1, 20D); data.CellValue(3, 1, 30D);
            seed.AddWorksheet("Formulas").CellValue(1, 1, 20D);
            foreach (ExcelFileFormat format in new[] { ExcelFileFormat.Xls, ExcelFileFormat.Xlsb }) {
                using ExcelDocument document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(format)));
                document.Sheets[1].CellFormula(1, 2, expression);
                byte[] bytes = document.ToBytes(format);
                byte[] tokens;
                if (format == ExcelFileFormat.Xls) {
                    LegacyXlsWorkbook native = LegacyXlsWorkbook.Load(bytes);
                    var formula = BiffRecordReader.ReadRecords(native.SourceCompoundFile!.Streams["Workbook"], new()).Single(record => record.Type == 6);
                    tokens = formula.Payload.Skip(22).Take(BiffRecordReader.ReadUInt16(formula.Payload, 20)).ToArray();
                    var scanned = new List<LegacyXlsFormulaTokenRecord>();
                    BiffFormulaTokenScanner.ScanTokens(tokens, "CellFormula", null, null, 0, 6, scanned);
                    Assert.Contains(scanned, token => token.Token == expectedReference);
                } else {
                    tokens = XlsbWorkbookReader.Load(bytes).Sheets[1].Cells.Single(cell => cell.Column == 2).FormulaBytes!;
                    var types = new List<byte>();
                    XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (_, token, _) => types.Add(token));
                    Assert.Contains(expectedReference, types);
                }
            }
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xls)]
        [InlineData(ExcelFileFormat.Xlsb)]
        public void A_range_returning_function_remains_a_reference_inside_SUM(ExcelFileFormat format) {
            using ExcelDocument seed = ExcelDocument.Create(); seed.AddWorksheet("Data").CellValue(1, 1, 10D);
            seed.AddWorksheet("Formulas");
            using ExcelDocument document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(format)));
            document.Sheets[1].CellFormula(1, 1, "SUM(OFFSET(Data!A1,0,0,3,1))");
            byte[] output = document.ToBytes(format);
            using ExcelDocument copy = ExcelDocument.Load(new MemoryStream(output));
            Assert.Equal("SUM(OFFSET('Data'!A1,0,0,3,1))", copy.Sheets[1].GetFormulaText(1, 1));
            byte[] tokens;
            if (format == ExcelFileFormat.Xls) {
                var native = LegacyXlsWorkbook.Load(output);
                var formula = BiffRecordReader.ReadRecords(native.SourceCompoundFile!.Streams["Workbook"], new()).Single(record => record.Type == 6);
                tokens = formula.Payload.Skip(22).Take(BiffRecordReader.ReadUInt16(formula.Payload, 20)).ToArray();
                var scanned = new List<LegacyXlsFormulaTokenRecord>();
                BiffFormulaTokenScanner.ScanTokens(tokens, "CellFormula", null, null, 0, 6, scanned);
                Assert.Contains(scanned, token => token.FunctionId == 0x4e && token.Token == 0x22);
            }
            else {
                tokens = XlsbWorkbookReader.Load(output).Sheets[1].Cells.Single().FormulaBytes!;
                var functions = new List<byte>();
                XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (_, token, operand) => {
                    if (token >= 0x20 && (token & 0x1f) == 2 && tokens[operand + 1] == 0x4e) functions.Add(token);
                });
                Assert.Contains((byte)0x22, functions);
            }
        }
    }
}
