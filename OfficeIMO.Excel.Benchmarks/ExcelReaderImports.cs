// The current comparison package targets .NET 10; .NET 8 uses its final compatible release.
// .NET 10 enumerates sheets from workbooks; .NET 8 retains the compatible reader API.
#if NET10_0_OR_GREATER
global using ExcelReader.Core.Reader.Schema;
global using XlsxReader = ExcelReader.Core.Reader.Xlsx.XlsxWorkbook;
global using XlsbReader = ExcelReader.Core.Reader.Xlsb.XlsbWorkbook;
global using XlsReader = ExcelReader.Core.Reader.Xls.XlsWorkbook;
global using ExcelReader.Core.Reader.Xls;
global using ExcelReader.Core.Reader.Xlsb;
global using ExcelReader.Core.Reader.Xlsx;
global using ExcelReader.Core.Writer.Xls;
global using ExcelReader.Core.Writer.Xlsb;
global using ExcelReader.Core.Writer.Xlsx;
#else
global using ExcelColumnType = ExcelReader.Core.Enums.ExcelColumnType;
#endif
