// The current comparison package targets .NET 10; .NET 8 uses its final compatible release.
// Version 4 reorganized namespaces without changing these benchmark operations.
#if NET10_0_OR_GREATER
global using ExcelReader.Core.Reader.Schema;
global using ExcelReader.Core.Reader.Xls;
global using ExcelReader.Core.Reader.Xlsb;
global using ExcelReader.Core.Reader.Xlsx;
global using ExcelReader.Core.Writer.Xls;
global using ExcelReader.Core.Writer.Xlsb;
global using ExcelReader.Core.Writer.Xlsx;
#else
global using ExcelColumnType = ExcelReader.Core.Enums.ExcelColumnType;
#endif
