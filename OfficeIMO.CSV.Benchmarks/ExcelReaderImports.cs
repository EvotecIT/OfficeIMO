// Keep current .NET 10 and final compatible .NET 8 comparisons on the same public operations.
#if NET10_0_OR_GREATER
global using ExcelReaderNetCsvReader = ExcelReader.Core.Reader.Csv.CsvReader;
global using ExcelReaderNetCsvRowWriter = ExcelReader.Core.Writer.Csv.CsvRowWriter;
global using ExcelReaderNetCsvWriter = ExcelReader.Core.Writer.Csv.CsvWorkbookWriter;
#else
global using ExcelReader.Core.ValueObjects;
global using ExcelReaderNetCsvReader = ExcelReader.Core.Reader.CsvReader;
global using ExcelReaderNetCsvRowWriter = ExcelReader.Core.Writer.CsvRowWriter;
global using ExcelReaderNetCsvWriter = ExcelReader.Core.Writer.CsvWriter;
#endif
