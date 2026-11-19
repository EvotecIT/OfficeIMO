using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Runs a 65K typed-read lane once before timing and records the result for the benchmark evidence. Each lane
/// already compares its row count, cell count and the checksum over every typed value against the hash-pinned
/// fixture; the recorder repeats that comparison so a passed entry always means the full observation matched.
/// </summary>
internal static class MarkPflug65KLaneValidation {
    internal static void Record(string format, string method, string library, string caseName, Func<ExcelReadObservation> lane) {
        ExcelReadObservation expected = MarkPflug65KXlsxBenchmarks.ExpectedObservation();
        ExcelReadObservation actual;
        try {
            actual = lane();
        } catch (Exception exception) {
            BenchmarkOutputValidationLog.Failed(library, $"{format} {method}: {exception.Message}", caseName);
            throw;
        }

        string detail = $"{format} {method}: rows={actual.Rows}; cells={actual.Cells}; checksum={actual.Checksum:x16}";
        if (actual != expected) {
            BenchmarkOutputValidationLog.Failed(library, detail + $"; expected {expected}", caseName);
            throw new InvalidDataException($"{method} did not produce the expected {format} observation {expected}; it produced {actual}.");
        }

        BenchmarkOutputValidationLog.Passed(library, detail, caseName);
    }
}

public partial class MarkPflug65KXlsxBenchmarks {
    private void ValidateLane(string method, string library, string caseName, Func<ExcelReadObservation> lane) {
        Setup();
        MarkPflug65KLaneValidation.Record("XLSX", method, library, caseName, lane);
    }

    [GlobalSetup(Target = nameof(OfficeIMO))]
    public void SetupOfficeIMO() => ValidateLane(nameof(OfficeIMO), "OfficeIMO", "Default", OfficeIMO);

    [GlobalSetup(Target = nameof(OfficeIMO_Prefetch))]
    public void SetupOfficeIMOPrefetch() => ValidateLane(nameof(OfficeIMO_Prefetch), "OfficeIMO", "Prefetch", OfficeIMO_Prefetch);

    [GlobalSetup(Target = nameof(ExcelReaderNet))]
    public void SetupExcelReaderNet() => ValidateLane(nameof(ExcelReaderNet), "ExcelReader.NET", "Default", ExcelReaderNet);

    [GlobalSetup(Target = nameof(ExcelReaderNet_Prefetch))]
    public void SetupExcelReaderNetPrefetch() => ValidateLane(nameof(ExcelReaderNet_Prefetch), "ExcelReader.NET", "Prefetch", ExcelReaderNet_Prefetch);

    [GlobalSetup(Target = nameof(Sylvan))]
    public void SetupSylvan() => ValidateLane(nameof(Sylvan), "Sylvan.Data.Excel", "Default", Sylvan);

    [GlobalSetup(Target = nameof(ExcelDataReader))]
    public void SetupExcelDataReader() => ValidateLane(nameof(ExcelDataReader), "ExcelDataReader", "Default", ExcelDataReader);

    [GlobalSetup(Target = nameof(ClosedXML))]
    public void SetupClosedXML() => ValidateLane(nameof(ClosedXML), "ClosedXML", "Default", ClosedXML);

    // The restricted-library filter runs this lane only when the caller enabled EPPlus.
    [GlobalSetup(Target = nameof(EPPlus))]
    public void SetupEPPlus() => ValidateLane(nameof(EPPlus), "EPPlus", "Default", EPPlus);

    [GlobalSetup(Target = nameof(MiniExcel))]
    public void SetupMiniExcel() => ValidateLane(nameof(MiniExcel), "MiniExcel", "Default", MiniExcel);

    [GlobalSetup(Target = nameof(OfficeIMO_DefaultNumbers))]
    public void SetupOfficeIMODefaultNumbers() => ValidateLane(nameof(OfficeIMO_DefaultNumbers), "OfficeIMO", "DefaultNumbers", OfficeIMO_DefaultNumbers);

    [GlobalSetup(Target = nameof(OfficeIMO_DefaultNumbers_Prefetch))]
    public void SetupOfficeIMODefaultNumbersPrefetch() =>
        ValidateLane(nameof(OfficeIMO_DefaultNumbers_Prefetch), "OfficeIMO", "DefaultNumbers_Prefetch", OfficeIMO_DefaultNumbers_Prefetch);

    [GlobalSetup(Target = nameof(OfficeIMO_Loaded))]
    public void SetupOfficeIMOLoaded() => ValidateLane(nameof(OfficeIMO_Loaded), "OfficeIMO", "Loaded", OfficeIMO_Loaded);

    [GlobalSetup(Target = nameof(OfficeIMO_SinglePass))]
    public void SetupOfficeIMOSinglePass() => ValidateLane(nameof(OfficeIMO_SinglePass), "OfficeIMO", "SinglePass", OfficeIMO_SinglePass);

    [GlobalSetup(Target = nameof(OfficeIMO_Memory))]
    public void SetupOfficeIMOMemory() => ValidateLane(nameof(OfficeIMO_Memory), "OfficeIMO", "Memory", OfficeIMO_Memory);

#if NET10_0_OR_GREATER
    [GlobalSetup(Target = nameof(ExcelReaderNet_Memory))]
    public void SetupExcelReaderNetMemory() => ValidateLane(nameof(ExcelReaderNet_Memory), "ExcelReader.NET", "Memory", ExcelReaderNet_Memory);

    [GlobalSetup(Target = nameof(ExcelReaderNet_InternStrings))]
    public void SetupExcelReaderNetInternStrings() =>
        ValidateLane(nameof(ExcelReaderNet_InternStrings), "ExcelReader.NET", "InternStrings", ExcelReaderNet_InternStrings);
#endif

    [GlobalSetup(Target = nameof(ExcelPrime))]
    public void SetupExcelPrime() => ValidateLane(nameof(ExcelPrime), "Excel-PRIME", "Default", ExcelPrime);

    [GlobalSetup(Target = nameof(ExcelPrime_TypedStyles))]
    public void SetupExcelPrimeTypedStyles() => ValidateLane(nameof(ExcelPrime_TypedStyles), "Excel-PRIME", "TypedStyles", ExcelPrime_TypedStyles);

    [GlobalSetup(Target = nameof(HypeLab))]
    public void SetupHypeLab() => ValidateLane(nameof(HypeLab), "HypeLab.IO.Excel", "Default", HypeLab);

    [GlobalSetup(Target = nameof(XLibur))]
    public void SetupXLibur() => ValidateLane(nameof(XLibur), "XLibur", "Default", XLibur);
}

public partial class MarkPflug65KXlsbBenchmarks {
    private void ValidateLane(string method, string library, string caseName, Func<ExcelReadObservation> lane) {
        Setup();
        MarkPflug65KLaneValidation.Record("XLSB", method, library, caseName, lane);
    }

    [GlobalSetup(Target = nameof(OfficeIMO))]
    public void SetupOfficeIMO() => ValidateLane(nameof(OfficeIMO), "OfficeIMO", "Default", OfficeIMO);

    [GlobalSetup(Target = nameof(ExcelReaderNet))]
    public void SetupExcelReaderNet() => ValidateLane(nameof(ExcelReaderNet), "ExcelReader.NET", "Default", ExcelReaderNet);

    [GlobalSetup(Target = nameof(Sylvan))]
    public void SetupSylvan() => ValidateLane(nameof(Sylvan), "Sylvan.Data.Excel", "Default", Sylvan);

    [GlobalSetup(Target = nameof(ExcelDataReader))]
    public void SetupExcelDataReader() => ValidateLane(nameof(ExcelDataReader), "ExcelDataReader", "Default", ExcelDataReader);

    [GlobalSetup(Target = nameof(OfficeIMO_SinglePass))]
    public void SetupOfficeIMOSinglePass() => ValidateLane(nameof(OfficeIMO_SinglePass), "OfficeIMO", "SinglePass", OfficeIMO_SinglePass);

    [GlobalSetup(Target = nameof(ExcelPrime))]
    public void SetupExcelPrime() => ValidateLane(nameof(ExcelPrime), "Excel-PRIME", "Default", ExcelPrime);

    [GlobalSetup(Target = nameof(ExcelPrime_TypedStyles))]
    public void SetupExcelPrimeTypedStyles() => ValidateLane(nameof(ExcelPrime_TypedStyles), "Excel-PRIME", "TypedStyles", ExcelPrime_TypedStyles);

#if NET10_0_OR_GREATER
    [GlobalSetup(Target = nameof(ExcelReaderNet_Prefetch))]
    public void SetupExcelReaderNetPrefetch() => ValidateLane(nameof(ExcelReaderNet_Prefetch), "ExcelReader.NET", "Prefetch", ExcelReaderNet_Prefetch);
#endif
}

public partial class MarkPflug65KXlsBenchmarks {
    private void ValidateLane(string method, string library, string caseName, Func<ExcelReadObservation> lane) {
        Setup();
        MarkPflug65KLaneValidation.Record("XLS", method, library, caseName, lane);
    }

    [GlobalSetup(Target = nameof(OfficeIMO))]
    public void SetupOfficeIMO() => ValidateLane(nameof(OfficeIMO), "OfficeIMO", "Default", OfficeIMO);

    [GlobalSetup(Target = nameof(ExcelReaderNet))]
    public void SetupExcelReaderNet() => ValidateLane(nameof(ExcelReaderNet), "ExcelReader.NET", "Default", ExcelReaderNet);

    [GlobalSetup(Target = nameof(Sylvan))]
    public void SetupSylvan() => ValidateLane(nameof(Sylvan), "Sylvan.Data.Excel", "Default", Sylvan);

    [GlobalSetup(Target = nameof(ExcelDataReader))]
    public void SetupExcelDataReader() => ValidateLane(nameof(ExcelDataReader), "ExcelDataReader", "Default", ExcelDataReader);
}
