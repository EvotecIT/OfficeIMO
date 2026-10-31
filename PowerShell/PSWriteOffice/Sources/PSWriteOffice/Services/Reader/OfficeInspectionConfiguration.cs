using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Reader;

/// <summary>Routes inspection and Reader policies without changing the caller's configuration.</summary>
internal static class OfficeInspectionConfiguration {
    internal static OfficeInspectOptions Bind(object[]? inputs, OfficeDocumentReader? reader = null) {
        object[] values = inputs?.Select(value => OfficeArgumentValue.Unwrap(value)!).ToArray() ?? Array.Empty<object>();
        var inspection = values.OfType<OfficeInspectOptions>().ToArray();
        if (inspection.Length > 1) throw new PSArgumentException("Supply only one OfficeInspectOptions policy.");
        var options = OfficeReadConfiguration.Copy(inspection.SingleOrDefault());
        object[] readOptions = values.Where(value => value is not (OfficeInspectOptions or PowerPointInspectionOptions or ExcelReaderOptions)).ToArray();
        var configuration = OfficeReadConfiguration.Bind(readOptions, reader?.DefaultOptions ?? options.Reader);
        options.Reader = configuration.Common;
        if (configuration.HasFormatOptions) options.Formats = configuration.Formats;
        if ((reader != null || options.DocumentReader != null) && configuration.HasFormatOptions)
            throw new PSArgumentException("Configure format options when creating a caller-provided immutable Reader.");
        options.DocumentReader = reader ?? options.DocumentReader;
        options.PowerPointInspection = values.OfType<PowerPointInspectionOptions>().LastOrDefault() ?? options.PowerPointInspection;
        options.ExcelInspection = values.OfType<ExcelReaderOptions>().LastOrDefault() ?? options.ExcelInspection;
        return options;
    }
}
