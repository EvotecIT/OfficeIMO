using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Shared;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        private static async Task<byte[]> ReadAllBytesCompatAsync(string path, CancellationToken ct) {
#if NETSTANDARD2_0 || NET472 || NET48
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 8192, FileOptions.Asynchronous))
            {
                var mem = new MemoryStream((int)Math.Max(0, fs.Length) + 8192);
                await fs.CopyToAsync(mem, 81920, ct).ConfigureAwait(false);
                return mem.ToArray();
            }
#else
            return await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
#endif
        }

        private static OpenSettings CreateOpenSettings(OpenSettings? openSettings) {
            if (openSettings is null) {
                return new OpenSettings { AutoSave = false };
            }

            return new OpenSettings {
                AutoSave = false,
                CompatibilityLevel = openSettings.CompatibilityLevel,
                MarkupCompatibilityProcessSettings = openSettings.MarkupCompatibilityProcessSettings,
                MaxCharactersInPart = openSettings.MaxCharactersInPart,
            };
        }
    }
}
