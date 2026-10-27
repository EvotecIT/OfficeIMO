using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Excel;

public sealed partial class ExcelWorkbookDataReader {
    private sealed class NativeOpenXmlLifetime(XlsxTabularWorkbook native, Func<ExcelDocumentReader> openFallback) : IDisposable {
        private ExcelDocumentReader? _fallback;

        internal DbDataReader OpenSheet(string name, ExcelReaderOptions options, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (_fallback != null) return OpenOpenXmlSheet(_fallback, name, options, token);
            try {
                return native.OpenTable(name, options.HasHeaderRow, token);
            } catch (XlsxTabularFastPathNotSupportedException) {
                // Metadata discovery remains lazy; unsupported worksheet shapes still use the existing SDK reader.
                _fallback = openFallback();
                return OpenOpenXmlSheet(_fallback, name, options, token);
            }
        }

        public void Dispose() {
            try { _fallback?.Dispose(); }
            finally { native.Dispose(); }
        }
    }
}
