using System.Collections;
using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderTableExportCancellationTests {
    [Theory]
    [InlineData("csv", false)]
    [InlineData("markdown", false)]
    [InlineData("json", false)]
    [InlineData("bundle", false)]
    [InlineData("csv", true)]
    [InlineData("markdown", true)]
    [InlineData("json", true)]
    [InlineData("bundle", true)]
    public void Table_exports_observe_cancellation_during_scan_or_final_row_write(string format, bool duringWrite) {
        using var cancellation = new CancellationTokenSource();
        var table = new ReaderTable {
            Columns = new[] { "Value" },
            Rows = duringWrite
                ? new[] { new CancelOnCellRead(cancellation) }
                : new CancelOnFirstRowRead(cancellation)
        };

        Assert.Throws<OperationCanceledException>(() => {
            switch (format) {
                case "csv": table.ToCsv(cancellation.Token); break;
                case "markdown": table.ToMarkdownTable(cancellation.Token); break;
                case "json": table.ToJson(indented: false, cancellationToken: cancellation.Token); break;
                case "bundle": new OfficeDocumentReaderBuilder().Build()
                    .ExportTables(new[] { table }, cancellationToken: cancellation.Token); break;
            }
        });
    }

    private sealed class CancelOnFirstRowRead : IReadOnlyList<IReadOnlyList<string>> {
        private readonly CancellationTokenSource _cancellation;

        internal CancelOnFirstRowRead(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        public int Count => 1;

        public IReadOnlyList<string> this[int index] {
            get {
                if (index == 0) _cancellation.Cancel();
                return new[] { index.ToString() };
            }
        }

        public IEnumerator<IReadOnlyList<string>> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class CancelOnCellRead : IReadOnlyList<string> {
        private readonly CancellationTokenSource _cancellation;
        internal CancelOnCellRead(CancellationTokenSource cancellation) { _cancellation = cancellation; }
        public int Count => 1;
        public string this[int index] { get { _cancellation.Cancel(); return "value"; } }
        public IEnumerator<string> GetEnumerator() { yield return this[0]; }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
