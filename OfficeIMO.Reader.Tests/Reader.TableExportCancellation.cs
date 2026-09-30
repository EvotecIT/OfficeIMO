using System.Collections;
using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderTableExportCancellationTests {
    [Fact]
    public void Markdown_export_observes_cancellation_during_row_scan() {
        using var cancellation = new CancellationTokenSource();
        var rows = new CancelOnFirstRowRead(cancellation);
        var table = new ReaderTable {
            Columns = new[] { "Value" },
            Rows = rows
        };

        Assert.Throws<OperationCanceledException>(() =>
            table.ToMarkdownTable(cancellation.Token));
    }

    private sealed class CancelOnFirstRowRead : IReadOnlyList<IReadOnlyList<string>> {
        private readonly CancellationTokenSource _cancellation;

        internal CancelOnFirstRowRead(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        public int Count => 2;

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
}
