using System;
using System.Diagnostics;
using System.IO;
using OfficeIMO.Word;
using Xunit;

#nullable enable

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void MultipleSectionsWarnComponentName() {
            string filePath = Path.Combine(_directoryWithFiles, "WarningTest.docx");
            using WordDocument document = WordDocument.Create(filePath);
            document.AddSection();

            AssertWarning(document, () => _ = document.Header, nameof(WordDocument.Header));
            AssertWarning(document, () => _ = document.Footer, nameof(WordDocument.Footer));
            AssertWarning(document, () => _ = document.DifferentFirstPage, nameof(WordDocument.DifferentFirstPage));
            AssertWarning(document, () => _ = document.DifferentOddAndEvenPages, nameof(WordDocument.DifferentOddAndEvenPages));
        }

        private static void AssertWarning(WordDocument document, Action action, string componentName) {
            using var listener = new TestTraceListener();
            Trace.Listeners.Add(listener);
            action();
            Trace.Listeners.Remove(listener);
            Assert.Contains($"Sections[wantedSection].{componentName}", listener.Messages);
        }

        private sealed class TestTraceListener : TraceListener {
            public string Messages = string.Empty;
            public override void Write(string? message) => Messages += message;
            public override void WriteLine(string? message) => Messages += message;
        }
    }
}

