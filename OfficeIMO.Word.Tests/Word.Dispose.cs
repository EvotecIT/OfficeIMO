using System;
using System.IO;
using OfficeIMO.Word;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains tests for disposing Word documents.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_DisposeMultipleTimes() {
            var filePath = Path.Combine(_directoryWithFiles, "DisposeTestingMultipleTimes.docx");
            File.Delete(filePath);

            var document = WordDocument.Create(filePath);
            document.AddParagraph("This is my test");
            document.Save();
            document.Dispose();
            document.Dispose();

            Assert.False(filePath.IsFileLocked());
        }

        [Fact]
        public void Test_DisposeDoesNotCaptureSynchronizationContext() {
            var filePath = Path.Combine(_directoryWithFiles, "DisposeTestingSynchronizationContext.docx");
            File.Delete(filePath);

            var document = WordDocument.Create(filePath, autoSave: true);
            document.AddParagraph("This is my test");

            var originalContext = SynchronizationContext.Current;
            var context = new ThrowingSynchronizationContext();

            try {
                SynchronizationContext.SetSynchronizationContext(context);
                document.Dispose();
            } finally {
                SynchronizationContext.SetSynchronizationContext(originalContext);
            }

            Assert.Equal(0, context.PostCount);
            Assert.False(filePath.IsFileLocked());
        }

        private sealed class ThrowingSynchronizationContext : SynchronizationContext {
            public int PostCount { get; private set; }

            public override void Post(SendOrPostCallback d, object? state) {
                PostCount++;
                throw new InvalidOperationException("Asynchronous continuations are not expected during synchronous disposal.");
            }
        }
    }
}
