using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Bounds every unloaded XML part before mutation planning can materialize it.
        /// The impact scan separately accounts for the elements it semantically inspects.
        /// </summary>
        private void EnsureMutationPlanPartsFitWithinBudget(
            ExcelMutationPlanOptions options) {
            var budget = new MutationPlanScanBudget(
                options.MaximumScannedElements,
                options.MaximumScannedCharacters);
            var visited = new HashSet<OpenXmlPart>();
            var pending = new Stack<OpenXmlPart>();
            foreach (IdPartPair relationship in _spreadSheetDocument.Parts) {
                budget.Consume();
                pending.Push(relationship.OpenXmlPart);
            }

            while (pending.Count > 0) {
                OpenXmlPart part = pending.Pop();
                if (!visited.Add(part)) {
                    continue;
                }

                EnsureMutationPlanPartXmlFitsWithinBudget(part, budget);
                foreach (IdPartPair relationship in part.Parts) {
                    budget.Consume();
                    pending.Push(relationship.OpenXmlPart);
                }
            }
        }

        private static void EnsureMutationPlanPartXmlFitsWithinBudget(
            OpenXmlPart part,
            MutationPlanScanBudget budget) {
            if (!part.IsRootElementLoaded && IsXmlContentType(part.ContentType)) {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                if (!stream.CanSeek || stream.Length > 0L) {
                    using var boundedStream = new MutationPlanBudgetedReadStream(
                        stream,
                        budget);
                    using XmlReader reader = XmlReader.Create(
                        boundedStream,
                        new XmlReaderSettings {
                            DtdProcessing = DtdProcessing.Prohibit,
                            XmlResolver = null,
                            IgnoreComments = true,
                            IgnoreProcessingInstructions = true,
                            MaxCharactersInDocument = budget.MaximumCharacters
                        });
                    while (reader.Read()) {
                        if (reader.NodeType == XmlNodeType.Element) {
                            budget.Consume();
                        }
                    }
                }
            }
        }

        private sealed class MutationPlanBudgetedReadStream : Stream {
            private readonly Stream _inner;
            private readonly MutationPlanScanBudget _budget;

            internal MutationPlanBudgetedReadStream(
                Stream inner,
                MutationPlanScanBudget budget) {
                _inner = inner;
                _budget = budget;
            }

            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => _inner.CanSeek;
            public override bool CanWrite => false;
            public override long Length => _inner.Length;
            public override long Position {
                get => _inner.Position;
                set => _inner.Position = value;
            }

            public override void Flush() => _inner.Flush();

            public override int Read(byte[] buffer, int offset, int count) {
                int read = _inner.Read(
                    buffer,
                    offset,
                    _budget.GetBudgetedReadSize(count));
                _budget.ConsumeXmlBytes(read);
                return read;
            }

            public override int Read(Span<byte> buffer) {
                int read = _inner.Read(
                    buffer[.._budget.GetBudgetedReadSize(buffer.Length)]);
                _budget.ConsumeXmlBytes(read);
                return read;
            }

            public override int ReadByte() {
                int value = _inner.ReadByte();
                if (value >= 0) {
                    _budget.ConsumeXmlBytes(1);
                }
                return value;
            }

            public override long Seek(long offset, SeekOrigin origin) =>
                _inner.Seek(offset, origin);

            public override void SetLength(long value) =>
                throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();
        }

        private static bool IsXmlContentType(string contentType) =>
            contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)
            || contentType.Equals("application/xml", StringComparison.OrdinalIgnoreCase)
            || contentType.Equals("text/xml", StringComparison.OrdinalIgnoreCase);
    }
}
