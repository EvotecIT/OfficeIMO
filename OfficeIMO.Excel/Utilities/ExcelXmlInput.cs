using System.Threading;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel.Utilities;

/// <summary>Shared bounded, nonresolving XML admission for the workbook and table adapters.</summary>
internal static class ExcelXmlInput {
    internal static XElement Read(Stream input, int maxBytes, int maxDepth, int maxElements, CancellationToken cancellation,
        XName? guardedElement = null, int maximumGuardedElements = int.MaxValue) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead) throw new ArgumentException("Input must be readable.", nameof(input));
        using var bytes = new MemoryStream(); var buffer = new byte[8192]; int count;
        while ((count = input.Read(buffer, 0, (int)Math.Min(buffer.Length, (long)maxBytes - bytes.Length + 1))) != 0) {
            cancellation.ThrowIfCancellationRequested();
            if (bytes.Length + count > maxBytes) throw new InvalidDataException("XML input exceeds MaxInputBytes.");
            bytes.Write(buffer, 0, count);
        }
        cancellation.ThrowIfCancellationRequested(); bytes.Position = 0;
        using var xml = XmlReader.Create(bytes, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = maxBytes, IgnoreComments = true, IgnoreProcessingInstructions = true });
        return Load(xml, maxDepth, maxElements, cancellation, guardedElement, maximumGuardedElements);
    }

    internal static XElement ReadText(string input, int maxBytes, int maxDepth, int maxElements, CancellationToken cancellation) {
        if (new UTF8Encoding(false, true).GetByteCount(input) > maxBytes) throw new InvalidDataException("XML input exceeds MaxInputBytes.");
        using var text = new StringReader(input);
        using var xml = XmlReader.Create(text, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = maxBytes, IgnoreComments = true, IgnoreProcessingInstructions = true });
        return Load(xml, maxDepth, maxElements, cancellation, null, int.MaxValue);
    }

    private static XElement Load(XmlReader xml, int maxDepth, int maxElements, CancellationToken cancellation, XName? guardedElement, int maximumGuardedElements) {
        using var limited = new OfficeXmlLimitingReader(xml, "Excel XML", maxDepth, maxElements,
            (int)Math.Min(int.MaxValue, 8L * maxElements), cancellation, guardedElement?.LocalName, guardedElement?.NamespaceName,
            maximumGuardedElements, "MaxCells");
        return XDocument.Load(limited, LoadOptions.PreserveWhitespace).Root ?? throw new InvalidDataException("XML input has no root element.");
    }
}
