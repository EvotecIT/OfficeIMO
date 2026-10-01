using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Internal.Invoicing;

internal enum InvoiceSourceField { DocumentId, TypeCode, GuidelineId, CurrencyCode, IssueDate, DueDate, BuyerReference, PaymentReference }

internal readonly struct InvoiceSourceScalarEdit {
    internal InvoiceSourceScalarEdit(InvoiceSourceField field, string value) { Field = field; Value = value; }
    internal InvoiceSourceField Field { get; }
    internal string Value { get; }
}

/// <summary>Immutable XML preservation owner shared by standalone invoices and the PDF CII bridge.</summary>
internal sealed class InvoicePreservedXml {
    private static readonly XNamespace Dsig = "http://www.w3.org/2000/09/xmldsig#";
    private readonly byte[] _bytes;
    private readonly XDocument _document;
    private InvoicePreservedXml(byte[] bytes, XDocument document) { _bytes = bytes; _document = document; }
    internal bool IsCii => _document.Root!.Name == InvoiceXml.Rsm + "CrossIndustryInvoice";
    internal bool IsCreditNote => _document.Root!.Name == InvoiceXml.UblCreditNote + "CreditNote";
    internal bool HasXmlSignature => _document.Descendants(Dsig + "Signature").Any();

    internal static InvoicePreservedXml Load(byte[] xml) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(xml);
#else
        if (xml == null) throw new ArgumentNullException(nameof(xml));
#endif
        if (xml.Length == 0 || xml.Length > InvoiceProfileDeclaration.MaximumXmlBytes)
            throw new InvalidDataException("Invoice XML must contain between 1 byte and 16 MiB.");
        byte[] bytes = (byte[])xml.Clone();
        XDocument document = InvoiceXml.Parse(bytes, InvoiceXml.MaximumNodes, InvoiceXml.MaximumAttributes);
        XName? root = document.Root?.Name;
        if (root != InvoiceXml.Rsm + "CrossIndustryInvoice" && root != InvoiceXml.UblInvoice + "Invoice" && root != InvoiceXml.UblCreditNote + "CreditNote")
            throw new InvalidDataException("Expected a namespace-100 CII invoice or UBL 2 invoice/credit-note root.");
        return new InvoicePreservedXml(bytes, document);
    }

    internal byte[] ToBytes() => (byte[])_bytes.Clone();
    internal static InvoicePreservedXml Load(Stream stream) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(stream);
#else
        if (stream == null) throw new ArgumentNullException(nameof(stream));
#endif
        if (!stream.CanRead) throw new ArgumentException("Invoice XML stream must be readable.", nameof(stream));
        using (var buffer = new MemoryStream()) {
            byte[] chunk = new byte[8192];
            while (true) {
                int count = stream.Read(chunk, 0, (int)Math.Min(chunk.Length, InvoiceProfileDeclaration.MaximumXmlBytes + 1L - buffer.Length));
                if (count == 0) break;
                buffer.Write(chunk, 0, count);
                if (buffer.Length > InvoiceProfileDeclaration.MaximumXmlBytes) throw new InvalidDataException("Invoice XML exceeds the maximum byte length.");
            }
            return Load(buffer.ToArray());
        }
    }
    internal string? Read(InvoiceSourceField field) {
        XElement? element = Find(_document, Path(field));
        if (element == null) return null;
        if (element.HasElements) throw new InvalidDataException("Invoice scalar field contains nested elements.");
        return element.Value;
    }

    internal DateTime? ReadDate(InvoiceSourceField field) {
        XElement? element = Find(_document, Path(field));
        if (element == null || IsCii && (string?)element.Attribute("format") != "102") return null;
        string? text = Read(field);
        return DateTime.TryParseExact(text, IsCii ? "yyyyMMdd" : "yyyy-MM-dd", CultureInfo.InvariantCulture,
            DateTimeStyles.None, out DateTime date) ? date : (DateTime?)null;
    }

    internal InvoicePreservedXml Apply(IEnumerable<InvoiceSourceScalarEdit> edits) {
        if (HasXmlSignature) throw new InvalidOperationException("Editing XML with a Signature element is not supported. Preserve the original bytes or use a signature-aware workflow.");
        var copy = new XDocument(_document);
        var targets = new List<(XElement Element, string Value)>();
        var fields = new HashSet<InvoiceSourceField>();
        foreach (InvoiceSourceScalarEdit edit in edits) {
            try {
                if (targets.Count >= 8 || !fields.Add(edit.Field)) throw new ArgumentException("Choose at most eight distinct supported scalar fields.", nameof(edits));
                if (string.IsNullOrWhiteSpace(edit.Value) || edit.Value.Length > InvoiceProfileDeclaration.MaximumXmlBytes)
                    throw new ArgumentException("An edited invoice value must be nonempty and within the XML size limit.", nameof(edits));
                XmlConvert.VerifyXmlChars(edit.Value);
                if (edit.Field is InvoiceSourceField.TypeCode or InvoiceSourceField.GuidelineId or InvoiceSourceField.CurrencyCode)
                    throw new NotSupportedException("Document type, profile and monetary currencies require a semantic editing workflow.");
                XElement element = Find(copy, Path(edit.Field)) ?? throw new InvalidOperationException("The field must already exist before it can be edited.");
                if (element.Nodes().Any(node => node is not XText)) throw new InvalidOperationException("The field contains structured XML content and cannot be replaced as text.");
                if (edit.Field is InvoiceSourceField.IssueDate or InvoiceSourceField.DueDate) {
                    if (IsCii && (string?)element.Attribute("format") != "102") throw new InvalidOperationException("Only existing format-102 CII dates can be edited.");
                    if (!IsCii && !IsPlainUblDate(element.Value)) throw new InvalidOperationException("Only existing plain yyyy-MM-dd UBL dates can be edited; timezone-bearing representations are retained.");
                }
                targets.Add((element, edit.Value));
            } catch (Exception exception) when (exception is ArgumentException || exception is InvalidDataException || exception is InvalidOperationException || exception is NotSupportedException || exception is XmlException) {
                throw new InvoiceSourceFieldException(edit.Field, exception);
            }
        }
        if (targets.Count == 0) return this;
        foreach (var target in targets) target.Element.Value = target.Value;
        using (var stream = new InvoiceXmlOutputStream()) {
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, NewLineHandling = NewLineHandling.Entitize })) copy.Save(writer);
            return Load(stream.ToArray());
        }
    }

    private static bool IsPlainUblDate(string text) => text.Length == 10 && text[4] == '-' && text[7] == '-' &&
        text.Where((_, index) => index != 4 && index != 7).All(character => character >= '0' && character <= '9');

    private XName[] Path(InvoiceSourceField field) {
        if (IsCii) return field switch {
            InvoiceSourceField.DocumentId => new[] { InvoiceXml.Rsm + "ExchangedDocument", InvoiceXml.Ram + "ID" },
            InvoiceSourceField.TypeCode => new[] { InvoiceXml.Rsm + "ExchangedDocument", InvoiceXml.Ram + "TypeCode" },
            InvoiceSourceField.GuidelineId => new[] { InvoiceXml.Rsm + "ExchangedDocumentContext", InvoiceXml.Ram + "GuidelineSpecifiedDocumentContextParameter", InvoiceXml.Ram + "ID" },
            InvoiceSourceField.CurrencyCode => Settlement("InvoiceCurrencyCode"),
            InvoiceSourceField.PaymentReference => Settlement("PaymentReference"),
            InvoiceSourceField.IssueDate => new[] { InvoiceXml.Rsm + "ExchangedDocument", InvoiceXml.Ram + "IssueDateTime", InvoiceXml.Udt + "DateTimeString" },
            InvoiceSourceField.DueDate => new[] { InvoiceXml.Rsm + "SupplyChainTradeTransaction", InvoiceXml.Ram + "ApplicableHeaderTradeSettlement", InvoiceXml.Ram + "SpecifiedTradePaymentTerms", InvoiceXml.Ram + "DueDateDateTime", InvoiceXml.Udt + "DateTimeString" },
            InvoiceSourceField.BuyerReference => new[] { InvoiceXml.Rsm + "SupplyChainTradeTransaction", InvoiceXml.Ram + "ApplicableHeaderTradeAgreement", InvoiceXml.Ram + "BuyerReference" },
            _ => throw new ArgumentOutOfRangeException(nameof(field))
        };
        return field switch {
            InvoiceSourceField.DocumentId => new[] { InvoiceXml.Cbc + "ID" },
            InvoiceSourceField.TypeCode => new[] { InvoiceXml.Cbc + (IsCreditNote ? "CreditNoteTypeCode" : "InvoiceTypeCode") },
            InvoiceSourceField.GuidelineId => new[] { InvoiceXml.Cbc + "CustomizationID" },
            InvoiceSourceField.CurrencyCode => new[] { InvoiceXml.Cbc + "DocumentCurrencyCode" },
            InvoiceSourceField.IssueDate => new[] { InvoiceXml.Cbc + "IssueDate" },
            InvoiceSourceField.DueDate when !IsCreditNote => new[] { InvoiceXml.Cbc + "DueDate" },
            InvoiceSourceField.DueDate => throw new NotSupportedException("UBL credit-note due-date editing is not supported."),
            InvoiceSourceField.BuyerReference => new[] { InvoiceXml.Cbc + "BuyerReference" },
            InvoiceSourceField.PaymentReference => new[] { InvoiceXml.Cac + "PaymentMeans", InvoiceXml.Cbc + "PaymentID" },
            _ => throw new ArgumentOutOfRangeException(nameof(field))
        };
    }
    private static XName[] Settlement(string field) => new[] { InvoiceXml.Rsm + "SupplyChainTradeTransaction", InvoiceXml.Ram + "ApplicableHeaderTradeSettlement", InvoiceXml.Ram + field };
    private static XElement? Find(XDocument document, XName[] path) {
        XElement? current = document.Root;
        foreach (XName name in path) { if (current == null) return null; current = InvoiceXml.Unique(current, name); }
        return current;
    }
}
