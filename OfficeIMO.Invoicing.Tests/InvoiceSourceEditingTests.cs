using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public sealed class InvoiceSourceEditingTests {
    private static readonly XNamespace Extension = "urn:example:preserved-business-data";

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void MultiFieldEditsRetainUnknownContentAndOriginalBytes(InvoiceSyntax syntax) {
        byte[] bytes = Source(syntax);
        InvoiceSourceDocument source = InvoiceSourceDocument.Load(bytes);
        Assert.Equal(bytes, source.ToBytes());
        Assert.False(InvoiceParser.Read(bytes).HasCompleteMapping);
        var result = InvoiceSourceEditor.Apply(source, new(number: "EDITED<&NUMBER", issueDate: new DateTime(2028, 2, 29),
            dueDate: new DateTime(2028, 3, 30), buyerReference: "BUYER-EDIT", paymentReference: "PAYMENT-EDIT"));
        Assert.True(result.Succeeded, string.Join("; ", result.Diagnostics.Select(d => d.Message)));
        InvoiceSourceDocument edited = result.Document!;
        Assert.Equal("EDITED<&NUMBER", edited.Number);
        Assert.Equal(new DateTime(2028, 2, 29), edited.IssueDate);
        Assert.Equal(new DateTime(2028, 3, 30), edited.DueDate);
        Assert.Equal("BUYER-EDIT", edited.BuyerReference); Assert.Equal("PAYMENT-EDIT", edited.PaymentReference);
        Assert.Equal("INV-2026-001", source.Number); Assert.Equal(bytes, source.ToBytes());
        using var beforeStream = new MemoryStream(bytes); using var afterStream = new MemoryStream(edited.ToBytes());
        XDocument before = XDocument.Load(beforeStream, LoadOptions.PreserveWhitespace), after = XDocument.Load(afterStream, LoadOptions.PreserveWhitespace);
        Assert.True(XNode.DeepEquals(before.Root!.Element(Extension + "Payload"), after.Root!.Element(Extension + "Payload")));
        Assert.Equal("retain", (string?)after.Root.Attribute(Extension + "flag"));
        Assert.Contains("<!-- retained -->", Encoding.UTF8.GetString(edited.ToBytes()));
        Assert.Equal(source.TypeCode, edited.TypeCode); Assert.Equal(source.Currency, edited.Currency); Assert.Equal(source.GuidelineId, edited.GuidelineId);
        Assert.False(InvoiceParser.Read(edited.ToBytes()).HasCompleteMapping);
        Assert.Equal(119m, InvoiceModelValidator.ValidateSource(InvoiceParser.Read(edited.ToBytes())).Calculation!.PayableAmount);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "INV-SOURCE-EDIT-VALIDATION-REQUIRED");
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void NumberOnlyEditRetainsPaymentReference(InvoiceSyntax syntax) {
        var source = InvoiceSourceDocument.Load(Source(syntax));
        var result = InvoiceSourceEditor.Apply(source, new(number: "NEW-NUMBER"));
        Assert.True(result.Succeeded);
        Assert.Equal(source.PaymentReference, result.Document!.PaymentReference);
    }

    [Fact]
    public void AmbiguousLaterFieldBlocksTheEntireEditAndReportsItsLocation() {
        XDocument xml = Parse(Source(InvoiceSyntax.Ubl));
        XNamespace cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
        XElement payment = xml.Root!.Element(cac + "PaymentMeans")!;
        payment.AddAfterSelf(new XElement(payment));
        byte[] bytes = Encode(xml); var source = InvoiceSourceDocument.Load(bytes);
        var result = InvoiceSourceEditor.Apply(source, new(number: "NEW-NUMBER", paymentReference: "NEW-PAYMENT"));
        Assert.False(result.Succeeded); Assert.Null(result.Document);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Location == "PaymentReference" && diagnostic.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.Equal("INV-2026-001", source.Number); Assert.Equal(bytes, source.ToBytes());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void MissingNestedOrAnnotatedTargetsAreRetainedWithoutReplacement(int variant) {
        XDocument xml = Parse(Source(InvoiceSyntax.Ubl));
        XElement number = xml.Root!.Element(XName.Get("ID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"))!;
        if (variant == 0) number.Remove();
        else if (variant == 1) number.Add(new XElement(Extension + "Nested", "retain"));
        else number.Add(new XComment("retain annotation"));
        byte[] bytes = Encode(xml); var source = InvoiceSourceDocument.Load(bytes);
        var result = InvoiceSourceEditor.Apply(source, new(number: "NEW-NUMBER"));
        Assert.False(result.Succeeded); Assert.Equal(bytes, source.ToBytes());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Location == "Number");
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void SignaturePassThroughIsAllowedButMutationIsBlocked(InvoiceSyntax syntax) {
        XDocument xml = Parse(Source(syntax)); xml.Root!.Add(new XElement(XName.Get("Signature", "http://www.w3.org/2000/09/xmldsig#")));
        byte[] bytes = Encode(xml); var source = InvoiceSourceDocument.Load(bytes);
        Assert.True(source.HasXmlSignature);
        Assert.True(InvoiceSourceEditor.Apply(source, new()).Succeeded);
        var result = InvoiceSourceEditor.Apply(source, new(number: "NEW-NUMBER"));
        Assert.False(result.Succeeded); Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Location == "Source.Signature");
        Assert.Equal(bytes, source.ToBytes());
    }

    [Fact]
    public void TimezoneDateAndCreditNoteDueDateCannotBeReinterpreted() {
        XDocument xml = Parse(Source(InvoiceSyntax.Ubl));
        xml.Root!.Element(XName.Get("IssueDate", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"))!.Value = "2026-09-10+02:00";
        var source = InvoiceSourceDocument.Load(Encode(xml)); Assert.Null(source.IssueDate);
        var result = InvoiceSourceEditor.Apply(source, new(issueDate: new DateTime(2028, 2, 29)));
        Assert.False(result.Succeeded); Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Location == "IssueDate");
        Invoice invoice = InvoiceFixture.Create(); invoice.TypeCode = "381"; invoice.DueDate = null;
        var credit = InvoiceSourceDocument.Load(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)));
        Assert.True(credit.IsUblCreditNote); Assert.Null(credit.DueDate);
        Assert.False(InvoiceSourceEditor.Apply(credit, new(dueDate: DateTime.Today)).Succeeded);
        Assert.True(InvoiceSourceEditor.Apply(credit, new(number: "CREDIT-EDIT", issueDate: new DateTime(2028, 2, 29))).Succeeded);
    }

    [Fact]
    public void StreamCaptureIsBoundedAndPreservesCallerOwnership() {
        byte[] bytes = Source(InvoiceSyntax.Ubl);
        using var stream = new MemoryStream(new byte[] { 1, 2 }.Concat(bytes).ToArray()); stream.Position = 2;
        Assert.Equal(bytes, InvoiceSourceDocument.Load(stream).ToBytes()); Assert.True(stream.CanRead);
        using var oversized = new MemoryStream(new byte[InvoiceSourceDocument.MaximumXmlBytes + 50]);
        Assert.Throws<InvalidDataException>(() => InvoiceSourceDocument.Load(oversized));
        Assert.Equal(InvoiceSourceDocument.MaximumXmlBytes + 1L, oversized.Position); Assert.True(oversized.CanRead);
        Assert.Throws<ArgumentException>(() => new InvoiceSourceEdits(number: new string('x', InvoiceSourceEdits.MaximumTextCharacters + 1)));
        Assert.Throws<XmlException>(() => new InvoiceSourceEdits(number: "invalid\0number"));
    }

    [Fact]
    public void ShallowFloodIsRejectedByThePreservationOwnerBeforeTreeMaterialization() {
        byte[] bytes = Encoding.UTF8.GetBytes("<Invoice xmlns='urn:oasis:names:specification:ubl:schema:xsd:Invoice-2'>" +
            string.Concat(Enumerable.Repeat("<a/>", InvoiceSourceDocument.MaximumXmlNodes)) + "</Invoice>");
        Assert.Throws<InvalidDataException>(() => InvoiceSourceDocument.Load(bytes));
    }

    private static byte[] Source(InvoiceSyntax syntax) {
        XDocument xml = Parse(InvoiceSerializer.Write(InvoiceFixture.Create(), InvoiceTestContracts.En16931(syntax)));
        xml.Root!.SetAttributeValue(Extension + "flag", "retain");
        xml.Root.Add(new XElement(Extension + "Payload", new XAttribute(Extension + "code", "unknown"), "A\rB", new XComment(" retained "), new XElement(Extension + "Inner", "unknown business data")));
        return Encode(xml);
    }
    private static XDocument Parse(byte[] bytes) { using var stream = new MemoryStream(bytes); return XDocument.Load(stream, LoadOptions.PreserveWhitespace); }
    private static byte[] Encode(XDocument document) {
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = Encoding.Unicode, NewLineHandling = NewLineHandling.Entitize })) document.Save(writer);
        return stream.ToArray();
    }
}
