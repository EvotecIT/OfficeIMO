using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public class PdfInvoiceDocumentTests {
    [Theory]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.XRechnung)]
    public void RecapturingEditedModelCanPreserveProfile(InvoiceProfile profile) {
        PdfInvoiceDocument original = PdfInvoiceDocument.Create(InvoiceFixture.Create(), profile);
        Invoice edited = original.ToInvoice();
        edited.Number = "EDITED-INVOICE";
        PdfInvoiceDocument updated = PdfInvoiceDocument.Create(edited, original.Profile);
        Assert.Equal(profile, original.Profile);
        Assert.Equal(profile, updated.Profile);
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(updated.ToXmlBytes()).Profile);
        Assert.Equal("EDITED-INVOICE", updated.ToInvoice().Number);
        Assert.NotEqual("EDITED-INVOICE", original.ToInvoice().Number);
        byte[] pdf = updated.ToPdfBytes(Options());
        Assert.Equal(updated.ToXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        var report = PdfDocument.Load(pdf).AssessCompliance(PdfComplianceProfile.FacturX);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied,
            report.Requirements.Single(r => r.Id == "readback-einvoice-profile-consistency").Status);
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.Extended)]
    [InlineData(InvoiceProfile.ExtendedCtcFr)]
    [InlineData(InvoiceProfile.PeppolBis)]
    [InlineData((InvoiceProfile)999)]
    public void UnsupportedAuthoringProfilesAreRejected(InvoiceProfile profile) =>
        Assert.Throws<NotSupportedException>(() => PdfInvoiceDocument.Create(InvoiceFixture.Create(), profile));

    [Theory]
    [InlineData("Delivery")]
    [InlineData("Payee")]
    [InlineData("TaxRepresentative")]
    public void LongDetailGroupsFlowAcrossPages(string group) {
        Invoice invoice = InvoiceFixture.Rich();
        string name = string.Join(" ", Enumerable.Repeat("Long business name for pagination", 250)) + " FINAL-DETAIL-MARKER";
        if (group == "Delivery") invoice.Delivery!.Name = name;
        else if (group == "Payee") invoice.Payee!.Name = name;
        else invoice.TaxRepresentative!.Name = name;
        byte[] pdf = PdfInvoiceDocument.Create(invoice).ToPdfBytes(Options());
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        Assert.True(document.Pages.Count > 1);
        Assert.Contains("FINAL-DETAIL-MARKER", document.ExtractText(), StringComparison.Ordinal);
        WriteEvidence("long-" + group, pdf);
    }

    [Theory]
    [InlineData(true, false, "From 2026-09-01")]
    [InlineData(false, true, "Until 2026-09-10")]
    [InlineData(true, true, "2026-09-01 to 2026-09-10")]
    public void PeriodsShowAvailableBoundaries(bool start, bool end, string expected) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Period = new InvoicePeriod { Start = start ? new DateTime(2026, 9, 1) : null, End = end ? new DateTime(2026, 9, 10) : null };
        invoice.Lines[0].Period = invoice.Period;
        byte[] pdf = PdfInvoiceDocument.Create(invoice).ToPdfBytes(Options());
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains(expected, text, StringComparison.Ordinal);
        Assert.Contains("Period: " + expected, text, StringComparison.Ordinal);
        WriteEvidence(start ? end ? "period-range" : "period-start" : "period-end", pdf);
    }

    [Fact]
    public void BusinessIdentifiersAndPriceDiscountRemainVisible() {
        Invoice invoice = InvoiceFixture.Rich();
        invoice.Seller.Identifiers[0].SchemeId = "0088";
        invoice.Buyer.Identifiers.Add(new InvoiceIdentifier("buyer-id", "0088"));
        invoice.Seller.LegalRegistration!.SchemeId = "0002";
        invoice.Payee!.Identifiers[0].SchemeId = "0088";
        invoice.Payee.LegalRegistration!.SchemeId = "0002";
        invoice.Delivery!.LocationIdentifier!.SchemeId = "0088";
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice);
        byte[] xml = snapshot.ToXmlBytes();
        byte[] pdf = snapshot.ToPdfBytes(Options());
        string text = PdfReadDocument.Open(pdf).ExtractText();
        foreach (string expected in new[] { "Identifier (0088): seller-1", "Identifier (0088): buyer-id",
            "Identifier (0088): payee-1", "Legal registration (0002): HRB 12345", "Legal registration (0002): payee-register",
            "Electronic address (EM):", "seller@example.test", "buyer@example.test", "Location (0088): location-1",
            "Gross price: 110 EUR / 1 C62", "Price discount: 10 EUR / 1 C62" })
            Assert.Contains(expected, text, StringComparison.Ordinal);
        Assert.Equal(xml, Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        WriteEvidence("identifiers-and-prices", pdf);
    }

    private static void WriteEvidence(string name, byte[] pdf) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output!);
        File.WriteAllBytes(Path.Combine(output!, name + ".pdf"), pdf);
    }

    [Fact]
    public void LineReferencesClassificationsAndAttributesRemainVisible() {
        Invoice invoice = InvoiceFixture.Rich();
        invoice.Lines[0].ObjectIdentifier!.SchemeId = "ABZ";
        invoice.ObjectIdentifier!.SchemeId = "ABZ";
        invoice.Lines[0].Classifications[0].ListVersion = "2026";
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice);
        byte[] pdf = snapshot.ToPdfBytes(Options());
        string text = string.Join(" ", PdfReadDocument.Open(pdf).ExtractText().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        foreach (string expected in new[] { "Order line: 10", "Accounting reference: project-cost", "Object (ABZ): line-object",
            "Standard item (0160): 1234567890128", "Origin: DE", "Classification (IB, version 2026): 0721-880X",
            "Service tier: Standard", "Identifier (ABZ): object-1", "Business process", invoice.BusinessProcessId!, "text/csv" })
            Assert.Contains(expected, text, StringComparison.Ordinal);
        Assert.Equal(snapshot.ToXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        WriteEvidence("line-metadata", pdf);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AdjustmentBasesPercentagesAndBothReasonsRemainVisible(bool documentLevel, bool charge) {
        Invoice invoice = InvoiceFixture.Create();
        var adjustment = new InvoiceAllowanceCharge { IsCharge = charge, Amount = 1m, BaseAmount = 100m,
            Percentage = 1.00001m, Reason = "Adjusted service", ReasonCode = charge ? "FC" : "95",
            Tax = documentLevel ? invoice.Lines[0].Tax : null };
        if (documentLevel) invoice.AllowancesAndCharges.Add(adjustment);
        else invoice.Lines[0].AllowancesAndCharges.Add(adjustment);
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice);
        byte[] pdf = snapshot.ToPdfBytes(Options());
        string text = string.Join(" ", PdfReadDocument.Open(pdf).ExtractText().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        foreach (string expected in new[] { "Base: 100.00 EUR", "Percentage: 1.00001%", "Adjusted service", "Reason code: " + adjustment.ReasonCode })
            Assert.Contains(expected, text, StringComparison.Ordinal);
        Assert.Equal(snapshot.ToXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        WriteEvidence((documentLevel ? "document-" : "line-") + (charge ? "charge" : "allowance"), pdf);
    }

    [Fact]
    public void FractionalUnitPricesRetainTheirPrecision() {
        Invoice invoice = InvoiceFixture.Create();
        InvoiceLine line = invoice.Lines[0];
        line.Quantity = 1000m;
        line.GrossPrice = 0.004m;
        line.PriceDiscount = 0.001m;
        line.UnitPrice = 0.003m;
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice);
        byte[] pdf = snapshot.ToPdfBytes(Options());
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("Gross price: 0.004 EUR / 1 C62", text, StringComparison.Ordinal);
        Assert.Contains("Price discount: 0.001 EUR / 1 C62", text, StringComparison.Ordinal);
        Assert.Contains("0.003 / 1 C62", text, StringComparison.Ordinal);
        Assert.Contains("3.57 EUR", text, StringComparison.Ordinal);
        Assert.Equal(snapshot.ToXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        WriteEvidence("fractional-prices", pdf);
    }
    [Theory]
    [InlineData("396")]
    [InlineData("384")]
    [InlineData("999")]
    public void UnsupportedPresentationTypesAreRejected(string typeCode) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TypeCode = typeCode;
        Assert.Throws<NotSupportedException>(() => PdfInvoiceDocument.Create(invoice));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OutsideScopeAdjustmentsPreserveAbsentTaxRate(bool charge) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller.VatIdentifier = null;
        invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "O", ExemptionReason = "Outside scope" };
        invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge {
            IsCharge = charge, Amount = 2m, Reason = "Outside scope adjustment",
            Tax = new InvoiceTaxCategory { Code = "O", ExemptionReason = "Outside scope" }
        });
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice);
        Assert.Null(snapshot.ToInvoice().AllowancesAndCharges[0].Tax!.Rate);
        byte[] pdf = snapshot.ToPdfBytes(Options());
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("Outside scope adjustment", text, StringComparison.Ordinal);
        Assert.Contains(charge ? "Charge" : "Allowance", text, StringComparison.Ordinal);
        Assert.DoesNotContain("0%", text, StringComparison.Ordinal);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output!);
            File.WriteAllBytes(Path.Combine(output!, charge ? "outside-scope-charge.pdf" : "outside-scope-allowance.pdf"), pdf);
        }
    }

    [Theory]
    [InlineData(false, false, true)]
    [InlineData(false, false, false)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    public void VisiblePdfAndAttachmentUseTheCapturedInvoice(bool credit, bool longInvoice, bool rich) {
        Invoice invoice = rich ? InvoiceFixture.Rich(credit) : InvoiceFixture.Create();
        if (longInvoice) {
            for (int index = 3; index <= 35; index++) invoice.Lines.Add(new InvoiceLine {
                Id = index.ToString(System.Globalization.CultureInfo.InvariantCulture), Name = "Service line " + index,
                Quantity = index, UnitPrice = 5.25m, Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
            });
        }
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice);
        byte[] xml = snapshot.ToXmlBytes();
        decimal due = InvoiceCalculator.Calculate(invoice).PayableAmount;
        invoice.Number = "MUTATED"; invoice.Lines[0].UnitPrice = 999m;
        snapshot.ToInvoice().Number = "SECOND-MUTATION";
        byte[] returned = snapshot.ToXmlBytes(); returned[0] ^= 1;
        byte[] pdf = snapshot.ToPdfBytes(Options());
        Assert.Equal(xml, snapshot.ToXmlBytes());
        var attachment = Assert.Single(PdfDocument.Load(pdf).Attachments.Extract());
        Assert.Equal("factur-x.xml", attachment.FileName);
        Assert.Equal(xml, attachment.Bytes);
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("INV-2026-001", text, StringComparison.Ordinal);
        Assert.Contains((credit ? "Credit note " : "Invoice ") + "INV-2026-001", text, StringComparison.Ordinal);
        Assert.DoesNotContain("MUTATED", text, StringComparison.Ordinal);
        Assert.Contains(due.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) + " EUR", text, StringComparison.Ordinal);
        Assert.Contains("Consulting", text, StringComparison.Ordinal);
        Assert.Contains("DE79000000001234567890", text, StringComparison.Ordinal);
        if (longInvoice) Assert.Contains("Service line 35", text, StringComparison.Ordinal);
        string name = !rich ? "simple-invoice" : longInvoice ? "multipage" : credit ? "credit-note" : "invoice";
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE") ?? Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output!);
            File.WriteAllBytes(Path.Combine(output!, name + ".pdf"), pdf);
            File.WriteAllBytes(Path.Combine(output!, name + ".xml"), xml);
        }
        foreach (PdfExternalValidator validator in new[] { PdfExternalValidator.VeraPdf(), PdfExternalValidator.Mustang() }) {
            if (!validator.IsAvailable) {
                if (!string.IsNullOrWhiteSpace(output)) File.WriteAllText(Path.Combine(output!, validator.Name + "-" + name + ".txt"), validator.Name + " was not configured.");
                PdfExternalValidator.SkipUnlessRequired(validator);
                continue;
            }
            PdfExternalProcessResult result = validator.Run(pdf, name + ".pdf");
            if (!string.IsNullOrWhiteSpace(output)) File.WriteAllText(Path.Combine(output!, validator.Name + "-" + name + ".txt"), result.GetDiagnosticText());
            Assert.True(result.ExitCode == 0, result.GetDiagnosticText());
        }
    }
    private static PdfOptions Options() {
        string fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!;
        Assert.NotNull(fontPath);
        byte[] font = File.ReadAllBytes(fontPath);
        return new PdfOptions { IncludeStandardFontToUnicodeMaps = true }
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
    }
}
