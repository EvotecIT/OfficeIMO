using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public class PdfInvoiceDocumentTests {
    [Fact]
    public void PresentationPdfKeepsVisibleInvoiceWithoutElectronicAttachment() {
        Invoice invoice = InvoiceFixture.Create();
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());

        byte[] pdf = snapshot.ToPresentationPdfBytes(Options());
        string text = string.Join(" ", PdfReadDocument.Open(pdf).ExtractText().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        Assert.Empty(PdfDocument.Load(pdf).Attachments.Extract());
        Assert.Contains("Invoice INV-2026-001", text, StringComparison.Ordinal);
        Assert.Contains("119.00 EUR", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PresentationPdfRejectsElectronicInvoiceOptions() {
        Invoice invoice = InvoiceFixture.Create();
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
        PdfOptions options = Options().UseFacturX(snapshot.ToXmlBytes());

        ArgumentException exception = Assert.Throws<ArgumentException>(() => snapshot.ToPresentationPdfBytes(options));
        Assert.Equal("options", exception.ParamName);
        Assert.Contains("ToPdfBytes", exception.Message, StringComparison.Ordinal);

        options.ElectronicInvoiceMetadata = null;
        Assert.Throws<ArgumentException>(() => snapshot.ToPresentationPdfBytes(options));

        PdfOptions genericAttachment = Options().AddEmbeddedFile(
            "FACTUR-X.XML",
            snapshot.ToXmlBytes(),
            "application/xml",
            PdfAssociatedFileRelationship.Alternative);
        Assert.Throws<ArgumentException>(() => snapshot.ToPresentationPdfBytes(genericAttachment));
    }

    [Fact]
    public void ModernPresentationCapturesThemeApprovalsAndLocalizedLabels() {
        Invoice invoice = InvoiceFixture.Create();
        var theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));
        var layout = InvoicePdfLayoutOptions.ForCultures("pl-PL");
        layout.Theme = theme;
        layout.Approvals.Add(new InvoicePdfApproval("Prepared by", "Marta Nowak", "Finance", invoice.IssueDate));
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract(), layout);

        byte[] first = snapshot.ToPresentationPdfBytes(Options());
        theme.Accent = PdfColor.Black;
        theme.CornerRadius = 0D;
        layout.Theme = null;
        layout.Approvals.Clear();
        byte[] second = snapshot.ToPresentationPdfBytes(Options());
        string text = PdfReadDocument.Open(first).ExtractText();

        Assert.Equal(first, second);
        Assert.Contains("Prepared by", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Marta Nowak", text, StringComparison.Ordinal);
        Assert.Contains("Finance", text, StringComparison.Ordinal);
        Assert.Contains("Referencja płatności", text, StringComparison.Ordinal);
        Assert.Contains("Akceptacje", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ModernPresentationUsesFirstAvailablePaymentReference() {
        Invoice invoice = InvoiceFixture.Create();
        InvoicePayment existing = invoice.Payments[0];
        existing.Reference = "SECOND-PAYMENT-REFERENCE";
        invoice.Payments.Insert(0, new InvoicePayment {
            MeansCode = existing.MeansCode,
            MeansText = existing.MeansText,
            Account = existing.Account
        });
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));

        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPresentationPdfBytes(Options());
        string text = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(3, text.Split(new[] { "SECOND-PAYMENT-REFERENCE" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void ModernPresentationPreservesLogoAspectRatioByDefault() {
        Invoice invoice = InvoiceFixture.Create();
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));
        layout.LogoBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        layout.LogoWidth = 150D;
        layout.LogoHeight = 30D;

        Assert.Equal(OfficeImageFit.Contain, layout.LogoFit);
        Assert.Throws<ArgumentOutOfRangeException>(() => layout.LogoFit = (OfficeImageFit)999);

        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPresentationPdfBytes(Options());
        PdfImagePlacement placement = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetImagePlacements());

        Assert.InRange(placement.Width, 29.99D, 30.01D);
        Assert.InRange(placement.Height, 29.99D, 30.01D);
    }

    [Fact]
    public void ModernPresentationKeepsFinancialBreakdownsAndTotalsVisible() {
        Invoice invoice = InvoiceFixture.Rich();
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));

        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPresentationPdfBytes(Options());
        string text = string.Join(" ", PdfReadDocument.Open(pdf).ExtractText().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        WriteEvidence("modern-financial-details", pdf);

        foreach (string expected in new[] {
            "Document adjustments", "Document discount", "Delivery", "VAT category",
            "S 19%", "S 7%", "Allowances", "5.00 EUR", "Charges", "2.00 EUR",
            "Total including", "254.64 EUR", "Rounding", "0.01 EUR"
        }) Assert.Contains(expected, text, StringComparison.Ordinal);
    }

    [Fact]
    public void ModernPresentationFlowsLongCardsAcrossPages() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Number = string.Join(" ", Enumerable.Repeat("LONG-INVOICE", 100)) + " FINAL-INVOICE-MARKER";
        invoice.Seller.LegalInformation = string.Join(" ", Enumerable.Repeat("Long seller disclosure", 160)) + " FINAL-PARTY-MARKER";
        invoice.PaymentTerms = string.Join(" ", Enumerable.Repeat("Long payment term", 180)) + " FINAL-PAYMENT-MARKER";
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));
        layout.Approvals.Add(new InvoicePdfApproval("Approved by", "Marta Nowak",
            string.Join(" ", Enumerable.Repeat("Long approval role", 180)) + " FINAL-APPROVAL-MARKER", invoice.IssueDate));

        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPresentationPdfBytes(Options());
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        string text = document.ExtractText();

        Assert.True(document.Pages.Count > 2);
        foreach (string expected in new[] { "FINAL-INVOICE-MARKER", "FINAL-PARTY-MARKER", "FINAL-PAYMENT-MARKER", "FINAL-APPROVAL-MARKER" })
            Assert.Contains(expected, text, StringComparison.Ordinal);
    }

    [Fact]
    public void ModernPresentationFlowsMultilinePaymentSummaryAcrossPages() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].Reference = string.Join("\n", Enumerable.Repeat("PAYMENT-REFERENCE-LINE", 70)) +
            "\nFINAL-PAYMENT-REFERENCE-MARKER";
        invoice.PaymentTerms = string.Join("\n", Enumerable.Repeat("TERM", 70)) +
            "\nFINAL-PAYMENT-TERMS-MARKER";
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));

        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPresentationPdfBytes(Options());
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        string text = document.ExtractText();

        Assert.True(document.Pages.Count > 2);
        Assert.Contains("FINAL-PAYMENT-REFERENCE-MARKER", text, StringComparison.Ordinal);
        Assert.Contains("FINAL-PAYMENT-TERMS-MARKER", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ClassicPresentationRendersRequestedApprovals() {
        Invoice invoice = InvoiceFixture.Create();
        var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
        layout.Approvals.Add(new InvoicePdfApproval("Approved by", "Marta Nowak", "Finance", invoice.IssueDate));

        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPresentationPdfBytes(Options());
        string text = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Approvals", text, StringComparison.Ordinal);
        Assert.Contains("Marta Nowak", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SnapshotSuppliesAttachmentDateForRequiredFacturXCompliance() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller.ElectronicAddress = new InvoiceIdentifier("1234567890128", "0088");
        invoice.Buyer.ElectronicAddress = new InvoiceIdentifier("1234567890135", "0088");
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
        byte[] pdf = snapshot.ToPdfBytes(Options().RequireCompliance(PdfComplianceProfile.FacturX));
        Assert.Equal(snapshot.ToXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        WriteEvidence("required-factur-x", pdf);
    }

    [Fact]
    public void LongInvoiceHeadingFlowsIntoVisiblePages() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Number = string.Join(" ", Enumerable.Repeat("INVOICE-NUMBER", 700)) + " FINAL-NUMBER-MARKER";
        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract()).ToPdfBytes(Options());
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        Assert.True(document.Pages.Count > 2);
        Assert.Contains("FINAL-NUMBER-MARKER", document.ExtractText(), StringComparison.Ordinal);
        WriteEvidence("long-invoice-number", pdf);
    }

    [Theory]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.XRechnung)]
    public void RecapturingEditedModelCanPreserveProfile(InvoiceProfile profile) {
        PdfInvoiceDocument original = PdfInvoiceDocument.Create(InvoiceFixture.Create(), Contract(profile));
        Invoice edited = original.ToInvoice();
        edited.Number = "EDITED-INVOICE";
        PdfInvoiceDocument updated = PdfInvoiceDocument.Create(edited, Contract(original.Profile));
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
    [InlineData(InvoiceProfile.Minimum, 0)]
    [InlineData(InvoiceProfile.BasicWithoutLines, 0)]
    [InlineData(InvoiceProfile.Basic, 1)]
    public void LowerFacturXPdfUsesTheProfileRetainedSnapshot(InvoiceProfile profile, int expectedLines) {
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(InvoiceFixture.Create(), Contract(profile));
        Invoice retained = snapshot.ToInvoice();

        Assert.Equal(expectedLines, retained.Lines.Count);
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(snapshot.ToXmlBytes()).Profile);
        retained.Number = "RECAPTURED-LOWER-PROFILE";
        PdfInvoiceDocument recaptured = PdfInvoiceDocument.Create(retained, Contract(profile));
        Assert.Equal("RECAPTURED-LOWER-PROFILE", recaptured.Number);
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(recaptured.ToXmlBytes()).Profile);

        byte[] pdf = snapshot.ToPdfBytes(Options());
        Assert.Equal(snapshot.ToXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("Invoice INV-2026-001", text, StringComparison.Ordinal);
        Assert.Equal(expectedLines > 0, text.IndexOf("Consulting", StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData(InvoiceProfile.ExtendedCtcFr)]
    [InlineData(InvoiceProfile.PeppolBis)]
    public void UnsupportedAuthoringProfilesAreRejected(InvoiceProfile profile) =>
        Assert.Throws<NotSupportedException>(() => PdfInvoiceDocument.Create(InvoiceFixture.Create(), Contract(profile)));

    [Fact]
    public void UnknownProfileIsRejectedByTheExplicitContract() =>
        Assert.Throws<ArgumentOutOfRangeException>(() => Contract((InvoiceProfile)999));

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
        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract()).ToPdfBytes(Options());
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
        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract()).ToPdfBytes(Options());
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
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
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
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
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
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
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
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
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
        Assert.Throws<NotSupportedException>(() => PdfInvoiceDocument.Create(invoice, Contract()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OutsideScopeAdjustmentsPreserveAbsentTaxRate(bool charge) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller.TaxRegistrations.Clear();
        invoice.Seller.LegalRegistration = new InvoiceIdentifier("HRB 12345");
        invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "O", ExemptionReason = "Outside scope" };
        invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge {
            IsCharge = charge, Amount = 2m, Reason = "Outside scope adjustment",
            Tax = new InvoiceTaxCategory { Code = "O", ExemptionReason = "Outside scope" }
        });
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
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
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract());
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

    private static InvoiceXmlOptions Contract(InvoiceProfile profile = InvoiceProfile.En16931) => profile switch {
        InvoiceProfile.XRechnung => new InvoiceXmlOptions(InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31, InvoiceSyntax.Cii, profile),
        InvoiceProfile.PeppolBis => new InvoiceXmlOptions(InvoiceSpecificationRelease.PeppolBis_3_0_21, InvoiceSyntax.Ubl, profile),
        _ => new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, profile,
            profile == InvoiceProfile.En16931 || profile == InvoiceProfile.Extended ? InvoiceProjectionPolicy.RejectDataLoss : InvoiceProjectionPolicy.AllowProfileDefinedDataLoss)
    };
}
