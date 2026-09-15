using System.Reflection;
using Xunit;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceSpecificationContractTests {
    [Theory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.Extended)]
    public void FacturXReleaseAcceptsItsFiveCiiProfiles(InvoiceProfile profile) {
        var options = new InvoiceXmlOptions(
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii,
            profile);
        Assert.Equal(profile, options.Profile);
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum, 0)]
    [InlineData(InvoiceProfile.BasicWithoutLines, 0)]
    [InlineData(InvoiceProfile.Basic, 1)]
    public void ReducedFacturXOutputCanBeReadAsItsRetainedModel(InvoiceProfile profile, int expectedLines) {
        var options = new InvoiceXmlOptions(
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii,
            profile,
            InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);

        InvoiceReadResult result = InvoiceParser.Read(InvoiceSerializer.Write(InvoiceFixture.Create(), options));

        Assert.True(result.HasCompleteMapping);
        Assert.Equal(profile, result.Declaration.Profile);
        Assert.Equal(expectedLines, result.Invoice.Lines.Count);
        Assert.Equal(119m, result.Invoice.DeclaredTotals!.PayableAmount);
        byte[] rewritten = result.Write(options);
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(rewritten).Profile);
        Assert.Equal(expectedLines, InvoiceParser.Read(rewritten).Invoice.Lines.Count);
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.Extended)]
    public void FacturXProfilesDoNotInventABusinessProcessIdentifier(InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.BusinessProcessId = null;
        var options = new InvoiceXmlOptions(
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii,
            profile,
            profile is InvoiceProfile.Minimum or InvoiceProfile.BasicWithoutLines or InvoiceProfile.Basic
                ? InvoiceProjectionPolicy.AllowProfileDefinedDataLoss
                : InvoiceProjectionPolicy.RejectDataLoss);

        byte[] xml = InvoiceSerializer.Write(invoice, options);

        Assert.DoesNotContain("BusinessProcessSpecifiedDocumentContextParameter", System.Text.Encoding.UTF8.GetString(xml));
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis)]
    public void CiusAndNetworkContractsStillRequireABusinessProcessIdentifier(InvoiceSyntax syntax, InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.BusinessProcessId = null;
        InvoiceXmlOptions options = InvoiceTestContracts.For(syntax, profile);

        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Location == "BusinessProcessId" && item.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Theory]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.Extended)]
    public void AggregateOnlyModelCannotBePromotedToALineBearingFacturXProfile(InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines.Clear();
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(profile);

        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, options), item => item.Code == "INV-LINES");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Fact]
    public void LowerProfileLineProjectionStatesThatDescriptionsAndTaxClassificationsAreOmitted() {
        Invoice invoice = InvoiceFixture.Create();
        IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(invoice,
            InvoiceTestContracts.FacturX(InvoiceProfile.Minimum, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss));

        InvoiceDiagnostic line = Assert.Single(diagnostics, item => item.Location == "Lines");
        Assert.Contains("descriptions and tax classifications", line.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("visible invoice", line.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AggregateOnlyRewriteRejectsInconsistentDeclaredTotals() {
        var options = InvoiceTestContracts.FacturX(
            InvoiceProfile.BasicWithoutLines,
            InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        InvoiceReadResult result = InvoiceParser.Read(InvoiceSerializer.Write(InvoiceFixture.Create(), options));
        result.Invoice.DeclaredTotals!.PayableAmount += 1m;

        IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(result.Invoice, options);

        Assert.Contains(diagnostics, item => item.Code == "INV-DECLARED-AMOUNT" && item.Location == "DeclaredTotals.PayableAmount");
        Assert.Throws<InvalidDataException>(() => result.Write(options));
    }

    [Theory]
    [InlineData(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Ubl, InvoiceProfile.En16931)]
    [InlineData(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31, InvoiceSyntax.Cii, InvoiceProfile.ExtendedCtcFr)]
    [InlineData(InvoiceSpecificationRelease.PeppolBis_3_0_21, InvoiceSyntax.Cii, InvoiceProfile.PeppolBis)]
    [InlineData(InvoiceSpecificationRelease.PeppolBis_3_0_21, InvoiceSyntax.Ubl, InvoiceProfile.En16931)]
    public void ReleaseSyntaxAndProfileCannotBeMixed(
        InvoiceSpecificationRelease release,
        InvoiceSyntax syntax,
        InvoiceProfile profile) =>
        Assert.Throws<NotSupportedException>(() => new InvoiceXmlOptions(release, syntax, profile));

    [Fact]
    public void XRechnungIsTheOnlyNationalCiusWithAnAuthoringContract() {
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            var contract = new InvoiceXmlOptions(
                InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31,
                syntax,
                InvoiceProfile.XRechnung);
            Assert.Equal(syntax, contract.Syntax);
        }
        foreach (InvoiceSpecificationRelease release in Enum.GetValues(typeof(InvoiceSpecificationRelease)))
            Assert.Throws<NotSupportedException>(() => new InvoiceXmlOptions(release, InvoiceSyntax.Cii, InvoiceProfile.ExtendedCtcFr));
    }

    [Fact]
    public void XmlAuthoringHasNoImplicitReleaseOverload() {
        MethodInfo[] writes = typeof(InvoiceSerializer).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.Name == nameof(InvoiceSerializer.Write)).ToArray();
        MethodInfo write = Assert.Single(writes);
        Assert.Equal(new[] { typeof(Invoice), typeof(InvoiceXmlOptions) }, write.GetParameters().Select(parameter => parameter.ParameterType));
    }
}
