using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookDistributionListTests {
    [Fact]
    public void OneOffEntryIdCodecRoundTripsUnicodeIdentityAndClassifiesIt() {
        var source = new EmailAddress("żółć@example.test", "Żaneta Żółć") { AddressType = "SMTP" };

        byte[] encoded = OutlookEntryIdCodec.EncodeOneOff(source);
        bool decoded = OutlookEntryIdCodec.TryDecodeOneOff(encoded,
            out EmailAddress? result, out string? error);

        Assert.True(decoded, error);
        Assert.Equal(OutlookEntryIdKind.OneOff, OutlookEntryIdCodec.Classify(encoded));
        Assert.Equal(source.Address, result!.Address);
        Assert.Equal(source.DisplayName, result.DisplayName);
        Assert.Equal("SMTP", result.AddressType);
        Assert.Equal(new byte[] { 0, 0, 0, 0, 0x81, 0x2B, 0x1F, 0xA4 }, encoded.Take(8));
    }

    [Fact]
    public void DistributionListRoundTripsSynchronizedMembersAndChecksumThroughMsg() {
        var list = new OutlookDistributionList { Name = "Project Team" };
        list.Add(new EmailAddress("alice@example.test", "Alice"));
        list.Add(new EmailAddress("bob@example.test", "Bob"));
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.DistributionList,
            Subject = "Project Team",
            DistributionList = list
        };

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument result = new EmailDocumentReader().Read(bytes).Document;

        Assert.Equal(OutlookItemKind.DistributionList, result.OutlookItemKind);
        Assert.Equal("IPM.DistList", result.MessageClass);
        Assert.NotNull(result.DistributionList);
        Assert.Equal("Project Team", result.DistributionList!.Name);
        Assert.Equal(new[] { "alice@example.test", "bob@example.test" },
            result.DistributionList.Members.Select(member => member.Address!.Address));
        Assert.All(result.DistributionList.Members, member => {
            Assert.Equal(OutlookEntryIdKind.OneOff, member.Kind);
            Assert.NotNull(member.EntryId);
            Assert.NotNull(member.OneOffEntryId);
            Assert.Null(member.DecodeError);
        });
        object[] members = result.Mapi.GetValueOrDefault(
            MapiKnownProperties.PidLid.DistributionListMembers)!;
        object[] oneOff = result.Mapi.GetValueOrDefault(
            MapiKnownProperties.PidLid.DistributionListOneOffMembers)!;
        Assert.Equal(2, members.Length);
        Assert.Equal(2, oneOff.Length);
        Assert.Equal(OutlookDistributionList.CalculateChecksum(members.Cast<byte[]>()),
            result.Mapi.GetNullableValue(MapiKnownProperties.PidLid.DistributionListChecksum));
        Assert.True(result.DistributionList.Validate().IsValid);
    }

    [Fact]
    public void DistributionListChecksumUsesTheSpecifiedSeedZeroIeeePolynomial() {
        int checksum = OutlookDistributionList.CalculateChecksum(new[] {
            Encoding.ASCII.GetBytes("123456789")
        });

        Assert.Equal(unchecked((int)0x2DFD2D88u), checksum);
    }

    [Fact]
    public void DistributionListValidationReportsSizeAndChecksumWithoutDiscardingRawEvidence() {
        var list = new OutlookDistributionList { Checksum = 123 };
        list.Members.Add(new OutlookDistributionListMember {
            EntryId = new byte[OutlookDistributionList.MaximumMemberPropertyBytes],
            OneOffEntryId = OutlookEntryIdCodec.EncodeOneOff(new EmailAddress("member@example.test"))
        });

        OutlookDistributionListValidationReport report = list.Validate();

        Assert.False(report.IsValid);
        Assert.Contains(report.Issues, issue => issue.Code == "OUTLOOK_DISTLIST_MEMBERS_TOO_LARGE" && issue.IsError);
        Assert.Contains(report.Issues, issue => issue.Code == "OUTLOOK_DISTLIST_CHECKSUM_MISMATCH" && !issue.IsError);
        Assert.Equal(OutlookDistributionList.MaximumMemberPropertyBytes, list.Members[0].EntryId!.Length);
    }

    [Fact]
    public void EntryIdCodecRejectsMalformedAndUnsupportedValuesWithoutThrowing() {
        bool decoded = OutlookEntryIdCodec.TryDecodeOneOff(
            new byte[] { 0, 1, 2, 3 }, out EmailAddress? address, out string? error);

        Assert.False(decoded);
        Assert.Null(address);
        Assert.NotNull(error);
        Assert.Equal(OutlookEntryIdKind.Unknown, OutlookEntryIdCodec.Classify(new byte[] { 0, 1, 2, 3 }));
    }
}
