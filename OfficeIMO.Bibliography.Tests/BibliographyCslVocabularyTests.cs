using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyCslVocabularyTests {
    // CSL input vocabulary, independent of the codec's private binding table.
    private const string NameProperties = "author chair collection-editor compiler composer container-author contributor curator director editor editorial-director executive-producer guest host interviewer illustrator narrator organizer original-author performer producer recipient reviewed-author script-writer series-creator translator";
    private const string ItemTypes = "article article-journal article-magazine article-newspaper bill book broadcast chapter classic collection dataset document entry entry-dictionary entry-encyclopedia event figure graphic hearing interview legal_case legislation manuscript map motion_picture musical_score pamphlet paper-conference patent performance periodical personal_communication post post-weblog regulation report review review-book software song speech standard thesis treaty webpage";

    [Fact]
    public void EveryStandardCslContributorRoleHasAnEditableTypedRoundTrip() {
        BibliographyDocument document = BibliographyDocument.Parse("[]", BibliographyFormat.CslJson).Document;
        var item = new BibliographyItem { Key = "work", Type = BibliographyItemType.MotionPicture };
        foreach (BibliographyContributorRole role in Enum.GetValues(typeof(BibliographyContributorRole))) {
            if (role != BibliographyContributorRole.Other) item.Contributors.Add(new BibliographyContributor(role, new BibliographyName { Family = role.ToString() }));
        }
        document.Items.Add(item);
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument json = JsonDocument.Parse(written.Content);
        string[] names = json.RootElement[0].EnumerateObject().Select(property => property.Name).Where(property => property != "id" && property != "type").OrderBy(name => name, StringComparer.Ordinal).ToArray();
        Assert.Equal(NameProperties.Split(' ').OrderBy(name => name, StringComparer.Ordinal), names);
        BibliographyItem reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items.Single();
        Assert.Equal(item.Contributors.Select(value => value.Role), reopened.Contributors.Select(value => value.Role));
        Assert.Empty(reopened.NativeFields);
        reopened.Contributors.Single(value => value.Role == BibliographyContributorRole.Director).Name.Given = "Edited";
        Assert.Equal("Edited", reopened.Contributors.Single(value => value.Role == BibliographyContributorRole.Director).Name.Given);
    }

    [Fact]
    public void StandardCslTypesRemainExactAfterNativeTypeEvidenceIsRemoved() {
        string[] types = ItemTypes.Split(' ');
        string source = "[" + string.Join(",", types.Select((type, index) => "{\"id\":\"" + index + "\",\"type\":\"" + type + "\"}")) + "]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        Assert.All(document.Items, item => { Assert.NotEqual(BibliographyItemType.Unknown, item.Type); item.NativeType = null; });
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument output = JsonDocument.Parse(written.Content);
        Assert.Equal(types, output.RootElement.EnumerateArray().Select(item => item.GetProperty("type").GetString()));
        Assert.Equal(document.Items.Select(item => item.Type), BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items.Select(item => item.Type));
    }

    [Fact]
    public void AvailabilityDatesRetainRangesAndNativeMetadataWhileOtherFormatsReportOmission() {
        const string source = "[{\"id\":\"work\",\"type\":\"book\",\"available-date\":{\"date-parts\":[[2020,3,1],[2020,4,2]],\"circa\":true}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        BibliographyDate date = document.Items.Single().GetDate(BibliographyDateRole.Available)!;
        Assert.Equal(3, date.Month); Assert.Equal(4, date.EndMonth);
        Assert.Equal("true", date.NativeFields.Single().Value);
        Assert.Equal(source, document.Write().Content);
        date.EndDay = 3;
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items.Single().GetDate(BibliographyDateRole.Available)!;
        Assert.Equal(3, reopened.EndDay); Assert.Equal("true", reopened.NativeFields.Single().Value);
        foreach (BibliographyFormat format in new[] { BibliographyFormat.BibTex, BibliographyFormat.BibLatex, BibliographyFormat.Ris, BibliographyFormat.Nbib, BibliographyFormat.EndNoteXml }) {
            BibliographyWriteResult converted = document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical });
            Assert.Contains(converted.Report.Diagnostics, issue => issue.Code == "BIBCONV202" && issue.Field == "dates.Available");
            Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = format, RequireNoLoss = true }));
        }
    }

    [Fact]
    public void NewlyTypedRolesPreserveMalformedNativeShapesAndRejectNativeOwnershipChanges() {
        const string source = "[{\"id\":\"work\",\"type\":\"book\",\"director\":17,\"available-date\":\"unknown\",\"chair\":[{\"family\":\"Doe\",\"custom\":true}]}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        BibliographyItem item = document.Items.Single();
        Assert.Equal(BibliographyContributorRole.Chair, item.Contributors.Single().Role);
        Assert.Equal("true", item.Contributors.Single().Name.NativeFields.Single().Value);
        Assert.Contains(item.NativeFields, field => field.Name == "director" && field.Value == "17");
        Assert.Contains(item.NativeFields, field => field.Name == "available-date" && field.Value == "unknown");
        BibliographyWriteResult exact = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Equal(2, BibliographyDocument.Parse(exact.Content, BibliographyFormat.CslJson).Document.Items.Single().NativeFields.Count);
        item.NativeFields.Single(field => field.Name == "director").Value = "[{\"family\":\"Promoted\"}]";
        BibliographyWriteResult changed = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        Assert.True(changed.Report.HasLoss);
        Assert.DoesNotContain(BibliographyDocument.Parse(changed.Content, BibliographyFormat.CslJson).Document.Items.Single().Contributors, value => value.Role == BibliographyContributorRole.Director);
    }

    [Fact]
    public void NewContributorRolesAreNotSilentlyConvertedToAuthorsInOtherFormats() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"work\",\"type\":\"motion_picture\",\"director\":[{\"family\":\"Doe\"}]}]", BibliographyFormat.CslJson).Document;
        foreach (BibliographyFormat format in new[] { BibliographyFormat.BibTex, BibliographyFormat.BibLatex, BibliographyFormat.Ris, BibliographyFormat.Nbib, BibliographyFormat.EndNoteXml }) {
            BibliographyWriteResult converted = document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical });
            Assert.Contains(converted.Report.Diagnostics, issue => issue.Code == "BIBCONV201" && issue.Field == "contributors.Director");
            Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = format, RequireNoLoss = true }));
        }
    }

    [Fact]
    public void InterleavedNewRolesReportRegroupingAndKeepEachRoleInItsOriginalRelativeOrder() {
        const string source = "[{\"id\":\"work\",\"type\":\"motion_picture\"}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        BibliographyItem item = document.Items.Single();
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Director, new BibliographyName { Family = "First" }));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Author" }));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Director, new BibliographyName { Family = "Second" }));
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        Assert.Contains(written.Report.Diagnostics, issue => issue.Code == "BIBCONV230");
        BibliographyItem reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items.Single();
        Assert.Equal(new[] { "Author", "First", "Second" }, reopened.Contributors.Select(value => value.Name.Family));
        Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
    }
}
