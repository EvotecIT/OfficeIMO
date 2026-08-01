using System.Collections;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void DigitalSignatureContentTypeOverridesResolveAcrossLargePackages() {
            const int partCount = 2_000;
            using var packageBytes = new MemoryStream();
            using (var zip = new ZipArchive(packageBytes, ZipArchiveMode.Create, leaveOpen: true)) {
                var contentTypes = new StringBuilder(
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                for (int index = 0; index < partCount; index++) {
                    string partName = "/parts/part" + index + ".xml";
                    contentTypes.Append("<Override PartName=\"").Append(partName)
                        .Append("\" ContentType=\"application/x-officeimo-").Append(index).Append("\" />");
                    using Stream part = zip.CreateEntry(partName.TrimStart('/')).Open();
                    part.WriteByte((byte)(index & 0xff));
                }
                contentTypes.Append("</Types>");
                using Stream contentTypePart = zip.CreateEntry("[Content_Types].xml").Open();
                byte[] xml = Encoding.UTF8.GetBytes(contentTypes.ToString());
                contentTypePart.Write(xml, 0, xml.Length);
            }

            using var archive = new OfficePackageSignatureArchive(packageBytes.ToArray(), partCount + 1);

            Assert.True(archive.TryGetContentType("/parts/part0.xml", out string first));
            Assert.Equal("application/x-officeimo-0", first);
            Assert.True(archive.TryGetContentType("/parts/part1999.xml", out string last));
            Assert.Equal("application/x-officeimo-1999", last);
        }

        [Fact]
        public void ComparisonDisclosureResourceLimitDoesNotClaimUndetectedShapes() {
            var result = (WordComparisonResult)(Activator.CreateInstance(
                typeof(WordComparisonResult),
                BindingFlags.Instance | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { "source", "target" },
                culture: null) ?? throw new InvalidOperationException("Comparison result could not be created."));
            Type scanResultType = typeof(WordDocumentComparer).GetNestedType(
                "ShapeScanResult", BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("Shape scan result type was not found.");
            MethodInfo addShapeLimitation = typeof(WordDocumentComparer).GetMethod(
                "AddShapeLimitation", BindingFlags.NonPublic | BindingFlags.Static)
                ?? throw new InvalidOperationException("Shape limitation helper was not found.");

            addShapeLimitation.Invoke(null, new[] {
                result,
                "EffectiveFormatting.ThemeResolution",
                "Theme limitation",
                Enum.Parse(scanResultType, "ResourceLimitExceeded"),
                Enum.Parse(scanResultType, "Absent")
            });

            WordComparisonLimitation limitation = Assert.Single(result.Limitations);
            Assert.Equal("EffectiveFormatting.ThemeResolution.ResourceLimit", limitation.Code);
            Assert.False(limitation.SourceContainsShape);
            Assert.False(limitation.TargetContainsShape);
        }

        [Fact]
        public void MailMergeOccurrenceDiscoveryDoesNotIndexEveryIrrelevantElement() {
            var body = new Body();
            for (int index = 0; index < 20_000; index++) {
                body.Append(new Paragraph(new Run(new Text("noise"))));
            }
            MethodInfo discover = typeof(WordMailMerge).GetMethod(
                "DiscoverMergeFieldOccurrences", BindingFlags.NonPublic | BindingFlags.Static)
                ?? throw new InvalidOperationException("Merge-field discovery helper was not found.");

#if !NET472
            long before = GC.GetAllocatedBytesForCurrentThread();
#endif
            var occurrences = (IEnumerable)(discover.Invoke(null, new object[] { body })
                ?? throw new InvalidOperationException("Merge-field discovery did not return a result."));
            int count = occurrences.Cast<object>().Count();
#if !NET472
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
#endif

            Assert.Equal(0, count);
#if !NET472
            Assert.True(allocated < 6L * 1024 * 1024,
                "Merge-field discovery allocated " + allocated + " bytes for irrelevant elements.");
#endif
        }

        [Fact]
        public void MacroValidationReturnsStructuredPackageByteLimitFailure() {
            string filePath = CreateMacroEnabledTestDocument("MacroValidationPackageByteLimit.docm");
            long packageLength = new FileInfo(filePath).Length;
            var options = new WordMacroProjectSignatureValidationOptions();
            options.Inspection.PackageSecurity.MaxPackageBytes = packageLength - 1;

            WordMacroProjectSignatureValidationResult result =
                WordDocument.ValidateMacroProjectSignature(filePath, options);

            Assert.False(result.IsValidUnderPolicy);
            Assert.Contains(result.Findings, finding => finding.Code == "PackageByteLimitExceeded");
        }
    }
}
