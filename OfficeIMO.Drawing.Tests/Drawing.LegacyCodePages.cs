using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Tests;

// .NET 8 and later have no in-box provider for the Windows ANSI code pages, and this test process registers none,
// so these decode the same text as .NET Framework only through OfficeIMO's own tables.
public sealed class DrawingLegacyCodePageTests {
    public static TheoryData<int, byte[], string> AnsiText => new() {
        { 1251, new byte[] { 0xCF, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2 }, "Привет" },
        { 874, new byte[] { 0xCA, 0xC7, 0xD1, 0xCA, 0xB4, 0xD5 }, "สวัสดี" }
    };

    [Theory]
    [MemberData(nameof(AnsiText))]
    public void PropertySetAnsiStringsUseTheSectionCodePage(int codePage, byte[] text, string expected) {
        byte[] propertySet = OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.SummaryInformationFormatId,
            OfficeOlePropertySetWriter.CreateSection(new[] {
                OfficeOleProperty.Int16(1, unchecked((short)codePage)),
                OfficeOleProperty.Blob(2, text)
            })));
        // Retype the blob as VT_LPSTR; both carry a 32-bit length followed by the bytes.
        int title = IndexOf(propertySet, text) - 8;
        propertySet[title] = 0x1E;

        OfficeOlePropertySection section = Assert.Single(OfficeOlePropertySetReader.ReadSections(propertySet));

        Assert.Equal(expected, section.Properties[2].Value);
    }

    [Theory]
    [MemberData(nameof(AnsiText))]
    public void VbaSourceUsesTheProjectCodePage(int codePage, byte[] text, string expected) {
        Assert.Equal(expected, OfficeVbaText.Decode(text, codePage));
        Assert.Equal(text, OfficeVbaText.Encode(expected, codePage));
    }

    private static int IndexOf(byte[] source, byte[] value) {
        for (int index = 0; index <= source.Length - value.Length; index++) {
            if (source.AsSpan(index, value.Length).SequenceEqual(value)) return index;
        }
        throw new InvalidOperationException("The value is not in the property set.");
    }
}
