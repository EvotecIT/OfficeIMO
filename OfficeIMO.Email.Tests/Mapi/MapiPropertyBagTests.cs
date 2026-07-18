using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MapiPropertyBagTests {
    private static readonly MapiPropertyKey<string> Subject = new MapiPropertyKey<string>(
        "PidTagSubject", 0x0037, MapiPropertyType.Unicode, MapiPropertyType.String8);

    [Fact]
    public void DocumentBagUsesExactRawCollectionAndLastCanonicalValueWins() {
        var document = new EmailDocument();
        document.MapiProperties.Add(new MapiProperty(0x0037, MapiPropertyType.String8, "first"));
        document.Mapi.Properties.Add(new MapiProperty(0x0037, MapiPropertyType.Unicode, "last"));

        Assert.Same(document.MapiProperties, document.Mapi.Properties);
        Assert.Equal("last", document.Mapi.GetValueOrDefault(Subject));
        Assert.Equal(2, document.Mapi.FindAll(Subject).Count);
    }

    [Fact]
    public void IncompatibleWireTypeDoesNotHideEarlierCompatibleValue() {
        var properties = new List<MapiProperty> {
            new MapiProperty(0x0037, MapiPropertyType.Unicode, "valid"),
            new MapiProperty(0x0037, MapiPropertyType.Integer32, 42)
        };
        var bag = new MapiPropertyBag(properties);

        Assert.True(bag.TryGetValue(Subject, out string? value));
        Assert.Equal("valid", value);
        Assert.Equal("valid", bag.Find(Subject)?.Value);
        Assert.Equal(42, bag.FindRaw(Subject)?.Value);
        Assert.Single(bag.FindAll(Subject));
        Assert.Equal(2, bag.FindAllRaw(Subject).Count);
    }

    [Fact]
    public void SetCollapsesOnlyMatchingIdentityAndClearsRawSerializationState() {
        var unknown = new MapiProperty(0x66AA, MapiPropertyType.Binary, new byte[] { 7 }) {
            RawData = new byte[] { 7 }
        };
        var properties = new List<MapiProperty> {
            unknown,
            new MapiProperty(0x0037, MapiPropertyType.String8, "old", 0x17) {
                RawData = new byte[] { 1, 2, 3 }
            },
            new MapiProperty(0x0037, MapiPropertyType.Unicode, "newer", 0x23) {
                RawData = new byte[] { 4, 5, 6 }
            }
        };
        var bag = new MapiPropertyBag(properties);

        MapiProperty replacement = bag.Set(Subject, "replacement");

        Assert.Equal(2, properties.Count);
        Assert.Same(unknown, properties[0]);
        Assert.Same(replacement, properties[1]);
        Assert.Equal(MapiPropertyType.Unicode, replacement.PropertyType);
        Assert.Equal((uint)0x23, replacement.Flags);
        Assert.Null(replacement.RawData);
    }

    [Fact]
    public void NamedSetIsCaseInsensitiveAndRetainsArtifactMappedPropertyId() {
        Guid set = Guid.NewGuid();
        var key = new MapiPropertyKey<string>("CustomProject", set, "Project",
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        var properties = new List<MapiProperty> {
            new MapiProperty(0x8123, MapiPropertyType.String8, "old", 0x31,
                new MapiNamedProperty(set, "PROJECT"))
        };
        var bag = new MapiPropertyBag(properties);

        MapiProperty replacement = bag.Set(key, "new");

        Assert.Equal((ushort)0x8123, replacement.PropertyId);
        Assert.Equal((uint)0x31, replacement.Flags);
        Assert.Equal("new", bag.GetValueOrDefault(key));
        Assert.Single(bag.FindAll(key));
    }

    [Fact]
    public void NumericNamedIdentityDoesNotAliasStringNamedIdentity() {
        Guid set = Guid.NewGuid();
        var numeric = new MapiPropertyKey<int>("Numeric", set, 0x1234, MapiPropertyType.Integer32);
        var text = new MapiPropertyKey<int>("Text", set, "1234", MapiPropertyType.Integer32);
        var properties = new List<MapiProperty> {
            new MapiProperty(0x8000, MapiPropertyType.Integer32, 1, name: new MapiNamedProperty(set, 0x1234)),
            new MapiProperty(0x8001, MapiPropertyType.Integer32, 2, name: new MapiNamedProperty(set, "1234"))
        };
        var bag = new MapiPropertyBag(properties);

        Assert.Equal(1, bag.GetValueOrDefault(numeric));
        Assert.Equal(2, bag.GetValueOrDefault(text));
        Assert.NotEqual(properties[0].Name, properties[1].Name);
    }

    [Fact]
    public void KnownVocabularyResolvesByCanonicalStandardAndNamedIdentity() {
        Assert.Same(MapiKnownProperties.PidTag.Subject, MapiKnownProperties.Find("pidtagsubject"));
        Assert.Same(MapiKnownProperties.PidTag.Subject, MapiKnownProperties.Find((ushort)0x0037));
        Assert.Same(MapiKnownProperties.PidTag.Subject,
            MapiKnownProperties.Find((ushort)0x0037, MapiPropertyType.String8));
        Assert.Same(MapiKnownProperties.PidTag.Subject,
            MapiKnownProperties.Find(new MapiProperty(0x0037, MapiPropertyType.Unicode)));
        Assert.Same(MapiKnownProperties.PidLid.ReminderSet,
            MapiKnownProperties.Find(MapiPropertySets.Common, 0x8503));
        Assert.Same(MapiKnownProperties.PidName.Keywords,
            MapiKnownProperties.Find(MapiPropertySets.PublicStrings, "keywords"));
    }

    [Fact]
    public void KnownVocabularyDisambiguatesSharedStandardIdByWireType() {
        Assert.Null(MapiKnownProperties.Find((ushort)0x3A1B));
        Assert.Same(MapiKnownProperties.PidTag.Business2TelephoneNumber,
            MapiKnownProperties.Find((ushort)0x3A1B, MapiPropertyType.Unicode));
        Assert.Same(MapiKnownProperties.PidTag.Business2TelephoneNumbers,
            MapiKnownProperties.Find((ushort)0x3A1B, MapiPropertyType.MultipleUnicode));
    }

    [Fact]
    public void PropertySetsAndStandardIdAccessUseCanonicalIdentities() {
        Assert.Equal(new Guid("00062040-0000-0000-C000-000000000046"), MapiPropertySets.Sharing);
        Assert.Equal((ushort)0x0037, MapiKnownProperties.PidTag.Subject.GetStandardPropertyId());
        Assert.Throws<InvalidOperationException>(() =>
            MapiKnownProperties.PidName.Keywords.GetStandardPropertyId());
    }

    [Fact]
    public void ExplicitWireTypeSupportsAttachmentBinaryAndObjectPlaceholderContracts() {
        var bag = new MapiPropertyBag(new List<MapiProperty>());

        MapiProperty binary = bag.Set(MapiKnownProperties.PidTag.AttachData,
            new byte[] { 1, 2, 3 }, MapiPropertyType.Binary);
        MapiProperty placeholder = bag.SetNull(MapiKnownProperties.PidTag.AttachData,
            MapiPropertyType.Object);

        Assert.Equal(MapiPropertyType.Binary, binary.PropertyType);
        Assert.Equal(MapiPropertyType.Object, placeholder.PropertyType);
        Assert.Null(placeholder.Value);
        Assert.Same(placeholder, bag.Find(MapiKnownProperties.PidTag.AttachData));
        Assert.Throws<ArgumentException>(() => bag.SetNull(MapiKnownProperties.PidTag.AttachData,
            MapiPropertyType.Binary));
        Assert.Throws<ArgumentException>(() => bag.Set(MapiKnownProperties.PidTag.AttachData,
            new byte[] { 4 }, MapiPropertyType.Unicode));
    }
}
