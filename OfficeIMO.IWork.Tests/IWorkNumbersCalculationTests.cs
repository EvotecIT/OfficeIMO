using System.Buffers.Binary;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNumbersCalculationTests {
    [Fact]
    public void Created_Numbers_tables_register_their_native_formula_owners_with_exact_identities_and_rectangular_bounds() {
        var source = IWorkSourceDocument.Open(IWorkNumbersWritingTests.Example().SaveBytes());
        var records = source.Records.ToDictionary(record => record.Identifier);
        var calculationRecord = Assert.Single(source.Records, record => record.MessageType == 4000);
        var tracker = source.Index.Message(calculationRecord).GetMessage(2)!;
        var registrations = tracker.GetRepeatedMessages(6).ToArray();
        var ownerMap = tracker.GetMessage(3)!.GetRepeatedMessages(1).ToArray();
        var tables = source.Records.Where(record => record.MessageType == 6000).ToArray();
        Assert.Equal(3, tables.Length);
        Assert.Equal(tables.Length, registrations.Length); Assert.Equal(tables.Length, ownerMap.Length);
        Assert.Equal(0UL, tracker.GetUnsigned(5)); // No authored formulas yet; native entry must still calculate.
        for (int index = 0; index < tables.Length; index++) {
            var info = source.Index.Message(tables[index]);
            var model = source.Index.Message(records[info.GetMessage(2)!.GetUnsigned(1)!.Value]);
            byte[] uuid = Convert.FromHexString(Guid.Parse(model.GetString(1)!).ToString("N"));
            ulong registrationId = registrations[index].GetUnsigned(1)!.Value;
            var registration = records[registrationId];
            Assert.Equal(4008u, registration.MessageType);
            Assert.Contains(registrationId, calculationRecord.ObjectReferences);
            Assert.Equal(new[] { tables[index].Identifier }, registration.ObjectReferences);
            var dependency = source.Index.Message(registration);
            Assert.Equal((ulong)index + 1, dependency.GetUnsigned(2));
            Assert.Equal(1UL, dependency.GetUnsigned(3)); // Native TABLE_MODEL owner, referring to TableInfo.
            Assert.Equal(tables[index].Identifier, dependency.GetMessage(11)!.GetUnsigned(1));
            Assert.Equal(BinaryPrimitives.ReadUInt64LittleEndian(uuid), dependency.GetMessage(1)!.GetUnsigned(1));
            Assert.Equal(BinaryPrimitives.ReadUInt64LittleEndian(uuid.AsSpan(8)), dependency.GetMessage(1)!.GetUnsigned(2));
            Assert.Equal((ulong)index + 1, ownerMap[index].GetUnsigned(1));
            for (int word = 0; word < 4; word++)
                Assert.Equal((ulong)BinaryPrimitives.ReadUInt32LittleEndian(uuid.AsSpan(word * 4)),
                    ownerMap[index].GetMessage(2)!.GetUnsigned(word + 2));
            foreach (int spanningField in new[] { 7, 8 }) {
                foreach (int rangeField in new[] { 2, 3 }) {
                    var range = dependency.GetMessage(spanningField)!.GetMessage(rangeField)!;
                    Assert.Equal(0UL, range.GetUnsigned(1)); Assert.Equal(0UL, range.GetUnsigned(2));
                    Assert.Equal(model.GetUnsigned(7)!.Value - 1, range.GetUnsigned(3));
                    Assert.Equal(model.GetUnsigned(6)!.Value - 1, range.GetUnsigned(4));
                }
            }
        }
        Assert.All(source.Records, record => Assert.All(record.ObjectReferences, id => Assert.Contains(id, records.Keys)));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Numbers, 0UL)]
    [InlineData(IWorkDocumentKind.Pages, 0UL)]
    [InlineData(IWorkDocumentKind.Keynote, 1UL)]
    public void Native_table_formula_coordinate_space_matches_the_receiving_application(IWorkDocumentKind kind, ulong expected) {
        var data = IWorkTableData.Create(2, 3).SetNumber(1, 1, 24);
        byte[] package;
        if (kind == IWorkDocumentKind.Numbers) {
            var document = IWorkNumbersDocument.Create(); document.AddSheet("Values").AddTable(data); package = document.SaveBytes();
        } else if (kind == IWorkDocumentKind.Pages) {
            var document = IWorkPagesDocument.Create(); document.AddTable(data); package = document.SaveBytes();
        } else {
            var document = IWorkKeynoteDocument.Create(); document.AddSlide().AddTable(data, 0, 0, 400, 200); package = document.SaveBytes();
        }
        var source = IWorkSourceDocument.Open(package);
        Assert.Equal(expected, source.Index.Message(Assert.Single(source.Records, record => record.MessageType == 6000)).GetUnsigned(10));
    }
}
