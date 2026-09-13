using OfficeIMO.Core.Internal;
using System.Text;
using System.Xml.Linq;

// Layout hypotheses qualified only by this producer corpus. This is deliberately
// outside the runtime package until version detection and wider records are proved.
internal static class NativeFixtureRecords {
    internal static object ExtractAndCompare(OfficeCompoundFile compound, string xmlPath) {
        var names = Names(compound, "TBkndTask", 0x0b40000e);
        var calendars = Names(compound, "TBkndCal", 0x0d400001);
        var assignments = AssignmentIdentities(compound);
        var expected = XDocument.Load(xmlPath);
        XNamespace ns = "http://schemas.microsoft.com/project";
        foreach (var name in names) {
            var task = expected.Root!.Element(ns + "Tasks")!.Elements(ns + "Task").Single(t => (int)t.Element(ns + "UID")! == name.Uid);
            if ((string?)task.Element(ns + "Name") != name.Name) throw new InvalidDataException("Native task name disagrees with producer XML.");
        }
        foreach (var calendar in calendars) {
            var expectedCalendar = expected.Root!.Element(ns + "Calendars")!.Elements(ns + "Calendar").SingleOrDefault(t => (int)t.Element(ns + "UID")! == calendar.Uid);
            // Project can retain an unused resource calendar in MPP while omitting it from XML.
            if (expectedCalendar != null && (string?)expectedCalendar.Element(ns + "Name") != calendar.Name)
                throw new InvalidDataException("Native calendar name disagrees with producer XML.");
        }
        var expectedAssignments = expected.Root!.Element(ns + "Assignments")?.Elements(ns + "Assignment").ToArray() ?? Array.Empty<XElement>();
        foreach (var assignment in expectedAssignments) {
            if (!assignments.Any(a => a.Uid == (int)assignment.Element(ns + "UID")! &&
                a.TaskUid == (int)assignment.Element(ns + "TaskUID")! && a.ResourceUid == (int)assignment.Element(ns + "ResourceUID")!))
                throw new InvalidDataException("Native assignment identity disagrees with producer XML.");
        }
        return new { names, calendars, assignments, producerXmlCompared = true,
            qualification = "MPP14 streams from Project 2024 build 16.0.20326.20144 only; names and assignment identities, not full records." };
    }

    private static byte[] Stream(OfficeCompoundFile file, string table, string name) => file.Streams.Single(s => s.Key.EndsWith("/" + table + "/" + name, StringComparison.Ordinal)).Value;

    private static List<NamedRecord> Names(OfficeCompoundFile file, string table, uint fieldId) {
        byte[] metadata = Stream(file, table, "VarMeta"), data = Stream(file, table, "Var2Data");
        if (metadata.Length < 24 || (metadata.Length - 24) % 12 != 0 || BitConverter.ToUInt32(metadata, 0) != 0xfadfadba)
            throw new InvalidDataException("Unqualified variable metadata layout.");
        var result = new List<NamedRecord>();
        for (int i = 24; i < metadata.Length; i += 12) {
            if (BitConverter.ToUInt32(metadata, i + 8) != fieldId) continue;
            int uid = BitConverter.ToInt32(metadata, i), offset = BitConverter.ToInt32(metadata, i + 4);
            if (offset < 0 || offset > data.Length - 4) throw new InvalidDataException("Variable record offset outside stream.");
            int length = BitConverter.ToInt32(data, offset);
            if (length < 0 || length > data.Length - offset - 4 || length % 2 != 0) throw new InvalidDataException("Invalid UTF-16 variable record length.");
            result.Add(new NamedRecord(uid, Encoding.Unicode.GetString(data, offset + 4, length).TrimEnd('\0')));
        }
        return result;
    }

    private static List<AssignmentRecord> AssignmentIdentities(OfficeCompoundFile file) {
        byte[] metadata = Stream(file, "TBkndAssn", "FixedMeta"), data = Stream(file, "TBkndAssn", "FixedData");
        if (metadata.Length < 16 || (metadata.Length - 16) % 34 != 0) throw new InvalidDataException("Unqualified assignment metadata layout.");
        var result = new List<AssignmentRecord>();
        for (int i = 16; i < metadata.Length; i += 34) {
            if ((BitConverter.ToUInt32(metadata, i) & 2) != 0) continue;
            int offset = BitConverter.ToInt32(metadata, i + 4);
            if (offset < 0 || offset > data.Length - 12) throw new InvalidDataException("Assignment offset outside stream.");
            result.Add(new AssignmentRecord(BitConverter.ToInt32(data, offset), BitConverter.ToInt32(data, offset + 4), BitConverter.ToInt32(data, offset + 8)));
        }
        return result;
    }
    private sealed record NamedRecord(int Uid, string Name);
    private sealed record AssignmentRecord(int Uid, int TaskUid, int ResourceUid);
}
