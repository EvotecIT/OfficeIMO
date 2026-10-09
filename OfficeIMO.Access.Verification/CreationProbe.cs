using OfficeIMO.Access;
using System.Text.Json;

namespace OfficeIMO.Access.Verification {
    /// <summary>Opt-in executable public model examples; independent DAO/Access qualification is separate.</summary>
    internal static class CreationProbe {
        internal static void Generate(string outputDirectory) {
            string root = Path.GetFullPath(outputDirectory);
            if (Directory.Exists(root) || File.Exists(root)) throw new IOException("Creation qualification requires a fresh output directory.");
            Directory.CreateDirectory(root); List<object> observations = new List<object>();
            foreach (AccessFileFormat format in new[] { AccessFileFormat.Mdb, AccessFileFormat.Accdb }) {
                string extension = format == AccessFileFormat.Mdb ? ".mdb" : ".accdb";
                using (AccessDocument empty = AccessDocument.Create(new AccessCreateOptions { Format = format })) empty.Save(Path.Combine(root, "empty" + extension));
                GenerateBoundaries(root, format, extension);
                using AccessDocument document = AccessDocument.Create(new AccessCreateOptions { Format = format, DatabaseTitle = "Synthetic native creation" });
                AccessTable groups = document.Tables.Add("Groups"); groups.Columns.Add("Id", AccessDataType.Int32); groups.Columns.Add("Label", AccessDataType.ShortText, 80); groups.Indexes.AddPrimaryKey("PK_Groups", "Id");
                groups.AppendRow(new AccessRowValues { ["Id"] = 1, ["Label"] = "Group A" });
                AccessTable contacts = document.Tables.Add("Contacts");
                contacts.Columns.Add("Id", AccessDataType.AutoNumber); contacts.Columns.Add("GroupId", AccessDataType.Int32); contacts.Columns.Add("DisplayName", AccessDataType.ShortText, 120);
                contacts.Columns.Add("Amount", AccessDataType.Currency); contacts.Columns.Add("CreatedAt", AccessDataType.DateTime); contacts.Columns.Add("Active", AccessDataType.Boolean);
                contacts.Columns.Add("Notes", AccessDataType.LongText); contacts.Columns.Add("Payload", AccessDataType.Binary);
                contacts.Columns.AddDecimal("Precise", 28, 9); contacts.Columns.Add("Identifier", AccessDataType.Guid); contacts.Columns.Add("Ratio", AccessDataType.Double); contacts.Columns.Add("SingleValue", AccessDataType.Single); contacts.Columns.Add("Small", AccessDataType.Int16); contacts.Columns.Add("Octet", AccessDataType.Byte);
                contacts.Indexes.AddPrimaryKey("PK_Contacts", "Id"); contacts.Indexes.AddUnique("NameUnique", "DisplayName");
                document.Relationships.Add("ContactGroups", groups.Columns["Id"], contacts.Columns["GroupId"]);
                string notes = string.Concat(Enumerable.Repeat("Ł🙂 Synthetic text\r\n", 1000));
                byte[] payload = Enumerable.Range(0, 20000).Select(x => (byte)(x * 17 % 251)).ToArray();
                for (int i = 0; i < 5000; i++) contacts.AppendRow(new AccessRowValues {
                    ["GroupId"] = 1, ["DisplayName"] = "Contact" + i.ToString("D5"), ["Amount"] = i == 0 ? -1.2345m : 12.3456m,
                    ["CreatedAt"] = i == 0 ? new DateTime(2026, 1, 2, 3, 4, 5) : null, ["Active"] = i % 2 == 0,
                    ["Notes"] = i == 0 ? notes : i == 1 ? "" : null, ["Payload"] = i == 0 ? payload : i == 1 ? Array.Empty<byte>() : null,
                    ["Precise"] = i == 0 ? 1234567890123456789.123456789m : -12.300000001m,
                    ["Identifier"] = Guid.Parse("01234567-89ab-cdef-0123-456789abcdef"), ["Ratio"] = -1.25d, ["SingleValue"] = 1.5f, ["Small"] = (short)-32768, ["Octet"] = (byte)255
                });
                string path = Path.Combine(root, "created" + extension);
                document.AssessSave(path).RequireNoLoss(); document.Save(path);
                using AccessDocument decoded = AccessDocument.Load(path);
                if (decoded.Tables["Contacts"].RowCount != 5000 || decoded.Relationships.Count != 1) throw new InvalidDataException("Native schema/row count verification failed.");
                using AccessDataReader reader = decoded.Tables["Contacts"].OpenDataReader(); if (!reader.Read() || (int)reader["Id"] != 1 || (decimal)reader["Precise"] != 1234567890123456789.123456789m || (string)reader["Notes"] != notes || !((byte[])reader["Payload"]).SequenceEqual(payload)) throw new InvalidDataException("Native scalar/long-value verification failed.");
                observations.Add(new { file = Path.GetFileName(path), rows = decoded.Tables["Contacts"].RowCount, profile = decoded.Profile.ToString(), bytes = new FileInfo(path).Length, sha256 = decoded.Inspection!.Sha256 });
            }
            File.WriteAllText(Path.Combine(root, "creation.json"), JsonSerializer.Serialize(observations, new JsonSerializerOptions { WriteIndented = true }));
        }
        private static void GenerateBoundaries(string root, AccessFileFormat format, string extension) {
            using (AccessDocument keys = AccessDocument.Create(new AccessCreateOptions { Format = format })) {
                AccessTable table = keys.Tables.Add("KeyValues");
                table.Columns.Add("Id", AccessDataType.Byte); table.Columns.Add("Small", AccessDataType.Int16); table.Columns.Add("Name", AccessDataType.ShortText, 100);
                table.Columns.Add("Owner", AccessDataType.Int32); table.Columns.Add("SID", AccessDataType.Guid);
                table.Indexes.AddPrimaryKey("PK_KeyValues", "Id"); table.Indexes.AddUnique("NameKey", "Name"); table.Indexes.AddUnique("CompositeKey", "Small", "Id");
                string[] names = { "", "_", "Alpha Beta", "Alpha_Beta", " alpha", "Gamma ", "Z9" };
                for (int i = 0; i < names.Length; i++) table.AppendRow(new AccessRowValues { ["Id"] = (byte)(i + 1), ["Small"] = (short)(short.MinValue + i), ["Name"] = names[i], ["Owner"] = 123, ["SID"] = Guid.Parse("01234567-89ab-cdef-0123-456789abcdef") });
                AccessTable tree = keys.Tables.Add("Tree"); tree.Columns.Add("Id", AccessDataType.Int32); tree.Columns.Add("ParentId", AccessDataType.Int32); tree.Indexes.AddPrimaryKey("PK_Tree", "Id");
                keys.Relationships.Add("TreeParent", tree.Columns["Id"], tree.Columns["ParentId"]);
                tree.AppendRow(new AccessRowValues { ["Id"] = 1, ["ParentId"] = null }); tree.AppendRow(new AccessRowValues { ["Id"] = 2, ["ParentId"] = 1 });
                keys.Save(Path.Combine(root, "keys" + extension));
            }
            using (AccessDocument columns = AccessDocument.Create(new AccessCreateOptions { Format = format })) {
                AccessTable table = columns.Tables.Add("LongColumns"); AccessRowValues row = new AccessRowValues();
                for (int i = 0; i < 255; i++) { string name = "Field" + i.ToString("D3"); table.Columns.Add(name, AccessDataType.Binary); row[name] = Enumerable.Repeat((byte)i, 300).ToArray(); }
                table.AppendRow(row); columns.Save(Path.Combine(root, "long-columns" + extension));
            }
            using (AccessDocument wide = AccessDocument.Create(new AccessCreateOptions { Format = format })) {
                AccessTable table = wide.Tables.Add("Wide"); AccessRowValues row = new AccessRowValues();
                for (int i = 0; i < 255; i++) { string name = "Field" + i.ToString("D3"); table.Columns.Add(name, AccessDataType.Byte); row[name] = (byte)i; }
                table.AppendRow(row); wide.Save(Path.Combine(root, "wide" + extension));
            }
            using (AccessDocument large = AccessDocument.Create(new AccessCreateOptions { Format = format })) {
                AccessTable table = large.Tables.Add("Allocated"); table.Columns.AddAutoNumber("Id", 1001); table.Columns.Add("Padding", AccessDataType.ShortText, 255); table.Indexes.AddPrimaryKey("PK_Allocated", "Id");
                for (int i = 0; i < 6000; i++) table.AppendRow(new AccessRowValues { ["Padding"] = new string('a', 255) });
                large.Save(Path.Combine(root, "allocation" + extension));
            }
            using (AccessDocument values = AccessDocument.Create(new AccessCreateOptions { Format = format })) {
                AccessTable table = values.Tables.Add("Lengths"); table.Columns.Add("Id", AccessDataType.Int32); table.Columns.Add("Payload", AccessDataType.Binary); table.Indexes.AddPrimaryKey("PK_Lengths", "Id");
                foreach (int length in new[] { 0, 1, 64, 65, 4072, 4076, 4077, 8144, 8145 }) table.AppendRow(new AccessRowValues { ["Id"] = length, ["Payload"] = Enumerable.Range(0, length).Select(i => (byte)(i % 251)).ToArray() });
                values.Save(Path.Combine(root, "lengths" + extension));
            }
        }

        internal static void VerifyAccessEdits(string directory) {
            foreach (string extension in new[] { ".mdb", ".accdb" }) {
                using AccessDocument document = AccessDocument.Load(Path.Combine(directory, "access-edited" + extension));
                AccessTable contacts = document.Tables["Contacts"]; int count = 0; bool changed = false, appended = false;
                using AccessDataReader reader = contacts.OpenDataReader();
                while (reader.Read()) {
                    int id = (int)reader["Id"]; string name = (string)reader["DisplayName"]; count++;
                    if (id == 4999) throw new InvalidDataException("Access deletion was not decoded.");
                    if (id == 5000) changed = name == "ChangedByAccess";
                    if (id > 5000) appended |= name == "AppendedByAccess";
                    if (id == 1 && (((byte[])reader["Payload"]).Length != 20000 || (decimal)reader["Precise"] != 1234567890123456789.123456789m)) throw new InvalidDataException("Access-resaved scalar/long values differ.");
                }
                if (count != 5000 || !changed || !appended || !document.Relationships.Any(r => r.Name == "ContactGroups")) throw new InvalidDataException($"Access-resaved model differs: {extension}, rows={count}, changed={changed}, appended={appended}, relationships={string.Join(",", document.Relationships.Select(r => r.Name))}.");
            }
        }
    }
}
