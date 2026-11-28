#region QuickStart
using OfficeIMO.Access;

using AccessDocument database = AccessDocument.Create(new AccessCreateOptions {
    Format = AccessFileFormat.Accdb,
    DatabaseTitle = "Contacts"
});
AccessTable contacts = database.Tables.Add("Contacts");
contacts.Columns.AddAutoNumber("Id", seed: 101);
contacts.Columns.Add("Name", AccessDataType.ShortText, maxLength: 120);
contacts.Indexes.AddPrimaryKey("PK_Contacts", "Id");

using (AccessUpdateScope edit = database.BeginUpdate()) {
    contacts.AppendRow(new AccessRowValues { ["Name"] = "Ada" });
    edit.Commit();
}
database.AssessSave("Contacts.accdb").ThrowIfLoss();
database.Save("Contacts.accdb");
#endregion
