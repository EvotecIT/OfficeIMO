using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreAssociatedDataTests {
    [Fact]
    public void Associated_information_round_trips_through_one_typed_catalog() {
        string path = TemporaryPstPath();
        DateTimeOffset stamp = new DateTimeOffset(2026, 7, 18, 10, 0, 0, TimeSpan.Zero);
        try {
            var categories = EmailStoreCategoryList.Create();
            Guid categoryId = new Guid("11111111-2222-3333-4444-555555555555");
            categories.Set("Customer", color: 7, keyboardShortcut: 1, id: categoryId,
                usageCount: 4, lastUsed: stamp);
            categories.DefaultCategoryName = "Customer";

            var dictionary = EmailStoreConfigurationDictionary.Create("Outlook.16");
            dictionary.Set("OLPrefsVersion", 1);
            dictionary.Set("piRemindDefault", true);
            dictionary.Set("piRemindDefaultDelta", 15);
            var calendarConfiguration = new EmailDocument {
                MessageClass = "IPM.Configuration.Calendar",
                Subject = "Calendar options"
            };
            calendarConfiguration.Mapi.Set(MapiKnownProperties.PidTag.RoamingDatatypes, 0x00000004);
            calendarConfiguration.Mapi.Set(MapiKnownProperties.PidTag.RoamingDictionary, dictionary.ToXml());
            calendarConfiguration.MessageMetadata.ModifiedDate = stamp;

            var view = new EmailDocument { MessageClass = "IPM.Microsoft.FolderDesign.NamedView" };
            view.Mapi.Set(MapiKnownProperties.PidTag.ViewDescriptorName, "Customer view");
            view.Mapi.Set(MapiKnownProperties.PidTag.ViewDescriptorVersion, 8);
            view.Mapi.Set(MapiKnownProperties.PidTag.ViewDescriptorBinary, ViewDescriptorHeader());
            view.Mapi.Set(MapiKnownProperties.PidTag.ViewDescriptorStrings, Array.Empty<byte>());

            var rule = new EmailDocument {
                MessageClass = "IPM.RuleOrganizer",
                Subject = "Outlook Rules Organizer"
            };
            rule.Mapi.Set(MapiKnownProperties.PidTag.RwRulesStream, new byte[] { 1, 3, 3, 7 });

            var search = new EmailDocument { MessageClass = "IPM.Microsoft.Wunderbar.SFInfo" };
            search.Mapi.Set(MapiKnownProperties.PidTag.SearchFolderTemplateId, 7);
            search.Mapi.Set(MapiKnownProperties.PidTag.SearchFolderId,
                new Guid("AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE").ToByteArray());
            search.Mapi.Set(MapiKnownProperties.PidTag.SearchFolderStorageType, 0);
            search.Mapi.Set(MapiKnownProperties.PidTag.SearchFolderDefinition, SearchDefinitionEnvelope());

            var folderFields = new EmailDocument {
                MessageClass = "IPM.Configuration.FolderFields",
                Subject = "Folder fields"
            };
            folderFields.UserProperties.Set("Case Number", "EVT-42", OutlookUserPropertyType.Text);

            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string calendar = writer.AddFolder("Calendar", EmailStoreSpecialFolderKind.Calendar,
                    containerClass: "IPF.Appointment");
                string inbox = writer.AddFolder("Inbox", EmailStoreSpecialFolderKind.Inbox,
                    containerClass: "IPF.Note");
                string commonViews = writer.AddFolder("Common Views", EmailStoreSpecialFolderKind.CommonViews,
                    containerClass: "IPF.Note");
                string projects = writer.AddFolder("Projects", containerClass: "IPF.Note");
                writer.AddItem(projects, new EmailDocument { MessageClass = "IPM.Note", Subject = "Visible" });
                writer.AddItem(calendar, categories.ToAssociatedDocument(stamp), isAssociated: true);
                writer.AddItem(calendar, calendarConfiguration, isAssociated: true);
                writer.AddItem(commonViews, view, isAssociated: true);
                writer.AddItem(inbox, rule, isAssociated: true);
                writer.AddItem(commonViews, search, isAssociated: true);
                writer.AddItem(projects, folderFields, isAssociated: true);
                writer.Complete();
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            EmailStoreItemReference[] associatedOnly = session.EnumerateItems(
                EmailStoreEnumerationOptions.ForAssociated(maxItems: 100)).ToArray();
            Assert.Equal(6, associatedOnly.Length);
            Assert.All(associatedOnly, reference => Assert.True(reference.IsAssociated));

            EmailStoreAssociatedDataCatalog catalog = session.ReadAssociatedData();
            Assert.True(catalog.IsComplete, string.Join(" | ", catalog.Diagnostics.Select(item => item.Code)));
            Assert.Equal(6, catalog.Items.Count);
            Assert.Equal(3, catalog.Configurations.Count);
            Assert.Single(catalog.CategoryLists);
            Assert.Single(catalog.Views);
            Assert.Single(catalog.RuleOrganizers);
            Assert.Single(catalog.SearchFolders);
            Assert.Single(catalog.FolderUserProperties);

            EmailStoreCategoryList categoryList = catalog.CategoryLists[0].CategoryList!;
            EmailStoreCategoryDefinition category = Assert.Single(categoryList.Categories);
            Assert.Equal("Customer", category.Name);
            Assert.Equal(7, category.Color);
            Assert.Equal(categoryId, category.Id);
            Assert.Equal("Customer", categoryList.DefaultCategoryName);

            EmailStoreAssociatedItem effective = catalog.FindEffectiveConfiguration(
                "IPM.Configuration.Calendar")!;
            Assert.True(effective.Configuration!.Dictionary!.TryGet("piRemindDefault", out EmailStoreConfigurationValue? remind));
            Assert.True(remind!.TryGet(out bool enabled));
            Assert.True(enabled);

            Assert.Equal((uint)8, catalog.Views[0].ViewDefinition!.BinaryVersion);
            Assert.True(catalog.Views[0].ViewDefinition!.IsProtocolEnvelopeValid);
            Assert.Equal(4, catalog.RuleOrganizers[0].RuleOrganizer!.RuleDataLength);
            Assert.NotNull(catalog.RuleOrganizers[0].RuleOrganizer!.FingerprintSha256);
            Assert.True(catalog.SearchFolders[0].SearchFolderDefinition!.IsProtocolEnvelopeValid);
            Assert.Contains(catalog.FolderUserProperties[0].FolderUserProperties!.Definitions,
                definition => definition.Name == "Case Number");
            Assert.Contains(catalog.SearchFolderContainers, container => container.Folder.IsSearchFolder);
        } finally {
            TryDelete(path);
        }
    }

    [Fact]
    public void Category_editor_preserves_unknown_xml_and_stable_identity() {
        Guid id = new Guid("11111111-2222-3333-4444-555555555555");
        string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<categories xmlns=\"CategoryList.xsd\" default=\"Alpha\" lastSavedSession=\"0\" " +
            "lastSavedTime=\"2026-07-18T10:00:00.000Z\" futureRoot=\"keep\">" +
            "<category name=\"Alpha\" color=\"1\" keyboardShortcut=\"0\" " +
            "lastTimeUsed=\"2026-07-18T10:00:00.000Z\" lastSessionUsed=\"0\" " +
            "guid=\"{11111111-2222-3333-4444-555555555555}\" future=\"keep\" />" +
            "<futureElement value=\"keep\" /></categories>";
        var document = new EmailDocument { MessageClass = "IPM.Configuration.CategoryList" };
        document.Mapi.Set(MapiKnownProperties.PidTag.RoamingDatatypes, 2);
        document.Mapi.Set(MapiKnownProperties.PidTag.RoamingXmlStream, Encoding.UTF8.GetBytes(xml));

        EmailStoreCategoryList list = EmailStoreCategoryList.Parse(document);
        EmailStoreCategoryDefinition updated = list.Set("Alpha", color: 24, keyboardShortcut: 2);
        string written = Encoding.UTF8.GetString(list.ToXml(
            new DateTimeOffset(2026, 7, 19, 10, 0, 0, TimeSpan.Zero)));

        Assert.Equal(id, updated.Id);
        Assert.Contains("futureRoot=\"keep\"", written);
        Assert.Contains("future=\"keep\"", written);
        Assert.Contains("futureElement", written);
    }

    [Fact]
    public void Reminder_queue_respects_domain_state_and_signal_evidence() {
        string path = TemporaryPstPath();
        DateTimeOffset asOf = new DateTimeOffset(2026, 7, 18, 12, 0, 0, TimeSpan.Zero);
        try {
            var overdue = new EmailDocument { MessageClass = "IPM.Note", Subject = "Overdue" };
            overdue.MessageMetadata.Reminder.IsSet = true;
            overdue.MessageMetadata.Reminder.Time = asOf.AddMinutes(-5);
            overdue.MessageMetadata.Reminder.SignalTime = asOf.AddMinutes(-5);

            var appointment = new EmailDocument {
                MessageClass = "IPM.Appointment",
                OutlookItemKind = OutlookItemKind.Appointment,
                Subject = "Pending appointment",
                Appointment = new OutlookAppointment {
                    Start = asOf.AddHours(2),
                    End = asOf.AddHours(3),
                    IsRecurring = false
                }
            };
            appointment.Appointment.Reminder.IsSet = true;
            appointment.Appointment.Reminder.DeltaMinutes = 30;
            appointment.Appointment.Reminder.Time = appointment.Appointment.Start;

            var excluded = new EmailDocument { MessageClass = "IPM.Note", Subject = "Draft reminder" };
            excluded.MessageMetadata.Reminder.IsSet = true;
            excluded.MessageMetadata.Reminder.SignalTime = asOf.AddMinutes(-10);

            var disabled = new EmailDocument { MessageClass = "IPM.Note", Subject = "Disabled" };
            disabled.MessageMetadata.Reminder.IsSet = false;
            disabled.MessageMetadata.Reminder.SignalTime = asOf.AddHours(1);

            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path)) {
                string inbox = writer.AddFolder("Inbox", EmailStoreSpecialFolderKind.Inbox,
                    containerClass: "IPF.Note");
                string calendar = writer.AddFolder("Calendar", EmailStoreSpecialFolderKind.Calendar,
                    containerClass: "IPF.Appointment");
                string drafts = writer.AddFolder("Drafts", EmailStoreSpecialFolderKind.Drafts,
                    containerClass: "IPF.Note");
                writer.AddItem(inbox, overdue);
                writer.AddItem(calendar, appointment);
                writer.AddItem(drafts, excluded);
                writer.AddItem(inbox, disabled);
                writer.Complete();
            }

            using EmailStoreSession session = EmailStoreSession.Open(path);
            EmailStoreReminderQueue active = session.GetReminders(
                new EmailStoreReminderQueryOptions(asOf: asOf));
            Assert.True(active.IsComplete);
            Assert.Equal(2, active.Items.Count);
            Assert.Equal("Overdue", active.Items[0].Summary.Subject);
            Assert.Equal(EmailStoreReminderState.Overdue, active.Items[0].State);
            Assert.Equal(EmailStoreReminderSignalSource.ReminderSignalTime, active.Items[0].SignalSource);
            Assert.Equal("Pending appointment", active.Items[1].Summary.Subject);
            Assert.Equal(EmailStoreReminderState.Pending, active.Items[1].State);
            Assert.Equal(asOf.AddMinutes(90), active.Items[1].SignalTime);
            Assert.Equal(EmailStoreReminderSignalSource.AppointmentStartMinusDelta, active.Items[1].SignalSource);
            Assert.DoesNotContain(active.Items, item => item.Summary.Subject == "Draft reminder");

            EmailStoreReminderQueue all = session.GetReminders(
                new EmailStoreReminderQueryOptions(asOf: asOf,
                    includeInactive: true, includeExcludedFolders: true));
            Assert.Equal(4, all.Items.Count);
            Assert.Contains(all.Items, item => item.Summary.Subject == "Disabled" &&
                item.State == EmailStoreReminderState.Disabled);
            Assert.Contains(all.Items, item => item.Summary.Subject == "Draft reminder");
        } finally {
            TryDelete(path);
        }
    }

    private static byte[] ViewDescriptorHeader() {
        var bytes = new byte[64];
        WriteUInt32LittleEndian(bytes, 8, 8);
        WriteUInt32LittleEndian(bytes, 12, 0);
        WriteUInt32LittleEndian(bytes, 20, 1);
        WriteUInt32LittleEndian(bytes, 24, 0);
        WriteUInt32LittleEndian(bytes, 28, 0);
        return bytes;
    }

    private static byte[] SearchDefinitionEnvelope() {
        var bytes = new byte[34];
        WriteUInt32BigEndian(bytes, 0, 0x04100000);
        WriteUInt32BigEndian(bytes, 4, 0);
        WriteUInt32BigEndian(bytes, 8, 0);
        // TextSearchLength=0, SkipBlock1=0, DeepSearch=0, FolderList1Length=0,
        // FolderList2Length=0, SkipBlock2=0, SkipBlock3=0.
        return bytes;
    }

    private static void WriteUInt32LittleEndian(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static void WriteUInt32BigEndian(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static string TemporaryPstPath() => Path.Combine(Path.GetTempPath(),
        string.Concat("officeimo-associated-data-", Guid.NewGuid().ToString("N"), ".pst"));

    private static void TryDelete(string path) {
        if (File.Exists(path)) File.Delete(path);
    }
}
