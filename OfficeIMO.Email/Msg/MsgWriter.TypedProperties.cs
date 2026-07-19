namespace OfficeIMO.Email;

internal static partial class MsgWriter {
    private static void AddTypedProperties(MsgPropertyBuilder properties, EmailDocument document) {
        if (document.Appointment != null) {
            OutlookAppointment item = document.Appointment;
            properties.Set(MapiKnownProperties.PidLid.GlobalObjectId, item.GlobalObjectId);
            properties.Set(MapiKnownProperties.PidLid.CleanGlobalObjectId, item.CleanGlobalObjectId);
            properties.Set(MapiKnownProperties.PidLid.AppointmentStartWhole, item.Start);
            properties.Set(MapiKnownProperties.PidLid.AppointmentEndWhole, item.End);
            properties.Set(MapiKnownProperties.PidLid.CommonStart, item.Start);
            properties.Set(MapiKnownProperties.PidLid.CommonEnd, item.End);
            properties.Set(MapiKnownProperties.PidLid.Location, item.Location);
            properties.Set(MapiKnownProperties.PidLid.AppointmentSubType, item.IsAllDay);
            properties.Set(MapiKnownProperties.PidLid.BusyStatus, item.BusyStatus);
            properties.Set(MapiKnownProperties.PidLid.AppointmentStateFlags, item.MeetingStatus);
            properties.Set(MapiKnownProperties.PidLid.ResponseStatus, item.ResponseStatus);
            properties.Set(MapiKnownProperties.PidLid.AppointmentSequence, item.Sequence);
            properties.Set(MapiKnownProperties.PidLid.AppointmentDuration,
                item.DurationMinutes ?? GetDurationMinutes(item.Start, item.End));
            properties.Set(MapiKnownProperties.PidLid.AllAttendeesString,
                item.AllAttendees ?? JoinAppointmentAttendees(document, null));
            properties.Set(MapiKnownProperties.PidLid.ToAttendeesString,
                item.RequiredAttendees ?? JoinAppointmentAttendees(document, EmailRecipientKind.To));
            properties.Set(MapiKnownProperties.PidLid.CcAttendeesString,
                item.OptionalAttendees ?? JoinAppointmentAttendees(document, EmailRecipientKind.Cc));
            properties.Set(MapiKnownProperties.PidLid.AppointmentNotAllowPropose, item.NotAllowPropose);
            properties.Set(MapiKnownProperties.PidLid.RecurrenceType, item.RecurrenceType);
            properties.Set(MapiKnownProperties.PidLid.RecurrencePattern, item.RecurrencePattern);
            byte[]? appointmentRecurrence = item.Recurrence == null
                ? item.RecurrenceState
                : OutlookRecurrenceBinary.EncodeAppointment(item.Recurrence, document.OutlookCodePage ?? 1252);
            properties.Set(MapiKnownProperties.PidLid.AppointmentRecur, appointmentRecurrence);
            properties.Set(MapiKnownProperties.PidLid.Recurring,
                item.IsRecurring ?? (appointmentRecurrence != null || item.RecurrenceType.GetValueOrDefault() != 0));
            properties.Set(MapiKnownProperties.PidLid.ClientIntent, item.ClientIntentFlags);
            OutlookMessageSemanticsWriter.WriteReminder(properties, item.Reminder);
            properties.Set(MapiKnownProperties.PidLid.TimeZoneStruct, item.LegacyTimeZone == null
                ? item.TimeZoneStructure
                : OutlookTimeZoneBinary.EncodeStructure(item.LegacyTimeZone));
            properties.Set(MapiKnownProperties.PidLid.TimeZoneDescription, item.TimeZoneDescription);
            properties.Set(MapiKnownProperties.PidLid.AppointmentTimeZoneDefinitionStartDisplay,
                item.StartTimeZone == null ? item.StartTimeZoneDefinition :
                    OutlookTimeZoneBinary.EncodeDefinition(item.StartTimeZone));
            properties.Set(MapiKnownProperties.PidLid.AppointmentTimeZoneDefinitionEndDisplay,
                item.EndTimeZone == null ? item.EndTimeZoneDefinition :
                    OutlookTimeZoneBinary.EncodeDefinition(item.EndTimeZone));
            properties.Set(MapiKnownProperties.PidLid.AppointmentTimeZoneDefinitionRecur,
                item.RecurrenceTimeZone == null ? item.RecurrenceTimeZoneDefinition :
                    OutlookTimeZoneBinary.EncodeDefinition(item.RecurrenceTimeZone));
        }
        if (document.MeetingCommunication != null) {
            OutlookMeetingCommunication item = document.MeetingCommunication;
            properties.Set(MapiKnownProperties.PidLid.MeetingType, item.RequestTypeValue);
            properties.Set(MapiKnownProperties.PidLid.IntendedBusyStatus, item.IntendedBusyStatus);
            properties.Set(MapiKnownProperties.PidLid.OwnerCriticalChange, item.OwnerCriticalChange);
            properties.Set(MapiKnownProperties.PidLid.AttendeeCriticalChange, item.AttendeeCriticalChange);
            properties.Set(MapiKnownProperties.PidLid.IsSilent, item.IsSilent);
            properties.Set(MapiKnownProperties.PidLid.AppointmentCounterProposal, item.IsCounterProposal);
            properties.Set(MapiKnownProperties.PidLid.AppointmentProposedStartWhole, item.ProposedStart);
            properties.Set(MapiKnownProperties.PidLid.AppointmentProposedEndWhole, item.ProposedEnd);
            properties.Set(MapiKnownProperties.PidLid.AppointmentProposedDuration, item.ProposedDurationMinutes);
            properties.Set(MapiKnownProperties.PidLid.AppointmentReplyTime, item.ReplyAt);
            properties.Set(MapiKnownProperties.PidLid.AppointmentReplyName, item.ReplyName);
        }
        if (document.Contact != null) {
            OutlookContact item = document.Contact;
            properties.Set(MapiKnownProperties.PidTag.DisplayName, item.DisplayName);
            properties.Set(MapiKnownProperties.PidTag.DisplayNamePrefix, item.Prefix);
            properties.Set(MapiKnownProperties.PidTag.Initials, item.Initials);
            properties.Set(MapiKnownProperties.PidTag.GivenName, item.GivenName);
            properties.Set(MapiKnownProperties.PidTag.MiddleName, item.MiddleName);
            properties.Set(MapiKnownProperties.PidTag.Surname, item.Surname);
            properties.Set(MapiKnownProperties.PidTag.Generation, item.Generation);
            properties.Set(MapiKnownProperties.PidTag.CompanyName, item.CompanyName);
            properties.Set(MapiKnownProperties.PidTag.Title, item.JobTitle);
            properties.Set(MapiKnownProperties.PidTag.DepartmentName, item.Department);
            properties.Set(MapiKnownProperties.PidTag.Nickname, item.NickName);
            properties.Set(MapiKnownProperties.PidTag.ManagerName, item.ManagerName);
            properties.Set(MapiKnownProperties.PidTag.Assistant, item.AssistantName);
            properties.Set(MapiKnownProperties.PidTag.SpouseName, item.SpouseName);
            properties.Set(MapiKnownProperties.PidTag.ChildrensNames, MapiPropertyType.Unicode,
                item.Children.Count == 0 ? null : string.Join(", ", item.Children));
            properties.Set(MapiKnownProperties.PidTag.Profession, item.Profession);
            properties.Set(MapiKnownProperties.PidTag.Language, item.Language);
            properties.Set(MapiKnownProperties.PidTag.MailUserLocation, item.Location);
            properties.Set(MapiKnownProperties.PidTag.OfficeLocation, item.OfficeLocation);
            properties.Set(MapiKnownProperties.PidTag.Birthday, item.Birthday);
            properties.Set(MapiKnownProperties.PidTag.WeddingAnniversary, item.WeddingAnniversary);
            properties.Set(MapiKnownProperties.PidTag.BusinessHomePage, item.BusinessHomePage);
            properties.Set(MapiKnownProperties.PidTag.PersonalHomePage, item.PersonalHomePage);
            properties.Set(MapiKnownProperties.PidLid.FileUnder, item.FileAs);
            properties.Set(MapiKnownProperties.PidLid.InstantMessagingAddress, item.InstantMessagingAddress);
            properties.Set(MapiKnownProperties.PidLid.BirthdayLocal, item.Birthday);
            properties.Set(MapiKnownProperties.PidLid.WeddingAnniversaryLocal, item.WeddingAnniversary);
            properties.Set(MapiKnownProperties.PidLid.Private, item.IsPrivate);
            properties.Set(MapiKnownProperties.PidLid.HasPicture,
                item.HasPicture ?? document.Attachments.Any(attachment => attachment.IsContactPhoto));
            properties.Set(MapiKnownProperties.PidLid.ContactHtml, item.Html);
            AddContactAddressProperties(properties, item);
            AddContactPhoneProperties(properties, item.Phones);
            AddContactEmailProperties(properties, item.Email1, MapiKnownProperties.PidLid.Email1DisplayName,
                MapiKnownProperties.PidLid.Email1AddressType, MapiKnownProperties.PidLid.Email1EmailAddress,
                MapiKnownProperties.PidLid.Email1OriginalDisplayName, MapiKnownProperties.PidLid.Email1OriginalEntryId);
            AddContactEmailProperties(properties, item.Email2, MapiKnownProperties.PidLid.Email2DisplayName,
                MapiKnownProperties.PidLid.Email2AddressType, MapiKnownProperties.PidLid.Email2EmailAddress,
                MapiKnownProperties.PidLid.Email2OriginalDisplayName, MapiKnownProperties.PidLid.Email2OriginalEntryId);
            AddContactEmailProperties(properties, item.Email3, MapiKnownProperties.PidLid.Email3DisplayName,
                MapiKnownProperties.PidLid.Email3AddressType, MapiKnownProperties.PidLid.Email3EmailAddress,
                MapiKnownProperties.PidLid.Email3OriginalDisplayName, MapiKnownProperties.PidLid.Email3OriginalEntryId);
        }
        if (document.DistributionList != null) {
            document.DistributionList.WriteTo(properties);
        }
        if (document.Task != null) {
            OutlookTask item = document.Task;
            properties.Set(MapiKnownProperties.PidLid.TaskStartDate, item.Start);
            properties.Set(MapiKnownProperties.PidLid.TaskDueDate, item.Due);
            properties.Set(MapiKnownProperties.PidLid.TaskStatus, item.Status);
            properties.Set(MapiKnownProperties.PidLid.PercentComplete, item.PercentComplete);
            properties.Set(MapiKnownProperties.PidLid.TaskComplete, item.IsComplete);
            properties.Set(MapiKnownProperties.PidLid.TaskOwner, item.Owner);
            properties.Set(MapiKnownProperties.PidLid.TaskActualEffort, ToMinutes(item.ActualEffort));
            properties.Set(MapiKnownProperties.PidLid.TaskEstimatedEffort, ToMinutes(item.EstimatedEffort));
            properties.Set(MapiKnownProperties.PidLid.TaskUpdates, item.SendUpdates);
            properties.Set(MapiKnownProperties.PidLid.TaskStatusOnComplete, item.SendStatusOnComplete);
            properties.Set(MapiKnownProperties.PidLid.TaskOwnership, item.Ownership);
            properties.Set(MapiKnownProperties.PidLid.TaskAcceptanceState, item.AcceptanceState);
            properties.Set(MapiKnownProperties.PidLid.TaskVersion, item.Version);
            properties.Set(MapiKnownProperties.PidLid.TaskState, item.State);
            properties.Set(MapiKnownProperties.PidLid.TaskAssigner, item.Assigner);
            properties.Set(MapiKnownProperties.PidLid.TeamTask, item.IsTeamTask);
            properties.Set(MapiKnownProperties.PidLid.TaskOrdinal, item.Ordinal);
            byte[]? taskRecurrence = item.Recurrence == null
                ? item.RecurrenceState
                : OutlookRecurrenceBinary.EncodeTask(item.Recurrence);
            properties.Set(MapiKnownProperties.PidLid.TaskRecurrence, taskRecurrence);
            properties.Set(MapiKnownProperties.PidLid.TaskFRecurring,
                item.IsRecurring ?? taskRecurrence != null);
            OutlookMessageSemanticsWriter.WriteReminder(properties, item.Reminder);
            properties.Set(MapiKnownProperties.PidLid.CommonStart, item.CommonStart);
            properties.Set(MapiKnownProperties.PidLid.CommonEnd, item.CommonEnd);
            properties.Set(MapiKnownProperties.PidLid.TaskMode, item.Mode);
            properties.Set(MapiKnownProperties.PidLid.TaskAccepted, item.IsAccepted);
            properties.Set(MapiKnownProperties.PidLid.TaskHistory, item.History);
            properties.Set(MapiKnownProperties.PidLid.TaskLastUpdate, item.LastUpdate);
            properties.Set(MapiKnownProperties.PidLid.TaskLastUser, item.LastUser);
            properties.Set(MapiKnownProperties.PidLid.TaskLastDelegate, item.LastDelegate);
            properties.Set(MapiKnownProperties.PidLid.TaskGlobalId,
                item.GlobalId.HasValue ? item.GlobalId.Value.ToByteArray() : null);
            properties.Set(MapiKnownProperties.PidLid.ToDoOrdinalDate, item.ToDoOrdinalDate);
            properties.Set(MapiKnownProperties.PidLid.ToDoSubOrdinal, item.ToDoSubOrdinal);
            properties.Set(MapiKnownProperties.PidLid.Contacts, ToObjectArray(item.Contacts));
            properties.Set(MapiKnownProperties.PidLid.Companies, ToObjectArray(item.Companies));
            properties.Set(MapiKnownProperties.PidLid.Billing, item.BillingInformation);
            properties.Set(MapiKnownProperties.PidLid.Mileage, item.Mileage);
            properties.Set(MapiKnownProperties.PidLid.TaskDateCompleted, item.CompletedAt);
        }
        if (document.Journal != null) {
            OutlookJournal item = document.Journal;
            properties.Set(MapiKnownProperties.PidLid.CommonStart, item.Start);
            properties.Set(MapiKnownProperties.PidLid.CommonEnd, item.End);
            properties.Set(MapiKnownProperties.PidLid.LogType, item.Type);
            properties.Set(MapiKnownProperties.PidLid.LogStart, item.Start);
            properties.Set(MapiKnownProperties.PidLid.LogDuration, item.DurationMinutes);
            properties.Set(MapiKnownProperties.PidLid.LogEnd, item.End);
            properties.Set(MapiKnownProperties.PidLid.LogFlags, item.Flags);
            properties.Set(MapiKnownProperties.PidLid.LogDocumentPrinted, item.DocumentPrinted);
            properties.Set(MapiKnownProperties.PidLid.LogDocumentSaved, item.DocumentSaved);
            properties.Set(MapiKnownProperties.PidLid.LogDocumentRouted, item.DocumentRouted);
            properties.Set(MapiKnownProperties.PidLid.LogDocumentPosted, item.DocumentPosted);
            properties.Set(MapiKnownProperties.PidLid.LogTypeDesc, item.TypeDescription);
        }
        if (document.Note != null) {
            OutlookNote item = document.Note;
            properties.Set(MapiKnownProperties.PidLid.NoteColor, item.Color);
            properties.Set(MapiKnownProperties.PidLid.NoteWidth, item.Width);
            properties.Set(MapiKnownProperties.PidLid.NoteHeight, item.Height);
            properties.Set(MapiKnownProperties.PidLid.NoteX, item.X);
            properties.Set(MapiKnownProperties.PidLid.NoteY, item.Y);
        }
    }

    private static void AddContactAddressProperties(MsgPropertyBuilder properties, OutlookContact contact) {
        AddFixedAddress(properties, contact.BusinessAddress, MapiKnownProperties.PidTag.StreetAddress,
            MapiKnownProperties.PidTag.Locality, MapiKnownProperties.PidTag.StateOrProvince,
            MapiKnownProperties.PidTag.PostalCode, MapiKnownProperties.PidTag.Country,
            MapiKnownProperties.PidTag.PostOfficeBox);
        AddFixedAddress(properties, contact.HomeAddress, MapiKnownProperties.PidTag.HomeAddressStreet,
            MapiKnownProperties.PidTag.HomeAddressCity, MapiKnownProperties.PidTag.HomeAddressStateOrProvince,
            MapiKnownProperties.PidTag.HomeAddressPostalCode, MapiKnownProperties.PidTag.HomeAddressCountry,
            MapiKnownProperties.PidTag.HomeAddressPostOfficeBox);
        AddFixedAddress(properties, contact.OtherAddress, MapiKnownProperties.PidTag.OtherAddressStreet,
            MapiKnownProperties.PidTag.OtherAddressCity, MapiKnownProperties.PidTag.OtherAddressStateOrProvince,
            MapiKnownProperties.PidTag.OtherAddressPostalCode, MapiKnownProperties.PidTag.OtherAddressCountry,
            MapiKnownProperties.PidTag.OtherAddressPostOfficeBox);
        properties.Set(MapiKnownProperties.PidTag.PostalAddress, contact.BusinessAddress.Formatted);
        properties.Set(MapiKnownProperties.PidLid.HomeAddress, contact.HomeAddress.Formatted);
        properties.Set(MapiKnownProperties.PidLid.WorkAddress,
            contact.WorkAddress.Formatted ?? contact.BusinessAddress.Formatted);
        properties.Set(MapiKnownProperties.PidLid.OtherAddress, contact.OtherAddress.Formatted);
        properties.Set(MapiKnownProperties.PidLid.WorkAddressStreet,
            contact.WorkAddress.Street ?? contact.BusinessAddress.Street);
        properties.Set(MapiKnownProperties.PidLid.WorkAddressCity,
            contact.WorkAddress.City ?? contact.BusinessAddress.City);
        properties.Set(MapiKnownProperties.PidLid.WorkAddressState,
            contact.WorkAddress.StateOrProvince ?? contact.BusinessAddress.StateOrProvince);
        properties.Set(MapiKnownProperties.PidLid.WorkAddressPostalCode,
            contact.WorkAddress.PostalCode ?? contact.BusinessAddress.PostalCode);
        properties.Set(MapiKnownProperties.PidLid.WorkAddressCountry,
            contact.WorkAddress.Country ?? contact.BusinessAddress.Country);
        properties.Set(MapiKnownProperties.PidLid.WorkAddressPostOfficeBox,
            contact.WorkAddress.PostOfficeBox ?? contact.BusinessAddress.PostOfficeBox);
        properties.Set(MapiKnownProperties.PidLid.WorkAddressCountryCode, contact.WorkAddress.CountryCode);
    }

    private static void AddFixedAddress(MsgPropertyBuilder properties, OutlookPostalAddress address,
        MapiPropertyKey<string> streetKey, MapiPropertyKey<string> cityKey, MapiPropertyKey<string> stateKey,
        MapiPropertyKey<string> postalKey, MapiPropertyKey<string> countryKey,
        MapiPropertyKey<string> postOfficeBoxKey) {
        properties.Set(streetKey, address.Street);
        properties.Set(cityKey, address.City);
        properties.Set(stateKey, address.StateOrProvince);
        properties.Set(postalKey, address.PostalCode);
        properties.Set(countryKey, address.Country);
        properties.Set(postOfficeBoxKey, address.PostOfficeBox);
    }

    private static void AddContactPhoneProperties(MsgPropertyBuilder properties, OutlookContactPhones phones) {
        properties.Set(MapiKnownProperties.PidTag.BusinessTelephoneNumber, phones.Business);
        properties.Set(MapiKnownProperties.PidTag.Business2TelephoneNumber, phones.Business2);
        properties.Set(MapiKnownProperties.PidTag.HomeTelephoneNumber, phones.Home);
        properties.Set(MapiKnownProperties.PidTag.Home2TelephoneNumber, phones.Home2);
        properties.Set(MapiKnownProperties.PidTag.MobileTelephoneNumber, phones.Mobile);
        properties.Set(MapiKnownProperties.PidTag.OtherTelephoneNumber, phones.Other);
        properties.Set(MapiKnownProperties.PidTag.PrimaryTelephoneNumber, phones.Primary);
        properties.Set(MapiKnownProperties.PidTag.BusinessFaxNumber, phones.BusinessFax);
        properties.Set(MapiKnownProperties.PidTag.HomeFaxNumber, phones.HomeFax);
        properties.Set(MapiKnownProperties.PidTag.PrimaryFaxNumber, phones.PrimaryFax);
        properties.Set(MapiKnownProperties.PidTag.AssistantTelephoneNumber, phones.Assistant);
        properties.Set(MapiKnownProperties.PidTag.CompanyMainPhoneNumber, phones.CompanyMain);
        properties.Set(MapiKnownProperties.PidTag.CarTelephoneNumber, phones.Car);
        properties.Set(MapiKnownProperties.PidTag.RadioTelephoneNumber, phones.Radio);
        properties.Set(MapiKnownProperties.PidTag.PagerTelephoneNumber, phones.Pager);
        properties.Set(MapiKnownProperties.PidTag.CallbackTelephoneNumber, phones.Callback);
        properties.Set(MapiKnownProperties.PidTag.TelexNumber, phones.Telex);
        properties.Set(MapiKnownProperties.PidTag.TtyTddPhoneNumber, phones.TextTelephone);
        properties.Set(MapiKnownProperties.PidTag.IsdnNumber, phones.Isdn);
    }

    private static void AddContactEmailProperties(MsgPropertyBuilder properties, OutlookContactEmailAddress email,
        MapiPropertyKey<string> displayKey, MapiPropertyKey<string> addressTypeKey,
        MapiPropertyKey<string> addressKey, MapiPropertyKey<string> originalDisplayKey,
        MapiPropertyKey<byte[]> entryKey) {
        properties.Set(displayKey, email.DisplayName);
        properties.Set(addressTypeKey,
            email.AddressType ?? (email.Address == null ? null : "SMTP"));
        properties.Set(addressKey, email.Address);
        properties.Set(originalDisplayKey,
            email.OriginalDisplayName ?? email.Address);
        byte[]? originalEntryId = email.OriginalEntryId;
        if (originalEntryId == null && email.Address != null) {
            originalEntryId = MsgIdentity.CreateOneOffEntryId(new EmailAddress(email.Address, email.DisplayName) {
                AddressType = email.AddressType ?? "SMTP"
            });
        }
        properties.Set(entryKey, originalEntryId);
    }

    private static int? GetDurationMinutes(DateTimeOffset? start, DateTimeOffset? end) {
        if (!start.HasValue || !end.HasValue) return null;
        double minutes = Math.Round((end.Value - start.Value).TotalMinutes);
        return minutes >= int.MinValue && minutes <= int.MaxValue ? (int)minutes : (int?)null;
    }

    private static string? JoinAppointmentAttendees(EmailDocument document, EmailRecipientKind? kind) {
        string[] attendees = document.Recipients
            .Where(recipient => recipient.Kind != EmailRecipientKind.ReplyTo &&
                (!kind.HasValue || recipient.Kind == kind.Value))
            .Select(recipient => recipient.Address.DisplayName ?? recipient.Address.Address)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Cast<string>()
            .ToArray();
        return attendees.Length == 0 ? null : string.Join("; ", attendees);
    }

    private static int? ToMinutes(TimeSpan? value) {
        if (!value.HasValue) return null;
        double minutes = Math.Round(value.Value.TotalMinutes);
        return minutes >= int.MinValue && minutes <= int.MaxValue ? (int)minutes : (int?)null;
    }

    private static object[]? ToObjectArray(IList<string> values) =>
        values.Count == 0 ? null : values.Cast<object>().ToArray();
}
