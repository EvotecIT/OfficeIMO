namespace OfficeIMO.Email;

internal static class OutlookMessageSemanticsProjection {
    internal static void Apply(EmailDocument document, MapiPropertyBag properties, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        EmailMessageMetadata metadata = document.MessageMetadata;
        ApplyFollowUp(metadata.FollowUp, properties);
        ApplyVoting(metadata.Voting, properties, codePage, diagnostics, location);
        if (document.OutlookItemKind != OutlookItemKind.Appointment &&
            document.OutlookItemKind != OutlookItemKind.Task) {
            ApplyReminder(metadata.Reminder, properties);
        }
    }

    internal static void ApplyReminder(OutlookReminder reminder, MapiPropertyBag properties) {
        reminder.IsSet = properties.GetNullableValue(MapiKnownProperties.PidLid.ReminderSet);
        reminder.DeltaMinutes = properties.GetNullableValue(MapiKnownProperties.PidLid.ReminderDelta);
        reminder.Time = properties.GetNullableValue(MapiKnownProperties.PidLid.ReminderTime);
        reminder.SignalTime = properties.GetNullableValue(MapiKnownProperties.PidLid.ReminderSignalTime);
        reminder.Override = properties.GetNullableValue(MapiKnownProperties.PidLid.ReminderOverride);
        reminder.PlaySound = properties.GetNullableValue(MapiKnownProperties.PidLid.ReminderPlaySound);
        reminder.SoundFile = properties.GetValueOrDefault(MapiKnownProperties.PidLid.ReminderFileParameter);
    }

    private static void ApplyFollowUp(OutlookFollowUp followUp, MapiPropertyBag properties) {
        followUp.RawStatus = properties.GetNullableValue(MapiKnownProperties.PidTag.FlagStatus);
        if (followUp.RawStatus.HasValue && Enum.IsDefined(typeof(OutlookFollowUpStatus), followUp.RawStatus.Value)) {
            followUp.Status = (OutlookFollowUpStatus)followUp.RawStatus.Value;
        }
        followUp.Request = properties.GetValueOrDefault(MapiKnownProperties.PidLid.FlagRequest);
        followUp.Title = properties.GetValueOrDefault(MapiKnownProperties.PidLid.ToDoTitle);
        followUp.Start = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskStartDate);
        followUp.Due = properties.GetNullableValue(MapiKnownProperties.PidLid.TaskDueDate);
        followUp.CompletedAt = properties.GetNullableValue(MapiKnownProperties.PidTag.FlagCompleteTime);
        int? icon = properties.GetNullableValue(MapiKnownProperties.PidTag.FollowupIcon);
        if (icon.HasValue && Enum.IsDefined(typeof(OutlookFollowUpIcon), icon.Value)) {
            followUp.Icon = (OutlookFollowUpIcon)icon.Value;
        }
        followUp.FlagString = properties.GetNullableValue(MapiKnownProperties.PidLid.FlagString);
        followUp.ValidRequestProof = properties.GetNullableValue(MapiKnownProperties.PidLid.ValidFlagStringProof);
        followUp.ToDoItemFlags = properties.GetNullableValue(MapiKnownProperties.PidTag.ToDoItemFlags);
    }

    private static void ApplyVoting(OutlookVoting voting, MapiPropertyBag properties, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        voting.Response = properties.GetValueOrDefault(MapiKnownProperties.PidLid.VerbResponse);
        voting.RawVerbStream = properties.GetValueOrDefault(MapiKnownProperties.PidLid.VerbStream);
        voting.OptionsDecoded = false;
        voting.ResetProjectionState();
        voting.Options.Clear();
        if (voting.RawVerbStream == null) return;
        if (OutlookVotingCodec.TryDecode(voting.RawVerbStream, codePage, voting.Options, out string? error)) {
            voting.OptionsDecoded = true;
            return;
        }
        diagnostics.Add(new EmailDiagnostic(
            "EMAIL_MSG_VOTING_VERB_STREAM_INVALID",
            string.Concat("The Outlook voting option stream was retained but could not be decoded: ", error),
            EmailDiagnosticSeverity.Warning,
            string.Concat(location, "/voting")));
    }
}
