using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookMessageSemanticsTests {
    [Fact]
    public void RoundTripsFollowUpReminderAndVotingThroughMsg() {
        DateTimeOffset start = new DateTimeOffset(2026, 10, 1, 8, 0, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Message,
            MessageClass = "IPM.Note",
            Subject = "Decision",
            OutlookCodePage = 1252
        };
        source.MessageMetadata.FollowUp.SetFlagged("Call", start, start.AddDays(2), OutlookFollowUpIcon.Red);
        source.MessageMetadata.FollowUp.Title = "Decision follow-up";
        source.MessageMetadata.FollowUp.FlagString = 1;
        source.MessageMetadata.FollowUp.ValidRequestProof = start.AddHours(-1);
        source.MessageMetadata.FollowUp.ToDoItemFlags = 1;
        source.MessageMetadata.Reminder.IsSet = true;
        source.MessageMetadata.Reminder.DeltaMinutes = 30;
        source.MessageMetadata.Reminder.Time = start;
        source.MessageMetadata.Reminder.SignalTime = start.AddMinutes(-30);
        source.MessageMetadata.Reminder.Override = true;
        source.MessageMetadata.Reminder.PlaySound = true;
        source.MessageMetadata.Reminder.SoundFile = "notify.wav";
        source.MessageMetadata.Voting.Options.Add(new OutlookVoteOption("Yes") {
            Id = 1,
            SendBehavior = OutlookVoteSendBehavior.Automatic,
            UseUsReplyHeaders = true
        });
        source.MessageMetadata.Voting.Options.Add(new OutlookVoteOption("No") {
            Id = 2,
            SendBehavior = OutlookVoteSendBehavior.Prompt
        });
        source.MessageMetadata.Voting.Response = "Yes";

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult read = new EmailDocumentReader().Read(bytes);
        EmailMessageMetadata result = read.Document.MessageMetadata;

        Assert.Equal(OutlookFollowUpStatus.Flagged, result.FollowUp.Status);
        Assert.Equal("Call", result.FollowUp.Request);
        Assert.Equal("Decision follow-up", result.FollowUp.Title);
        Assert.Equal(start, result.FollowUp.Start);
        Assert.Equal(start.AddDays(2), result.FollowUp.Due);
        Assert.Equal(OutlookFollowUpIcon.Red, result.FollowUp.Icon);
        Assert.Equal(1, result.FollowUp.FlagString);
        Assert.Equal(start.AddHours(-1), result.FollowUp.ValidRequestProof);
        Assert.True(result.Reminder.IsSet);
        Assert.Equal(30, result.Reminder.DeltaMinutes);
        Assert.True(result.Reminder.Override);
        Assert.True(result.Reminder.PlaySound);
        Assert.Equal("notify.wav", result.Reminder.SoundFile);
        Assert.True(result.Voting.OptionsDecoded);
        Assert.Equal("Yes", result.Voting.Response);
        Assert.Collection(result.Voting.Options,
            option => {
                Assert.Equal("Yes", option.DisplayName);
                Assert.Equal(1, option.Id);
                Assert.Equal(OutlookVoteSendBehavior.Automatic, option.SendBehavior);
                Assert.True(option.UseUsReplyHeaders);
            },
            option => {
                Assert.Equal("No", option.DisplayName);
                Assert.Equal(2, option.Id);
                Assert.Equal(OutlookVoteSendBehavior.Prompt, option.SendBehavior);
                Assert.False(option.UseUsReplyHeaders);
            });
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void CompletedFollowUpWritesConsistentCompletionProperties() {
        DateTimeOffset completed = new DateTimeOffset(2026, 10, 2, 9, 30, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            MessageClass = "IPM.Note",
            Subject = "Completed"
        };
        source.MessageMetadata.FollowUp.MarkComplete(completed);

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal(OutlookFollowUpStatus.Complete, result.MessageMetadata.FollowUp.Status);
        Assert.Equal(completed, result.MessageMetadata.FollowUp.CompletedAt);
        Assert.Null(result.MessageMetadata.FollowUp.Icon);
        Assert.Equal(completed, result.Mapi.GetNullableValue(MapiKnownProperties.PidLid.TaskDateCompleted));
        Assert.True(result.Mapi.GetNullableValue(MapiKnownProperties.PidLid.TaskComplete));
        Assert.Equal(2, result.Mapi.GetNullableValue(MapiKnownProperties.PidLid.TaskStatus));
        Assert.Equal(1d, result.Mapi.GetNullableValue(MapiKnownProperties.PidLid.PercentComplete));
    }

    [Fact]
    public void InvalidVotingStreamIsRetainedAndReported() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            MessageClass = "IPM.Note",
            Subject = "Malformed voting"
        };
        source.MessageMetadata.Voting.RawVerbStream = new byte[] { 0x02, 0x01, 0x01 };

        EmailReadResult read = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg));

        Assert.Equal(new byte[] { 0x02, 0x01, 0x01 }, read.Document.MessageMetadata.Voting.RawVerbStream);
        Assert.False(read.Document.MessageMetadata.Voting.OptionsDecoded);
        Assert.Empty(read.Document.MessageMetadata.Voting.Options);
        Assert.Contains(read.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MSG_VOTING_VERB_STREAM_INVALID" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
    }

    [Fact]
    public void TaskCompletionUsesTaskDateCompletedNotMessageFlagCompletion() {
        DateTimeOffset completed = new DateTimeOffset(2026, 10, 3, 0, 0, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            OutlookItemKind = OutlookItemKind.Task,
            MessageClass = "IPM.Task",
            Subject = "Task",
            Task = new OutlookTask { CompletedAt = completed, Status = 2, IsComplete = true, PercentComplete = 1 }
        };

        EmailDocument result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg)).Document;

        Assert.Equal(completed, result.Task!.CompletedAt);
        Assert.Equal(completed, result.Mapi.GetNullableValue(MapiKnownProperties.PidLid.TaskDateCompleted));
        Assert.Null(result.Mapi.Find(MapiKnownProperties.PidTag.FlagCompleteTime));
    }
}
