---
title: "Calendars, contacts, and Outlook items"
description: "Read and update iCalendar, vCalendar, vCard, Outlook appointments, contacts, tasks, recurrence, reminders, and meeting lifecycle data in .NET."
meta.seo_title: "Read ICS, vCard, and Outlook items in .NET"
order: 39
---

OfficeIMO.Email uses typed Outlook models and ordered, loss-preserving content-line models for calendar and contact data. It supports standalone ICS, legacy vCalendar, and vCard 2.1, 3.0, and 4.0 documents as well as their projections inside MSG, TNEF, and EML.

## Update an iCalendar event

```csharp
using OfficeIMO.Email;

IcsDocument calendar = IcsDocument.Load("meeting.ics");
ContentLineComponent meeting =
    calendar.GetComponents("VEVENT").Single();

meeting.SetProperty("SUMMARY", "Updated planning meeting");
meeting.SetTemporalValue(
    "DTSTART",
    IcsTemporalValue.Zoned(
        new DateTime(2026, 7, 20, 9, 0, 0),
        "Europe/Warsaw"));

IReadOnlyList<ContentLineValidationIssue> issues =
    calendar.Validate();

calendar.Save("updated.ics");
```

Temporal helpers retain date-only, floating, UTC, and `TZID`-local forms. RRULE helpers update known recurrence fields while preserving unknown fields.

## Update vCard contacts

```csharp
VCardDocument contacts = VCardDocument.Load("contacts.vcf");
ContentLineComponent contact = contacts.Cards[0];

contact.SetVCardText("FN", "Ada Lovelace");
contact.AddProperty("EMAIL", "ada@example.test")
    .SetParameter("TYPE", "work");

IReadOnlyList<ContentLineValidationIssue> issues =
    contacts.Validate();

contacts.Save("updated.vcf");
```

Repeated, grouped, vendor, IANA, and `X-` properties remain available instead of being discarded by the typed conveniences.

## Expand Outlook recurrence

Decoded appointment and task recurrence uses the same model across MSG, TNEF, and PST. Expansion is bounded and remains in local-clock space until the caller supplies a time zone:

```csharp
EmailDocument item = EmailDocument.Load("recurring-meeting.msg");

OutlookRecurrenceExpansionResult occurrences =
    OutlookRecurrenceExpander.Expand(
        item.Appointment!.Recurrence!,
        new OutlookRecurrenceExpansionOptions {
            WindowStart = new DateTime(2026, 7, 1),
            WindowEnd = new DateTime(2026, 8, 1),
            MaxOccurrences = 500
        });
```

Modified and deleted occurrences remain explicit. ICS import and export return a report when a recurrence pattern, exception, or time-zone constraint cannot be represented.

## Conversion behavior

Appointments and tasks written as EML become `text/calendar` parts. Contacts become vCard attachments, reminders become `VALARM`, and OfficeIMO-specific task fields use valid `X-OFFICEIMO-*` extensions. Journal and sticky-note semantics do not have a standard EML representation, so EML output requires an explicit conversion-loss decision.

Use [messages and conversion](/docs/email/messages-and-conversion/) to inspect conversion diagnostics before writing.
