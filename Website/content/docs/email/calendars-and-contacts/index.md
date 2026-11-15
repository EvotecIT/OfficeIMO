---
title: "Read and Edit Calendars, Contacts, and Outlook Items"
description: "Read and update iCalendar (ICS), vCalendar, and vCard files, and expand Outlook appointment recurrence with OfficeIMO.Email."
meta.seo_title: "Read ICS, vCard, and Outlook items in .NET"
order: 39
packages:
  - OfficeIMO.Email
limits: "Recurrence expansion is bounded by a date window and a maximum occurrence count, and stays in local-clock time until you supply a time zone."
search_terms:
  - ics
  - icalendar
  - vcard
  - vcf
  - outlook recurrence
verified: 2026-11-14
---

`IcsDocument` and `VCardDocument` read standalone iCalendar, vCalendar, and vCard 2.1, 3.0, and 4.0 files into ordered content-line components. Typed helpers update common properties, while repeated, grouped, vendor, and `X-` properties stay in place. Outlook appointments and tasks inside MSG, TNEF, and PST use `OutlookRecurrenceExpander` for their recurrence.

## Update an iCalendar event

### C#

```csharp
using OfficeIMO.Email;

IcsDocument calendar = IcsDocument.Load("meeting.ics");
ContentLineComponent meeting = calendar.GetComponents("VEVENT").Single();

meeting.SetProperty("SUMMARY", "Updated planning meeting");
meeting.SetTemporalValue(
    "DTSTART",
    IcsTemporalValue.Zoned(new DateTime(2026, 7, 20, 9, 0, 0), "Europe/Warsaw"));

IReadOnlyList<ContentLineValidationIssue> issues = calendar.Validate();
Console.WriteLine($"{issues.Count} validation issues");

calendar.Save("updated.ics");
```

Temporal helpers keep date-only, floating, UTC, and `TZID`-local forms. RRULE helpers change known recurrence fields and keep unknown ones.

## Update a vCard contact

### C#

```csharp
using OfficeIMO.Email;

VCardDocument contacts = VCardDocument.Load("contacts.vcf");
ContentLineComponent contact = contacts.Cards[0];

contact.SetVCardText("FN", "Ada Lovelace");
contact.AddProperty("EMAIL", "ada@example.test").SetParameter("TYPE", "work");

IReadOnlyList<ContentLineValidationIssue> issues = contacts.Validate();
Console.WriteLine($"{issues.Count} validation issues");

contacts.Save("updated.vcf");
```

## Expand Outlook recurrence

Modified and deleted occurrences stay explicit in the result. The same recurrence model is used for MSG, TNEF, and PST items.

### C#

```csharp
using OfficeIMO.Email;

EmailDocument item = EmailDocument.Load("recurring-meeting.msg");

if (item.Appointment?.Recurrence is { } recurrence) {
    OutlookRecurrenceExpansionResult occurrences = OutlookRecurrenceExpander.Expand(
        recurrence,
        new OutlookRecurrenceExpansionOptions {
            WindowStart = new DateTime(2026, 7, 1),
            WindowEnd = new DateTime(2026, 8, 1),
            MaxOccurrences = 500
        });

    foreach (OutlookRecurrenceOccurrence occurrence in occurrences.Occurrences) {
        string note = occurrence.IsException ? " (modified)" : "";
        Console.WriteLine($"{occurrence.Start:g} - {occurrence.End:t}{note}");
    }
}
```

## How items convert to EML

Appointments and tasks written as EML become `text/calendar` parts. Contacts become vCard attachments, reminders become `VALARM`, and task fields without a standard property use `X-OFFICEIMO-*` extensions. Journals and sticky notes have no standard EML form, so writing them as EML needs an explicit conversion-loss policy; see [Read and Convert Email Messages](/docs/email/messages-and-conversion/).
