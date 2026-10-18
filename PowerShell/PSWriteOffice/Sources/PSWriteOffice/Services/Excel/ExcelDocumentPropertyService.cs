using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Models.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelDocumentPropertyService
{
    private static readonly IReadOnlyDictionary<string, Func<ExcelBuiltinDocumentProperties, object?>> BuiltInReaders =
        new Dictionary<string, Func<ExcelBuiltinDocumentProperties, object?>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = properties => properties.Title,
            ["Subject"] = properties => properties.Subject,
            ["Creator"] = properties => properties.Creator,
            ["Author"] = properties => properties.Creator,
            ["Keywords"] = properties => properties.Keywords,
            ["Description"] = properties => properties.Description,
            ["Category"] = properties => properties.Category,
            ["Revision"] = properties => properties.Revision,
            ["LastModifiedBy"] = properties => properties.LastModifiedBy,
            ["Version"] = properties => properties.Version,
            ["Created"] = properties => properties.Created,
            ["Modified"] = properties => properties.Modified,
            ["LastPrinted"] = properties => properties.LastPrinted
        };

    private static readonly IReadOnlyDictionary<string, Action<ExcelBuiltinDocumentProperties, object?>> BuiltInWriters =
        new Dictionary<string, Action<ExcelBuiltinDocumentProperties, object?>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = (properties, value) => properties.Title = ConvertToString(value),
            ["Subject"] = (properties, value) => properties.Subject = ConvertToString(value),
            ["Creator"] = (properties, value) => properties.Creator = ConvertToString(value),
            ["Author"] = (properties, value) => properties.Creator = ConvertToString(value),
            ["Keywords"] = (properties, value) => properties.Keywords = ConvertToString(value),
            ["Description"] = (properties, value) => properties.Description = ConvertToString(value),
            ["Category"] = (properties, value) => properties.Category = ConvertToString(value),
            ["Revision"] = (properties, value) => properties.Revision = ConvertToString(value),
            ["LastModifiedBy"] = (properties, value) => properties.LastModifiedBy = ConvertToString(value),
            ["Version"] = (properties, value) => properties.Version = ConvertToString(value),
            ["Created"] = (properties, value) => properties.Created = ConvertToDateTime(value),
            ["Modified"] = (properties, value) => properties.Modified = ConvertToDateTime(value),
            ["LastPrinted"] = (properties, value) => properties.LastPrinted = ConvertToDateTime(value)
        };

    private static readonly IReadOnlyDictionary<string, Func<ExcelApplicationProperties, object?>> ApplicationReaders =
        new Dictionary<string, Func<ExcelApplicationProperties, object?>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Company"] = properties => properties.Company,
            ["Manager"] = properties => properties.Manager,
            ["ApplicationName"] = properties => properties.ApplicationName,
            ["Application"] = properties => properties.ApplicationName
        };

    private static readonly IReadOnlyDictionary<string, Action<ExcelApplicationProperties, object?>> ApplicationWriters =
        new Dictionary<string, Action<ExcelApplicationProperties, object?>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Company"] = (properties, value) => properties.Company = ConvertToString(value) ?? string.Empty,
            ["Manager"] = (properties, value) => properties.Manager = ConvertToString(value) ?? string.Empty,
            ["ApplicationName"] = (properties, value) => properties.ApplicationName = ConvertToString(value) ?? string.Empty,
            ["Application"] = (properties, value) => properties.ApplicationName = ConvertToString(value) ?? string.Empty
        };

    public static IEnumerable<ExcelDocumentPropertyInfo> GetProperties(ExcelDocument document, bool includeBuiltIn, bool includeApplication, bool includeCustom)
    {
        if (includeBuiltIn)
        {
            foreach (var property in BuiltInReaders)
            {
                var value = property.Value(document.BuiltinDocumentProperties);
                yield return new ExcelDocumentPropertyInfo(property.Key, "BuiltIn", value, value?.GetType().FullName);
            }
        }

        if (includeApplication)
        {
            foreach (var property in ApplicationReaders)
            {
                var value = property.Value(document.ApplicationProperties);
                yield return new ExcelDocumentPropertyInfo(property.Key, "Application", value, value?.GetType().FullName);
            }
        }

        if (includeCustom)
        {
            foreach (var property in document.CustomDocumentProperties)
            {
                var value = property.Value.Value;
                yield return new ExcelDocumentPropertyInfo(
                    property.Key,
                    "Custom",
                    value,
                    value?.GetType().FullName,
                    property.Value.PropertyType.ToString());
            }
        }
    }

    public static void SetProperty(ExcelDocument document, string name, object? value, bool custom = false)
    {
        if (custom)
        {
            document.SetCustomDocumentProperty(name, UnwrapValue(value));
            return;
        }

        if (BuiltInWriters.TryGetValue(name, out var builtInWriter))
        {
            builtInWriter(document.BuiltinDocumentProperties, UnwrapValue(value));
            return;
        }

        if (ApplicationWriters.TryGetValue(name, out var applicationWriter))
        {
            applicationWriter(document.ApplicationProperties, UnwrapValue(value));
            return;
        }

        throw new PSArgumentException($"'{name}' is not a supported Excel document property.");
    }

    public static void ApplyCommonProperties(
        ExcelDocument document,
        string? title,
        string? author,
        string? subject,
        string? keywords,
        string? description,
        string? category,
        string? company,
        string? manager,
        string? applicationName,
        string? lastModifiedBy)
    {
        if (title != null) SetProperty(document, "Title", title);
        if (author != null) SetProperty(document, "Creator", author);
        if (subject != null) SetProperty(document, "Subject", subject);
        if (keywords != null) SetProperty(document, "Keywords", keywords);
        if (description != null) SetProperty(document, "Description", description);
        if (category != null) SetProperty(document, "Category", category);
        if (company != null) SetProperty(document, "Company", company);
        if (manager != null) SetProperty(document, "Manager", manager);
        if (applicationName != null) SetProperty(document, "ApplicationName", applicationName);
        if (lastModifiedBy != null) SetProperty(document, "LastModifiedBy", lastModifiedBy);
    }

    private static object? UnwrapValue(object? value)
    {
        return value is PSObject psObject ? psObject.BaseObject : value;
    }

    private static string? ConvertToString(object? value)
    {
        value = UnwrapValue(value);
        return value == null ? null : LanguagePrimitives.ConvertTo<string>(value);
    }

    private static DateTime? ConvertToDateTime(object? value)
    {
        value = UnwrapValue(value);
        if (value == null)
        {
            return null;
        }

        if (value is DateTime dateTime)
        {
            return dateTime;
        }

        if (value is DateTimeOffset dateTimeOffset)
        {
            return dateTimeOffset.DateTime;
        }

        return LanguagePrimitives.ConvertTo<DateTime>(value);
    }
}
