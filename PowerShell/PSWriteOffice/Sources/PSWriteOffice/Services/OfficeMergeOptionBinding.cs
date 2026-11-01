using System;
using System.Collections;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

/// <summary>Binds shell option bags to the existing native merge options for one selected family.</summary>
internal static class OfficeMergeOptionBinding {
    internal static object Capture(EngineIntrinsics engine, OfficeMergeFamily family, object[]? values) {
        Type type = family switch {
            OfficeMergeFamily.Word => typeof(WordDocumentMergeOptions),
            OfficeMergeFamily.Excel => typeof(ExcelWorkbookMergeOptions),
            _ => typeof(PdfMergeOptions)
        };
        object result = Activator.CreateInstance(type)!;
        bool typed = false;
        foreach (object item in values ?? Array.Empty<object>()) {
            object? value = Unwrap(item);
            if (value is IDictionary) continue;
            if (value == null || !type.IsInstanceOfType(value))
                throw new PSArgumentException("The selected merge accepts only " + type.Name + " or property hashtables.");
            if (typed) throw new PSArgumentException("Provide at most one typed merge options object.");
            typed = true;
            // Override tables apply to a separate bag rather than modifying the caller's options object.
            result = CopyProperties(value);
            if (result is PdfMergeOptions pdfCopy && pdfCopy.Policy != null)
                pdfCopy.Policy = (PdfMergePolicy)CopyProperties(pdfCopy.Policy);
        }
        var binder = new OfficeOptionsTransformationAttribute(type);
        foreach (object item in values ?? Array.Empty<object>()) {
            if (Unwrap(item) is IDictionary table) result = binder.ApplyTo(engine, result, table);
        }
        ValidateEnums(result);
        if (result is ExcelWorkbookMergeOptions excel) {
            if (excel.MaxDefinedNames <= 0 || excel.MaxDefinedNameCharacters <= 0 || excel.MaxInCellImages <= 0 ||
                excel.MaxInCellImageBytes <= 0 || excel.MaxTotalInCellImageBytes <= 0)
                throw new PSArgumentException("Excel merge resource limits must be positive.");
        }
        if (result is PdfMergeOptions pdf) {
            if (pdf.Policy == null) throw new PSArgumentException("PDF merge Policy cannot be null.");
            ValidateEnums(pdf.Policy);
            if (pdf.ResizePages is PdfPageResizeOptions resize) {
                ValidateEnums(resize);
                double width = resize.PageSize.Width.ToPoints(), height = resize.PageSize.Height.ToPoints(), margin = resize.Margin.ToPoints();
                if (!Finite(width) || !Finite(height) || width <= 0 || height <= 0 || !Finite(margin) || margin < 0 ||
                    margin * 2 >= width || margin * 2 >= height)
                    throw new PSArgumentException("PDF resize settings require finite positive dimensions and a margin leaving positive content area.");
            }
        }
        return result;
    }

    internal static object? Unwrap(object? value) {
        while (value is PSObject wrapper && !ReferenceEquals(value, wrapper.BaseObject)) value = wrapper.BaseObject;
        return value;
    }

    // Merge option records contain scalar settings; PDF's separately mutable policy needs its own copy.
    private static object CopyProperties(object source) {
        Type type = source.GetType();
        object copy = Activator.CreateInstance(type)!;
        foreach (PropertyInfo property in type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            if (property.CanRead && property.SetMethod?.IsPublic == true) property.SetValue(copy, property.GetValue(source));
        return copy;
    }

    private static void ValidateEnums(object options) {
        foreach (PropertyInfo property in options.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
            if (property.PropertyType.IsEnum && !Enum.IsDefined(property.PropertyType, property.GetValue(options)!))
                throw new PSArgumentException("Invalid merge option " + property.Name + ".");
        }
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
