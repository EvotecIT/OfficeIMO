using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    private static PdfSignatureProfile ResolveSignatureProfile(PdfExternalSignatureOptions options) {
        if (options.SubFilter == PdfExternalSignatureSubFilter.DocumentTimestamp) {
            if (options.Profile == PdfSignatureProfile.Certification) {
                throw new ArgumentException("Certification signatures cannot use the document timestamp subfilter.", nameof(options));
            }

            return PdfSignatureProfile.DocumentTimestamp;
        }

        int profile = (int)options.Profile;
        if (profile < (int)PdfSignatureProfile.Approval || profile > (int)PdfSignatureProfile.DocumentTimestamp) {
            throw new ArgumentOutOfRangeException(nameof(options), options.Profile, "Unsupported PDF signature profile.");
        }

        if (options.Profile == PdfSignatureProfile.Certification) {
            int permission = (int)options.CertificationPermission;
            if (permission < 1 || permission > 3) {
                throw new ArgumentOutOfRangeException(nameof(options), options.CertificationPermission, "Certification permission must map to DocMDP /P 1, 2, or 3.");
            }
        }

        return options.Profile;
    }

    private static PdfExternalSignatureSubFilter ResolveSignatureSubFilter(PdfExternalSignatureOptions options) {
        PdfSignatureProfile profile = ResolveSignatureProfile(options);
        if (profile == PdfSignatureProfile.DocumentTimestamp) {
            return PdfExternalSignatureSubFilter.DocumentTimestamp;
        }

        return options.SubFilter;
    }

    private static void ApplySignatureProfile(
        byte[] sourcePdf,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        PdfDictionary signatureField,
        int signatureObjectNumber,
        PdfExternalSignatureOptions options,
        PdfSignatureProfile profile,
        ref int nextObjectNumber,
        ref bool catalogChanged,
        HashSet<int> changedObjects) {
        if (profile == PdfSignatureProfile.Certification) {
            ApplyCertificationReference(
                objects,
                catalog,
                signatureObjectNumber,
                ref catalogChanged,
                changedObjects);
        }

        if (options.VisibleAppearance is not null) {
            if (profile == PdfSignatureProfile.DocumentTimestamp) {
                throw new ArgumentException("Document timestamp signatures do not support visible approval widgets.", nameof(options));
            }

            ApplyVisibleSignatureAppearance(
                sourcePdf,
                objects,
                signatureField,
                options.FieldName,
                options.VisibleAppearance,
                ref nextObjectNumber,
                changedObjects);
        }
    }

    private static void ApplyCertificationReference(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        int signatureObjectNumber,
        ref bool catalogChanged,
        HashSet<int> changedObjects) {
        PdfDictionary permissions;
        if (catalog.Items.TryGetValue("Perms", out PdfObject? permissionsObject) &&
            permissionsObject is PdfReference permissionsReference &&
            ResolveDictionary(objects, permissionsReference) is PdfDictionary referencedPermissions) {
            permissions = referencedPermissions;
            changedObjects.Add(permissionsReference.ObjectNumber);
        } else if (catalog.Items.TryGetValue("Perms", out permissionsObject) &&
            ResolveDictionary(objects, permissionsObject) is PdfDictionary directPermissions) {
            permissions = directPermissions;
            catalogChanged = true;
        } else {
            permissions = new PdfDictionary();
            catalog.Items["Perms"] = permissions;
            catalogChanged = true;
        }

        permissions.Items["DocMDP"] = new PdfReference(signatureObjectNumber, 0);
    }

    private static void ApplyVisibleSignatureAppearance(
        byte[] sourcePdf,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary signatureField,
        string fieldName,
        PdfVisibleSignatureAppearanceOptions options,
        ref int nextObjectNumber,
        HashSet<int> changedObjects) {
        ValidateVisibleAppearance(options);
        PdfReadDocument document = PdfReadDocument.Open(sourcePdf);
        if (options.PageNumber > document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(options), "Visible signature page exceeds the document page count.");
        }

        int pageObjectNumber = document.Pages[options.PageNumber - 1].ObjectNumber;
        if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? pageObject) ||
            pageObject.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("Visible signature target page object could not be resolved.");
        }

        int appearanceObjectNumber = nextObjectNumber++;
        string appearanceText = string.IsNullOrWhiteSpace(options.Text) ? fieldName : options.Text!;
        if (!PdfWinAnsiEncoding.CanEncode(appearanceText, out int unsupportedIndex)) {
            char unsupportedCharacter = appearanceText[unsupportedIndex];
            throw new ArgumentException("Visible signature text contains unsupported Helvetica character '" + unsupportedCharacter + "'.", nameof(options));
        }

        objects[appearanceObjectNumber] = new PdfIndirectObject(
            appearanceObjectNumber,
            0,
            BuildVisibleAppearanceStream(options, appearanceText));
        changedObjects.Add(appearanceObjectNumber);

        PdfArray annotations = EnsurePageAnnotations(objects, pageDictionary, changedObjects);
        PdfReference signatureFieldReference = FindSignatureFieldReference(signatureField, objects, changedObjects);
        annotations.Items.Add(signatureFieldReference);
        changedObjects.Add(pageObjectNumber);

        signatureField.Items["Type"] = new PdfName("Annot");
        signatureField.Items["Subtype"] = new PdfName("Widget");
        signatureField.Items["Rect"] = CreateRectangle(options.X, options.Y, options.Width, options.Height);
        signatureField.Items["P"] = new PdfReference(pageObjectNumber, pageObject.Generation);
        signatureField.Items["F"] = new PdfNumber(4);
        var appearanceDictionary = new PdfDictionary();
        appearanceDictionary.Items["N"] = new PdfReference(appearanceObjectNumber, 0);
        signatureField.Items["AP"] = appearanceDictionary;
    }

    private static PdfReference FindSignatureFieldReference(
        PdfDictionary signatureField,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> changedObjects) {
        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects) {
            if (ReferenceEquals(entry.Value.Value, signatureField)) {
                changedObjects.Add(entry.Key);
                return new PdfReference(entry.Key, entry.Value.Generation);
            }
        }

        throw new InvalidOperationException("Visible signature field object was not registered before page annotation wiring.");
    }

    private static PdfArray EnsurePageAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        HashSet<int> changedObjects) {
        if (pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotationsObject)) {
            if (annotationsObject is PdfReference annotationsReference &&
                ResolveObject(objects, annotationsReference) is PdfArray referencedAnnotations) {
                changedObjects.Add(annotationsReference.ObjectNumber);
                return referencedAnnotations;
            }

            if (ResolveObject(objects, annotationsObject) is PdfArray directAnnotations) {
                return directAnnotations;
            }
        }

        var annotations = new PdfArray();
        pageDictionary.Items["Annots"] = annotations;
        return annotations;
    }

    private static PdfStream BuildVisibleAppearanceStream(
        PdfVisibleSignatureAppearanceOptions options,
        string text) {
        var font = new PdfDictionary();
        font.Items["Type"] = new PdfName("Font");
        font.Items["Subtype"] = new PdfName("Type1");
        font.Items["BaseFont"] = new PdfName("Helvetica");
        var fonts = new PdfDictionary();
        fonts.Items["F1"] = font;
        var resources = new PdfDictionary();
        resources.Items["Font"] = fonts;
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["FormType"] = new PdfNumber(1);
        dictionary.Items["BBox"] = CreateRectangle(0, 0, options.Width, options.Height);
        dictionary.Items["Resources"] = resources;

        double textY = Math.Max(2, (options.Height - options.FontSize) / 2);
        string content =
            "q\n" +
            FormatColor(options.BackgroundColor) + " rg 0 0 " + Format(options.Width) + " " + Format(options.Height) + " re f\n" +
            FormatColor(options.BorderColor) + " RG 1 w 0.5 0.5 " + Format(Math.Max(0, options.Width - 1)) + " " + Format(Math.Max(0, options.Height - 1)) + " re S\n" +
            "BT " + FormatColor(options.TextColor) + " rg /F1 " + Format(options.FontSize) + " Tf 6 " + Format(textY) + " Td " + PdfSyntaxEscaper.LiteralString(text) + " Tj ET\nQ\n";
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfArray CreateRectangle(double x, double y, double width, double height) {
        var rectangle = new PdfArray();
        rectangle.Items.Add(new PdfNumber(x));
        rectangle.Items.Add(new PdfNumber(y));
        rectangle.Items.Add(new PdfNumber(x + width));
        rectangle.Items.Add(new PdfNumber(y + height));
        return rectangle;
    }

    private static void ValidateVisibleAppearance(PdfVisibleSignatureAppearanceOptions options) {
        ValidateFinite(options.X, nameof(options.X));
        ValidateFinite(options.Y, nameof(options.Y));
        ValidateFinite(options.Width, nameof(options.Width));
        ValidateFinite(options.Height, nameof(options.Height));
        ValidateFinite(options.FontSize, nameof(options.FontSize));
    }

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Visible signature geometry must be finite.");
        }
    }

    private static string FormatColor(PdfColor color) =>
        Format(color.R) + " " + Format(color.G) + " " + Format(color.B);

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
}
