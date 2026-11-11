using System.Globalization;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private readonly List<(string Location, PdfDictionary State)> _overprintStates = new();
    private readonly List<(string Location, PdfStream Profile, byte[] Data)> _iccCmykProfiles = new();
    private bool _usesIccCmyk;

    private void CheckResources(PdfDictionary resources, string location, int depth) {
        if (!_checkedResources.Add(resources)) return;
        if (ResolveDictionary(resources.Items.TryGetValue("ExtGState", out PdfObject? states) ? states : null) is PdfDictionary graphicsStates) {
            foreach (KeyValuePair<string, PdfObject> entry in graphicsStates.Items) {
                if (ResolveDictionary(entry.Value) is PdfDictionary state) CheckExtGState(state, resources, location + ", graphics state " + entry.Key, depth);
            }
        }
        if (ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? spaces) ? spaces : null) is PdfDictionary colorSpaces) {
            foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
                PdfObject? space = Resolve(entry.Value);
                if (space is PdfArray) CheckColorSpace(space, resources, location + ", colour space " + entry.Key, space);
            }
        }
        if (ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? patterns) ? patterns : null) is PdfDictionary patternResources) {
            foreach (KeyValuePair<string, PdfObject> entry in patternResources.Items) {
                PdfObject? pattern = Resolve(entry.Value);
                if (pattern is not null) CheckPattern(pattern, resources, location + ", pattern " + entry.Key, depth);
            }
        }
        if (ResolveDictionary(resources.Items.TryGetValue("Shading", out PdfObject? shadings) ? shadings : null) is PdfDictionary shadingResources) {
            foreach (KeyValuePair<string, PdfObject> entry in shadingResources.Items) {
                PdfObject? shading = Resolve(entry.Value);
                if (shading is not null) CheckShading(shading, resources, location + ", shading " + entry.Key);
            }
        }
    }

    private void CheckExtGState(PdfDictionary state, PdfDictionary? resources, string location, int depth) {
        if (state.Items.TryGetValue("TR", out PdfObject? transfer)) {
            if (NameOf(Resolve(transfer)) == "Identity") {
                if (Change(PdfAConversionChangeKind.GraphicsStateUpdated, location, "Removed the identity transfer function, which has no effect and PDF/A forbids.", state)) {
                    state.Items.Remove("TR");
                }
            } else {
                AddViolation(PdfAConversionRefusalReason.ForbiddenGraphicsState, location,
                    "The graphics state uses a transfer function (TR), which PDF/A forbids and which changes how colours are output.", state);
            }
        }
        if (state.Items.TryGetValue("TR2", out PdfObject? transfer2)) {
            string? name = NameOf(Resolve(transfer2));
            if (name == "Identity") {
                if (Change(PdfAConversionChangeKind.GraphicsStateUpdated, location, "Replaced the identity TR2 transfer function with /Default.", state)) {
                    state.Items["TR2"] = new PdfName("Default");
                }
            } else if (name != "Default") {
                AddViolation(PdfAConversionRefusalReason.ForbiddenGraphicsState, location,
                    "The graphics state uses a TR2 transfer function, which PDF/A forbids and which changes how colours are output.", state);
            }
        }
        if (state.Items.ContainsKey("HTP")) {
            if (Change(PdfAConversionChangeKind.GraphicsStateUpdated, location, "Removed the halftone phase (HTP), which applies only to printing devices.", state)) {
                state.Items.Remove("HTP");
            }
        }
        if (state.Items.ContainsKey("HTO")) {
            if (Change(PdfAConversionChangeKind.GraphicsStateUpdated, location, "Removed the halftone origin (HTO), which applies only to printing devices.", state)) {
                state.Items.Remove("HTO");
            }
        }
        if (state.Items.TryGetValue("HT", out PdfObject? halftone) && !IsCompliantHalftone(Resolve(halftone))) {
            if (Change(PdfAConversionChangeKind.GraphicsStateUpdated, location,
                    "Removed a halftone that PDF/A forbids. Halftones apply only to printing devices; the page renders the same.", state,
                    OfficeConversionLossKind.Approximation)) {
                state.Items.Remove("HT");
            }
        }
        if (state.Items.TryGetValue("BM", out PdfObject? blend)) {
            PdfObject? resolved = Resolve(blend);
            IEnumerable<PdfObject> modes = resolved is PdfArray array ? array.Items : new[] { resolved ?? PdfNull.Instance };
            foreach (PdfObject mode in modes) {
                string? modeName = NameOf(Resolve(mode));
                if (modeName is null || !BlendModes.Contains(modeName)) {
                    AddViolation(PdfAConversionRefusalReason.ForbiddenGraphicsState, location,
                        "The graphics state uses the blend mode " + (modeName ?? "(not a name)") + ", which PDF/A does not define.", state);
                }
            }
        }
        if (_part != 4 && state.Items.TryGetValue("RI", out PdfObject? intent) && NameOf(Resolve(intent)) is string intentName && !RenderingIntents.Contains(intentName)) {
            AddViolation(PdfAConversionRefusalReason.ForbiddenGraphicsState, location,
                "The rendering intent /" + intentName + " is not one PDF/A allows.", state);
        }
        if (Resolve(state.Items.TryGetValue("OPM", out PdfObject? overprintMode) ? overprintMode : null) is PdfNumber { Value: 1D } &&
            (Resolve(state.Items.TryGetValue("OP", out PdfObject? strokeOverprint) ? strokeOverprint : null) is PdfBoolean { Value: true } ||
             Resolve(state.Items.TryGetValue("op", out PdfObject? fillOverprint) ? fillOverprint : null) is PdfBoolean { Value: true })) {
            _overprintStates.Add((location, state));
        }
        if (ResolveDictionary(state.Items.TryGetValue("SMask", out PdfObject? mask) ? mask : null) is PdfDictionary softMask &&
            softMask.Items.TryGetValue("G", out PdfObject? group) && Resolve(group) is PdfStream groupForm) {
            WalkForm(groupForm, resources, location + ", soft mask", depth + 1);
        }
    }

    private bool IsCompliantHalftone(PdfObject? halftone) {
        if (halftone is PdfName name) return name.Name == "Default";
        PdfDictionary? dictionary = halftone is PdfStream stream ? stream.Dictionary : halftone as PdfDictionary;
        if (dictionary is null) return false;
        if (Resolve(dictionary.Items.TryGetValue("HalftoneType", out PdfObject? type) ? type : null) is not PdfNumber kind ||
            (kind.Value != 1 && kind.Value != 5)) return false;
        if (dictionary.Items.ContainsKey("HalftoneName")) return false;
        if (kind.Value == 1) return !dictionary.Items.ContainsKey("TransferFunction");
        foreach (KeyValuePair<string, PdfObject> component in dictionary.Items) {
            if (component.Key is "Type" or "HalftoneType" or "HalftoneName") continue;
            PdfObject? resolved = Resolve(component.Value);
            PdfDictionary? child = resolved is PdfStream childStream ? childStream.Dictionary : resolved as PdfDictionary;
            if (child is null) continue;
            if (child.Items.ContainsKey("HalftoneName")) return false;
            bool primary = component.Key is "Cyan" or "Magenta" or "Yellow" or "Black" or "Red" or "Green" or "Blue" or "Gray" or "Default";
            if (!primary && child.Items.ContainsKey("TransferFunction")) continue;
            if (primary && child.Items.ContainsKey("TransferFunction")) return false;
        }
        return true;
    }

    private void CheckColorSpaceName(string name, PdfDictionary? resources, string location, PdfObject owner) {
        switch (name) {
            case "DeviceRGB":
            case "RGB":
                RecordDeviceColor("DeviceRGB", location, owner, resources);
                return;
            case "DeviceCMYK":
            case "CMYK":
                RecordDeviceColor("DeviceCMYK", location, owner, resources);
                return;
            case "DeviceGray":
            case "G":
                RecordDeviceColor("DeviceGray", location, owner, resources);
                return;
            case "Pattern":
                return;
        }
        PdfObject? space = ResolveResourceObject(resources, "ColorSpace", name);
        if (space is null) return;
        CheckColorSpace(space, resources, location, owner, recordUse: true);
    }

    /// <summary>Checks one colour space; device families are recorded only when content actually paints with them.</summary>
    private void CheckColorSpace(PdfObject? space, PdfDictionary? resources, string location, PdfObject? owner, bool recordUse = true) {
        space = Resolve(space);
        if (space is PdfName name) {
            if (recordUse && name.Name is "DeviceRGB" or "DeviceCMYK" or "DeviceGray") RecordDeviceColor(name.Name, location, owner, resources);
            return;
        }
        if (space is not PdfArray array || array.Items.Count == 0 || Resolve(array.Items[0]) is not PdfName family) return;
        bool firstCheck = _checkedColorSpaces.Add(array);
        switch (family.Name) {
            case "ICCBased":
                if (firstCheck && array.Items.Count > 1 && Resolve(array.Items[1]) is PdfStream profile) CheckIccBased(profile, location);
                break;
            case "Indexed":
            case "I":
                if (array.Items.Count > 1) CheckColorSpace(array.Items[1], resources, location, owner, recordUse);
                break;
            case "Pattern":
                if (array.Items.Count > 1) CheckColorSpace(array.Items[1], resources, location, owner, recordUse);
                break;
            case "Separation":
                if (array.Items.Count > 2) CheckColorSpace(array.Items[2], resources, location, owner, recordUse);
                break;
            case "DeviceN":
                if (_part != 4 && array.Items.Count > 1 && Resolve(array.Items[1]) is PdfArray colorants && colorants.Items.Count > 32 && firstCheck) {
                    AddViolation(PdfAConversionRefusalReason.ImplementationLimitExceeded, location,
                        "A DeviceN colour space has " + colorants.Items.Count.ToString(CultureInfo.InvariantCulture) + " colourants; PDF/A allows 32.", array);
                }
                if (array.Items.Count > 2) CheckColorSpace(array.Items[2], resources, location, owner, recordUse);
                if (array.Items.Count > 4 && ResolveDictionary(array.Items[4]) is PdfDictionary attributes) {
                    if (ResolveDictionary(attributes.Items.TryGetValue("Colorants", out PdfObject? separations) ? separations : null) is PdfDictionary separationSpaces) {
                        foreach (PdfObject separation in separationSpaces.Items.Values) CheckColorSpace(separation, resources, location, owner, recordUse);
                    }
                    if (ResolveDictionary(attributes.Items.TryGetValue("Process", out PdfObject? process) ? process : null) is PdfDictionary processSpace &&
                        processSpace.Items.TryGetValue("ColorSpace", out PdfObject? processColorSpace)) {
                        CheckColorSpace(processColorSpace, resources, location, owner, recordUse);
                    }
                }
                break;
        }
    }

    private void CheckIccBased(PdfStream profile, string location) {
        if (!StreamDecoder.TryDecode(profile, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] data, _objects)) {
            AddViolation(PdfAConversionRefusalReason.InvalidIccProfile, location, "The ICC profile stream cannot be decoded.", profile);
            return;
        }
        if (!PdfAIccProfile.TryRead(data, out PdfAIccProfile header, out string? error)) {
            AddViolation(PdfAConversionRefusalReason.InvalidIccProfile, location, "The ICC profile is not valid for PDF/A: " + error + ".", profile);
            return;
        }
        if (header.Components == 4) {
            _usesIccCmyk = true;
            _iccCmykProfiles.Add((location, profile, data));
        }
        if (!header.IsColorSpaceClass) {
            AddViolation(PdfAConversionRefusalReason.InvalidIccProfile, location,
                "The ICC profile device class '" + header.DeviceClass + "' is not an input, display, output or colour-space profile.", profile);
            return;
        }
        if (Resolve(profile.Dictionary.Items.TryGetValue("N", out PdfObject? n) ? n : null) is PdfNumber count && (int)count.Value != header.Components) {
            AddViolation(PdfAConversionRefusalReason.InvalidIccProfile, location,
                "The ICC profile declares " + ((int)count.Value).ToString(CultureInfo.InvariantCulture) + " components but describes " +
                header.Components.ToString(CultureInfo.InvariantCulture) + ".", profile);
        }
    }

    private void CheckShading(PdfObject shading, PdfDictionary? resources, string location) {
        PdfDictionary? dictionary = ResolveDictionary(shading);
        if (dictionary is null) return;
        if (dictionary.Items.TryGetValue("ColorSpace", out PdfObject? space)) CheckColorSpace(space, resources, location, ObjectOwner(shading));
    }

    private void CheckPattern(PdfObject pattern, PdfDictionary? resources, string location, int depth) {
        if (pattern is PdfStream tiling) {
            WalkForm(tiling, resources, location, depth + 1);
            return;
        }
        if (pattern is PdfDictionary shadingPattern) {
            if (shadingPattern.Items.TryGetValue("Shading", out PdfObject? shading)) CheckShading(shading, resources, location);
            if (ResolveDictionary(shadingPattern.Items.TryGetValue("ExtGState", out PdfObject? state) ? state : null) is PdfDictionary graphicsState) {
                CheckExtGState(graphicsState, resources, location, depth);
            }
        }
    }

    private PdfObject? ObjectOwner(PdfObject value) => Resolve(value);

    private void CheckXObject(PdfStream xobject, PdfDictionary? resources, string location, int depth) {
        switch (GetName(xobject.Dictionary, "Subtype")) {
            case "Form":
                if (xobject.Dictionary.Items.ContainsKey("Ref")) {
                    AddViolation(PdfAConversionRefusalReason.ForbiddenXObject, location,
                        "A reference XObject imports its content from another file, which PDF/A forbids.", xobject);
                }
                WalkForm(xobject, resources, location, depth + 1);
                break;
            case "Image":
                CheckImage(xobject, resources, location);
                break;
            case "PS":
                AddViolation(PdfAConversionRefusalReason.ForbiddenXObject, location,
                    "A PostScript XObject is drawn here; PDF/A forbids PostScript.", xobject);
                break;
        }
    }

    private void CheckImage(PdfStream image, PdfDictionary? resources, string location) {
        PdfDictionary dictionary = image.Dictionary;
        bool isMask = Resolve(dictionary.Items.TryGetValue("ImageMask", out PdfObject? imageMask) ? imageMask : null) is PdfBoolean { Value: true };
        if (!isMask && dictionary.Items.TryGetValue("ColorSpace", out PdfObject? space)) CheckColorSpace(space, resources, location, image);
        if (_part != 4 && dictionary.Items.TryGetValue("Intent", out PdfObject? intent) && NameOf(Resolve(intent)) is string intentName && !RenderingIntents.Contains(intentName)) {
            AddViolation(PdfAConversionRefusalReason.ContentNotCompliant, location, "The image rendering intent /" + intentName + " is not one PDF/A allows.", image);
        }
        if (Resolve(dictionary.Items.TryGetValue("SMask", out PdfObject? softMask) ? softMask : null) is PdfStream mask) CheckImage(mask, resources, location + " soft mask");
    }

    private void CheckInlineImage(PdfDictionary image, PdfDictionary? resources, string location, PdfObject owner) {
        PdfObject? space = image.Items.TryGetValue("ColorSpace", out PdfObject? longSpace) ? longSpace :
            image.Items.TryGetValue("CS", out PdfObject? shortSpace) ? shortSpace : null;
        bool isMask = (image.Items.TryGetValue("ImageMask", out PdfObject? mask) ? mask : image.Items.TryGetValue("IM", out PdfObject? shortMask) ? shortMask : null)
            is PdfBoolean { Value: true };
        if (space is PdfName name && !isMask) {
            CheckColorSpaceName(name.Name, resources, location, owner);
        } else if (space is PdfArray && !isMask) {
            CheckColorSpace(space, resources, location, owner);
        }
        PdfObject? filter = image.Items.TryGetValue("Filter", out PdfObject? longFilter) ? longFilter :
            image.Items.TryGetValue("F", out PdfObject? shortFilter) ? shortFilter : null;
        IEnumerable<PdfObject> filters = filter is PdfArray filterArray ? filterArray.Items : filter is null ? Array.Empty<PdfObject>() : new[] { filter };
        foreach (PdfObject item in filters) {
            if (NameOf(item) is "LZWDecode" or "LZW") {
                AddViolation(PdfAConversionRefusalReason.InlineImageNotCompliant, location,
                    "An inline image is LZW-compressed, which PDF/A forbids; re-encoding it would require rewriting the page content.", owner);
            }
        }
        PdfObject? interpolate = image.Items.TryGetValue("Interpolate", out PdfObject? longInterpolate) ? longInterpolate :
            image.Items.TryGetValue("I", out PdfObject? shortInterpolate) ? shortInterpolate : null;
        if (interpolate is PdfBoolean { Value: true }) {
            AddViolation(PdfAConversionRefusalReason.InlineImageNotCompliant, location,
                "An inline image requests interpolation, which PDF/A forbids; changing it would require rewriting the page content.", owner);
        }
    }

    private void CheckTransparencyGroup(PdfDictionary owner, string location, PdfObject ownerObject, PdfDictionary? resources) {
        if (ResolveDictionary(owner.Items.TryGetValue("Group", out PdfObject? group) ? group : null) is not PdfDictionary groupDictionary) return;
        if (groupDictionary.Items.TryGetValue("CS", out PdfObject? space)) {
            CheckColorSpace(space, resources, location + " transparency group", ownerObject);
        }
    }

    /// <summary>Chooses the output intent and checks every device colour use against it.</summary>
    private void ProcessColor(PdfDictionary catalog) {
        var intents = new List<PdfDictionary>();
        PdfArray? intentArray = Resolve(catalog.Items.TryGetValue("OutputIntents", out PdfObject? existing) ? existing : null) as PdfArray;
        if (intentArray is not null) {
            foreach (PdfObject item in intentArray.Items) {
                if (ResolveDictionary(item) is PdfDictionary intent) intents.Add(intent);
            }
        }

        if (_part == 4) RemoveOutputProfileReferences(intents, "catalog output intent");
        var profiles = new List<PdfObject>();
        foreach (PdfDictionary intent in intents) {
            if (intent.Items.TryGetValue("DestOutputProfile", out PdfObject? profileReference)) {
                PdfObject? resolved = Resolve(profileReference);
                if (resolved is not null && !profiles.Any(profile => ReferenceEquals(profile, resolved))) profiles.Add(resolved);
            }
        }

        int components;
        if (_options.OutputIntentProfile is byte[] suppliedProfile && !_verifyOnly) {
            if (profiles.Count > 0) {
                AddViolation(PdfAConversionRefusalReason.ConflictingOutputIntents, "catalog output intents",
                    "The PDF already has an output intent profile; PDF/A requires every output intent to use the same profile as the supplied one.", intentArray);
                return;
            }
            PdfAIccProfile.TryRead(suppliedProfile, out PdfAIccProfile header, out _);
            components = header.Components;
            Change(PdfAConversionChangeKind.OutputIntentAdded, "catalog", "Added a PDF/A output intent with the supplied ICC profile.", catalog);
            AddOutputIntent(catalog, intentArray, suppliedProfile, header.Components,
                string.IsNullOrWhiteSpace(_options.OutputConditionIdentifier) ? "Custom" : _options.OutputConditionIdentifier!, null);
        } else if (profiles.Count > 1) {
            AddViolation(PdfAConversionRefusalReason.ConflictingOutputIntents, "catalog output intents",
                "The PDF has " + profiles.Count.ToString(CultureInfo.InvariantCulture) +
                " output intents with different destination profiles; PDF/A requires them to share one profile.", intentArray);
            return;
        } else if (profiles.Count == 1) {
            if (profiles[0] is not PdfStream profileStream ||
                !StreamDecoder.TryDecode(profileStream, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] profileData, _objects) ||
                !PdfAIccProfile.TryRead(profileData, out PdfAIccProfile header, out string? error) || !header.IsOutputDevice) {
                AddViolation(PdfAConversionRefusalReason.InvalidIccProfile, "catalog output intent",
                    "The existing output intent profile is not a valid output or display ICC profile.", profiles[0]);
                return;
            }
            components = header.Components;
            bool hasPdfAIntent = intents.Any(intent => GetName(intent, "S") == "GTS_PDFA1");
            if (!hasPdfAIntent) {
                PdfDictionary template = intents.First(intent => intent.Items.ContainsKey("DestOutputProfile"));
                if (Change(PdfAConversionChangeKind.OutputIntentAdded, "catalog",
                        "Added a PDF/A output intent that shares the existing output intent's destination profile.", catalog)) {
                    var intent = new PdfDictionary();
                    intent.Items["Type"] = new PdfName("OutputIntent");
                    intent.Items["S"] = new PdfName("GTS_PDFA1");
                    foreach (string key in new[] { "OutputConditionIdentifier", "OutputCondition", "RegistryName", "Info" }) {
                        if (template.Items.TryGetValue(key, out PdfObject? value)) intent.Items[key] = value;
                    }
                    if (!intent.Items.ContainsKey("OutputConditionIdentifier")) intent.Items["OutputConditionIdentifier"] = new PdfStringObj("Custom");
                    intent.Items["DestOutputProfile"] = template.Items["DestOutputProfile"];
                    intentArray!.Items.Add(AddObject(intent));
                }
            }
        } else {
            components = 3;
            if (Change(PdfAConversionChangeKind.OutputIntentAdded, "catalog", "Added the sRGB IEC61966-2.1 PDF/A output intent.", catalog)) {
                AddOutputIntent(catalog, intentArray, PdfIccProfiles.SrgbIec6196621, 3,
                    PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, "sRGB IEC61966-2.1");
            }
        }

        if (_usesIccCmyk || components == 4) {
            foreach ((string location, PdfDictionary state) in _overprintStates) {
                AddViolation(PdfAConversionRefusalReason.ForbiddenGraphicsState, location,
                    "The graphics state sets overprint mode 1 with overprinting while CMYK colour is ICC-based, which PDF/A forbids; changing it would change printed output.", state);
            }
        }
        Dictionary<int, int> pageComponents = _part == 4 ? ProcessPageOutputIntents() : new Dictionary<int, int>();
        if (_part == 4 && (components == 4 || pageComponents.ContainsValue(4))) CheckCmykProfilesDifferFromOutputIntent(profiles);
        foreach (DeviceColorUse use in _deviceColors) {
            if (use.Family == "DeviceGray") continue;
            int needed = use.Family == "DeviceRGB" ? 3 : 4;
            int available = pageComponents.TryGetValue(use.PageIndex, out int pageIntent) ? pageIntent : components;
            if (available == needed) continue;
            string defaultKey = use.Family == "DeviceRGB" ? "DefaultRGB" : "DefaultCMYK";
            if (use.Resources is not null && ResolveResourceObject(use.Resources, "ColorSpace", defaultKey) is not null) continue;
            AddViolation(PdfAConversionRefusalReason.DeviceColorNotCharacterized, use.Location,
                use.Family + " colour is painted here, but the output intent is " + DescribeComponents(available) +
                ". PDF/A requires device colour to match the output intent" +
                (use.Family == "DeviceCMYK" ? "; supply a CMYK profile in PdfAConversionOptions.OutputIntentProfile." : "."),
                use.Owner);
        }
    }

    /// <summary>PDF/A-4 forbids DestOutputProfileRef (6.2.3); it points at an external profile that no viewer uses for display.</summary>
    private void RemoveOutputProfileReferences(IEnumerable<PdfDictionary> intents, string location) {
        foreach (PdfDictionary intent in intents) {
            if (intent.Items.ContainsKey("DestOutputProfileRef") &&
                Change(PdfAConversionChangeKind.OutputIntentAdded, location,
                    "Removed the external output profile reference (DestOutputProfileRef), which PDF/A-4 forbids.", intent, OfficeConversionLossKind.Omission)) {
                intent.Items.Remove("DestOutputProfileRef");
            }
        }
    }

    /// <summary>
    /// PDF 2.0 pages can carry their own output intents. PDF/A-4 characterizes device colour on such a page by its PDF/A
    /// output intent, so each page intent is checked like the document's; returns the colour components per page.
    /// </summary>
    private Dictionary<int, int> ProcessPageOutputIntents() {
        var components = new Dictionary<int, int>();
        IReadOnlyList<PdfReadPage> pages = _document.Pages;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfDictionary page = pages[pageIndex].PageDictionary;
            if (Resolve(page.Items.TryGetValue("OutputIntents", out PdfObject? value) ? value : null) is not PdfArray array) continue;
            string location = "page " + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + " output intents";
            var intents = array.Items.Select(ResolveDictionary).Where(static intent => intent is not null).Cast<PdfDictionary>().ToList();
            RemoveOutputProfileReferences(intents, location);
            var profiles = new List<PdfObject>();
            foreach (PdfDictionary intent in intents) {
                PdfObject? profile = intent.Items.TryGetValue("DestOutputProfile", out PdfObject? reference) ? Resolve(reference) : null;
                if (profile is not null && !profiles.Any(existing => ReferenceEquals(existing, profile))) profiles.Add(profile);
            }
            if (profiles.Count > 1) {
                AddViolation(PdfAConversionRefusalReason.ConflictingOutputIntents, location,
                    "The page has output intents with different destination profiles; PDF/A-4 requires them to share one profile.", array);
                continue;
            }
            if (profiles.Count == 0) continue;
            if (profiles[0] is not PdfStream profileStream ||
                !StreamDecoder.TryDecode(profileStream, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] data, _objects) ||
                !PdfAIccProfile.TryRead(data, out PdfAIccProfile header, out _) || !header.IsOutputDevice) {
                AddViolation(PdfAConversionRefusalReason.InvalidIccProfile, location,
                    "The page output intent profile is not a valid output or display ICC profile.", profiles[0]);
                continue;
            }
            if (intents.Any(intent => GetName(intent, "S") == "GTS_PDFA1" && intent.Items.ContainsKey("DestOutputProfile"))) {
                components[pageIndex] = header.Components;
            }
        }
        return components;
    }

    /// <summary>PDF/A-4 (6.2.4.2) forbids an ICCBased CMYK space whose profile is the CMYK output intent profile itself.</summary>
    private void CheckCmykProfilesDifferFromOutputIntent(List<PdfObject> documentProfiles) {
        var outputProfiles = new List<(PdfObject? Profile, byte[]? Data)>();
        if (_options.OutputIntentProfile is byte[] supplied) outputProfiles.Add((null, supplied));
        IEnumerable<PdfObject> candidates = documentProfiles.Concat(_document.Pages
            .Select(page => Resolve(page.PageDictionary.Items.TryGetValue("OutputIntents", out PdfObject? value) ? value : null) as PdfArray)
            .Where(static array => array is not null)
            .SelectMany(array => array!.Items.Select(ResolveDictionary))
            .Where(static intent => intent is not null && intent.Items.ContainsKey("DestOutputProfile"))
            .Select(intent => Resolve(intent!.Items["DestOutputProfile"])!));
        foreach (PdfObject profile in candidates) {
            byte[]? data = profile is PdfStream stream && StreamDecoder.TryDecode(stream, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] decoded, _objects)
                ? decoded : null;
            outputProfiles.Add((profile, data));
        }
        foreach ((string location, PdfStream profile, byte[] data) in _iccCmykProfiles) {
            if (outputProfiles.Any(output => ReferenceEquals(output.Profile, profile) || output.Data is not null && output.Data.SequenceEqual(data))) {
                AddViolation(PdfAConversionRefusalReason.InvalidIccProfile, location,
                    "An ICCBased CMYK colour space uses the CMYK output intent profile itself, which PDF/A-4 forbids; replacing it would change how the colour is defined.", profile);
            }
        }
    }

    private static string DescribeComponents(int components) => components switch {
        1 => "Gray",
        3 => "RGB",
        4 => "CMYK",
        _ => components.ToString(CultureInfo.InvariantCulture) + "-component"
    };

    private void AddOutputIntent(PdfDictionary catalog, PdfArray? intentArray, byte[] profile, int components, string identifier, string? info) {
        var profileDictionary = new PdfDictionary();
        profileDictionary.Items["N"] = new PdfNumber(components);
        profileDictionary.Items["Filter"] = new PdfName("FlateDecode");
        PdfReference profileReference = AddObject(new PdfStream(profileDictionary, PdfFlateEncoder.Compress(profile, _cancellationToken)));
        var intent = new PdfDictionary();
        intent.Items["Type"] = new PdfName("OutputIntent");
        intent.Items["S"] = new PdfName("GTS_PDFA1");
        intent.Items["OutputConditionIdentifier"] = new PdfStringObj(identifier);
        if (info is not null) {
            intent.Items["RegistryName"] = new PdfStringObj("http://www.color.org");
            intent.Items["Info"] = new PdfStringObj(info);
        }
        intent.Items["DestOutputProfile"] = profileReference;
        if (intentArray is null) {
            intentArray = new PdfArray();
            catalog.Items["OutputIntents"] = intentArray;
        }
        intentArray.Items.Add(AddObject(intent));
    }
}
