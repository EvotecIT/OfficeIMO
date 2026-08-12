using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAcroFormAuthoringTests {
    [Fact]
    public void Edit_AuthorsChromiumEditorFieldKindsAppearancesAndWidgetScripts() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Chromium form authoring acceptance"))
            .ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Create(new PdfFormFieldCreateOptions {
                Name = "notes",
                Kind = PdfFormFieldCreationKind.Text,
                X = 72,
                Y = 650,
                Width = 220,
                Height = 60,
                Value = "line one\nline two",
                JavaScript = "this.getField('total').value = 2;",
                Style = new PdfFormFieldStyle { IsMultiline = true }
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "agree",
                Kind = PdfFormFieldCreationKind.CheckBox,
                X = 72,
                Y = 610,
                Width = 18,
                Height = 18,
                Value = "Yes",
                JavaScript = "app.alert('checked');"
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "country",
                Kind = PdfFormFieldCreationKind.Choice,
                X = 72,
                Y = 565,
                Width = 180,
                Height = 24,
                ChoiceOptions = new[] { "Poland", "Germany", "Australia" },
                Value = "Germany",
                IsComboBox = true,
                JavaScript = "app.alert('country');"
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "size",
                Kind = PdfFormFieldCreationKind.RadioButtonGroup,
                X = 72,
                Y = 430,
                Width = 180,
                Height = 100,
                ChoiceOptions = new[] { "Small", "Medium", "Large" },
                Value = "Medium",
                JavaScript = "app.alert('size');"
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "calculate",
                Kind = PdfFormFieldCreationKind.PushButton,
                X = 72,
                Y = 385,
                Width = 110,
                Height = 26,
                Caption = "Calculate",
                JavaScript = "this.getField('total').value = 42;"
            }));

        PdfDocumentInfo info = result.ToDocument().Inspect();
        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Equal(5, info.FormFields.Count);

        PdfFormField notes = info.FormFieldsByName["notes"];
        Assert.True(notes.IsMultiline);
        Assert.Equal("line one\nline two", notes.Value);
        Assert.Equal("U", Assert.Single(Assert.Single(notes.Widgets).Actions).TriggerName);
        Assert.Equal("this.getField('total').value = 2;", notes.JavaScript);

        PdfFormField country = info.FormFieldsByName["country"];
        Assert.True(country.IsCombo);
        Assert.Equal(new[] { "Poland", "Germany", "Australia" }, country.Options.Select(static option => option.ExportValue));
        Assert.Equal("Germany", country.Value);

        PdfFormField radio = info.FormFieldsByName["size"];
        Assert.True(radio.IsRadioButton);
        Assert.Equal("Medium", radio.Value);
        Assert.Equal(3, radio.WidgetCount);
        Assert.All(radio.Widgets, static widget => Assert.True(widget.HasNormalAppearanceStates));
        Assert.All(radio.Widgets, static widget => Assert.Equal(180D, widget.Width, 3));
        Assert.All(radio.Widgets, static widget => Assert.Equal("U", Assert.Single(widget.Actions).TriggerName));

        PdfFormField button = info.FormFieldsByName["calculate"];
        Assert.True(button.IsPushButton);
        PdfFormWidgetAction action = Assert.Single(Assert.Single(button.Widgets).Actions);
        Assert.True(action.IsPrimary);
        Assert.Equal("JavaScript", action.ActionType);
        Assert.Equal("this.getField('total').value = 42;", action.JavaScript);

        Assert.All(info.FormFields.SelectMany(static field => field.Widgets), static widget => Assert.True(widget.IsPrint));
        Assert.Contains(PdfSanitizer.Analyze(result.ToBytes()), static finding =>
            finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "JavaScript");
    }

    [Fact]
    public void Fill_PreservesAuthoredWidgetActionsAndUpdatesRadioAndChoiceValues() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Fill scripted fields")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Create(new PdfFormFieldCreateOptions {
                Name = "country",
                Kind = PdfFormFieldCreationKind.Choice,
                X = 72,
                Y = 600,
                Width = 160,
                Height = 24,
                ChoiceOptions = new[] { "One", "Two" },
                JavaScript = "app.alert('country');"
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "option",
                Kind = PdfFormFieldCreationKind.RadioButtonGroup,
                X = 72,
                Y = 480,
                Width = 160,
                Height = 90,
                ChoiceOptions = new[] { "A", "B" },
                JavaScript = "app.alert('radio');"
            })).ToBytes();

        PdfDocument filled = PdfDocument.Open(authored).Forms.Fill(new Dictionary<string, string> {
            ["country"] = "Two",
            ["option"] = "B"
        });

        PdfDocumentInfo info = filled.Inspect();
        Assert.Equal("Two", info.FormFieldsByName["country"].Value);
        Assert.Equal("B", info.FormFieldsByName["option"].Value);
        Assert.Equal("app.alert('country');", info.FormFieldsByName["country"].JavaScript);
        Assert.Equal("app.alert('radio');", info.FormFieldsByName["option"].JavaScript);
    }

    [Fact]
    public void Fill_RejectsAuthoredPushButtonsWithoutReplacingTheirCaptionAppearance() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Push button fill guard")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "calculate",
            Kind = PdfFormFieldCreationKind.PushButton,
            Caption = "Calculate"
        })).ToBytes();

        PdfDocument opened = PdfDocument.Open(authored);
        Assert.Throws<ArgumentException>(() => opened.Forms.Fill(new Dictionary<string, string> {
            ["calculate"] = "On"
        }));

        PdfFormField button = PdfInspector.Inspect(authored).FormFieldsByName["calculate"];
        Assert.True(button.IsPushButton);
        Assert.Equal(authored, opened.ToBytes());
    }

    [Fact]
    public void Edit_AllowsSubsequentEditsWhenActiveContentBelongsOnlyToFormWidgets() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Repeated scripted edits")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "calculate",
            Kind = PdfFormFieldCreationKind.PushButton,
            Caption = "Calculate",
            JavaScript = "app.alert('kept');"
        })).ToBytes();

        PdfAcroFormEditResult renamed = PdfDocument.Open(authored).Forms.Edit(edit => edit.Rename("calculate", "recalculate"));

        PdfFormField field = Assert.Single(renamed.Fields);
        Assert.Equal("recalculate", field.Name);
        Assert.Equal("app.alert('kept');", field.JavaScript);
        Assert.True(renamed.PreservationReport.IsPreserved);
    }

    [Fact]
    public void Sanitize_RemovesAuthoredWidgetJavaScriptWithoutRemovingFields() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Sanitize widget script")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "submit",
            Kind = PdfFormFieldCreationKind.PushButton,
            X = 72,
            Y = 600,
            Width = 100,
            Height = 24,
            Caption = "Submit",
            JavaScript = "app.alert('submit');"
        })).ToBytes();

        PdfSanitizationResult sanitized = PdfSanitizer.Sanitize(authored);
        PdfFormField field = Assert.Single(sanitized.ToDocument().Inspect().FormFields);

        Assert.Equal("submit", field.Name);
        Assert.False(field.HasJavaScript);
        Assert.False(Assert.Single(field.Widgets).HasActions);
        Assert.True(sanitized.PreservationReport.IsPreserved);
        Assert.Contains(sanitized.RemovedFindings, static finding => finding.Detail == "JavaScript");
    }

    [Fact]
    public void Create_SnapshotsMutableOptionsAndEnforcesWidgetJavaScriptBudget() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Snapshot options")).ToBytes();
        var choices = new List<string> { "One", "Two" };
        var options = new PdfFormFieldCreateOptions {
            Name = "choice",
            Kind = PdfFormFieldCreationKind.Choice,
            X = 72,
            Y = 600,
            Width = 160,
            Height = 24,
            ChoiceOptions = choices,
            JavaScript = "app.alert('kept');"
        };

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => {
            edit.Create(options);
            options.Name = "mutated";
            choices[0] = "Changed";
            options.Style = new PdfFormFieldStyle { IsReadOnly = true };
        });

        PdfFormField field = Assert.Single(result.Fields);
        Assert.Equal("choice", field.Name);
        Assert.Equal(new[] { "One", "Two" }, field.Options.Select(static option => option.ExportValue));
        Assert.False(field.IsReadOnly);

        var readOptions = new PdfReadOptions { Limits = new PdfReadLimits { MaxJavaScriptBytes = 8 } };
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Open(source, readOptions).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
                Name = "limited",
                Kind = PdfFormFieldCreationKind.PushButton,
                JavaScript = "app.alert('too large');"
            })));
        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void Reader_DecodesExistingWidgetJavaScriptStreamAndAppliesLimits() {
        byte[] source = BuildWidgetJavaScriptStreamPdf("app.alert('stream');");

        PdfFormField field = Assert.Single(PdfDocument.Open(source).Inspect().FormFields);
        PdfFormWidgetAction action = Assert.Single(Assert.Single(field.Widgets).Actions);
        Assert.Equal("A", action.TriggerName);
        Assert.Equal("app.alert('stream');", action.JavaScript);

        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxJavaScriptBytes = 8 } };
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfDocument.Open(source, options).Inspect());
        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void Reader_TraversesChainedWidgetActionsWithCyclesAndAggregateBudgets() {
        byte[] source = BuildWidgetActionGraphPdf(includeOpenAction: false);

        PdfFormWidget widget = Assert.Single(Assert.Single(PdfDocument.Open(source).Inspect().FormFields).Widgets);
        Assert.Equal(new[] { "A", "A.Next.0", "A.Next.1" }, widget.Actions.Select(static action => action.TriggerName));
        Assert.Equal(new[] { "GoTo", "JavaScript", "JavaScript" }, widget.Actions.Select(static action => action.ActionType));
        Assert.Equal("app.alert('one');", widget.JavaScript);

        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxJavaScripts = 1 } };
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfDocument.Open(source, options).Inspect());
        Assert.Equal(PdfReadLimitKind.JavaScripts, exception.Kind);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void Reader_TraversesSingleIndirectNextWidgetAction() {
        byte[] source = BuildWidgetActionGraphPdf(includeOpenAction: false, useSingleIndirectNext: true);

        PdfFormWidget widget = Assert.Single(Assert.Single(PdfDocument.Open(source).Inspect().FormFields).Widgets);

        Assert.Equal(new[] { "A", "A.Next" }, widget.Actions.Select(static action => action.TriggerName));
        Assert.Equal(new[] { "GoTo", "JavaScript" }, widget.Actions.Select(static action => action.ActionType));
        Assert.Equal("app.alert('one');", widget.JavaScript);
    }

    [Fact]
    public void Flatten_PrunesIndirectWidgetActionGraphButKeepsPublicRedactionAnalysisClean() {
        byte[] source = BuildWidgetActionGraphPdf(includeOpenAction: false);

        PdfDocument flattened = PdfDocument.Open(source).Forms.Flatten();
        IReadOnlyList<PdfSanitizationFinding> findings = PdfSanitizer.Analyze(flattened.ToBytes());

        Assert.Empty(flattened.Inspect().FormFields);
        Assert.DoesNotContain(findings, static finding => finding.Kind == PdfSanitizationFindingKind.ActiveAction);
        Assert.DoesNotContain("app.alert", Encoding.ASCII.GetString(flattened.ToBytes()), StringComparison.Ordinal);
    }

    [Fact]
    public void FormPreflightAndPlannerBothRejectUnrelatedCatalogOpenAction() {
        byte[] source = BuildWidgetActionGraphPdf(includeOpenAction: true);

        PdfDocumentPreflight preflight = PdfDocument.Open(source).Preflight();

        Assert.False(preflight.CanFillSimpleFormFields);
        Assert.False(preflight.CanFlattenSimpleFormFields);
        Assert.Contains(preflight.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields), message => message.Contains("open action", StringComparison.OrdinalIgnoreCase));
        Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Open(source).Forms.Fill(new Dictionary<string, string> { ["run"] = "updated" }));
        Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Open(source).Forms.Edit(edit => edit.Rename("run", "renamed")));

        byte[] nonWidgetAction = Encoding.ASCII.GetBytes(Encoding.ASCII.GetString(BuildWidgetActionGraphPdf(includeOpenAction: false))
            .Replace("/Subtype /Widget /FT /Tx", "/Subtype /Text /FT /Tx", StringComparison.Ordinal));
        Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Open(nonWidgetAction).Forms.Edit(edit => edit.Rename("run", "renamed")));
    }

    [Fact]
    public void WidgetJavaScript_RoundTripsUnicodeAndPdfDocSensitiveCharacters() {
        const string script = "app.alert('€ •');\f";
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Unicode widget action")).ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "unicode",
            Kind = PdfFormFieldCreationKind.PushButton,
            X = 72,
            Y = 600,
            Width = 100,
            Height = 24,
            Caption = "Run",
            JavaScript = script
        }));

        PdfFormWidgetAction action = Assert.Single(Assert.Single(Assert.Single(result.Fields).Widgets).Actions);
        Assert.Equal(script, action.JavaScript);
    }

    [Fact]
    public void Create_AccountsForExistingWidgetJavaScriptInAggregateLimits() {
        const string existingScript = "app.alert('existing');";
        const string newScript = "app.alert('new');";
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Existing widget action budget")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "existing",
            Kind = PdfFormFieldCreationKind.PushButton,
            X = 72,
            Y = 600,
            Width = 100,
            Height = 24,
            JavaScript = existingScript
        })).ToBytes();

        var countOptions = new PdfReadOptions { Limits = new PdfReadLimits { MaxJavaScripts = 1 } };
        PdfReadLimitException countException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Open(authored, countOptions).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
                Name = "new",
                Kind = PdfFormFieldCreationKind.PushButton,
                JavaScript = newScript
            })));
        Assert.Equal(PdfReadLimitKind.JavaScripts, countException.Kind);
        Assert.Equal(2, countException.Actual);

        long existingBytes = PdfJavaScriptStringEncoding.EncodeUnicode(existingScript, nameof(existingScript)).LongLength;
        long newBytes = PdfJavaScriptStringEncoding.EncodeUnicode(newScript, nameof(newScript)).LongLength;
        var byteOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxJavaScripts = 2,
                MaxTotalJavaScriptBytes = existingBytes + newBytes - 1L
            }
        };
        PdfReadLimitException byteException = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Open(authored, byteOptions).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
                Name = "new",
                Kind = PdfFormFieldCreationKind.PushButton,
                JavaScript = newScript
            })));
        Assert.Equal(PdfReadLimitKind.JavaScriptBytes, byteException.Kind);
        Assert.Equal(existingBytes + newBytes, byteException.Actual);
    }

    [Fact]
    public void ChoiceCreation_PreservesEmptyOptionsAndHonorsLegacyRawFlags() {
        const int comboFlag = 131072;
        const int editFlag = 262144;
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Choice compatibility")).ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Create(new PdfFormFieldCreateOptions {
                Name = "empty",
                Kind = PdfFormFieldCreationKind.Choice,
                X = 72,
                Y = 600,
                Width = 140,
                Height = 24
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "legacyCombo",
                Kind = PdfFormFieldCreationKind.Choice,
                X = 72,
                Y = 560,
                Width = 140,
                Height = 24,
                FieldFlags = comboFlag,
                ChoiceOptions = new[] { "One", "Two" },
                Value = "Two"
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "legacyEditableCombo",
                Kind = PdfFormFieldCreationKind.Choice,
                X = 72,
                Y = 520,
                Width = 140,
                Height = 24,
                FieldFlags = comboFlag | editFlag,
                ChoiceOptions = new[] { "One", "Two" },
                Value = "Custom"
            }));

        PdfFormField empty = result.Fields.Single(static field => field.Name == "empty");
        Assert.Empty(empty.Options);
        Assert.Equal(string.Empty, empty.Value);
        Assert.False(empty.IsCombo);
        Assert.True(result.Fields.Single(static field => field.Name == "legacyCombo").IsCombo);
        PdfFormField editable = result.Fields.Single(static field => field.Name == "legacyEditableCombo");
        Assert.True(editable.IsCombo);
        Assert.True(editable.IsEditableChoice);
        Assert.Equal("Custom", editable.Value);
    }

    [Fact]
    public void ChoiceCreation_PreservesAnExplicitlyUnselectedValueWhenOptionsExist() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Unselected choice")).ToBytes();

        PdfFormField field = Assert.Single(PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "country",
            Kind = PdfFormFieldCreationKind.Choice,
            ChoiceOptions = new[] { "Poland", "Germany" },
            Value = string.Empty
        })).Fields);

        Assert.Equal(string.Empty, field.Value);
        Assert.Equal(new[] { "Poland", "Germany" }, field.Options.Select(static option => option.ExportValue));
    }

    [Fact]
    public void Create_RejectsCombOutsideItsCompatibleTextFieldContract() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Comb validation")).ToBytes();

        Assert.Throws<ArgumentException>(() => PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "multiline",
            Kind = PdfFormFieldCreationKind.Text,
            Style = new PdfFormFieldStyle { IsComb = true, IsMultiline = true, MaxLength = 4 }
        })));
        Assert.Throws<ArgumentException>(() => PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "password",
            Kind = PdfFormFieldCreationKind.Text,
            Style = new PdfFormFieldStyle { IsComb = true, IsPassword = true, MaxLength = 4 }
        })));
        Assert.Throws<ArgumentException>(() => PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "file",
            Kind = PdfFormFieldCreationKind.Text,
            Style = new PdfFormFieldStyle { IsComb = true, IsFileSelect = true, MaxLength = 4 }
        })));
        Assert.Throws<ArgumentException>(() => PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "button",
            Kind = PdfFormFieldCreationKind.PushButton,
            Style = new PdfFormFieldStyle { IsComb = true, MaxLength = 4 }
        })));
    }

    [Theory]
    [InlineData(PdfFormFieldCreationKind.CheckBox, 32768)]
    [InlineData(PdfFormFieldCreationKind.CheckBox, 65536)]
    [InlineData(PdfFormFieldCreationKind.RadioButtonGroup, 65536)]
    [InlineData(PdfFormFieldCreationKind.PushButton, 32768)]
    public void Create_RejectsConflictingRawButtonKindFlags(PdfFormFieldCreationKind kind, int fieldFlags) {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Button flag validation")).ToBytes();

        Assert.Throws<ArgumentException>(() => PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "button",
            Kind = kind,
            Caption = kind == PdfFormFieldCreationKind.PushButton ? "Run" : null,
            FieldFlags = fieldFlags,
            ChoiceOptions = kind == PdfFormFieldCreationKind.RadioButtonGroup ? new[] { "One" } : Array.Empty<string>(),
            Value = kind == PdfFormFieldCreationKind.RadioButtonGroup ? "One" : string.Empty,
            Width = kind == PdfFormFieldCreationKind.RadioButtonGroup ? 120D : 100D,
            Height = 24D
        })));
    }

    [Fact]
    public void Reader_AccountsWidgetActionsEvenWhenWidgetGeometryIsUnreadable() {
        byte[] source = Encoding.ASCII.GetBytes(Encoding.ASCII.GetString(BuildWidgetActionGraphPdf(includeOpenAction: false))
            .Replace(" /Rect [20 20 160 48]", string.Empty, StringComparison.Ordinal));
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxJavaScripts = 1 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfDocument.Open(source, options).Inspect());

        Assert.Equal(PdfReadLimitKind.JavaScripts, exception.Kind);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void CreateReadback_HonorsLaterFlagEditsInsteadOfRequiringTheInitialPresentation() {
        const int comboFlag = 131072;
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Final flags")).ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Create(new PdfFormFieldCreateOptions {
                Name = "country",
                Kind = PdfFormFieldCreationKind.Choice,
                ChoiceOptions = new[] { "Poland", "Germany" },
                IsComboBox = true
            })
            .SetFlags("country", 0));

        PdfFormField field = Assert.Single(result.Fields);
        Assert.False(field.IsCombo);
        Assert.Equal(0, field.Flags);

        PdfAcroFormEditResult promoted = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Create(new PdfFormFieldCreateOptions {
                Name = "country",
                Kind = PdfFormFieldCreationKind.Choice,
                ChoiceOptions = new[] { "Poland", "Germany" }
            })
            .SetFlags("country", comboFlag));
        Assert.True(Assert.Single(promoted.Fields).IsCombo);
    }

    [Fact]
    public void ButtonCaptionsIgnoreTextOnlyPasswordAndMultilineStyleFlags() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Button captions")).ToBytes();

        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Create(new PdfFormFieldCreateOptions {
                Name = "calculate",
                Kind = PdfFormFieldCreationKind.PushButton,
                Caption = "Calculate",
                Style = new PdfFormFieldStyle { IsPassword = true, IsMultiline = true }
            })
            .Create(new PdfFormFieldCreateOptions {
                Name = "size",
                Kind = PdfFormFieldCreationKind.RadioButtonGroup,
                Y = 100,
                Width = 180,
                Height = 60,
                ChoiceOptions = new[] { "Small", "Large" },
                Style = new PdfFormFieldStyle { IsPassword = true, IsMultiline = true }
            })).ToBytes();

        string raw = PdfEncoding.Latin1GetString(authored);
        Assert.Contains("<43616C63756C617465> Tj", raw, StringComparison.Ordinal);
        Assert.Contains("<536D616C6C> Tj", raw, StringComparison.Ordinal);
        Assert.Contains("<4C61726765> Tj", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("••", raw, StringComparison.Ordinal);
    }

    private static byte[] BuildWidgetJavaScriptStreamPdf(string source) {
        byte[] script = PdfJavaScriptStringEncoding.EncodeUnicode(source, nameof(source));
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [6 0 R] >>\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Fields [6 0 R] >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (run) /Rect [20 20 120 44] /P 3 0 R /A << /S /JavaScript /JS 7 0 R >> >>\nendobj\n");
        WriteAscii(output, "7 0 obj\n<< /Length " + script.Length + " >>\nstream\n");
        output.Write(script, 0, script.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 8 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildWidgetActionGraphPdf(bool includeOpenAction, bool useSingleIndirectNext = false) {
        using var output = new MemoryStream();
        string openAction = includeOpenAction ? " /OpenAction 11 0 R" : string.Empty;
        string nextAction = useSingleIndirectNext ? "9 0 R" : "[9 0 R 10 0 R]";
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R" + openAction + " >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /Helv 12 0 R >> >> /Annots [6 0 R] >>\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Fields [6 0 R] /DA (/Helv 10 Tf 0 g) /DR << /Font << /Helv 12 0 R >> >> >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Tx /T (run) /V (before) /DA (/Helv 10 Tf 0 g) /Rect [20 20 160 48] /P 3 0 R /A 8 0 R >>\nendobj\n");
        WriteAscii(output, "8 0 obj\n<< /S /GoTo /D [3 0 R /Fit] /Next " + nextAction + " >>\nendobj\n");
        WriteAscii(output, "9 0 obj\n<< /S /JavaScript /JS (app.alert\\('one'\\);) >>\nendobj\n");
        WriteAscii(output, "10 0 obj\n<< /S /JavaScript /JS (app.alert\\('two'\\);) /Next 8 0 R >>\nendobj\n");
        WriteAscii(output, "11 0 obj\n<< /S /JavaScript /JS (app.alert\\('open'\\);) >>\nendobj\n");
        WriteAscii(output, "12 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 13 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
