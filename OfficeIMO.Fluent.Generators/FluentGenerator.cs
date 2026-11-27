using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;

namespace OfficeIMO.Fluent.Generators {
    /// <summary>
    /// Writes the property-named configuration extensions of every type marked <c>[OfficeIMO.Fluent.Fluent]</c>:
    /// <list type="bullet">
    /// <item><c>X(value)</c> for a public settable property, <c>X()</c> / <c>X(false)</c> for a <c>bool</c> or <c>bool?</c>;</item>
    /// <item><c>X(...)</c> calling the instance method marked <c>[FluentMutator("X")]</c>, the canonical write path for
    /// state without a public setter and for value-type handles;</item>
    /// <item><c>X(o =&gt; ...)</c> for an option object held by a property, configuring it in place;</item>
    /// <item><c>Configure(x =&gt; ...)</c> when the type has delegate-valued properties, so they are assigned, not invoked.</item>
    /// </list>
    /// Each method sets one property or calls one mutator and returns the receiver. A class that can be derived from gets
    /// methods generic over <c>TSelf</c>, so a chain keeps the derived type. No reflection is generated.
    /// </summary>
    [Generator(LanguageNames.CSharp)]
    public sealed class FluentGenerator : IIncrementalGenerator {
        private static readonly SymbolDisplayFormat TypeFormat = SymbolDisplayFormat.FullyQualifiedFormat
            .AddMiscellaneousOptions(SymbolDisplayMiscellaneousOptions.IncludeNullableReferenceTypeModifier);

        /// <inheritdoc />
        public void Initialize(IncrementalGeneratorInitializationContext context) {
            context.RegisterPostInitializationOutput(output => output.AddSource("OfficeIMO.Fluent.Attributes.g.cs", SourceText.From(FluentAttributesSource.Source, Encoding.UTF8)));
            var models = context.SyntaxProvider.ForAttributeWithMetadataName(
                FluentAttributesSource.Fluent,
                (node, _) => node is Microsoft.CodeAnalysis.CSharp.Syntax.TypeDeclarationSyntax,
                (attributeContext, _) => Build((INamedTypeSymbol)attributeContext.TargetSymbol, attributeContext.SemanticModel.Compilation));
            context.RegisterSourceOutput(models, (output, model) => {
                foreach (var diagnostic in model.Diagnostics) output.ReportDiagnostic(diagnostic.Create());
                if (model.Source != null) output.AddSource(model.HintName, SourceText.From(model.Source, Encoding.UTF8));
            });
        }

        private static TypeModel Build(INamedTypeSymbol type, Compilation compilation) {
            var diagnostics = new List<DiagnosticInfo>();
            string hint = type.ToDisplayString().Replace('<', '_').Replace('>', '_') + ".Fluent.g.cs";
            if (type.ContainingType != null || type.IsGenericType || type.IsStatic || type.TypeKind is not (TypeKind.Class or TypeKind.Struct)) {
                diagnostics.Add(DiagnosticInfo.From(FluentDiagnostics.UnsupportedType, type, type.Name));
                return new TypeModel(hint, null, diagnostics);
            }
            bool nullable = compilation is CSharpCompilation csharp && csharp.LanguageVersion >= LanguageVersion.CSharp8;
            var writer = new Writer(type, nullable);
            var mutated = new HashSet<string>(StringComparer.Ordinal);

            foreach (var method in type.GetMembers().OfType<IMethodSymbol>()) {
                foreach (var attribute in method.GetAttributes().Where(a => a.AttributeClass?.ToDisplayString() == FluentAttributesSource.Mutator)) {
                    if (attribute.ConstructorArguments.Length != 1 || attribute.ConstructorArguments[0].Value is not string name) continue;
                    if (method.IsStatic || method.DeclaredAccessibility is Accessibility.Private or Accessibility.Protected or Accessibility.ProtectedAndInternal
                        || method.IsGenericMethod || method.Parameters.Any(p => p.RefKind != RefKind.None)) {
                        diagnostics.Add(DiagnosticInfo.From(FluentDiagnostics.InvalidMutator, method, method.Name));
                        continue;
                    }
                    mutated.Add(name);
                    writer.Mutator(name, method, diagnostics);
                }
            }

            bool delegates = false;
            foreach (var property in type.GetMembers().OfType<IPropertySymbol>()) {
                if (property.IsStatic || property.IsIndexer || property.DeclaredAccessibility != Accessibility.Public || HasAttribute(property, FluentAttributesSource.NoFluent)) continue;
                string name = FluentName(property);
                if (mutated.Contains(name)) continue; // The canonical mutator owns this name.
                bool settable = property.SetMethod is { DeclaredAccessibility: Accessibility.Public, IsInitOnly: false };
                if (settable && property.Type.TypeKind == TypeKind.Delegate) { delegates = true; continue; }
                if (settable && type.TypeKind == TypeKind.Struct) {
                    diagnostics.Add(DiagnosticInfo.From(FluentDiagnostics.ValueTypeSetter, property, property.Name, type.Name));
                    continue;
                }
                if (settable) writer.Setter(name, property, diagnostics);
                if (IsOptionObject(property.Type) && property.GetMethod?.DeclaredAccessibility == Accessibility.Public && type.TypeKind != TypeKind.Struct)
                    writer.Configure(name, property, settable, diagnostics);
            }
            if (delegates) writer.ConfigureSelf(diagnostics);
            return new TypeModel(hint, writer.Methods == 0 ? null : writer.Finish(), diagnostics);
        }

        private static string FluentName(IPropertySymbol property) {
            var renamed = property.GetAttributes().FirstOrDefault(a => a.AttributeClass?.ToDisplayString() == FluentAttributesSource.Fluent);
            if (renamed is { ConstructorArguments.Length: 1 } && renamed.ConstructorArguments[0].Value is string name) return name;
            string original = property.Name;
            if (IsBoolean(property.Type) && original.Length > 2 && original.StartsWith("Is", StringComparison.Ordinal) && char.IsUpper(original[2])) return original.Substring(2);
            return original;
        }

        private static bool HasAttribute(ISymbol symbol, string name) => symbol.GetAttributes().Any(a => a.AttributeClass?.ToDisplayString() == name);

        internal static bool IsBoolean(ITypeSymbol type) =>
            type.SpecialType == SpecialType.System_Boolean
            || type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable && nullable.TypeArguments[0].SpecialType == SpecialType.System_Boolean;

        private static bool IsOptionObject(ITypeSymbol type) =>
            type is INamedTypeSymbol { TypeKind: TypeKind.Class, SpecialType: SpecialType.None, IsAbstract: false } named
            && !named.AllInterfaces.Any(i => i.ToDisplayString() == "System.Collections.IEnumerable")
            && (HasAttribute(named, FluentAttributesSource.Fluent)
                || named.GetMembers().OfType<IPropertySymbol>().Any(p => !p.IsStatic && p.SetMethod is { DeclaredAccessibility: Accessibility.Public, IsInitOnly: false }));

        /// <summary>Emits one static partial <c>&lt;Type&gt;Fluent</c> class in the receiver's namespace.</summary>
        private sealed class Writer {
            private readonly INamedTypeSymbol _type;
            private readonly bool _nullable;
            private readonly bool _open;
            private readonly string _typeName;
            private readonly StringBuilder _body = new StringBuilder();
            private readonly HashSet<string> _signatures = new HashSet<string>(StringComparer.Ordinal);

            internal int Methods { get; private set; }

            internal Writer(INamedTypeSymbol type, bool nullable) {
                _type = type;
                _nullable = nullable;
                _open = type.TypeKind == TypeKind.Class && !type.IsSealed;
                _typeName = type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
            }

            private string Receiver => _open ? "TSelf" : _typeName;

            internal void Mutator(string name, IMethodSymbol method, List<DiagnosticInfo> diagnostics) {
                var parameters = method.Parameters.Select((p, i) => Parameter(p, defaultTrue: method.Parameters.Length == 1 && IsBoolean(p.Type) && !p.HasExplicitDefaultValue)).ToArray();
                string call = "receiver." + method.Name + "(" + string.Join(", ", method.Parameters.Select(p => "@" + p.Name)) + ");";
                Emit(name, method, parameters, method.Parameters.Select(p => p.Type).ToArray(),
                    $"Calls <see cref=\"{Cref(method)}\"/>, the canonical write path, and returns the same receiver.", call, diagnostics);
            }

            internal void Setter(string name, IPropertySymbol property, List<DiagnosticInfo> diagnostics) {
                string type = property.Type.ToDisplayString(_nullable ? TypeFormat : SymbolDisplayFormat.FullyQualifiedFormat);
                string parameter = IsBoolean(property.Type) ? type + " value = true" : type + " value";
                Emit(name, property, new[] { new ParameterText(parameter, "value", "The value to assign.") }, new[] { property.Type },
                    $"Sets <see cref=\"{Cref(property)}\"/> and returns the same receiver.", $"receiver.{property.Name} = value;", diagnostics);
            }

            internal void Configure(string name, IPropertySymbol property, bool settable, List<DiagnosticInfo> diagnostics) {
                var option = (INamedTypeSymbol)property.Type.WithNullableAnnotation(NullableAnnotation.NotAnnotated);
                string optionType = option.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                bool create = settable && option.InstanceConstructors.Any(c => c.Parameters.Length == 0 && c.DeclaredAccessibility == Accessibility.Public);
                string fetch = create
                    ? $"var value = receiver.{property.Name};\n            if (value == null) {{\n                value = new {optionType}();\n                receiver.{property.Name} = value;\n            }}"
                    : $"var value = receiver.{property.Name};\n            if (value == null) throw new global::System.InvalidOperationException(\"{_type.Name}.{property.Name} is not set.\");";
                Emit(name, property, new[] { new ParameterText($"global::System.Action<{optionType}> configure", "configure", "Configures the option object in place.") },
                    Array.Empty<ITypeSymbol>(),
                    $"Configures <see cref=\"{Cref(property)}\"/> in place{(create ? ", creating it when unset," : "")} and returns the same receiver.",
                    "if (configure == null) throw new global::System.ArgumentNullException(nameof(configure));\n            " + fetch + "\n            configure(value);", diagnostics,
                    signatureKey: name + "(Action<" + option.ToDisplayString() + ">)");
            }

            internal void ConfigureSelf(List<DiagnosticInfo> diagnostics) {
                Emit("Configure", _type, new[] { new ParameterText($"global::System.Action<{Receiver}> configure", "configure", "Assigns values such as delegate-valued properties.") },
                    Array.Empty<ITypeSymbol>(), "Runs a callback on the receiver, for assignments such as delegate-valued properties, and returns it.",
                    "if (configure == null) throw new global::System.ArgumentNullException(nameof(configure));\n            configure(receiver);", diagnostics, signatureKey: "Configure(Action<self>)");
            }

            private void Emit(string name, ISymbol source, IReadOnlyList<ParameterText> parameters, IReadOnlyList<ITypeSymbol> parameterTypes, string summary, string body,
                List<DiagnosticInfo> diagnostics, string? signatureKey = null) {
                if (!SyntaxFacts.IsValidIdentifier(name)) {
                    diagnostics.Add(DiagnosticInfo.From(FluentDiagnostics.Collision, source, name, "a name that is not an identifier"));
                    return;
                }
                string key = signatureKey ?? name + "(" + string.Join(",", parameterTypes.Select(t => t.WithNullableAnnotation(NullableAnnotation.NotAnnotated).ToDisplayString())) + ")";
                if (!_signatures.Add(key)) {
                    diagnostics.Add(DiagnosticInfo.From(FluentDiagnostics.Duplicate, source, key));
                    return;
                }
                if (ExistingMethod(name, parameters.Count) is { } existing) {
                    diagnostics.Add(DiagnosticInfo.From(FluentDiagnostics.Collision, source, name, existing));
                    return;
                }
                Methods++;
                string typeParameter = _open ? "<TSelf>" : "";
                string constraint = _open ? $" where TSelf : {_typeName}" : "";
                string guard = _type.TypeKind == TypeKind.Struct ? "" : "            if (receiver == null) throw new global::System.ArgumentNullException(nameof(receiver));\n";
                _body.Append("        /// <summary>").Append(summary).Append("</summary>\n");
                if (_open) _body.Append("        /// <typeparam name=\"TSelf\">The receiver's own type, kept through the chain.</typeparam>\n");
                _body.Append("        /// <param name=\"receiver\">The object to configure.</param>\n");
                foreach (var parameter in parameters) _body.Append("        /// <param name=\"").Append(parameter.Name).Append("\">").Append(parameter.Summary).Append("</param>\n");
                _body.Append("        /// <returns>The same receiver.</returns>\n");
                _body.Append("        public static ").Append(Receiver).Append(' ').Append(name).Append(typeParameter).Append("(this ").Append(Receiver).Append(" receiver");
                foreach (var parameter in parameters) _body.Append(", ").Append(parameter.Text);
                _body.Append(')').Append(constraint).Append(" {\n").Append(guard).Append("            ").Append(body).Append("\n            return receiver;\n        }\n\n");
            }

            /// <summary>An instance method of the same name and arity would win over, or clash with, the generated extension.</summary>
            private string? ExistingMethod(string name, int parameters) {
                for (var type = _type; type != null; type = type.BaseType) {
                    foreach (var method in type.GetMembers(name).OfType<IMethodSymbol>()) {
                        if (method.DeclaredAccessibility == Accessibility.Private && !SymbolEqualityComparer.Default.Equals(type, _type)) continue;
                        if (method.Parameters.Length == parameters || method.Parameters.Count(p => !p.IsOptional) <= parameters && parameters <= method.Parameters.Length)
                            return "method " + type.Name + "." + method.Name;
                    }
                }
                return null;
            }

            internal string Finish() {
                string ns = _type.ContainingNamespace.IsGlobalNamespace ? "" : _type.ContainingNamespace.ToDisplayString();
                string accessibility = _type.DeclaredAccessibility == Accessibility.Public ? "public" : "internal";
                var text = new StringBuilder("// <auto-generated/>\n").Append(_nullable ? "#nullable enable\n\n" : "\n");
                if (ns.Length > 0) text.Append("namespace ").Append(ns).Append(" {\n");
                text.Append("    /// <summary>Property-named fluent configuration for <see cref=\"").Append(_typeName).Append("\"/>.</summary>\n");
                text.Append("    ").Append(accessibility).Append(" static partial class ").Append(_type.Name).Append("Fluent {\n");
                text.Append(_body.ToString().TrimEnd('\n')).Append('\n');
                text.Append("    }\n");
                if (ns.Length > 0) text.Append("}\n");
                return text.ToString();
            }

            private ParameterText Parameter(IParameterSymbol parameter, bool defaultTrue) {
                var text = new StringBuilder();
                if (parameter.IsParams) text.Append("params ");
                text.Append(parameter.Type.ToDisplayString(_nullable ? TypeFormat : SymbolDisplayFormat.FullyQualifiedFormat)).Append(" @").Append(parameter.Name);
                if (parameter.HasExplicitDefaultValue) text.Append(" = ").Append(DefaultValue(parameter));
                else if (defaultTrue) text.Append(" = true");
                return new ParameterText(text.ToString(), parameter.Name, "Passed to the mutator.");
            }

            private static string DefaultValue(IParameterSymbol parameter) {
                object? value = parameter.ExplicitDefaultValue;
                var type = parameter.Type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable ? nullable.TypeArguments[0] : parameter.Type;
                if (value == null) return type.IsValueType && parameter.Type.OriginalDefinition.SpecialType != SpecialType.System_Nullable_T ? "default" : "null";
                if (type.TypeKind == TypeKind.Enum) return "(" + type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat) + ")(" + SymbolDisplay.FormatPrimitive(value, false, false) + ")";
                string literal = SymbolDisplay.FormatPrimitive(value, true, false);
                return value switch {
                    float => literal + "F",
                    double => literal.Contains(".") || literal.Contains("E") ? literal : literal + "D",
                    decimal => literal + "M",
                    long => literal + "L",
                    ulong => literal + "UL",
                    uint => literal + "U",
                    _ => literal
                };
            }

            private static readonly SymbolDisplayFormat CrefFormat = SymbolDisplayFormat.CSharpErrorMessageFormat
                .RemoveMiscellaneousOptions(SymbolDisplayMiscellaneousOptions.IncludeNullableReferenceTypeModifier);

            private static string Cref(ISymbol symbol) => symbol.ToDisplayString(CrefFormat).Replace('<', '{').Replace('>', '}');
        }

        private sealed class ParameterText {
            internal ParameterText(string text, string name, string summary) { Text = text; Name = name; Summary = summary; }
            internal string Text { get; }
            internal string Name { get; }
            internal string Summary { get; }
        }
    }
}
