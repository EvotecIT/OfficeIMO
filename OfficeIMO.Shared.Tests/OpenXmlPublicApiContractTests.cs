using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace OfficeIMO.Tests;

public class OpenXmlPublicApiContractTests {
    private static readonly string[] RequiredOpenXmlBackedAssemblies = {
        "OfficeIMO.Word",
        "OfficeIMO.Word.GoogleDocs",
        "OfficeIMO.Word.Html",
        "OfficeIMO.Word.Markdown",
        "OfficeIMO.Word.OpenDocument",
        "OfficeIMO.Word.Pdf",
        "OfficeIMO.Word.Rtf",
        "OfficeIMO.PowerPoint",
        "OfficeIMO.PowerPoint.GoogleSlides",
        "OfficeIMO.PowerPoint.Html",
        "OfficeIMO.PowerPoint.OpenDocument",
        "OfficeIMO.PowerPoint.Pdf",
        "OfficeIMO.Excel",
        "OfficeIMO.Excel.Csv",
        "OfficeIMO.Excel.GoogleSheets",
        "OfficeIMO.Excel.Html",
        "OfficeIMO.Excel.OpenDocument",
        "OfficeIMO.Excel.Pdf",
        "OfficeIMO.Markup.Word",
        "OfficeIMO.Markup.PowerPoint",
        "OfficeIMO.Markup.Excel",
        "OfficeIMO.Reader.Word",
        "OfficeIMO.Reader.PowerPoint",
        "OfficeIMO.Reader.Excel"
    };

    [Fact]
    public void PublicApisDoNotExposeOpenXmlGeneratedValueStructs() {
        var leaks = new List<string>();
        string[] assemblyNames = Directory.EnumerateFiles(AppContext.BaseDirectory, "OfficeIMO*.dll")
            .Select(path => AssemblyName.GetAssemblyName(path).Name)
            .Where(name => name != null && !name.EndsWith(".Tests", StringComparison.Ordinal))
            .Cast<string>()
            .Distinct(StringComparer.Ordinal)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();

        foreach (string requiredAssembly in RequiredOpenXmlBackedAssemblies) {
            Assert.Contains(requiredAssembly, assemblyNames);
        }

        foreach (string assemblyName in assemblyNames) {
            Assembly assembly = Assembly.Load(new AssemblyName(assemblyName));
            foreach (Type type in assembly.GetExportedTypes()) {
                InspectTypeShape(type.BaseType, $"{type.FullName} base type", leaks);
                foreach (Type implementedInterface in type.GetInterfaces()) {
                    InspectTypeShape(implementedInterface, $"{type.FullName} interface", leaks);
                }

                foreach (ConstructorInfo constructor in type.GetConstructors(BindingFlags.Public | BindingFlags.Instance)) {
                    InspectParameters(constructor.GetParameters(), $"{type.FullName}.{constructor.Name}", leaks);
                }

                foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)) {
                    string member = $"{type.FullName}.{method.Name}";
                    InspectTypeShape(method.ReturnType, $"{member} return type", leaks);
                    InspectParameters(method.GetParameters(), member, leaks);
                    foreach (Type genericArgument in method.GetGenericArguments()) {
                        foreach (Type constraint in genericArgument.GetGenericParameterConstraints()) {
                            InspectTypeShape(constraint, $"{member} generic constraint", leaks);
                        }
                    }
                }

                foreach (PropertyInfo property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)) {
                    string member = $"{type.FullName}.{property.Name}";
                    InspectTypeShape(property.PropertyType, $"{member} property type", leaks);
                    InspectParameters(property.GetIndexParameters(), member, leaks);
                }

                foreach (FieldInfo field in type.GetFields(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)) {
                    InspectTypeShape(field.FieldType, $"{type.FullName}.{field.Name} field type", leaks);
                }

                foreach (EventInfo eventInfo in type.GetEvents(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)) {
                    InspectTypeShape(eventInfo.EventHandlerType, $"{type.FullName}.{eventInfo.Name} event type", leaks);
                }
            }
        }

        Assert.True(leaks.Count == 0,
            "OfficeIMO public APIs expose OpenXML SDK generated value structs:" + Environment.NewLine +
            string.Join(Environment.NewLine, leaks.OrderBy(value => value, StringComparer.Ordinal)));
    }

    [Fact]
    public void OfficeEnumsRoundTripEveryOpenXmlToken() {
        string[] coreAssemblies = { "OfficeIMO.Word", "OfficeIMO.PowerPoint", "OfficeIMO.Excel" };
        var testedEnums = new List<string>();

        foreach (string assemblyName in coreAssemblies) {
            Assembly assembly = Assembly.Load(new AssemblyName(assemblyName));
            MethodInfo[] extensionMethods = assembly.GetTypes()
                .SelectMany(type => type.GetMethods(BindingFlags.Static | BindingFlags.NonPublic))
                .ToArray();
            MethodInfo[] forwardMappings = extensionMethods
                .Where(method => method.Name == "ToOpenXml" && method.GetParameters().Length == 1)
                .Where(method => method.GetParameters()[0].ParameterType.IsEnum)
                .Where(method => IsOpenXmlGeneratedValueStruct(method.ReturnType))
                .ToArray();

            foreach (MethodInfo forward in forwardMappings) {
                Type officeEnum = forward.GetParameters()[0].ParameterType;
                Type openXmlType = forward.ReturnType;
                MethodInfo? reverse = extensionMethods.SingleOrDefault(method =>
                    (method.Name == "ToOfficeEnum" || method.Name == "ToOfficeIMO") &&
                    method.ReturnType == officeEnum &&
                    method.GetParameters().Length == 1 &&
                    method.GetParameters()[0].ParameterType == openXmlType);
                Assert.True(reverse != null,
                    $"{assemblyName}.{officeEnum.Name} does not have a reverse OpenXML mapping.");

                Array values = Enum.GetValues(officeEnum);
                var serializedValues = new HashSet<object>();
                foreach (object value in values) {
                    object? serialized = forward.Invoke(null, new[] { value });
                    Assert.NotNull(serialized);
                    Assert.True(serializedValues.Add(serialized!),
                        $"{assemblyName}.{officeEnum.Name} maps more than one enum member to {serialized}.");
                    object? roundTrip = reverse!.Invoke(null, new[] { serialized });
                    Assert.Equal(value, roundTrip);
                }

                testedEnums.Add($"{assemblyName}.{officeEnum.Name}");
            }
        }

        Assert.True(testedEnums.Count >= 70,
            $"Expected the generated and manual Word, PowerPoint, and Excel mappings; found {testedEnums.Count}.");
    }

    private static void InspectParameters(IEnumerable<ParameterInfo> parameters, string member, ICollection<string> leaks) {
        foreach (ParameterInfo parameter in parameters) {
            InspectTypeShape(parameter.ParameterType, $"{member} parameter '{parameter.Name}'", leaks);
        }
    }

    private static void InspectTypeShape(Type? type, string location, ICollection<string> leaks) {
        if (type is null || type.IsGenericParameter) return;

        if (IsOpenXmlGeneratedValueStruct(type)) {
            leaks.Add($"{location}: {type.FullName}");
            return;
        }

        if (type.HasElementType) {
            InspectTypeShape(type.GetElementType(), location, leaks);
        }

        if (type.IsGenericType) {
            foreach (Type argument in type.GetGenericArguments()) {
                InspectTypeShape(argument, location, leaks);
            }
        }
    }

    private static bool IsOpenXmlGeneratedValueStruct(Type type) =>
        type.IsValueType &&
        !type.IsEnum &&
        type.GetInterfaces().Any(implementedInterface =>
            implementedInterface.FullName == "DocumentFormat.OpenXml.IEnumValue");
}
