using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace OfficeIMO.Build.ApiMap.Fluent;

/// <summary>
/// What a method body returns, read from its source declaration: the receiver itself (a mutation that
/// continues the chain), a new instance (a copy), something else, or nothing.
/// </summary>
internal static class BodyFacts {
    internal const string ReturnsReceiver = "returns-receiver";
    internal const string ReturnsNew = "returns-new";
    internal const string Mixed = "mixed";
    internal const string Delegates = "delegates";
    internal const string NoSource = "no-source";
    internal const string Void = "void";

    internal static string Semantics(IReadOnlyList<SyntaxNode> declarations, IMethodSymbol method) {
        if (method.ReturnsVoid) return Void;
        var methods = declarations.OfType<MethodDeclarationSyntax>().ToArray();
        if (methods.Length == 0) return NoSource;
        string receiver = method.IsExtensionMethod ? method.Parameters[0].Name : "this";
        bool self = false, fresh = false;
        foreach (var declaration in methods) {
            foreach (var expression in Returned(declaration)) {
                var value = Unwrap(expression);
                if (value is ThisExpressionSyntax || value is IdentifierNameSyntax identifier && identifier.Identifier.ValueText == receiver) self = true;
                else if (IsFresh(value)) fresh = true;
            }
        }
        return self && fresh ? Mixed : self ? ReturnsReceiver : fresh ? ReturnsNew : Delegates;
    }

    /// <summary>What a read-only property's getter hands out.</summary>
    internal enum GetterResult {
        /// <summary>A stored instance, or nothing the source shows to be new.</summary>
        Stored,
        /// <summary>A new wrapper over the receiver or a reference it holds (<c>new X(this)</c>, <c>new X(_options)</c>): a live view.</summary>
        View,
        /// <summary>A new object over receiver state whose constructor could be a copy constructor: syntax cannot tell a view from a copy.</summary>
        Wraps,
        /// <summary>A copy or a detached instance (<c>Clone()</c>, <c>with</c>, <c>new X { ... }</c>, <c>new X(_width)</c>): configuring it changes nothing.</summary>
        Copy
    }

    private static readonly HashSet<string> ValueTypes = new(StringComparer.Ordinal) {
        "bool", "byte", "sbyte", "short", "ushort", "int", "uint", "long", "ulong", "float", "double", "decimal", "char", "string", "object" };

    /// <summary>
    /// Classifies what a property getter returns; the least configurable result of any return wins. <paramref name="type"/>
    /// is the property type's simple name and <paramref name="constructors"/> the parameter types of its source constructors.
    /// </summary>
    internal static GetterResult Getter(IReadOnlyList<SyntaxNode> declarations, string type, IReadOnlyList<string[]> constructors) {
        var result = GetterResult.Stored;
        foreach (var property in declarations.OfType<PropertyDeclarationSyntax>()) {
            var getter = property.AccessorList?.Accessors.FirstOrDefault(accessor => accessor.Keyword.ValueText == "get");
            IEnumerable<ExpressionSyntax> expressions = property.ExpressionBody != null ? [property.ExpressionBody.Expression]
                : getter?.ExpressionBody != null ? [getter.ExpressionBody.Expression]
                : getter?.Body != null ? OwnedReturns(getter.Body) : [];
            foreach (var expression in expressions) {
                var value = Unwrap(expression);
                var kind = value is BaseObjectCreationExpressionSyntax creation ? Created(creation, type, constructors)
                    : IsFresh(value) ? GetterResult.Copy : GetterResult.Stored;
                if (kind > result) result = kind;
            }
        }
        return result;
    }

    private static GetterResult Created(BaseObjectCreationExpressionSyntax creation, string type, IReadOnlyList<string[]> constructors) {
        var arguments = creation.ArgumentList?.Arguments ?? default;
        if (arguments.Any(argument => Unwrap(argument.Expression) is ThisExpressionSyntax)) return GetterResult.View;
        bool reads = arguments.Any(argument => HoldsReference(Unwrap(argument.Expression)));
        if (!reads) return GetterResult.Copy;
        bool known = creation is ImplicitObjectCreationExpressionSyntax
            || creation is ObjectCreationExpressionSyntax { Type: var created } && SimpleName(created) == type;
        var candidates = known ? constructors.Where(parameters => parameters.Length >= arguments.Count).ToArray() : [];
        if (candidates.Length == 0 || candidates.Any(parameters => parameters.Contains(type, StringComparer.Ordinal))) return GetterResult.Wraps;
        // A reference the receiver holds, passed where every candidate constructor takes a reference: the result wraps it.
        for (int i = 0; i < arguments.Count; i++) {
            if (arguments[i].NameColon != null || !HoldsReference(Unwrap(arguments[i].Expression))) continue;
            if (candidates.All(parameters => !ValueTypes.Contains(parameters[i]))) return GetterResult.View;
        }
        return GetterResult.Copy;
    }

    /// <summary>A field, property or local the receiver holds (<c>_options</c>, <c>Element</c>, <c>this.Element</c>), as opposed to a literal, a constant or a value read out of it.</summary>
    private static bool HoldsReference(ExpressionSyntax argument) =>
        argument is IdentifierNameSyntax || argument is MemberAccessExpressionSyntax { Expression: ThisExpressionSyntax };

    private static string SimpleName(TypeSyntax type) => type switch {
        QualifiedNameSyntax qualified => SimpleName(qualified.Right),
        GenericNameSyntax generic => generic.Identifier.ValueText,
        IdentifierNameSyntax identifier => identifier.Identifier.ValueText,
        NullableTypeSyntax nullable => SimpleName(nullable.ElementType),
        _ => type.ToString()
    };

    /// <summary>
    /// True when a getter hands out a stored instance (an auto-property, a field or a member of one), false when it computes
    /// the value each time (<c>ToList()</c>, <c>new List&lt;T&gt;(...)</c>, a query), so adding to the result would be lost.
    /// </summary>
    internal static bool GetterStored(IReadOnlyList<SyntaxNode> declarations) {
        foreach (var property in declarations.OfType<PropertyDeclarationSyntax>()) {
            var getter = property.AccessorList?.Accessors.FirstOrDefault(accessor => accessor.Keyword.ValueText == "get");
            IEnumerable<ExpressionSyntax> expressions = property.ExpressionBody != null ? [property.ExpressionBody.Expression]
                : getter?.ExpressionBody != null ? [getter.ExpressionBody.Expression]
                : getter?.Body != null ? OwnedReturns(getter.Body) : [];
            if (expressions.Any(expression => Unwrap(expression) is var value && (IsFresh(value) || value is InvocationExpressionSyntax
                    or CollectionExpressionSyntax or ArrayCreationExpressionSyntax or ImplicitArrayCreationExpressionSyntax or ConditionalExpressionSyntax or SwitchExpressionSyntax)))
                return false;
        }
        return true;
    }

    private static bool IsFresh(ExpressionSyntax value) =>
        value is BaseObjectCreationExpressionSyntax or WithExpressionSyntax
        || value is InvocationExpressionSyntax invocation && InvokedName(invocation) is { } name
            && (name.Contains("Clone", StringComparison.Ordinal) || name.Contains("Copy", StringComparison.Ordinal) || name.StartsWith("With", StringComparison.Ordinal));

    private static IEnumerable<ExpressionSyntax> Returned(MethodDeclarationSyntax declaration) {
        if (declaration.ExpressionBody != null) return [declaration.ExpressionBody.Expression];
        return declaration.Body == null ? [] : OwnedReturns(declaration.Body);
    }

    /// <summary>Returned expressions of a body, leaving out returns that belong to lambdas and local functions inside it.</summary>
    private static IEnumerable<ExpressionSyntax> OwnedReturns(BlockSyntax body) => body
        .DescendantNodes(node => node is not (AnonymousFunctionExpressionSyntax or LocalFunctionStatementSyntax))
        .OfType<ReturnStatementSyntax>()
        .Select(statement => statement.Expression)
        .OfType<ExpressionSyntax>();

    private static ExpressionSyntax Unwrap(ExpressionSyntax expression) {
        while (true) {
            switch (expression) {
                case ParenthesizedExpressionSyntax parenthesized: expression = parenthesized.Expression; continue;
                case CastExpressionSyntax cast: expression = cast.Expression; continue;
                case BinaryExpressionSyntax { RawKind: (int)Microsoft.CodeAnalysis.CSharp.SyntaxKind.AsExpression } binary: expression = binary.Left; continue;
                case PostfixUnaryExpressionSyntax { RawKind: (int)Microsoft.CodeAnalysis.CSharp.SyntaxKind.SuppressNullableWarningExpression } suppressed: expression = suppressed.Operand; continue;
                default: return expression;
            }
        }
    }

    private static string? InvokedName(InvocationExpressionSyntax invocation) => invocation.Expression switch {
        MemberAccessExpressionSyntax member => member.Name.Identifier.ValueText,
        IdentifierNameSyntax identifier => identifier.Identifier.ValueText,
        GenericNameSyntax generic => generic.Identifier.ValueText,
        _ => null
    };
}
