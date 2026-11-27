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

    /// <summary>True when a property getter hands out a new instance or a copy, so configuring the result changes nothing.</summary>
    internal static bool GetterReturnsCopy(IReadOnlyList<SyntaxNode> declarations) {
        foreach (var property in declarations.OfType<PropertyDeclarationSyntax>()) {
            var getter = property.AccessorList?.Accessors.FirstOrDefault(accessor => accessor.Keyword.ValueText == "get");
            var expressions = property.ExpressionBody != null ? [property.ExpressionBody.Expression]
                : getter?.ExpressionBody != null ? [getter.ExpressionBody.Expression]
                : getter?.Body?.DescendantNodes().OfType<ReturnStatementSyntax>().Select(statement => statement.Expression).OfType<ExpressionSyntax>().ToArray() ?? [];
            if (expressions.Any(expression => IsFresh(Unwrap(expression)))) return true;
        }
        return false;
    }

    private static bool IsFresh(ExpressionSyntax value) =>
        value is BaseObjectCreationExpressionSyntax or WithExpressionSyntax
        || value is InvocationExpressionSyntax invocation && InvokedName(invocation) is { } name
            && (name.Contains("Clone", StringComparison.Ordinal) || name.Contains("Copy", StringComparison.Ordinal) || name.StartsWith("With", StringComparison.Ordinal));

    private static IEnumerable<ExpressionSyntax> Returned(MethodDeclarationSyntax declaration) {
        if (declaration.ExpressionBody != null) {
            yield return declaration.ExpressionBody.Expression;
            yield break;
        }
        if (declaration.Body == null) yield break;
        foreach (var statement in declaration.Body.DescendantNodes().OfType<ReturnStatementSyntax>()) {
            if (statement.Expression == null) continue;
            // Returns inside lambdas and local functions belong to them, not to the method.
            if (statement.Ancestors().TakeWhile(node => node != declaration).Any(node => node is AnonymousFunctionExpressionSyntax or LocalFunctionStatementSyntax)) continue;
            yield return statement.Expression;
        }
    }

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
