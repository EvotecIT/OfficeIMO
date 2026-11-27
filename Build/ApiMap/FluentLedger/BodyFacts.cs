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
                else if (value is BaseObjectCreationExpressionSyntax or WithExpressionSyntax) fresh = true;
                else if (value is InvocationExpressionSyntax invocation && InvokedName(invocation) is { } name
                    && (name.Contains("Clone", StringComparison.Ordinal) || name.Contains("Copy", StringComparison.Ordinal) || name.StartsWith("With", StringComparison.Ordinal))) fresh = true;
            }
        }
        return self && fresh ? Mixed : self ? ReturnsReceiver : fresh ? ReturnsNew : Delegates;
    }

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
