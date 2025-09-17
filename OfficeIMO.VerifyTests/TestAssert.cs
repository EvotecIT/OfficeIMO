using System.Collections;
using System.Collections.Generic;

namespace OfficeIMO.VerifyTests;

internal static class TestAssert {
    public static T NotNull<T>(T? value) where T : class {
        global::Xunit.Assert.NotNull(value);
        return value!;
    }

    public static T IsType<T>(object? value) {
        return global::Xunit.Assert.IsType<T>(value);
    }

    public static void Single<T>(IEnumerable<T> collection) {
        global::Xunit.Assert.Single(collection);
    }

    public static void Single(IEnumerable collection) {
        global::Xunit.Assert.Single(collection);
    }

    public static void Empty<T>(IEnumerable<T> collection) {
        global::Xunit.Assert.Empty(collection);
    }

    public static void Empty(IEnumerable collection) {
        global::Xunit.Assert.Empty(collection);
    }
}
