namespace OfficeIMO.Workflows;

/// <summary>School and college grading systems with ONIX-defined ordered grade codes.</summary>
public enum BookOnixGradeSystem {
    /// <summary>United States; list 30 code 11, values from list 77.</summary>
    UnitedStates,
    /// <summary>Canada excluding Québec; list 30 code 26, values from list 77.</summary>
    CanadaExcludingQuebec,
    /// <summary>China; list 30 code 29, values from list 227.</summary>
    China
}

/// <summary>
/// Ordered grade codes shared by ONIX lists 77 and 227. Meanings and typical ages
/// depend on the selected grading system; grades 13–17 denote tertiary levels.
/// </summary>
public enum BookOnixGrade {
    /// <summary>Preschool (P).</summary>
    Preschool = -1,
    /// <summary>Kindergarten (K).</summary>
    Kindergarten = 0,
    /// <summary>Grade code 1; interpret within the selected grading system.</summary>
    Grade1 = 1,
    /// <summary>Grade code 2; interpret within the selected grading system.</summary>
    Grade2 = 2,
    /// <summary>Grade code 3; interpret within the selected grading system.</summary>
    Grade3 = 3,
    /// <summary>Grade code 4; interpret within the selected grading system.</summary>
    Grade4 = 4,
    /// <summary>Grade code 5; interpret within the selected grading system.</summary>
    Grade5 = 5,
    /// <summary>Grade code 6; interpret within the selected grading system.</summary>
    Grade6 = 6,
    /// <summary>Grade code 7; interpret within the selected grading system.</summary>
    Grade7 = 7,
    /// <summary>Grade code 8; interpret within the selected grading system.</summary>
    Grade8 = 8,
    /// <summary>Grade code 9; interpret within the selected grading system.</summary>
    Grade9 = 9,
    /// <summary>Grade code 10; interpret within the selected grading system.</summary>
    Grade10 = 10,
    /// <summary>Grade code 11; interpret within the selected grading system.</summary>
    Grade11 = 11,
    /// <summary>Grade code 12; interpret within the selected grading system.</summary>
    Grade12 = 12,
    /// <summary>Grade code 13; interpret within the selected grading system.</summary>
    Grade13 = 13,
    /// <summary>Grade code 14; interpret within the selected grading system.</summary>
    Grade14 = 14,
    /// <summary>Grade code 15; interpret within the selected grading system.</summary>
    Grade15 = 15,
    /// <summary>Grade code 16; interpret within the selected grading system.</summary>
    Grade16 = 16,
    /// <summary>Grade code 17; interpret within the selected grading system.</summary>
    Grade17 = 17
}

/// <summary>
/// Explicit school or college grade bounds. At least one bound is required, equal bounds mean
/// an exact grade, and the order is preschool, kindergarten, then grades 1–17.
/// </summary>
public sealed record BookOnixGradeRange(BookOnixGradeSystem System, BookOnixGrade? Minimum = null, BookOnixGrade? Maximum = null);
