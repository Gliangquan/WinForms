namespace PaperFormat.Domain;

public sealed class RuleProfile
{
    public string ProfileName { get; init; } = "Default Rule Profile";

    public string Version { get; init; } = "1.0.0";

    public string Description { get; init; } = string.Empty;

    public DocumentFormattingRule Document { get; init; } = new();

    public List<ParagraphFormattingRule> ParagraphRules { get; init; } = [];
}

public sealed class DocumentFormattingRule
{
    public bool EnforceA4Paper { get; init; }

    public MarginSpec? MarginsCentimeters { get; init; }
}

public sealed class MarginSpec
{
    public double Top { get; init; }

    public double Bottom { get; init; }

    public double Left { get; init; }

    public double Right { get; init; }
}

public sealed class ParagraphFormattingRule
{
    public string RuleId { get; init; } = Guid.NewGuid().ToString("N");

    public string DisplayName { get; init; } = "Unnamed Rule";

    public List<string> StyleNames { get; init; } = [];

    public List<int> OutlineLevels { get; init; } = [];

    public List<string> StoryScopes { get; init; } = [];

    public bool ApplyToPageNumberParagraph { get; init; }

    public bool ApplyToAnyNonEmptyParagraph { get; init; }

    public bool IgnoreEmptyParagraph { get; init; } = true;

    public IssueSeverity Severity { get; init; } = IssueSeverity.Warning;

    public ParagraphFormattingExpectation Expected { get; init; } = new();
}

public sealed class ParagraphFormattingExpectation
{
    public string? FontName { get; init; }

    public string? FontColor { get; init; }

    public double? FontSize { get; init; }

    public bool? Bold { get; init; }

    public ParagraphAlignmentKind? Alignment { get; init; }

    public double? FirstLineIndentChars { get; init; }

    public double? LineSpacingMultiple { get; init; }
}
