namespace PaperFormat.Domain;

public sealed class DocumentProcessRequest
{
    public string DocumentPath { get; init; } = string.Empty;

    public string RuleProfilePath { get; init; } = string.Empty;

    public string OutputDirectory { get; init; } = string.Empty;

    public bool ApplyAutoFixes { get; init; } = true;

    public bool ShowOfficeWindow { get; init; }

    public Action<string>? Progress { get; init; }
}

public sealed class NodeLocator
{
    public NodeKind NodeKind { get; init; } = NodeKind.Paragraph;

    public string StoryScope { get; init; } = string.Empty;

    public int SectionIndex { get; init; }

    public int ParagraphIndex { get; init; }

    public int RangeStart { get; init; }

    public int RangeEnd { get; init; }

    public int? PageNumber { get; init; }

    public string StyleName { get; init; } = string.Empty;

    public string Snippet { get; init; } = string.Empty;
}

public sealed class IssueRecord
{
    public string IssueId { get; init; } = Guid.NewGuid().ToString("N");

    public string RuleId { get; init; } = string.Empty;

    public string RuleName { get; init; } = string.Empty;

    public IssueSeverity Severity { get; init; }

    public string PropertyName { get; init; } = string.Empty;

    public string ExpectedValue { get; init; } = string.Empty;

    public string ActualValue { get; init; } = string.Empty;

    public bool CanAutoFix { get; init; }

    public IssueStatus Status { get; init; }

    public string? FailureReason { get; init; }

    public NodeLocator Location { get; init; } = new();
}

public sealed class ProcessingSummary
{
    public int TotalIssues { get; init; }

    public int FixedIssues { get; init; }

    public int RemainingIssues { get; init; }

    public int FailedFixes { get; init; }
}

public sealed class DocumentProcessingResult
{
    public string InputDocumentPath { get; init; } = string.Empty;

    public string RuleProfilePath { get; init; } = string.Empty;

    public string OutputDirectory { get; init; } = string.Empty;

    public string RuleProfileName { get; set; } = string.Empty;

    public string RuleProfileVersion { get; set; } = string.Empty;

    public string? BackupDocumentPath { get; set; }

    public string? FixedDocumentPath { get; set; }

    public string? JsonReportPath { get; set; }

    public string? MarkdownReportPath { get; set; }

    public bool AutoFixRequested { get; init; }

    public DateTimeOffset StartedAtUtc { get; init; }

    public DateTimeOffset CompletedAtUtc { get; set; }

    public string? ErrorMessage { get; set; }

    public List<IssueRecord> Issues { get; } = [];

    public ProcessingSummary Summary =>
        new()
        {
            TotalIssues = Issues.Count,
            FixedIssues = Issues.Count(issue => issue.Status == IssueStatus.Fixed),
            RemainingIssues = Issues.Count(issue => issue.Status == IssueStatus.Detected),
            FailedFixes = Issues.Count(issue => issue.Status == IssueStatus.Failed)
        };
}
