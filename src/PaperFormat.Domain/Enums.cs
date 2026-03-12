namespace PaperFormat.Domain;

public enum IssueSeverity
{
    Info,
    Warning,
    Error
}

public enum ParagraphAlignmentKind
{
    Left,
    Center,
    Right,
    Justify
}

public enum IssueStatus
{
    Detected,
    Fixed,
    Failed
}

public enum NodeKind
{
    Document,
    Section,
    Paragraph
}
