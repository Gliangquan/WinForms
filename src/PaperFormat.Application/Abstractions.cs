using PaperFormat.Domain;

namespace PaperFormat.Application;

public interface IRuleProfileStore
{
    Task<RuleProfile> LoadAsync(string path, CancellationToken cancellationToken = default);
}

public interface IWordDocumentProcessor
{
    Task<DocumentProcessingResult> ProcessAsync(
        DocumentProcessRequest request,
        RuleProfile profile,
        CancellationToken cancellationToken = default);
}

public interface IProcessingReportStore
{
    Task WriteAsync(DocumentProcessingResult result, CancellationToken cancellationToken = default);
}
