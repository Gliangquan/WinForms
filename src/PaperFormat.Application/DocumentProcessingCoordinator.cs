using PaperFormat.Domain;

namespace PaperFormat.Application;

public sealed class DocumentProcessingCoordinator
{
    private readonly IRuleProfileStore _ruleProfileStore;
    private readonly IWordDocumentProcessor _documentProcessor;
    private readonly IProcessingReportStore _reportStore;

    public DocumentProcessingCoordinator(
        IRuleProfileStore ruleProfileStore,
        IWordDocumentProcessor documentProcessor,
        IProcessingReportStore reportStore)
    {
        _ruleProfileStore = ruleProfileStore;
        _documentProcessor = documentProcessor;
        _reportStore = reportStore;
    }

    public async Task<DocumentProcessingResult> ProcessAsync(
        DocumentProcessRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.DocumentPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.RuleProfilePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.OutputDirectory);

        Directory.CreateDirectory(request.OutputDirectory);

        request.Progress?.Invoke($"Loading rule profile: {request.RuleProfilePath}");
        var profile = await _ruleProfileStore.LoadAsync(request.RuleProfilePath, cancellationToken);
        request.Progress?.Invoke($"Rule profile loaded: {profile.ProfileName} {profile.Version}");

        request.Progress?.Invoke($"Starting document processing: {request.DocumentPath}");
        var result = await _documentProcessor.ProcessAsync(request, profile, cancellationToken);

        request.Progress?.Invoke("Writing reports");
        await _reportStore.WriteAsync(result, cancellationToken);
        request.Progress?.Invoke("Processing finished");
        return result;
    }
}
