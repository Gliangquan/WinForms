namespace PaperFormat.Infrastructure.Word;

internal static class StaTaskRunner
{
    public static Task<T> RunAsync<T>(Func<T> func, CancellationToken cancellationToken = default)
    {
        var completionSource = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                completionSource.SetResult(func());
            }
            catch (OperationCanceledException ex)
            {
                completionSource.SetCanceled(ex.CancellationToken);
            }
            catch (Exception ex)
            {
                completionSource.SetException(ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();

        return completionSource.Task;
    }
}
