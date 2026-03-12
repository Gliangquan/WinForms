using System.Text.Json;
using PaperFormat.Application;
using PaperFormat.Domain;

namespace PaperFormat.Infrastructure.Word;

public sealed class JsonRuleProfileStore : IRuleProfileStore
{
    private readonly JsonSerializerOptions _serializerOptions = JsonOptionsFactory.Create();

    public async Task<RuleProfile> LoadAsync(string path, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("未找到规则文件。", path);
        }

        await using var stream = File.OpenRead(path);
        var profile = await JsonSerializer.DeserializeAsync<RuleProfile>(stream, _serializerOptions, cancellationToken);

        if (profile is null)
        {
            throw new InvalidOperationException("规则文件解析失败。");
        }

        if (profile.ParagraphRules.Count == 0)
        {
            throw new InvalidOperationException("规则文件至少需要包含一条段落规则。");
        }

        return profile;
    }
}
