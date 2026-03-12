using System.Text;
using System.Text.Json;
using PaperFormat.Application;
using PaperFormat.Domain;

namespace PaperFormat.Infrastructure.Word;

public sealed class JsonProcessingReportStore : IProcessingReportStore
{
    private readonly JsonSerializerOptions _serializerOptions = JsonOptionsFactory.Create();

    public async Task WriteAsync(DocumentProcessingResult result, CancellationToken cancellationToken = default)
    {
        var baseDirectory = result.OutputDirectory;
        if (string.IsNullOrWhiteSpace(baseDirectory))
        {
            baseDirectory = Path.GetDirectoryName(result.FixedDocumentPath)
                ?? Path.GetDirectoryName(result.BackupDocumentPath)
                ?? throw new InvalidOperationException("无法确定报告输出目录。");
        }

        Directory.CreateDirectory(baseDirectory);

        var stamp = result.StartedAtUtc.ToString("yyyyMMdd_HHmmss");
        var baseName = Path.GetFileNameWithoutExtension(result.InputDocumentPath);

        result.JsonReportPath ??= Path.Combine(baseDirectory, $"{baseName}.{stamp}.report.json");
        result.MarkdownReportPath ??= Path.Combine(baseDirectory, $"{baseName}.{stamp}.report.md");

        var json = JsonSerializer.Serialize(result, _serializerOptions);
        await File.WriteAllTextAsync(result.JsonReportPath, json, cancellationToken);

        var markdown = BuildMarkdown(result);
        await File.WriteAllTextAsync(result.MarkdownReportPath, markdown, cancellationToken);
    }

    private static string BuildMarkdown(DocumentProcessingResult result)
    {
        var summary = result.Summary;
        var builder = new StringBuilder();

        builder.AppendLine("# 论文格式检测/修复报告");
        builder.AppendLine();
        builder.AppendLine($"- 输入文档：`{result.InputDocumentPath}`");
        builder.AppendLine($"- 规则包：`{result.RuleProfileName}` `{result.RuleProfileVersion}`");
        builder.AppendLine($"- 开始时间：`{result.StartedAtUtc:O}`");
        builder.AppendLine($"- 结束时间：`{result.CompletedAtUtc:O}`");
        builder.AppendLine($"- 自动修复：`{(result.AutoFixRequested ? "已启用" : "未启用")}`");
        builder.AppendLine($"- 备份文件：`{result.BackupDocumentPath ?? "未生成"}`");
        builder.AppendLine($"- 修复输出：`{result.FixedDocumentPath ?? "未生成"}`");
        builder.AppendLine();

        if (!string.IsNullOrWhiteSpace(result.ErrorMessage))
        {
            builder.AppendLine("## 处理错误");
            builder.AppendLine();
            builder.AppendLine(result.ErrorMessage);
            builder.AppendLine();
        }

        builder.AppendLine("## 汇总");
        builder.AppendLine();
        builder.AppendLine($"- 问题总数：`{summary.TotalIssues}`");
        builder.AppendLine($"- 已修复：`{summary.FixedIssues}`");
        builder.AppendLine($"- 待处理：`{summary.RemainingIssues}`");
        builder.AppendLine($"- 修复失败：`{summary.FailedFixes}`");
        builder.AppendLine();

        builder.AppendLine("## 明细");
        builder.AppendLine();
        builder.AppendLine("|规则|属性|状态|位置|实际值|期望值|片段|");
        builder.AppendLine("|---|---|---|---|---|---|---|");

        foreach (var issue in result.Issues)
        {
            var location = issue.Location.NodeKind == NodeKind.Paragraph
                ? $"P{issue.Location.ParagraphIndex}"
                : $"S{issue.Location.SectionIndex}";

            builder.Append('|').Append(Escape(issue.RuleName)).Append('|')
                .Append(Escape(issue.PropertyName)).Append('|')
                .Append(issue.Status).Append('|')
                .Append(location).Append('|')
                .Append(Escape(issue.ActualValue)).Append('|')
                .Append(Escape(issue.ExpectedValue)).Append('|')
                .Append(Escape(issue.Location.Snippet)).AppendLine("|");
        }

        return builder.ToString();
    }

    private static string Escape(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "-";
        }

        return value.Replace("\r", " ").Replace("\n", " ").Replace("|", "\\|").Trim();
    }
}
