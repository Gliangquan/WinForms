using PaperFormat.Application;
using PaperFormat.Domain;
using PaperFormat.Infrastructure.Word;

namespace PaperFormat.WinForms;

public sealed class MainForm : Form
{
    private readonly DocumentProcessingCoordinator _coordinator;
    private readonly TextBox _documentPathTextBox = new() { Dock = DockStyle.Fill };
    private readonly TextBox _rulePathTextBox = new() { Dock = DockStyle.Fill };
    private readonly TextBox _outputPathTextBox = new() { Dock = DockStyle.Fill };
    private readonly CheckBox _autoFixCheckBox = new() { Text = "Enable auto-fix", Checked = true, AutoSize = true };
    private readonly CheckBox _showOfficeWindowCheckBox = new() { Text = "Show Office window", AutoSize = true };
    private readonly Button _executeButton = new() { Text = "Run analysis", Dock = DockStyle.Fill, Height = 34 };
    private readonly DataGridView _issueGrid = new() { Dock = DockStyle.Fill, ReadOnly = true, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
    private readonly TextBox _summaryTextBox = new() { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true };
    private readonly TextBox _logTextBox = new() { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true };
    private readonly Label _statusLabel = new() { Dock = DockStyle.Fill, AutoSize = true, Text = "Ready" };

    public MainForm()
    {
        Text = "Paper Format Analyzer";
        MinimumSize = new Size(1280, 820);
        StartPosition = FormStartPosition.CenterScreen;

        _coordinator = new DocumentProcessingCoordinator(
            new JsonRuleProfileStore(),
            new WordDocumentProcessor(),
            new JsonProcessingReportStore());

        BuildLayout();
        BindEvents();
        ApplyDefaultPaths();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(12)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 25));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 25));

        var inputPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 3,
            RowCount = 4,
            AutoSize = true
        };
        inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

        inputPanel.Controls.Add(CreateLabel("Document"), 0, 0);
        inputPanel.Controls.Add(_documentPathTextBox, 1, 0);
        inputPanel.Controls.Add(CreateButton("Browse", OnBrowseDocument), 2, 0);

        inputPanel.Controls.Add(CreateLabel("Rule file"), 0, 1);
        inputPanel.Controls.Add(_rulePathTextBox, 1, 1);
        inputPanel.Controls.Add(CreateButton("Browse", OnBrowseRule), 2, 1);

        inputPanel.Controls.Add(CreateLabel("Output"), 0, 2);
        inputPanel.Controls.Add(_outputPathTextBox, 1, 2);
        inputPanel.Controls.Add(CreateButton("Browse", OnBrowseOutput), 2, 2);

        var optionsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            WrapContents = false
        };
        optionsPanel.Controls.Add(_autoFixCheckBox);
        optionsPanel.Controls.Add(_showOfficeWindowCheckBox);

        inputPanel.Controls.Add(CreateLabel("Options"), 0, 3);
        inputPanel.Controls.Add(optionsPanel, 1, 3);
        inputPanel.Controls.Add(_executeButton, 2, 3);

        var statusPanel = new Panel { Dock = DockStyle.Top, Height = 32 };
        _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
        statusPanel.Controls.Add(_statusLabel);

        var issuePanel = new GroupBox { Text = "Issues", Dock = DockStyle.Fill };
        issuePanel.Controls.Add(_issueGrid);

        var summaryPanel = new GroupBox { Text = "Summary", Dock = DockStyle.Fill };
        summaryPanel.Controls.Add(_summaryTextBox);

        var logPanel = new GroupBox { Text = "Live log", Dock = DockStyle.Fill };
        logPanel.Controls.Add(_logTextBox);

        root.Controls.Add(inputPanel, 0, 0);
        root.Controls.Add(statusPanel, 0, 1);
        root.Controls.Add(issuePanel, 0, 2);
        root.Controls.Add(summaryPanel, 0, 3);
        root.Controls.Add(logPanel, 0, 4);

        Controls.Add(root);
    }

    private void BindEvents()
    {
        _executeButton.Click += async (_, _) => await ExecuteAsync();
    }

    private void ApplyDefaultPaths()
    {
        var root = FindWorkspaceRoot();
        _rulePathTextBox.Text = Path.Combine(root, "rules", "sample.paperformat.json");
        _outputPathTextBox.Text = Path.Combine(root, "output");
    }

    private static string FindWorkspaceRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null)
        {
            var rulePath = Path.Combine(directory.FullName, "rules", "sample.paperformat.json");
            if (File.Exists(rulePath))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
    }

    private async Task ExecuteAsync()
    {
        ToggleBusyState(true);
        _statusLabel.Text = "Processing...";
        _summaryTextBox.Clear();
        _logTextBox.Clear();
        _issueGrid.DataSource = null;

        try
        {
            var request = new DocumentProcessRequest
            {
                DocumentPath = _documentPathTextBox.Text.Trim(),
                RuleProfilePath = _rulePathTextBox.Text.Trim(),
                OutputDirectory = _outputPathTextBox.Text.Trim(),
                ApplyAutoFixes = _autoFixCheckBox.Checked,
                ShowOfficeWindow = _showOfficeWindowCheckBox.Checked,
                Progress = AppendLogFromWorker
            };

            var result = await _coordinator.ProcessAsync(request);
            RenderResult(result);

            _statusLabel.Text = string.IsNullOrWhiteSpace(result.ErrorMessage)
                ? "Completed"
                : "Failed";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "Failed";
            _summaryTextBox.Text = ex.ToString();
            AppendLog($"Unhandled error: {ex}");
            MessageBox.Show(this, ex.Message, "Run failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            ToggleBusyState(false);
        }
    }

    private void RenderResult(DocumentProcessingResult result)
    {
        _issueGrid.DataSource = result.Issues.Select(issue => new
        {
            Rule = issue.RuleName,
            Property = issue.PropertyName,
            Status = issue.Status.ToString(),
            Severity = issue.Severity.ToString(),
            Location = issue.Location.NodeKind == NodeKind.Paragraph ? $"P{issue.Location.ParagraphIndex}" : $"S{issue.Location.SectionIndex}",
            Page = issue.Location.PageNumber,
            Style = issue.Location.StyleName,
            Actual = issue.ActualValue,
            Expected = issue.ExpectedValue,
            Snippet = issue.Location.Snippet
        }).ToList();

        var summary = result.Summary;
        _summaryTextBox.Text = string.Join(
            Environment.NewLine,
            [
                $"Rule profile: {result.RuleProfileName} {result.RuleProfileVersion}",
                $"Input document: {result.InputDocumentPath}",
                $"Auto-fix: {(result.AutoFixRequested ? "enabled" : "disabled")}",
                $"Total issues: {summary.TotalIssues}",
                $"Fixed issues: {summary.FixedIssues}",
                $"Remaining issues: {summary.RemainingIssues}",
                $"Failed fixes: {summary.FailedFixes}",
                $"Backup file: {result.BackupDocumentPath ?? "not generated"}",
                $"Fixed output: {result.FixedDocumentPath ?? "not generated"}",
                $"JSON report: {result.JsonReportPath ?? "not generated"}",
                $"Markdown report: {result.MarkdownReportPath ?? "not generated"}",
                string.IsNullOrWhiteSpace(result.ErrorMessage) ? "Result: success" : $"Result: failed - {result.ErrorMessage}"
            ]);
    }

    private void AppendLogFromWorker(string message)
    {
        if (IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(AppendLog), message);
            return;
        }

        AppendLog(message);
    }

    private void AppendLog(string message)
    {
        _logTextBox.AppendText(message + Environment.NewLine);
    }

    private void ToggleBusyState(bool isBusy)
    {
        _executeButton.Enabled = !isBusy;
        UseWaitCursor = isBusy;
    }

    private void OnBrowseDocument(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Word Document|*.docx;*.doc|All Files|*.*"
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _documentPathTextBox.Text = dialog.FileName;
        }
    }

    private void OnBrowseRule(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Rule File|*.json|All Files|*.*"
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _rulePathTextBox.Text = dialog.FileName;
        }
    }

    private void OnBrowseOutput(object? sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog
        {
            SelectedPath = Directory.Exists(_outputPathTextBox.Text) ? _outputPathTextBox.Text : string.Empty
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _outputPathTextBox.Text = dialog.SelectedPath;
        }
    }

    private static Label CreateLabel(string text)
    {
        return new Label
        {
            Text = text,
            TextAlign = ContentAlignment.MiddleLeft,
            Dock = DockStyle.Fill,
            AutoSize = true,
            Padding = new Padding(0, 8, 0, 0)
        };
    }

    private static Button CreateButton(string text, EventHandler handler)
    {
        var button = new Button
        {
            Text = text,
            Dock = DockStyle.Fill,
            Height = 32
        };
        button.Click += handler;
        return button;
    }
}
