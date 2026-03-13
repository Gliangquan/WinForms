using PaperFormat.Application;
using PaperFormat.Domain;
using PaperFormat.Infrastructure.Word;
using MsWord = Microsoft.Office.Interop.Word;
using AntButton = AntdUI.Button;
using AntInput = AntdUI.Input;
using AntPageHeader = AntdUI.PageHeader;
using AntSwitch = AntdUI.Switch;
using AntTable = AntdUI.Table;

namespace PaperFormat.WinForms;

public sealed class MainForm : Form
{
    private readonly DocumentProcessingCoordinator _coordinator;
    private readonly AntPageHeader _pageHeader = new() { Dock = DockStyle.Fill, Text = "\u8bba\u6587\u683c\u5f0f\u68c0\u6d4b", SubText = "AntdUI", Description = "\u9876\u90e8\u76f4\u63a5\u9009\u6587\u6863\u3001\u89c4\u5219\u548c\u8f93\u51fa\u76ee\u5f55\u3002", UseTitleFont = true };
    private readonly AntInput _documentPathTextBox = CreateInput();
    private readonly AntInput _rulePathTextBox = CreateInput();
    private readonly AntInput _outputPathTextBox = CreateInput();
    private readonly AntSwitch _autoFixCheckBox = new() { Text = "\u5728\u526f\u672c\u4e0a\u81ea\u52a8\u4fee\u590d", Checked = true, Dock = DockStyle.Left };
    private readonly AntSwitch _showOfficeWindowCheckBox = new() { Text = "\u663e\u793a Word \u7a97\u53e3", Dock = DockStyle.Left };
    private readonly AntButton _runButton = CreateButton("\u5f00\u59cb\u68c0\u6d4b", 120, 36);
    private readonly AntButton _openInWordButton = CreateButton("\u5728 Word \u4e2d\u5b9a\u4f4d", 130, 32);
    private readonly Label _statusLabel = new() { AutoSize = true, Text = "\u5c31\u7eea" };
    private readonly AntTable _issueGrid = new() { Dock = DockStyle.Fill };
    private readonly AntInput _summaryTextBox = CreateTextArea();
    private readonly AntInput _detailTextBox = CreateTextArea();
    private readonly AntInput _logTextBox = CreateTextArea();

    private List<IssueRecord> _issues = [];
    private string? _currentReviewDocumentPath;

    public MainForm()
    {
        Text = "\u8bba\u6587\u683c\u5f0f\u68c0\u6d4b\u5de5\u5177";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1320, 820);
        MinimumSize = new Size(1000, 680);
        Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular);

        _coordinator = new DocumentProcessingCoordinator(
            new JsonRuleProfileStore(),
            new WordDocumentProcessor(),
            new JsonProcessingReportStore());

        BuildLayout();
        BindEvents();
        ApplyDefaultPaths();
        ResetIssueGrid();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 84));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 270));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        root.Controls.Add(_pageHeader, 0, 0);
        root.Controls.Add(BuildSetupPanel(), 0, 1);
        root.Controls.Add(BuildContentPanel(), 0, 2);
        Controls.Add(root);
    }

    private Control BuildSetupPanel()
    {
        var group = new GroupBox { Text = "\u7b2c\u4e00\u6b65\uff1a\u9009\u62e9\u6587\u6863", Dock = DockStyle.Fill, Padding = new Padding(10) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 5, Padding = new Padding(10, 12, 10, 10) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 55));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        AddPathRow(layout, 0, "\u6587\u6863", _documentPathTextBox, OnBrowseDocument);
        AddPathRow(layout, 1, "\u89c4\u5219", _rulePathTextBox, OnBrowseRule);
        AddPathRow(layout, 2, "\u8f93\u51fa", _outputPathTextBox, OnBrowseOutput);

        var optionPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = true, Padding = new Padding(0, 6, 0, 0) };
        optionPanel.Controls.Add(_autoFixCheckBox);
        optionPanel.Controls.Add(_showOfficeWindowCheckBox);
        layout.Controls.Add(new Label { Text = "\u9009\u9879", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 8, 0, 0) }, 0, 3);
        layout.Controls.Add(optionPanel, 1, 3);
        layout.SetColumnSpan(optionPanel, 2);

        var actionPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            AutoSize = true
        };
        actionPanel.Controls.Add(new Label
        {
            Text = "\u5148\u9009\u6587\u6863\uff0c\u518d\u70b9\u201c\u5f00\u59cb\u68c0\u6d4b\u201d\u3002",
            AutoSize = true,
            Margin = new Padding(0, 10, 16, 0)
        });
        actionPanel.Controls.Add(new Label
        {
            Text = "\u72b6\u6001\uff1a",
            AutoSize = true,
            Margin = new Padding(0, 10, 4, 0)
        });
        actionPanel.Controls.Add(_statusLabel);
        _statusLabel.Margin = new Padding(0, 10, 16, 0);
        actionPanel.Controls.Add(_runButton);
        layout.Controls.Add(actionPanel, 0, 4);
        layout.SetColumnSpan(actionPanel, 3);

        group.Controls.Add(layout);
        return group;
    }

    private Control BuildContentPanel()
    {
        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 340, FixedPanel = FixedPanel.Panel1 };

        var left = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        left.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
        left.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
        left.Controls.Add(BuildSummaryPanel(), 0, 0);
        left.Controls.Add(BuildDetailPanel(), 0, 1);

        var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 64));
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 36));
        right.Controls.Add(BuildIssuePanel(), 0, 0);
        right.Controls.Add(BuildLogPanel(), 0, 1);

        split.Panel1.Controls.Add(left);
        split.Panel2.Controls.Add(right);
        return split;
    }

    private Control BuildSummaryPanel()
    {
        var group = new GroupBox { Text = "\u8fd0\u884c\u7ed3\u679c", Dock = DockStyle.Fill };
        group.Controls.Add(_summaryTextBox);
        return group;
    }

    private Control BuildDetailPanel()
    {
        var group = new GroupBox { Text = "\u95ee\u9898\u8be6\u60c5", Dock = DockStyle.Fill };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(6) };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.Controls.Add(_openInWordButton, 0, 0);
        layout.Controls.Add(_detailTextBox, 0, 1);
        group.Controls.Add(layout);
        return group;
    }

    private Control BuildIssuePanel()
    {
        var group = new GroupBox { Text = "\u95ee\u9898\u5217\u8868", Dock = DockStyle.Fill };
        _issueGrid.Dock = DockStyle.Fill;
        _issueGrid.Bordered = true;
        _issueGrid.VisibleHeader = true;
        group.Controls.Add(_issueGrid);
        return group;
    }

    private Control BuildLogPanel()
    {
        var group = new GroupBox { Text = "\u8fd0\u884c\u65e5\u5fd7", Dock = DockStyle.Fill };
        group.Controls.Add(_logTextBox);
        return group;
    }

    private void BindEvents()
    {
        _runButton.Click += async (_, _) => await ExecuteAsync();
        _issueGrid.SelectIndexChanged += (_, _) => RenderSelectedIssue();
        _issueGrid.CellClick += (_, _) => RenderSelectedIssue();
        _issueGrid.DoubleClick += (_, _) => LocateSelectedIssue();
        _openInWordButton.Click += (_, _) => LocateSelectedIssue();
    }

    private void AddPathRow(TableLayoutPanel layout, int row, string label, AntInput textBox, EventHandler browseHandler)
    {
        var browseButton = CreateButton("\u9009\u62e9", 70, 30);
        browseButton.Dock = DockStyle.Fill;
        browseButton.Click += browseHandler;

        layout.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 8, 0, 0) }, 0, row);
        layout.Controls.Add(textBox, 1, row);
        layout.Controls.Add(browseButton, 2, row);
    }

    private static AntInput CreateInput()
    {
        return new AntInput
        {
            Dock = DockStyle.Fill,
            PlaceholderText = "\u8bf7\u9009\u62e9\u6587\u4ef6\u6216\u76ee\u5f55",
            Height = 38,
            Margin = new Padding(0, 2, 8, 2)
        };
    }

    private static AntInput CreateTextArea()
    {
        return new AntInput
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            Multiline = true
        };
    }

    private static AntButton CreateButton(string text, int width, int height)
    {
        return new AntButton
        {
            Text = text,
            Width = width,
            Height = height,
            Margin = new Padding(0, 4, 0, 4)
        };
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
        ToggleBusy(true);
        _statusLabel.Text = "\u5904\u7406\u4e2d";
        _summaryTextBox.Text = string.Empty;
        _detailTextBox.Text = string.Empty;
        _logTextBox.Text = string.Empty;
        _issues = [];
        _currentReviewDocumentPath = null;
        ResetIssueGrid();

        try
        {
            var result = await _coordinator.ProcessAsync(new DocumentProcessRequest
            {
                DocumentPath = _documentPathTextBox.Text.Trim(),
                RuleProfilePath = _rulePathTextBox.Text.Trim(),
                OutputDirectory = _outputPathTextBox.Text.Trim(),
                ApplyAutoFixes = _autoFixCheckBox.Checked,
                ShowOfficeWindow = _showOfficeWindowCheckBox.Checked,
                Progress = AppendLogFromWorker
            });

            _issues = result.Issues.ToList();
            _currentReviewDocumentPath = result.FixedDocumentPath ?? result.InputDocumentPath;
            _statusLabel.Text = string.IsNullOrWhiteSpace(result.ErrorMessage) ? "\u5b8c\u6210" : "\u5931\u8d25";

            _summaryTextBox.Text = string.Join(Environment.NewLine, new[]
            {
                $"\u89c4\u5219\u96c6: {result.RuleProfileName} {result.RuleProfileVersion}",
                $"\u8f93\u5165\u6587\u6863: {result.InputDocumentPath}",
                $"\u95ee\u9898\u603b\u6570: {result.Summary.TotalIssues}",
                $"\u5df2\u4fee\u590d: {result.Summary.FixedIssues}",
                $"\u5f85\u5904\u7406: {result.Summary.RemainingIssues}",
                $"\u4fee\u590d\u540e\u6587\u6863: {result.FixedDocumentPath ?? "\u672a\u751f\u6210"}",
                $"\u62a5\u544a JSON: {result.JsonReportPath ?? "\u672a\u751f\u6210"}",
                $"\u62a5\u544a Markdown: {result.MarkdownReportPath ?? "\u672a\u751f\u6210"}",
                string.IsNullOrWhiteSpace(result.ErrorMessage) ? "\u7ed3\u679c: \u6210\u529f" : $"\u7ed3\u679c: \u5931\u8d25 - {result.ErrorMessage}"
            });

            BindIssueGrid();
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "\u5931\u8d25";
            _summaryTextBox.Text = ex.ToString();
            AppendLog($"Unhandled error: {ex}");
            MessageBox.Show(this, ex.Message, "\u8fd0\u884c\u5931\u8d25", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            ToggleBusy(false);
        }
    }

    private void BindIssueGrid()
    {
        _issueGrid.DataSource = _issues.Select(issue => new IssueGridRow(issue)).ToList();
        if (_issues.Count > 0)
        {
            _issueGrid.SelectedIndex = 0;
            RenderSelectedIssue();
        }
        else
        {
            _detailTextBox.Text = "\u5f53\u524d\u89c4\u5219\u6ca1\u6709\u8bc6\u522b\u51fa\u95ee\u9898\u3002";
        }
    }

    private void ResetIssueGrid()
    {
        _issueGrid.DataSource = null;
        _detailTextBox.Text = string.Empty;
    }

    private void RenderSelectedIssue()
    {
        var issue = GetSelectedIssue();
        if (issue is null)
        {
            return;
        }

        _detailTextBox.Text = string.Join(Environment.NewLine, new[]
        {
            $"\u89c4\u5219: {issue.RuleName}",
            $"\u533a\u57df: {TranslateScope(issue.Location.StoryScope)}",
            $"\u72b6\u6001: {issue.Status}",
            $"\u7ea7\u522b: {issue.Severity}",
            $"\u5c5e\u6027: {issue.PropertyName}",
            $"\u671f\u671b\u503c: {issue.ExpectedValue}",
            $"\u5b9e\u9645\u503c: {issue.ActualValue}",
            $"\u9875\u7801: {issue.Location.PageNumber?.ToString() ?? "-"}",
            $"\u6837\u5f0f: {issue.Location.StyleName}",
            $"\u6587\u672c\u7247\u6bb5: {issue.Location.Snippet}"
        });
    }

    private void LocateSelectedIssue()
    {
        var issue = GetSelectedIssue();
        if (issue is null)
        {
            MessageBox.Show(this, "\u8bf7\u5148\u9009\u4e2d\u4e00\u6761\u95ee\u9898\u3002", "\u5728 Word \u4e2d\u5b9a\u4f4d", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (string.IsNullOrWhiteSpace(_currentReviewDocumentPath) || !File.Exists(_currentReviewDocumentPath))
        {
            MessageBox.Show(this, "\u5f53\u524d\u590d\u6838\u6587\u6863\u4e0d\u53ef\u7528\uff0c\u8bf7\u5148\u8fd0\u884c\u68c0\u6d4b\u3002", "\u5728 Word \u4e2d\u5b9a\u4f4d", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            var word = new MsWord.Application { Visible = true, ScreenUpdating = true };
            var document = word.Documents.Open(_currentReviewDocumentPath, ReadOnly: false);
            var start = Math.Max(0, issue.Location.RangeStart);
            var end = Math.Max(start + 1, issue.Location.RangeEnd);
            document.Range(start, end).Select();
            word.Activate();
        }
        catch (Exception ex)
        {
            AppendLog($"Word navigation failed: {ex.Message}");
            MessageBox.Show(this, ex.Message, "\u5b9a\u4f4d\u5931\u8d25", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private IssueRecord? GetSelectedIssue()
    {
        var index = _issueGrid.SelectedIndex;
        if (index < 0 || index >= _issues.Count)
        {
            return null;
        }

        return _issues[index];
    }

    private static string TranslateScope(string scope)
    {
        return scope switch
        {
            "Header" => "\u9875\u7709",
            "Footer" => "\u9875\u811a",
            "Main" => "\u6b63\u6587",
            _ => scope
        };
    }

    private void ToggleBusy(bool busy)
    {
        _runButton.Enabled = !busy;
        UseWaitCursor = busy;
    }

    private void AppendLogFromWorker(string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(AppendLog), message);
            return;
        }

        AppendLog(message);
    }

    private void AppendLog(string message)
    {
        _logTextBox.Text += message + Environment.NewLine;
    }

    private void OnBrowseDocument(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog { Filter = "Word \u6587\u6863|*.docx;*.doc|\u6240\u6709\u6587\u4ef6|*.*" };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _documentPathTextBox.Text = dialog.FileName;
        }
    }

    private void OnBrowseRule(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog { Filter = "\u89c4\u5219\u6587\u4ef6|*.json|\u6240\u6709\u6587\u4ef6|*.*" };
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

    private sealed record IssueGridRow(IssueRecord Issue)
    {
        public string IssueId => Issue.IssueId;
        public string Rule => Issue.RuleName;
        public string Scope => Issue.Location.StoryScope;
        public string Property => Issue.PropertyName;
        public string Status => Issue.Status.ToString();
        public string Severity => Issue.Severity.ToString();
        public int? Page => Issue.Location.PageNumber;
        public string Snippet => Issue.Location.Snippet;
    }
}
