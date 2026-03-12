using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using PaperFormat.Application;
using PaperFormat.Domain;
using MsWord = Microsoft.Office.Interop.Word;

namespace PaperFormat.Infrastructure.Word;

public sealed class WordDocumentProcessor : IWordDocumentProcessor
{
    private static readonly WdHeaderFooterIndex[] HeaderFooterIndexes =
    [
        WdHeaderFooterIndex.wdHeaderFooterPrimary,
        WdHeaderFooterIndex.wdHeaderFooterFirstPage,
        WdHeaderFooterIndex.wdHeaderFooterEvenPages
    ];

    public Task<DocumentProcessingResult> ProcessAsync(
        DocumentProcessRequest request,
        RuleProfile profile,
        CancellationToken cancellationToken = default)
    {
        return StaTaskRunner.RunAsync(() => ProcessCore(request, profile), cancellationToken);
    }

    private static DocumentProcessingResult ProcessCore(DocumentProcessRequest request, RuleProfile profile)
    {
        void Log(string message)
        {
            request.Progress?.Invoke($"[{DateTime.Now:HH:mm:ss}] {message}");
        }

        var result = new DocumentProcessingResult
        {
            InputDocumentPath = request.DocumentPath,
            RuleProfilePath = request.RuleProfilePath,
            OutputDirectory = request.OutputDirectory,
            RuleProfileName = profile.ProfileName,
            RuleProfileVersion = profile.Version,
            StartedAtUtc = DateTimeOffset.UtcNow,
            AutoFixRequested = request.ApplyAutoFixes
        };

        MsWord.Application? rawApplication = null;
        MsWord.Document? rawDocument = null;
        var editableDocumentPath = request.DocumentPath;

        try
        {
            Log("Preparing output directory");
            Directory.CreateDirectory(request.OutputDirectory);

            if (request.ApplyAutoFixes)
            {
                result.FixedDocumentPath = BuildFixedOutputPath(request.DocumentPath, request.OutputDirectory);
                File.Copy(request.DocumentPath, result.FixedDocumentPath, overwrite: true);
                editableDocumentPath = result.FixedDocumentPath;
                Log($"Working copy created: {editableDocumentPath}");
            }

            Log("Launching Office COM application");
            rawApplication = new MsWord.Application();
            ConfigureApplication(rawApplication, request.ShowOfficeWindow);
            Log($"Office application: {rawApplication.Name} {rawApplication.Version}, visible={rawApplication.Visible}");

            Log("Opening document");
            rawDocument = rawApplication.Documents.Open(editableDocumentPath, ReadOnly: !request.ApplyAutoFixes);
            Log($"Document opened: {rawDocument.Name}");

            Log("Evaluating document-level rules");
            EvaluateDocumentRules(rawDocument, profile, request.ApplyAutoFixes, result.Issues);
            Log("Document-level rules completed");

            Log("Evaluating paragraph rules");
            EvaluateParagraphRules(rawDocument, profile, request.ApplyAutoFixes, result.Issues, Log);
            Log($"Paragraph rules completed, issues={result.Issues.Count}");

            if (request.ApplyAutoFixes && !string.IsNullOrWhiteSpace(result.FixedDocumentPath))
            {
                Log("Updating tables of contents");
                UpdateTablesOfContents(rawDocument);
                Log("Saving fixed document");
                rawDocument.Save();
                Log("Fixed document saved");
            }
        }
        catch (Exception ex)
        {
            result.ErrorMessage = ex.Message;
            Log($"Processing failed: {ex}");
        }
        finally
        {
            try
            {
                Log("Closing document");
                rawDocument?.Close(MsWord.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch
            {
            }

            try
            {
                if (rawApplication is not null)
                {
                    Log("Quitting Office application");
                    rawApplication.ScreenUpdating = true;
                    rawApplication.DisplayAlerts = WdAlertLevel.wdAlertsAll;
                    rawApplication.Quit(WdSaveOptions.wdDoNotSaveChanges);
                }
            }
            catch
            {
            }
        }

        result.CompletedAtUtc = DateTimeOffset.UtcNow;
        Log("Worker finished");
        return result;
    }

    private static void ConfigureApplication(MsWord.Application rawApplication, bool showOfficeWindow)
    {
        rawApplication.Visible = showOfficeWindow;
        rawApplication.ScreenUpdating = showOfficeWindow;
        rawApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        rawApplication.Options.SaveNormalPrompt = false;
    }

    private static void EvaluateDocumentRules(
        MsWord.Document document,
        RuleProfile profile,
        bool applyFixes,
        ICollection<IssueRecord> issues)
    {
        if (document.Sections.Count == 0)
        {
            return;
        }

        var pageRule = profile.Document;
        if (!pageRule.EnforceA4Paper && pageRule.MarginsCentimeters is null)
        {
            return;
        }

        for (var index = 1; index <= document.Sections.Count; index++)
        {
            MsWord.Section? section = null;

            try
            {
                section = document.Sections[index];
                var pageSetup = section.PageSetup;

                if (pageRule.EnforceA4Paper)
                {
                    var expectedWidth = document.Application.CentimetersToPoints(21F);
                    var expectedHeight = document.Application.CentimetersToPoints(29.7F);
                    var actualWidth = (double)pageSetup.PageWidth;
                    var actualHeight = (double)pageSetup.PageHeight;

                    if (!ApproximatelyEquals(actualWidth, expectedWidth, 1) ||
                        !ApproximatelyEquals(actualHeight, expectedHeight, 1))
                    {
                        if (applyFixes)
                        {
                            pageSetup.PaperSize = WdPaperSize.wdPaperA4;
                        }

                        issues.Add(CreateDocumentIssue(
                            index,
                            "document.paper",
                            "Paper size",
                            IssueSeverity.Error,
                            "PaperSize",
                            "A4",
                            $"{PointsToCentimeters(actualWidth):0.##}cm x {PointsToCentimeters(actualHeight):0.##}cm",
                            applyFixes));
                    }
                }

                if (pageRule.MarginsCentimeters is not null)
                {
                    EvaluateMargin(index, "TopMargin", pageRule.MarginsCentimeters.Top, pageSetup.TopMargin, applyFixes, value => pageSetup.TopMargin = value, issues);
                    EvaluateMargin(index, "BottomMargin", pageRule.MarginsCentimeters.Bottom, pageSetup.BottomMargin, applyFixes, value => pageSetup.BottomMargin = value, issues);
                    EvaluateMargin(index, "LeftMargin", pageRule.MarginsCentimeters.Left, pageSetup.LeftMargin, applyFixes, value => pageSetup.LeftMargin = value, issues);
                    EvaluateMargin(index, "RightMargin", pageRule.MarginsCentimeters.Right, pageSetup.RightMargin, applyFixes, value => pageSetup.RightMargin = value, issues);
                }
            }
            finally
            {
                ReleaseComObject(section);
            }
        }
    }

    private static void EvaluateMargin(
        int sectionIndex,
        string propertyName,
        double expectedCentimeters,
        float actualPoints,
        bool applyFixes,
        Action<float> setAction,
        ICollection<IssueRecord> issues)
    {
        var expectedPoints = CentimetersToPoints(expectedCentimeters);
        if (ApproximatelyEquals(actualPoints, expectedPoints, 0.8))
        {
            return;
        }

        if (applyFixes)
        {
            setAction(expectedPoints);
        }

        issues.Add(CreateDocumentIssue(
            sectionIndex,
            $"document.margin.{propertyName}",
            $"{propertyName} does not match the profile",
            IssueSeverity.Warning,
            propertyName,
            $"{expectedCentimeters:0.##} cm",
            $"{PointsToCentimeters(actualPoints):0.##} cm",
            applyFixes));
    }

    private static void EvaluateParagraphRules(
        MsWord.Document document,
        RuleProfile profile,
        bool applyFixes,
        ICollection<IssueRecord> issues,
        Action<string>? log)
    {
        var stats = new ParagraphScanStats();

        ProcessParagraphCollection(
            document.Paragraphs,
            profile,
            applyFixes,
            issues,
            "Main",
            document,
            stats);

        for (var sectionIndex = 1; sectionIndex <= document.Sections.Count; sectionIndex++)
        {
            MsWord.Section? section = null;

            try
            {
                section = document.Sections[sectionIndex];
                ProcessHeaderFooterCollection(section.Headers, "Header", profile, applyFixes, issues, document, stats);
                ProcessHeaderFooterCollection(section.Footers, "Footer", profile, applyFixes, issues, document, stats);
            }
            finally
            {
                ReleaseComObject(section);
            }
        }

        log?.Invoke(
            $"Paragraph scan summary: scanned={stats.Scanned}, matched={stats.Matched}, stories=[{string.Join(", ", stats.StoryScopes.OrderBy(scope => scope))}], styles=[{string.Join(", ", stats.StyleNames.OrderBy(name => name))}]");
    }

    private static void ProcessHeaderFooterCollection(
        HeadersFooters headerFooters,
        string storyScope,
        RuleProfile profile,
        bool applyFixes,
        ICollection<IssueRecord> issues,
        MsWord.Document document,
        ParagraphScanStats stats)
    {
        foreach (var index in HeaderFooterIndexes)
        {
            HeaderFooter? headerFooter = null;

            try
            {
                headerFooter = headerFooters[index];
                if (!headerFooter.Exists)
                {
                    continue;
                }

                ProcessParagraphCollection(
                    headerFooter.Range.Paragraphs,
                    profile,
                    applyFixes,
                    issues,
                    storyScope,
                    document,
                    stats);
            }
            finally
            {
                ReleaseComObject(headerFooter);
            }
        }
    }

    private static void ProcessParagraphCollection(
        Paragraphs paragraphs,
        RuleProfile profile,
        bool applyFixes,
        ICollection<IssueRecord> issues,
        string storyScope,
        MsWord.Document document,
        ParagraphScanStats stats)
    {
        for (var index = 1; index <= paragraphs.Count; index++)
        {
            MsWord.Paragraph? paragraph = null;

            try
            {
                paragraph = paragraphs[index];
                var range = paragraph.Range;
                var paragraphKey = $"{storyScope}:{range.Start}:{range.End}";
                if (!stats.ProcessedParagraphKeys.Add(paragraphKey))
                {
                    continue;
                }

                stats.Scanned++;
                stats.StoryScopes.Add(storyScope);

                var text = NormalizeText(range.Text);
                var styleName = paragraph.get_Style() switch
                {
                    string value => value,
                    MsWord.Style style => style.NameLocal,
                    _ => string.Empty
                };

                if (!string.IsNullOrWhiteSpace(styleName))
                {
                    stats.StyleNames.Add(styleName);
                }

                var outlineLevel = ConvertOutlineLevel(paragraph.OutlineLevel);
                var hasPageNumberField = ContainsPageNumberField(range);
                var rule = ResolveRule(profile, storyScope, styleName, outlineLevel, text, hasPageNumberField);
                if (rule is null)
                {
                    continue;
                }

                stats.Matched++;

                if (rule.IgnoreEmptyParagraph && string.IsNullOrWhiteSpace(text) && !hasPageNumberField)
                {
                    continue;
                }

                var locator = CreateParagraphLocator(paragraph, stats.Scanned, styleName, text);
                var font = range.Font;
                var format = paragraph.Format;

                ApplyParagraphRule(rule, range, font, format, document, applyFixes, hasPageNumberField, locator, issues);
            }
            finally
            {
                ReleaseComObject(paragraph);
            }
        }
    }

    private static void ApplyParagraphRule(
        ParagraphFormattingRule rule,
        MsWord.Range range,
        MsWord.Font font,
        MsWord.ParagraphFormat format,
        MsWord.Document document,
        bool applyFixes,
        bool hasPageNumberField,
        NodeLocator locator,
        ICollection<IssueRecord> issues)
    {
        if (!string.IsNullOrWhiteSpace(rule.Expected.FontName))
        {
            var actualFont = ResolveFontName(font);
            if (!StringEquals(actualFont, rule.Expected.FontName))
            {
                var status = IssueStatus.Detected;
                if (applyFixes)
                {
                    font.NameFarEast = rule.Expected.FontName;
                    font.Name = rule.Expected.FontName;
                    status = IssueStatus.Fixed;
                }

                issues.Add(CreateParagraphIssue(rule, "FontName", rule.Expected.FontName, actualFont, status, locator));
            }
        }

        if (!string.IsNullOrWhiteSpace(rule.Expected.FontColor) &&
            TryResolveColor(rule.Expected.FontColor, out var expectedColor))
        {
            ApplyFontColorRule(rule, range, font, applyFixes, hasPageNumberField, expectedColor, locator, issues);
        }

        if (rule.Expected.FontSize is double expectedFontSize &&
            !ApproximatelyEquals(font.Size, expectedFontSize, 0.2))
        {
            var status = IssueStatus.Detected;
            if (applyFixes)
            {
                font.Size = (float)expectedFontSize;
                status = IssueStatus.Fixed;
            }

            issues.Add(CreateParagraphIssue(rule, "FontSize", expectedFontSize.ToString("0.##"), font.Size.ToString("0.##"), status, locator));
        }

        if (rule.Expected.Bold is bool expectedBold)
        {
            var actualBold = Convert.ToInt32(font.Bold) != 0;
            if (actualBold != expectedBold)
            {
                var status = IssueStatus.Detected;
                if (applyFixes)
                {
                    font.Bold = expectedBold ? -1 : 0;
                    status = IssueStatus.Fixed;
                }

                issues.Add(CreateParagraphIssue(rule, "Bold", expectedBold ? "true" : "false", actualBold ? "true" : "false", status, locator));
            }
        }

        if (rule.Expected.Alignment is ParagraphAlignmentKind expectedAlignment)
        {
            var actualAlignment = ConvertAlignment(format.Alignment);
            if (actualAlignment != expectedAlignment)
            {
                var status = IssueStatus.Detected;
                if (applyFixes)
                {
                    format.Alignment = ConvertAlignment(expectedAlignment);
                    status = IssueStatus.Fixed;
                }

                issues.Add(CreateParagraphIssue(rule, "Alignment", expectedAlignment.ToString(), actualAlignment.ToString(), status, locator));
            }
        }

        if (rule.Expected.FirstLineIndentChars is double expectedIndentChars)
        {
            var actualIndentChars = (double)format.CharacterUnitFirstLineIndent;
            if (!ApproximatelyEquals(actualIndentChars, expectedIndentChars, 0.2))
            {
                var status = IssueStatus.Detected;
                if (applyFixes)
                {
                    format.CharacterUnitFirstLineIndent = (float)expectedIndentChars;
                    status = IssueStatus.Fixed;
                }

                issues.Add(CreateParagraphIssue(rule, "FirstLineIndentChars", expectedIndentChars.ToString("0.##"), actualIndentChars.ToString("0.##"), status, locator));
            }
        }

        if (rule.Expected.LineSpacingMultiple is double expectedLineSpacingMultiple)
        {
            var expectedSpacingPoints = document.Application.LinesToPoints((float)expectedLineSpacingMultiple);
            var actualSpacingPoints = (double)format.LineSpacing;
            var spacingMatches = format.LineSpacingRule == WdLineSpacing.wdLineSpaceMultiple &&
                                 ApproximatelyEquals(actualSpacingPoints, expectedSpacingPoints, 0.8);

            if (!spacingMatches)
            {
                var status = IssueStatus.Detected;
                if (applyFixes)
                {
                    format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                    format.LineSpacing = expectedSpacingPoints;
                    status = IssueStatus.Fixed;
                }

                issues.Add(CreateParagraphIssue(rule, "LineSpacing", $"{expectedLineSpacingMultiple:0.##}x", $"{actualSpacingPoints:0.##} pt / {format.LineSpacingRule}", status, locator));
            }
        }
    }

    private static void ApplyFontColorRule(
        ParagraphFormattingRule rule,
        MsWord.Range range,
        MsWord.Font font,
        bool applyFixes,
        bool hasPageNumberField,
        MsWord.WdColor expectedColor,
        NodeLocator locator,
        ICollection<IssueRecord> issues)
    {
        if (rule.ApplyToPageNumberParagraph && hasPageNumberField)
        {
            var actualColor = GetPageNumberColor(range);
            if (actualColor == expectedColor)
            {
                return;
            }

            var status = IssueStatus.Detected;
            if (applyFixes)
            {
                ApplyColorToPageFields(range, expectedColor);
                status = IssueStatus.Fixed;
            }

            issues.Add(CreateParagraphIssue(
                rule,
                "PageNumberColor",
                DescribeColor(expectedColor),
                DescribeColor(actualColor),
                status,
                locator));

            return;
        }

        var actualParagraphColor = (MsWord.WdColor)font.Color;
        if (actualParagraphColor == expectedColor)
        {
            return;
        }

        var paragraphStatus = IssueStatus.Detected;
        if (applyFixes)
        {
            font.Color = expectedColor;
            paragraphStatus = IssueStatus.Fixed;
        }

        issues.Add(CreateParagraphIssue(
            rule,
            "FontColor",
            DescribeColor(expectedColor),
            DescribeColor(actualParagraphColor),
            paragraphStatus,
            locator));
    }

    private static ParagraphFormattingRule? ResolveRule(
        RuleProfile profile,
        string storyScope,
        string styleName,
        int? outlineLevel,
        string text,
        bool hasPageNumberField)
    {
        foreach (var rule in profile.ParagraphRules)
        {
            if (!MatchesStoryScope(rule, storyScope))
            {
                continue;
            }

            if (rule.ApplyToPageNumberParagraph && hasPageNumberField)
            {
                return rule;
            }

            if (rule.StyleNames.Exists(candidate => StringEquals(candidate, styleName)))
            {
                return rule;
            }

            if (outlineLevel is int level && rule.OutlineLevels.Contains(level))
            {
                return rule;
            }
        }

        return profile.ParagraphRules.FirstOrDefault(rule =>
            MatchesStoryScope(rule, storyScope) &&
            !rule.ApplyToPageNumberParagraph &&
            rule.ApplyToAnyNonEmptyParagraph &&
            !string.IsNullOrWhiteSpace(text));
    }

    private static bool MatchesStoryScope(ParagraphFormattingRule rule, string storyScope)
    {
        return rule.StoryScopes.Count == 0 ||
               rule.StoryScopes.Exists(scope => StringEquals(scope, storyScope));
    }

    private static bool ContainsPageNumberField(MsWord.Range range)
    {
        foreach (MsWord.Field field in range.Fields)
        {
            try
            {
                if (field.Type == WdFieldType.wdFieldPage || field.Type == WdFieldType.wdFieldNumPages)
                {
                    return true;
                }
            }
            finally
            {
                ReleaseComObject(field);
            }
        }

        return false;
    }

    private static MsWord.WdColor GetPageNumberColor(MsWord.Range range)
    {
        foreach (MsWord.Field field in range.Fields)
        {
            try
            {
                if (field.Type == WdFieldType.wdFieldPage || field.Type == WdFieldType.wdFieldNumPages)
                {
                    return (MsWord.WdColor)field.Result.Font.Color;
                }
            }
            finally
            {
                ReleaseComObject(field);
            }
        }

        return WdColor.wdColorAutomatic;
    }

    private static void ApplyColorToPageFields(MsWord.Range range, MsWord.WdColor color)
    {
        foreach (MsWord.Field field in range.Fields)
        {
            try
            {
                if (field.Type == WdFieldType.wdFieldPage || field.Type == WdFieldType.wdFieldNumPages)
                {
                    field.Result.Font.Color = color;
                }
            }
            finally
            {
                ReleaseComObject(field);
            }
        }
    }

    private static NodeLocator CreateParagraphLocator(MsWord.Paragraph paragraph, int paragraphIndex, string styleName, string text)
    {
        return new NodeLocator
        {
            NodeKind = NodeKind.Paragraph,
            ParagraphIndex = paragraphIndex,
            RangeStart = paragraph.Range.Start,
            RangeEnd = paragraph.Range.End,
            PageNumber = TryGetPageNumber(paragraph.Range),
            StyleName = styleName,
            Snippet = text.Length > 60 ? text[..60] : text
        };
    }

    private static int? TryGetPageNumber(MsWord.Range range)
    {
        try
        {
            return Convert.ToInt32(range.get_Information(WdInformation.wdActiveEndPageNumber));
        }
        catch
        {
            return null;
        }
    }

    private static IssueRecord CreateDocumentIssue(
        int sectionIndex,
        string ruleId,
        string ruleName,
        IssueSeverity severity,
        string propertyName,
        string expectedValue,
        string actualValue,
        bool fixedByProcessor)
    {
        return new IssueRecord
        {
            RuleId = ruleId,
            RuleName = ruleName,
            Severity = severity,
            PropertyName = propertyName,
            ExpectedValue = expectedValue,
            ActualValue = actualValue,
            CanAutoFix = true,
            Status = fixedByProcessor ? IssueStatus.Fixed : IssueStatus.Detected,
            Location = new NodeLocator
            {
                NodeKind = NodeKind.Section,
                SectionIndex = sectionIndex,
                Snippet = $"Section {sectionIndex}"
            }
        };
    }

    private static IssueRecord CreateParagraphIssue(
        ParagraphFormattingRule rule,
        string propertyName,
        string expectedValue,
        string actualValue,
        IssueStatus status,
        NodeLocator locator)
    {
        return new IssueRecord
        {
            RuleId = rule.RuleId,
            RuleName = rule.DisplayName,
            Severity = rule.Severity,
            PropertyName = propertyName,
            ExpectedValue = expectedValue,
            ActualValue = actualValue,
            CanAutoFix = true,
            Status = status,
            Location = locator
        };
    }

    private static string ResolveFontName(MsWord.Font font)
    {
        return string.IsNullOrWhiteSpace(font.NameFarEast) ? font.Name : font.NameFarEast;
    }

    private static int? ConvertOutlineLevel(MsWord.WdOutlineLevel outlineLevel)
    {
        return outlineLevel switch
        {
            WdOutlineLevel.wdOutlineLevel1 => 1,
            WdOutlineLevel.wdOutlineLevel2 => 2,
            WdOutlineLevel.wdOutlineLevel3 => 3,
            WdOutlineLevel.wdOutlineLevel4 => 4,
            WdOutlineLevel.wdOutlineLevel5 => 5,
            WdOutlineLevel.wdOutlineLevel6 => 6,
            WdOutlineLevel.wdOutlineLevel7 => 7,
            WdOutlineLevel.wdOutlineLevel8 => 8,
            WdOutlineLevel.wdOutlineLevel9 => 9,
            _ => null
        };
    }

    private static bool TryResolveColor(string value, out MsWord.WdColor color)
    {
        var normalized = value.Trim();

        if (TryResolveNamedColor(normalized, out color))
        {
            return true;
        }

        try
        {
            var parsed = ColorTranslator.FromHtml(normalized);
            color = (MsWord.WdColor)ColorTranslator.ToOle(parsed);
            return true;
        }
        catch
        {
            color = WdColor.wdColorAutomatic;
            return false;
        }
    }

    private static bool TryResolveNamedColor(string value, out MsWord.WdColor color)
    {
        switch (value.Trim().ToLowerInvariant())
        {
            case "red":
            case "红":
            case "红色":
                color = WdColor.wdColorRed;
                return true;
            case "blue":
            case "蓝":
            case "蓝色":
                color = WdColor.wdColorBlue;
                return true;
            case "green":
            case "绿":
            case "绿色":
                color = WdColor.wdColorGreen;
                return true;
            case "orange":
            case "橙":
            case "橙色":
                color = (MsWord.WdColor)ColorTranslator.ToOle(Color.Orange);
                return true;
            case "purple":
            case "紫":
            case "紫色":
                color = (MsWord.WdColor)ColorTranslator.ToOle(Color.Purple);
                return true;
            case "teal":
                color = (MsWord.WdColor)ColorTranslator.ToOle(Color.Teal);
                return true;
            case "brown":
                color = (MsWord.WdColor)ColorTranslator.ToOle(Color.Brown);
                return true;
            case "gray":
            case "grey":
            case "灰":
            case "灰色":
                color = (MsWord.WdColor)ColorTranslator.ToOle(Color.Gray);
                return true;
            case "black":
            case "黑":
            case "黑色":
                color = WdColor.wdColorBlack;
                return true;
            default:
                color = WdColor.wdColorAutomatic;
                return false;
        }
    }

    private static string DescribeColor(MsWord.WdColor color)
    {
        try
        {
            var drawingColor = ColorTranslator.FromOle((int)color);
            return $"#{drawingColor.R:X2}{drawingColor.G:X2}{drawingColor.B:X2}";
        }
        catch
        {
            return color.ToString();
        }
    }

    private static ParagraphAlignmentKind ConvertAlignment(MsWord.WdParagraphAlignment alignment)
    {
        return alignment switch
        {
            WdParagraphAlignment.wdAlignParagraphCenter => ParagraphAlignmentKind.Center,
            WdParagraphAlignment.wdAlignParagraphRight => ParagraphAlignmentKind.Right,
            WdParagraphAlignment.wdAlignParagraphJustify => ParagraphAlignmentKind.Justify,
            _ => ParagraphAlignmentKind.Left
        };
    }

    private static MsWord.WdParagraphAlignment ConvertAlignment(ParagraphAlignmentKind alignment)
    {
        return alignment switch
        {
            ParagraphAlignmentKind.Center => WdParagraphAlignment.wdAlignParagraphCenter,
            ParagraphAlignmentKind.Right => WdParagraphAlignment.wdAlignParagraphRight,
            ParagraphAlignmentKind.Justify => WdParagraphAlignment.wdAlignParagraphJustify,
            _ => WdParagraphAlignment.wdAlignParagraphLeft
        };
    }

    private static void UpdateTablesOfContents(MsWord.Document document)
    {
        foreach (MsWord.TableOfContents tableOfContents in document.TablesOfContents)
        {
            try
            {
                tableOfContents.Update();
            }
            finally
            {
                ReleaseComObject(tableOfContents);
            }
        }
    }

    private static string BuildFixedOutputPath(string inputPath, string outputDirectory)
    {
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(outputDirectory, $"{Path.GetFileNameWithoutExtension(inputPath)}.{stamp}.fixed.docx");
    }

    private static string NormalizeText(string? text)
    {
        return (text ?? string.Empty)
            .Replace("\r", " ")
            .Replace("\a", " ")
            .Replace("\n", " ")
            .Trim();
    }

    private static bool StringEquals(string? left, string? right)
    {
        return string.Equals(left?.Trim(), right?.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static bool ApproximatelyEquals(double actual, double expected, double tolerance)
    {
        return Math.Abs(actual - expected) <= tolerance;
    }

    private static float CentimetersToPoints(double centimeters)
    {
        return (float)(centimeters * 28.3464567);
    }

    private static double PointsToCentimeters(double points)
    {
        return points / 28.3464567;
    }

    private static void ReleaseComObject(object? instance)
    {
        if (instance is not null && Marshal.IsComObject(instance))
        {
            Marshal.ReleaseComObject(instance);
        }
    }

    private sealed class ParagraphScanStats
    {
        public HashSet<string> ProcessedParagraphKeys { get; } = new(StringComparer.Ordinal);

        public HashSet<string> StyleNames { get; } = new(StringComparer.OrdinalIgnoreCase);

        public HashSet<string> StoryScopes { get; } = new(StringComparer.OrdinalIgnoreCase);

        public int Scanned { get; set; }

        public int Matched { get; set; }
    }
}
