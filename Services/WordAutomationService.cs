using System.IO;
using System.Runtime.InteropServices;
using WordImageExtractor.Models;

namespace WordImageExtractor.Services;

public sealed class WordAutomationService
{
    // Word / Office COM constants used without an Interop NuGet dependency.
    private const int WdStatisticPages = 2;
    private const int WdActiveEndPageNumber = 3; // absolute physical page number
    private const int WdFormatDocumentDefault = 16; // .docx
    private const int MsoAutomationSecurityForceDisable = 3;

    private static readonly HashSet<int> InlinePictureTypes = [3, 4, 7, 8];
    private static readonly HashSet<int> FloatingPictureTypes = [11, 13];

    public bool IsWordInstalled()
    {
        try
        {
            return OperatingSystem.IsWindows() && Type.GetTypeFromProgID("Word.Application") is not null;
        }
        catch
        {
            return false;
        }
    }

    public Task<PreparedDocument> PrepareDocumentAsync(string inputPath, CancellationToken cancellationToken)
    {
        var extension = Path.GetExtension(inputPath).ToLowerInvariant();
        if (extension is ".docx" or ".docm")
        {
            return Task.FromResult(new PreparedDocument
            {
                OriginalPath = inputPath,
                WorkingPath = inputPath,
                IsTemporary = false,
            });
        }

        if (extension != ".doc")
            throw new NotSupportedException("Supported Word formats are .docx, .docm, and .doc.");

        if (!IsWordInstalled())
            throw new InvalidOperationException("Legacy .doc files require Microsoft Word to convert safely to .docx first.");

        var tempPath = Path.Combine(Path.GetTempPath(), $"WordImageExtractor_{Guid.NewGuid():N}.docx");
        return RunStaAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            dynamic? word = null;
            dynamic? doc = null;
            try
            {
                var wordType = Type.GetTypeFromProgID("Word.Application")
                    ?? throw new InvalidOperationException("Microsoft Word is not installed.");
                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0;
                try { word.AutomationSecurity = MsoAutomationSecurityForceDisable; } catch { }

                doc = word.Documents.Open(
                    FileName: inputPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: true,
                    NoEncodingDialog: true,
                    UpdateLinksAtOpen: false);

                cancellationToken.ThrowIfCancellationRequested();
                doc.SaveAs2(FileName: tempPath, FileFormat: WdFormatDocumentDefault, AddToRecentFiles: false);
                return new PreparedDocument
                {
                    OriginalPath = inputPath,
                    WorkingPath = tempPath,
                    IsTemporary = true,
                };
            }
            finally
            {
                try { if (doc is not null) doc.Close(SaveChanges: false); } catch { }
                try { if (word is not null) word.Quit(SaveChanges: false); } catch { }
                ReleaseCom(doc);
                ReleaseCom(word);
                ForceComCleanup();
            }
        }, cancellationToken);
    }

    public Task<PaginationResult> MapBodyImagePagesAsync(string documentPath, int expectedBodyImageCount, CancellationToken cancellationToken)
    {
        if (!IsWordInstalled())
        {
            return Task.FromResult(new PaginationResult
            {
                Success = false,
                Warning = "Microsoft Word is not installed. Accurate page mapping is unavailable."
            });
        }

        return RunStaAsync(() =>
        {
            dynamic? word = null;
            dynamic? doc = null;
            var objects = new List<(int start, int order, int page)>();
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                var wordType = Type.GetTypeFromProgID("Word.Application")
                    ?? throw new InvalidOperationException("Microsoft Word is not installed.");
                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0;
                word.ScreenUpdating = false;
                try { word.AutomationSecurity = MsoAutomationSecurityForceDisable; } catch { }

                doc = word.Documents.Open(
                    FileName: documentPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: true,
                    NoEncodingDialog: true,
                    UpdateLinksAtOpen: false);

                doc.Repaginate();
                var pageCount = Convert.ToInt32(doc.ComputeStatistics(WdStatisticPages));
                var ordinal = 0;

                dynamic? inlineShapes = null;
                try
                {
                    inlineShapes = doc.InlineShapes;
                    var count = Convert.ToInt32(inlineShapes.Count);
                    for (var i = 1; i <= count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        dynamic? shape = null;
                        dynamic? range = null;
                        try
                        {
                            shape = inlineShapes.Item(i);
                            var type = Convert.ToInt32(shape.Type);
                            if (!InlinePictureTypes.Contains(type)) continue;
                            range = shape.Range;
                            var start = Convert.ToInt32(range.Start);
                            var page = Convert.ToInt32(range.Information[WdActiveEndPageNumber]);
                            objects.Add((start, ordinal++, page));
                        }
                        finally
                        {
                            ReleaseCom(range);
                            ReleaseCom(shape);
                        }
                    }
                }
                finally
                {
                    ReleaseCom(inlineShapes);
                }

                dynamic? shapes = null;
                try
                {
                    shapes = doc.Shapes;
                    var count = Convert.ToInt32(shapes.Count);
                    for (var i = 1; i <= count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        dynamic? shape = null;
                        dynamic? anchor = null;
                        try
                        {
                            shape = shapes.Item(i);
                            var type = Convert.ToInt32(shape.Type);
                            if (!FloatingPictureTypes.Contains(type)) continue;
                            anchor = shape.Anchor;
                            var start = Convert.ToInt32(anchor.Start);
                            var page = Convert.ToInt32(anchor.Information[WdActiveEndPageNumber]);
                            objects.Add((start, ordinal++, page));
                        }
                        finally
                        {
                            ReleaseCom(anchor);
                            ReleaseCom(shape);
                        }
                    }
                }
                finally
                {
                    ReleaseCom(shapes);
                }

                var ordered = objects
                    .OrderBy(x => x.start)
                    .ThenBy(x => x.order)
                    .Select(x => x.page)
                    .ToArray();

                if (ordered.Length != expectedBodyImageCount)
                {
                    return new PaginationResult
                    {
                        Success = false,
                        PageCount = pageCount,
                        DetectedPictureCount = ordered.Length,
                        Pages = ordered,
                        Warning = $"Safe page mapping was not applied because Word detected {ordered.Length} picture objects while the DOCX package contains {expectedBodyImageCount} body image instances. The extractor will preserve document order instead of guessing page numbers."
                    };
                }

                return new PaginationResult
                {
                    Success = true,
                    PageCount = pageCount,
                    DetectedPictureCount = ordered.Length,
                    Pages = ordered,
                };
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                return new PaginationResult
                {
                    Success = false,
                    DetectedPictureCount = objects.Count,
                    Warning = $"Word pagination failed safely: {ex.Message}"
                };
            }
            finally
            {
                try { if (doc is not null) doc.Close(SaveChanges: false); } catch { }
                try { if (word is not null) word.Quit(SaveChanges: false); } catch { }
                ReleaseCom(doc);
                ReleaseCom(word);
                ForceComCleanup();
            }
        }, cancellationToken);
    }

    private static Task<T> RunStaAsync<T>(Func<T> action, CancellationToken cancellationToken)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                tcs.SetResult(action());
            }
            catch (OperationCanceledException oce)
            {
                tcs.SetCanceled(oce.CancellationToken);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        })
        {
            IsBackground = true,
            Name = "WordImageExtractor.OfficeSTA"
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return tcs.Task;
    }

    private static void ReleaseCom(object? value)
    {
        if (value is null) return;
        try
        {
            if (Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
        catch { }
    }

    private static void ForceComCleanup()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
