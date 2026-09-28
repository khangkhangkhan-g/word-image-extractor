using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using WordImageExtractor.Models;
using WordImageExtractor.Services;

namespace WordImageExtractor;

public partial class MainWindow : Window
{
    private readonly ExtractionService _service = new();
    private CancellationTokenSource? _cts;
    private CancellationTokenSource? _scanCts;
    private string? _selectedFile;
    private string? _lastOutput;
    private bool _isBusy;
    private bool _isVietnamese;
    private int _scanGeneration;

    public MainWindow()
    {
        InitializeComponent();
        Loaded += MainWindow_Loaded;
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        UpdateWordStatus();
        ApplyLanguage();
    }

    private void UpdateWordStatus()
    {
        var installed = _service.IsWordInstalled;
        WordStatusText.Text = _isVietnamese
            ? installed ? "Word: đã phát hiện" : "Word: chưa phát hiện"
            : installed ? "Word: detected" : "Word: not detected";

        if (!installed && PaginationCombo.SelectedItem is ComboBoxItem item && (string?)item.Tag == "AccurateWord")
            PaginationCombo.SelectedIndex = 1;
    }

    private async void AddFileButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = _isVietnamese ? "Chọn file Word" : "Select a Word document",
            Filter = "Word documents (*.docx;*.docm;*.doc)|*.docx;*.docm;*.doc",
            Multiselect = false,
        };
        if (dialog.ShowDialog() == true)
            await LoadDocumentAsync(dialog.FileName);
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            e.Effects = files.Length == 1 && IsSupported(files[0]) ? DragDropEffects.Copy : DragDropEffects.None;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }
        e.Handled = true;
    }

    private async void Window_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        var files = (string[])e.Data.GetData(DataFormats.FileDrop);
        if (files.Length == 1 && IsSupported(files[0]))
            await LoadDocumentAsync(files[0]);
    }

    private async Task LoadDocumentAsync(string path)
    {
        if (_isBusy) return;
        _scanCts?.Cancel();
        _scanCts?.Dispose();
        _scanCts = new CancellationTokenSource();
        var scanToken = _scanCts.Token;
        var generation = ++_scanGeneration;
        _selectedFile = path;
        _lastOutput = null;
        OpenFolderButton.IsEnabled = false;
        FileValueText.Text = Path.GetFileName(path);
        ExtractButton.IsEnabled = false;
        ProgressBar.Value = 0;
        ProgressText.Text = string.Empty;

        var fi = new FileInfo(path);
        InfoValueText.Text = _isVietnamese
            ? $"{FormatBytes(fi.Length)} · đang quét cấu trúc..."
            : $"{FormatBytes(fi.Length)} · scanning structure...";
        StatusText.Text = _isVietnamese ? "Đang phân tích file..." : "Analyzing document...";

        try
        {
            if (Path.GetExtension(path).Equals(".doc", StringComparison.OrdinalIgnoreCase))
            {
                if (generation != _scanGeneration) return;
                InfoValueText.Text = _isVietnamese
                    ? $"{FormatBytes(fi.Length)} · file .doc sẽ được chuyển đổi tạm thời bằng Microsoft Word khi trích xuất"
                    : $"{FormatBytes(fi.Length)} · legacy .doc will be converted temporarily by Microsoft Word during extraction";
                if (_service.IsWordInstalled)
                {
                    StatusText.Text = _isVietnamese ? "Sẵn sàng." : "Ready.";
                    ExtractButton.IsEnabled = true;
                }
                else
                {
                    StatusText.Text = _isVietnamese
                        ? "File .doc cần Microsoft Word để chuyển đổi. Hãy cài Word hoặc dùng file .docx/.docm."
                        : "Legacy .doc requires Microsoft Word for conversion. Install Word or use .docx/.docm.";
                    ExtractButton.IsEnabled = false;
                }
                return;
            }

            var scan = await _service.QuickScanAsync(path, HeaderFooterCheck.IsChecked == true, scanToken);
            if (generation != _scanGeneration) return;
            InfoValueText.Text = _isVietnamese
                ? $"{FormatBytes(fi.Length)} · {scan.BodyImageCount:N0} ảnh trong nội dung" + (scan.StaticImageCount > 0 ? $" · {scan.StaticImageCount:N0} ảnh header/footer" : string.Empty)
                : $"{FormatBytes(fi.Length)} · {scan.BodyImageCount:N0} body images" + (scan.StaticImageCount > 0 ? $" · {scan.StaticImageCount:N0} header/footer assets" : string.Empty);
            StatusText.Text = _isVietnamese ? "Sẵn sàng." : "Ready.";
            ExtractButton.IsEnabled = true;
        }
        catch (OperationCanceledException)
        {
            if (generation == _scanGeneration)
                StatusText.Text = _isVietnamese ? "Đã dừng quét file." : "Document scan cancelled.";
        }
        catch (Exception ex)
        {
            if (generation != _scanGeneration) return;
            InfoValueText.Text = FormatBytes(fi.Length);
            StatusText.Text = (_isVietnamese ? "Không thể quét file: " : "Could not scan file: ") + ex.Message;
            ExtractButton.IsEnabled = false;
        }
    }

    private async void HeaderFooterCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (_selectedFile is not null && !_isBusy && IsLoaded)
            await LoadDocumentAsync(_selectedFile);
    }

    private async void ExtractButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedFile is null || _isBusy) return;

        if (!File.Exists(_selectedFile))
        {
            MessageBox.Show(_isVietnamese ? "File nguồn không còn tồn tại." : "The source file no longer exists.");
            return;
        }

        if (GetPaginationMode() == PaginationMode.AccurateWord && !_service.IsWordInstalled)
        {
            MessageBox.Show(
                _isVietnamese
                    ? "Chế độ dò trang chính xác cần Microsoft Word. Hãy chọn chế độ Portable hoặc cài Microsoft Word."
                    : "Accurate page detection requires Microsoft Word. Choose Portable mode or install Microsoft Word.",
                "Word Image Extractor",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var save = new SaveFileDialog
        {
            Title = _isVietnamese ? "Lưu ZIP ảnh" : "Save image ZIP",
            Filter = "ZIP archive (*.zip)|*.zip",
            AddExtension = true,
            DefaultExt = ".zip",
            FileName = $"{Path.GetFileNameWithoutExtension(_selectedFile)}_Images.zip",
        };
        if (save.ShowDialog() != true) return;

        var options = new ExtractionOptions
        {
            InputPath = _selectedFile,
            OutputZipPath = save.FileName,
            OutputFormat = GetOutputFormat(),
            Organization = GetOrganizationMode(),
            Pagination = GetPaginationMode(),
            IncludeHeadersFooters = HeaderFooterCheck.IsChecked == true,
            IncludeMetadata = MetadataCheck.IsChecked == true,
            IncludeCsvManifest = CsvCheck.IsChecked == true,
            IncludeJsonManifest = JsonCheck.IsChecked == true,
            JpegQuality = (int)Math.Round(QualitySlider.Value),
        };

        SetBusy(true);
        _cts = new CancellationTokenSource();
        var progress = new Progress<ExtractionProgress>(p =>
        {
            var pct = p.Total > 0 ? Math.Clamp(100.0 * p.Processed / p.Total, 0, 100) : 0;
            ProgressBar.Value = pct;
            StatusText.Text = LocalizeStage(p.Stage);
            ProgressText.Text = p.Total > 0
                ? $"{p.Processed:N0} / {p.Total:N0}" + (string.IsNullOrWhiteSpace(p.CurrentItem) ? string.Empty : $" · {p.CurrentItem}")
                : p.CurrentItem;
        });

        try
        {
            var result = await _service.ExtractAsync(options, progress, _cts.Token);
            _lastOutput = result.OutputPath;
            OpenFolderButton.IsEnabled = true;
            ProgressBar.Value = 100;

            var warningPart = result.Warnings.Count > 0
                ? (_isVietnamese ? $" · {result.Warnings.Count} cảnh báo" : $" · {result.Warnings.Count} warning(s)")
                : string.Empty;

            StatusText.Text = _isVietnamese
                ? $"Hoàn tất. Đã ghi {result.Written:N0}/{result.TotalInstances:N0} ảnh{warningPart}."
                : $"Complete. Wrote {result.Written:N0}/{result.TotalInstances:N0} images{warningPart}.";
            ProgressText.Text = result.OutputPath;

            if (result.Failed > 0 || result.Warnings.Count > 0)
            {
                var firstWarnings = string.Join(Environment.NewLine, result.Warnings.Take(5).Select(x => "- " + x));
                MessageBox.Show(
                    (_isVietnamese
                        ? $"Hoàn tất với {result.Failed} lỗi và {result.Warnings.Count} cảnh báo.\n\n"
                        : $"Completed with {result.Failed} failure(s) and {result.Warnings.Count} warning(s).\n\n") + firstWarnings,
                    "Word Image Extractor",
                    MessageBoxButton.OK,
                    result.Failed > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = _isVietnamese ? "Đã hủy. File ZIP chưa hoàn tất đã được dọn dẹp." : "Cancelled. Incomplete ZIP was cleaned up.";
            ProgressText.Text = string.Empty;
            ProgressBar.Value = 0;
        }
        catch (Exception ex)
        {
            StatusText.Text = (_isVietnamese ? "Trích xuất thất bại: " : "Extraction failed: ") + ex.Message;
            MessageBox.Show(ex.Message, "Word Image Extractor", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _cts?.Dispose();
            _cts = null;
            SetBusy(false);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => _cts?.Cancel();

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (_lastOutput is null || !File.Exists(_lastOutput)) return;
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{_lastOutput}\"") { UseShellExecute = true });
        }
        catch
        {
            Process.Start(new ProcessStartInfo(Path.GetDirectoryName(_lastOutput)!) { UseShellExecute = true });
        }
    }

    private void FormatCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!IsLoaded) return;
        QualitySlider.IsEnabled = GetOutputFormat() == OutputImageFormat.Jpeg;
    }

    private void QualitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (QualityValueText is not null)
            QualityValueText.Text = ((int)Math.Round(e.NewValue)).ToString(CultureInfo.InvariantCulture);
    }

    private async void LanguageCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LanguageCombo.SelectedItem is ComboBoxItem item)
            _isVietnamese = string.Equals((string?)item.Tag, "vi", StringComparison.OrdinalIgnoreCase);
        if (IsLoaded)
        {
            ApplyLanguage();
            UpdateWordStatus();
            if (_selectedFile is not null && !_isBusy)
                await LoadDocumentAsync(_selectedFile);
        }
    }

    private void ApplyLanguage()
    {
        SubtitleText.Text = _isVietnamese ? "Trích xuất ảnh chất lượng gốc từ tài liệu Word" : "Extract original-quality images from Word documents";
        DocumentTitleText.Text = _isVietnamese ? "Tài liệu" : "Document";
        DropText.Text = _isVietnamese ? "Thả file Word vào đây" : "Drop a Word document here";
        OrText.Text = _isVietnamese ? "hoặc" : "or";
        AddFileButton.Content = _isVietnamese ? "Thêm file" : "Add file";
        FileLabelText.Text = _isVietnamese ? "File" : "File";
        InfoLabelText.Text = _isVietnamese ? "Thông tin" : "Info";
        if (_selectedFile is null) FileValueText.Text = _isVietnamese ? "Chưa chọn file" : "No file selected";
        OutputTitleText.Text = _isVietnamese ? "Đầu ra" : "Output";
        FormatLabelText.Text = _isVietnamese ? "Định dạng ảnh" : "Output format";
        OrganizationLabelText.Text = _isVietnamese ? "Sắp xếp" : "Organization";
        PaginationLabelText.Text = _isVietnamese ? "Dò số trang" : "Page detection";
        QualityLabelText.Text = _isVietnamese ? "Chất lượng JPEG" : "JPEG quality";
        OptionsTitleText.Text = _isVietnamese ? "Tùy chọn" : "Options";
        MetadataCheck.Content = _isVietnamese ? "Kèm metadata ảnh" : "Include image metadata";
        CsvCheck.Content = _isVietnamese ? "Tạo manifest CSV" : "Create manifest CSV";
        JsonCheck.Content = _isVietnamese ? "Tạo manifest JSON" : "Create manifest JSON";
        HeaderFooterCheck.Content = _isVietnamese ? "Kèm ảnh header/footer một lần" : "Include header/footer image assets once";
        HeaderFooterHintText.Text = _isVietnamese
            ? "Ảnh header/footer chỉ xuất một lần vì Word có thể lặp lại cùng một asset trên nhiều trang."
            : "Header/footer assets are exported once because Word may repeat the same asset across many pages.";
        StatusTitleText.Text = _isVietnamese ? "Trạng thái" : "Status";
        if (!_isBusy && _selectedFile is null) StatusText.Text = _isVietnamese ? "Sẵn sàng." : "Ready.";
        ExtractButton.Content = _isVietnamese ? "Trích xuất ảnh" : "Extract images";
        CancelButton.Content = _isVietnamese ? "Hủy" : "Cancel";
        OpenFolderButton.Content = _isVietnamese ? "Mở thư mục" : "Open folder";
        SafetyText.Text = _isVietnamese
            ? "Chế độ Original giữ nguyên byte ảnh nhúng. Nếu chuyển định dạng thất bại, app giữ file gốc thay vì làm mất ảnh."
            : "Original mode preserves embedded bytes. Failed conversions fall back to the original file instead of dropping the image.";

        SetComboText(FormatCombo, 0, _isVietnamese ? "Original - Chất lượng tốt nhất" : "Original - Best Quality");
        SetComboText(FormatCombo, 1, _isVietnamese ? "PNG - Không mất dữ liệu" : "PNG - Lossless");
        SetComboText(OrganizationCombo, 0, _isVietnamese ? "Chia theo trang" : "Group by page");
        SetComboText(OrganizationCombo, 1, _isVietnamese ? "Danh sách tuần tự" : "Flat sequential list");
        SetComboText(PaginationCombo, 0, _isVietnamese ? "Chính xác - Microsoft Word" : "Accurate - Microsoft Word");
        SetComboText(PaginationCombo, 1, _isVietnamese ? "Portable - theo thứ tự tài liệu" : "Portable - document order");
    }

    private string LocalizeStage(string stage)
    {
        if (!_isVietnamese) return stage;
        return stage switch
        {
            "Preparing document" => "Đang chuẩn bị tài liệu",
            "Scanning images" => "Đang quét ảnh",
            "Mapping Word pages" => "Đang xác định số trang bằng Word",
            "Extracting images" => "Đang trích xuất ảnh",
            _ => stage,
        };
    }

    private void SetBusy(bool busy)
    {
        _isBusy = busy;
        AddFileButton.IsEnabled = !busy;
        ExtractButton.IsEnabled = !busy && CanExtractSelectedFile();
        CancelButton.IsEnabled = busy;
        LanguageCombo.IsEnabled = !busy;
        FormatCombo.IsEnabled = !busy;
        OrganizationCombo.IsEnabled = !busy;
        PaginationCombo.IsEnabled = !busy;
        HeaderFooterCheck.IsEnabled = !busy;
        MetadataCheck.IsEnabled = !busy;
        CsvCheck.IsEnabled = !busy;
        JsonCheck.IsEnabled = !busy;
        QualitySlider.IsEnabled = !busy && GetOutputFormat() == OutputImageFormat.Jpeg;
    }

    private OutputImageFormat GetOutputFormat() => Enum.Parse<OutputImageFormat>((string)((ComboBoxItem)FormatCombo.SelectedItem).Tag);
    private OrganizationMode GetOrganizationMode() => Enum.Parse<OrganizationMode>((string)((ComboBoxItem)OrganizationCombo.SelectedItem).Tag);
    private PaginationMode GetPaginationMode() => Enum.Parse<PaginationMode>((string)((ComboBoxItem)PaginationCombo.SelectedItem).Tag);

    private bool CanExtractSelectedFile()
    {
        var selectedFile = _selectedFile;
        if (string.IsNullOrWhiteSpace(selectedFile) || !File.Exists(selectedFile)) return false;
        var ext = Path.GetExtension(selectedFile).ToLowerInvariant();
        if (ext == ".doc" && !_service.IsWordInstalled) return false;
        return IsSupported(selectedFile);
    }

    private static bool IsSupported(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext is ".docx" or ".docm" or ".doc";
    }

    private static void SetComboText(ComboBox combo, int index, string text)
    {
        if (index >= 0 && index < combo.Items.Count && combo.Items[index] is ComboBoxItem item)
            item.Content = text;
    }

    private static string FormatBytes(long bytes)
    {
        string[] units = ["B", "KB", "MB", "GB", "TB"];
        var value = (double)bytes;
        var i = 0;
        while (value >= 1024 && i < units.Length - 1)
        {
            value /= 1024;
            i++;
        }
        return $"{value:0.##} {units[i]}";
    }

    private void Copyright_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        e.Handled = true;
    }
}
