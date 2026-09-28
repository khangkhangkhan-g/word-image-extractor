using System.IO;
using System.Windows.Media.Imaging;
using WordImageExtractor.Models;

namespace WordImageExtractor.Services;

public sealed class ImageConversionService
{
    public sealed record ConversionResult(bool Success, byte[]? Bytes, string Extension, int? PixelWidth, int? PixelHeight, string? Error);

    public ConversionResult Convert(byte[] sourceBytes, OutputImageFormat format, int jpegQuality, string originalExtension)
    {
        if (format == OutputImageFormat.Original)
        {
            var dims = TryGetDimensions(sourceBytes);
            return new ConversionResult(true, sourceBytes, NormalizeExtension(originalExtension), dims.width, dims.height, null);
        }

        try
        {
            using var input = new MemoryStream(sourceBytes, writable: false);
            var decoder = BitmapDecoder.Create(input, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
            var frame = decoder.Frames[0];

            BitmapEncoder encoder = format switch
            {
                OutputImageFormat.Png => new PngBitmapEncoder(),
                OutputImageFormat.Jpeg => new JpegBitmapEncoder { QualityLevel = Math.Clamp(jpegQuality, 1, 100) },
                OutputImageFormat.Tiff => new TiffBitmapEncoder(),
                OutputImageFormat.Bmp => new BmpBitmapEncoder(),
                _ => throw new NotSupportedException($"Unsupported output format: {format}")
            };

            encoder.Frames.Add(BitmapFrame.Create(frame));
            using var output = new MemoryStream();
            encoder.Save(output);

            return new ConversionResult(
                true,
                output.ToArray(),
                ExtensionFor(format),
                frame.PixelWidth,
                frame.PixelHeight,
                null);
        }
        catch (Exception ex)
        {
            var dims = TryGetDimensions(sourceBytes);
            return new ConversionResult(
                false,
                null,
                NormalizeExtension(originalExtension),
                dims.width,
                dims.height,
                ex.Message);
        }
    }

    public (int? width, int? height) TryGetDimensions(byte[] sourceBytes)
    {
        try
        {
            using var input = new MemoryStream(sourceBytes, writable: false);
            var decoder = BitmapDecoder.Create(input, BitmapCreateOptions.DelayCreation, BitmapCacheOption.OnLoad);
            var frame = decoder.Frames[0];
            return (frame.PixelWidth, frame.PixelHeight);
        }
        catch
        {
            return (null, null);
        }
    }

    public static string ExtensionFor(OutputImageFormat format) => format switch
    {
        OutputImageFormat.Png => ".png",
        OutputImageFormat.Jpeg => ".jpg",
        OutputImageFormat.Tiff => ".tiff",
        OutputImageFormat.Bmp => ".bmp",
        _ => string.Empty,
    };

    public static string NormalizeExtension(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension)) return ".bin";
        return extension.StartsWith('.') ? extension.ToLowerInvariant() : $".{extension.ToLowerInvariant()}";
    }
}
