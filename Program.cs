using ClipImage;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics.Imaging;
using Windows.Storage.Streams;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        if (!UI.Initialize())
        {
            return;
        }

        try
        {
            string location = GetSaveLocation(args);
            ExecuteAsync(location).Wait();
        }
        catch (Exception e)
        {
            UI.ShowError($"予期せぬエラーが発生しました。{e.Message}");
            UI.ShowError(e.StackTrace ?? string.Empty);
        }
        finally
        {
            UI.UnInitialize();
        }
    }

    static string GetSaveLocation(string[] args)
    {
        if (args.Length > 0 && Directory.Exists(args[0]))
            return args[0];

        return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
    }

    static async Task ExecuteAsync(string location)
    {
        var content = Clipboard.GetContent();
        if (!content.Contains(StandardDataFormats.Bitmap))
        {
            UI.ShowError("クリップボードに画像データが存在しません。");
            return;
        }

        if (!UI.ShowSaveFiledialog(location, out var filePath) || string.IsNullOrEmpty(filePath))
            return;

        var bitmapRef = await content.GetBitmapAsync();
        using var stream = await bitmapRef.OpenReadAsync();
        var decoder = await BitmapDecoder.CreateAsync(stream);
        var pixelData = await decoder.GetPixelDataAsync();

        using var fileStream = new FileStream(filePath, FileMode.Create);
        using var randomAccessStream = new InMemoryRandomAccessStream();
        var encoder = await BitmapEncoder.CreateAsync(GetImageFormat(filePath), randomAccessStream);

        encoder.SetPixelData(
            decoder.BitmapPixelFormat,
            decoder.BitmapAlphaMode,
            decoder.PixelWidth,
            decoder.PixelHeight,
            decoder.DpiX,
            decoder.DpiY,
            pixelData.DetachPixelData());

        await encoder.FlushAsync();

        randomAccessStream.Seek(0);
        await randomAccessStream.AsStream().CopyToAsync(fileStream);
    }

    static Guid GetImageFormat(string filePath)
    {
        return Path.GetExtension(filePath).ToLower() switch
        {
            ".png" => BitmapEncoder.PngEncoderId,
            ".jpg" or ".jpeg" => BitmapEncoder.JpegEncoderId,
            _ => BitmapEncoder.PngEncoderId,
        };
    }
}
