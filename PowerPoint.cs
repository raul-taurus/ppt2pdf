using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

class PowerPoint : IDisposable
{
    ApplicationClass _app;
    Presentations _presentations;

    public PowerPoint()
    {
        Console.WriteLine("Start PowerPoint Application...");
        _app = new ApplicationClass();
        _presentations = _app.Presentations;
    }

    public void ConvertToPdf(string pptFile)
    {
        Console.WriteLine("Open PPT file...");
        var ppt = _presentations.Open(pptFile,
            ReadOnly: MsoTriState.msoTrue,
            Untitled: MsoTriState.msoFalse,
            WithWindow: MsoTriState.msoTrue);

        Console.WriteLine("Export as PDF...");
        string pdfFile = Path.ChangeExtension(pptFile, ".pdf");
        ppt.ExportAsFixedFormat2(pdfFile,
            PpFixedFormatType.ppFixedFormatTypePDF,
            Intent: PpFixedFormatIntent.ppFixedFormatIntentScreen,
            FrameSlides: MsoTriState.msoFalse,
            HandoutOrder: PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
            OutputType: PpPrintOutputType.ppPrintOutputSlides,
            PrintHiddenSlides: MsoTriState.msoFalse,
            PrintRange: null,
            RangeType: PpPrintRangeType.ppPrintAll,
            SlideShowName: "",
            IncludeDocProperties: true,
            KeepIRMSettings: true,
            DocStructureTags: false,
            BitmapMissingFonts: true,
            UseISO19005_1: true,
            IncludeMarkup: true
        );

        ppt.Close();

#pragma warning disable CA1416 // Validate platform compatibility
        Marshal.ReleaseComObject(ppt);
#pragma warning restore CA1416 // Validate platform compatibility
    }

    public void Dispose()
    {
#pragma warning disable CA1416 // Validate platform compatibility
        Marshal.ReleaseComObject(_presentations);
#pragma warning restore CA1416 // Validate platform compatibility
        Console.WriteLine("Quit PowerPoint Application...");
        _app.Quit();
#pragma warning disable CA1416 // Validate platform compatibility
        Marshal.ReleaseComObject(_app);
#pragma warning restore CA1416 // Validate platform compatibility
    }
}
