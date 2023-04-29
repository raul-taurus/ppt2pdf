using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

if (args.Length > 0)
{
    var path = Path.GetFullPath(args[0]);
    if (Directory.Exists(path))
    {
        using var app = new PowerPoint();
        foreach (var pptFile in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
        {
            var ext = Path.GetExtension(pptFile).ToUpperInvariant();
            if (ext == ".PPTX" || ext == ".PPT")
            {
                Console.WriteLine($"ConvertToPdf: {pptFile}");
                app.ConvertToPdf(pptFile);
            }
        }
    }
    else
    {
        if (File.Exists(path))
        {
            using var app = new PowerPoint();
            app.ConvertToPdf(path);
        }
        else
        {
            Console.Error.WriteLine($"Cannot find file: {path}");
            return;
        }
    }
}
else
{
    Console.WriteLine($"Usae: exec <ppt filename>");
    return;
}

Console.WriteLine("Bye!");
