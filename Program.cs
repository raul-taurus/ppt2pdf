using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

if (args.Length > 0)
{
    if (Directory.Exists(args[0]))
    {
        using var app = new PowerPoint();
        foreach (var pptFile in Directory.EnumerateFiles(args[0], "*", SearchOption.AllDirectories))
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
        if (File.Exists(args[0]))
        {
            string pptFile = args[0];
            using var app = new PowerPoint();
            app.ConvertToPdf(pptFile);
        }
        else
        {
            Console.Error.WriteLine($"Cannot find file: {args[0]}");
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
