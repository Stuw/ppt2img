using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;


namespace ppt2img
{
    class Program
    {
        static int width = 0;
        static int height = 0;
        static String imgType = "png";
        static String outDir = ".";
        static String inPpt = "";
        static String baseName = "test";

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine(@"Usage: ppt2img <ppt|pptx> [options]
Option:
    -t|--type <png|jpg>
    -o|--output <dir>");
                return;
            }

            try
            {
                for (int i = 0; i < args.Length; ++i)
                {
                    if (args[i] == "--type" || args[i] == "-t")
                    {
                        ++i;
                        imgType = args[i];
                    }
                    else if (args[i] == "--output" || args[i] == "-o")
                    {
                        ++i;
                        outDir = args[i];
                    }
                    else if (inPpt.Length == 0)
                        inPpt = args[i];
                    else
                        throw new Exception("Unknow option '" + args[i] + "'");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Invalid args");
                Console.WriteLine("{0}", e.Message);
                return;
            }

            outDir = Path.GetFullPath(outDir);
            inPpt = Path.GetFullPath(inPpt);
            baseName = Path.GetFileNameWithoutExtension(inPpt);

            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(inPpt,
                                                                                                     MsoTriState.msoTrue /* ReadOnly=true */,
                                                                                                     MsoTriState.msoTrue /* Untitled=true */,
                                                                                                     MsoTriState.msoFalse /* WithWindow=false */);

            int count = presentation.Slides.Count;
            for (int i = 0; i < count; i++)
            {
                Console.WriteLine("Saving slide {0} of {1}...", i + 1, count);
                String outName = String.Format(@"{0}\{1}_slide{2}.{3}", outDir, baseName, i, imgType);
                try
                {
                    presentation.Slides[i + 1].Export(outName, imgType, width, height);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Failed to export slide {0}", i + 1);
                    Console.WriteLine("{0}", e.Message);
                    break;
                }
            }

            PowerPoint_App.Quit();

            Console.WriteLine("Done");
        }
    }
}
