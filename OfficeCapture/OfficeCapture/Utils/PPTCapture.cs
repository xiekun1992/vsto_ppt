using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeCapture.Models;
using Office = Microsoft.Office.Core;
using System.Diagnostics;

namespace OfficeCapture.Utils
{

    class PPTCapture
    {
        private static PowerPoint.Application application = new PowerPoint.Application();
        public static List<PPTInfo> filenames = new List<PPTInfo>();
        private static string exportRootDir = null;

        public static void Capture(int width, int height)
        {
            try
            {
                exportRootDir = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "export");
                if (!System.IO.Directory.Exists(exportRootDir))
                {
                    System.IO.Directory.CreateDirectory(exportRootDir);
                }
                PowerPoint.Presentations pres = application.Presentations;
                foreach (PPTInfo filename in filenames)
                {
                    PowerPoint.Presentation pre = pres.Open(filename.FullPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);

                    string exportFileDir = System.IO.Path.Combine(exportRootDir, filename.Filename);
                    if (!System.IO.Directory.Exists(exportFileDir))
                    {
                        System.IO.Directory.CreateDirectory(exportFileDir);
                    }
                    foreach (PowerPoint.Slide slide in pre.Slides)
                    {
                        slide.Export(System.IO.Path.Combine(exportFileDir, $"{slide.SlideIndex}.jpg"), "jpg", width, height);
                    }
                    pre.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static void AddFilename(string filename, string fullPath)
        {
            filenames.Add(new PPTInfo(filename, fullPath));
        }
        public static void recurseFolderForPPT(string folder)
        {
            File.searchPPTInFolder(folder, filenames);
        }
        public static void Clear()
        {
            filenames.Clear();

        }
        public static void openExportDir()
        {
            if (exportRootDir == null)
            {
                exportRootDir = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "export");
                if (!System.IO.Directory.Exists(exportRootDir))
                {
                    System.IO.Directory.CreateDirectory(exportRootDir);
                }
            }
            Process.Start(exportRootDir);
        }
    }
}
