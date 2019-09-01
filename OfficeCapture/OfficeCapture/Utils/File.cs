using OfficeCapture.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace OfficeCapture.Utils
{
    class File
    {
        public static void searchPPTInFolder(string filename, List<PPTInfo> filenames)
        {
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(filename);
                FileInfo[] info = dirInfo.GetFiles("*.ppt");
                foreach (FileInfo f in info)
                {
                    // exclude hidden files
                    if (!f.Attributes.HasFlag(FileAttributes.Hidden))
                    {
                        filenames.Add(new PPTInfo(f.Name, f.FullName));
                    }
                }
                DirectoryInfo[] subDir = dirInfo.GetDirectories();
                foreach (DirectoryInfo dir in subDir)
                {
                    searchPPTInFolder(dir.FullName, filenames);
                }
            }
            catch (Exception ex){
                MessageBox.Show(ex.Message);
            }
        }
    }
}
