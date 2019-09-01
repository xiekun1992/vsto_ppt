using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace OfficeCapture.Models
{
    class PPTInfo
    {
        private string filename;
        private string fullPath;

        public PPTInfo(string filename, string fullPath)
        {
            Filename = filename;
            FullPath = fullPath;
        }

        public string Filename { get => filename; set => filename = value; }
        public string FullPath { get => fullPath; set => fullPath = value; }
    }
}
