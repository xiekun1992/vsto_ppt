using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using System.Windows.Forms;
using System.IO;
using OfficeCapture.Utils;
using MessageBox = System.Windows.MessageBox;
using OfficeCapture.Models;

namespace OfficeCapture
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(width.Text) && !string.IsNullOrEmpty(height.Text))
            {
                PPTCapture.Capture(Int32.Parse(width.Text), Int32.Parse(height.Text));
            } else
            {
                PPTCapture.Capture(0, 0);
            }
        }

        private void Select_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult res = dialog.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                PPTCapture.AddFilename(dialog.SafeFileName , dialog.FileName);
                updateListBox();
            }
        }

        private void Select_Dir_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult res = dialog.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK)
            {
                PPTCapture.recurseFolderForPPT(dialog.SelectedPath);
                updateListBox();
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            PPTCapture.Clear();
            lbox.Items.Clear();
        }

        private void updateListBox()
        {
            for (int i = 0; i < PPTCapture.filenames.Count; i++)
            {
                PPTInfo info = PPTCapture.filenames[i];
                TextBlock textBlock = new TextBlock();
                textBlock.Text = $"{i + 1}. {info.Filename} - {info.FullPath}";
                lbox.Items.Add(textBlock);
            }
        }

        private void Open_Explorer_Click(object sender, RoutedEventArgs e)
        {
            PPTCapture.openExportDir();
        }
    }
}
