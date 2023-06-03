using Microsoft.Win32;
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
using System.IO.Compression;
using System.IO;
using SharpCompress.Archives;
using SharpCompress.Common;

namespace ArchiveDecompresser
{

    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }


        private void buttonZip_Click(object sender, RoutedEventArgs e)
        {
            OpenExplorer(".zip");
        }

        private void button7z_Click(object sender, RoutedEventArgs e)
        {
            OpenExplorer(".7z");
        }

        private void buttonTar_Click(object sender, RoutedEventArgs e)
        {
            OpenExplorer(".tar");
        }

        private void buttonRar_Click(object sender, RoutedEventArgs e)
        {
            OpenExplorer(".rar");
        }

        private void OpenExplorer(string extension)
        {

            //open the file explorer window to select the compressed file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Compressed Files (*.zip;*.7z;*.rar)|*.zip;*.7z;*.rar";
            openFileDialog.Filter = extension + " Files (*" + extension + ")|*" + extension;
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openFileDialog.ShowDialog() == true)
            {

                //determine the file format and call the decompression method
                string compressedFilePath = openFileDialog.FileName;
                DecompressFile(compressedFilePath, extension);

            }

        }

        private void DecompressFile(string compressedFilePath, string extension)
        {

            string extractionPath = System.IO.Path.GetDirectoryName(compressedFilePath);
            if (extractionPath != null)
            {

                try
                {
                    var archive = ArchiveFactory.Open(compressedFilePath);

                    foreach (var entry in archive.Entries)
                    {
                        if (!entry.IsDirectory)
                        {
                            entry.WriteToDirectory(extractionPath, new ExtractionOptions()
                            {
                                ExtractFullPath = true,
                                Overwrite = true
                            });
                        }
                    }
                    MessageBox.Show(extension + " file decompressed successfully!");
                }
                catch
                {
                    MessageBox.Show("Error during decompression " + extension + " file!");
                }

            }

        }

    }

}
