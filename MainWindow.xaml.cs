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

        private void buttonDecompress_Click(object sender, RoutedEventArgs e)
        {
            OpenExplorer();
        }

        private void OpenExplorer()
        {

            //open the file explorer window to select the compressed file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Compressed Files (*.zip;*.7z;*.rar;*.tar)|*.zip;*.7z;*.rar;*.tar";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openFileDialog.ShowDialog() == true)
            {

                //determine the file format and call the decompression method
                string compressedFilePath = openFileDialog.FileName;
                DecompressFile(compressedFilePath);

            }

        }

        private void DecompressFile(string compressedFilePath)
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
                    MessageBox.Show("File decompressed successfully!");
                }
                catch
                {
                    MessageBox.Show("Error during file decompression!");
                }

            }

        }

    }

}
