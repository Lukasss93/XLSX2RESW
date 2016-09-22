using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using Microsoft.Win32;
using XLSX2RESW.Classes;

namespace XLSX2RESW.Windows
{
    public partial class Home : Window
    {
        public Home()
        {
            InitializeComponent();

            dropzone.AllowDrop = true;
            dropzone.Drop += Dropzone_Drop;
            dropzone_browse.Click += Dropzone_browse_Click;
        }

        private void Dropzone_browse_Click(object sender, RoutedEventArgs e)
        {
            //create openfiledialog 
            OpenFileDialog dlg = new OpenFileDialog();

            //set filter for file extension and default file extension 
            dlg.DefaultExt = Constants.InputExtension;
            dlg.Filter = Constants.InputExtensionFilter;
            dlg.Multiselect = true;

            //display openfiledialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            //get the selected file name and display in a TextBox 
            if(result == true)
            {
                //open document 
                var files = dlg.FileNames;

                //convert files
                Elaborator.Convert(files);
            }
        }

        private void Dropzone_Drop(object sender, DragEventArgs e)
        {
            if(e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                //get files from dropzone
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);

                //convert files
                Elaborator.Convert(files);
            }
        }
    }
}
