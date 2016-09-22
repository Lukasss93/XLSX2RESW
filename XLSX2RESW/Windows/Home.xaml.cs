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

namespace XLSX2RESW.Windows
{
    public partial class Home : Window
    {
        public Home()
        {
            InitializeComponent();
            dropzone.AllowDrop = true;
            dropzone.Drop += Dropzone_Drop;
        }

        private void Dropzone_Drop(object sender, DragEventArgs e)
        {
            if(e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                
                foreach(var file in files)
                {
                    Debug.WriteLine(file);
                }
            }
        }
    }
}
