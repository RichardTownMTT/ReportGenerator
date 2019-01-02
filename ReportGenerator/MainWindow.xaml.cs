using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace ReportGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void cmdImperialMetSum_Click(object sender, RoutedEventArgs e)
        {
            string filename = OpenFile("Word Document(*.docx)|*.docx", "Select Metrology Summary", "No metrology summary document selected");
            if (string.IsNullOrEmpty(filename))
            {
                //No file selected.  Message has already been displayed.
                return;
            }

            bool success = MetrologySummary.CreateImperialMetSum(filename);
            if (success)
            {
                MessageBox.Show("File created");
            }
        }

        private string OpenFile(string filterString, string openFileTitle, string noDocumentSelectedMessage)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = filterString;
            openFile.Title = openFileTitle;

            string retval = "";

            if (openFile.ShowDialog() == true)
            {
                retval = openFile.FileName;
            }
            else
            {
                MessageBox.Show(noDocumentSelectedMessage);
            }

            return retval;
        }
    }
}
