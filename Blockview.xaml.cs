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

namespace MyAuto
{
    /// <summary>
    /// Interaction logic for Blockview.xaml
    /// </summary>
    public partial class Blockview : Window
    {
        public Blockview(string _Filter)
        {
            InitializeComponent();
            Filter = _Filter;
        }

        public string UserName => nameBox.Text;
        string Filter = "";

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            openFileDlg.Filter = Filter;

            // Launch OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = openFileDlg.ShowDialog();
            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                nameBox.Text = openFileDlg.FileName;
                DialogResult = true;
            }
            else { DialogResult = false; }
        }
    }
}
