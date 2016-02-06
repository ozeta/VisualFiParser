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
using System.Windows.Shapes;

namespace VisualFiParser
{
    /// <summary>
    /// Logica di interazione per Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
        float release;
        string changelog;
        string download_path;

        public Window2(float release, string download_path)
        {

            InitializeComponent();
            this.release = release;
            this.download_path = download_path;
            Uri newuri = new Uri(download_path);
            update_link.NavigateUri = newuri;

        }
        public Window2(float release, string download_path, string changelog)
            :this(release, download_path)
        {
            this.changelog = changelog;
        }
        private void update_link_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                System.Diagnostics.Process.Start(update_link.NavigateUri.ToString());

            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
                Trace.WriteLine(ex.StackTrace);
            }
        }
    }
}
