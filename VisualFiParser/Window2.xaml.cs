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
        public Window2()
        {
            InitializeComponent();
        }
        public Window2(float release, string download_path) : this()
        {
            this.release = release;
            this.download_path = download_path;
        }
        public Window2(float release, string download_path, string changelog)
            :this(release, download_path)
        {
            this.changelog = changelog;
        }
        private void update_link_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(download_path);
        }
    }
}
