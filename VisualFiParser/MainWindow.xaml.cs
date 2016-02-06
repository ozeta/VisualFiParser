using DocumentFormat.OpenXml.Packaging;
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

namespace VisualFiParser
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    /// 


    public partial class MainWindow : Window
    {

        float release = 0.4f;

        Team team;
        TextBox[] tb0 = new TextBox[4];
        TextBox[] tb1 = new TextBox[3];
        TextBox[] tb2 = new TextBox[3];
        TextBox[] tb3 = new TextBox[2];
        TextBox[] tb00;
        TextBox[] tb01;
        String[][] team_arr = new string[2][];
        String[][] team_hom = new string[2][];

        String caption = "Fiparser";
        static String path_desktop = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        static String path_exe = @AppDomain.CurrentDomain.BaseDirectory;
        static String base_path = path_exe + @"files\";
        String settings = base_path + @"settings.txt";
        String file_moduloD = base_path + @"templateD.docx";
        String file_moduloD2 = base_path + @"templateD2.docx";
        String update_path = base_path + "update.txt";
        string updateUrl = @"https://raw.githubusercontent.com/ozeta/VisualFiParser/master/update.txt";
        string setupUrl = @"https://github.com/ozeta/VisualFiParser/releases/download/release_0.4/setup.exe";
        String excel_full_path = null;
        String sheetName = "FIPARSER";
        String output_path = null;
        Update currentVersion;
        const int COLUMNS = 15;
        private void checkOutputDirectory(string path)
        {
            string outstr = @"\Moduli compilati\";
            bool exists = System.IO.Directory.Exists(path + outstr);
            if (!exists)
                System.IO.Directory.CreateDirectory(path + outstr);
        }
        public string Output_path
        {
            get
            {
                return output_path;
            }

            set
            {
                output_path = System.IO.Path.GetDirectoryName(value) + @"\Moduli compilati\";
            }
        }

        private void fileSelect()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.FileName = "Elenco atleti"; // Default file name
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Foglio Excel (.xlsx)|*.xlsx";

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                try
                {
                    // imposta path+filename
                    this.excel_full_path = dlg.FileName;
                    file_atleti.Text = dlg.FileName;
                    team.Excel_path = dlg.FileName;
                    this.Output_path = dlg.FileName;

                    //legge la lista atleti
                    team.Athlete_list = (new openXML(this.excel_full_path, sheetName)).parseSpreadSheet(COLUMNS);
                }
                catch (Exception ex)
                {
                    DialogBox.write(caption, ex.Message, MessageBoxImage.Stop);
                }

            }
            else throw new Exception("Devi selezionare un file Atleti!");
        }

        /// <summary>
        /// legge i campi e li associa all'oggetto team
        /// </summary>
        /// <returns>se esiste, il vecchio team aggiornato. altrimenti un nuovo team compilato</returns>
        private Team parseTexboxToTeam()
        {
            for (int i = 0; i < tb00.Length; i++)
            {
                team_arr[1][i] = tb00[i].Text;
            }
            for (int i = 0; i < tb01.Length; i++)
            {
                team_hom[1][i] = tb01[i].Text;
            }

            try
            {
                checkInput(team_hom[1][3]);
                if (!file_atleti.Text.Equals(excel_full_path))
                {
                    excel_full_path = file_atleti.Text;
                }
                if (team == null)
                    return new Team(team_arr, team_hom, excel_full_path);
                else
                {
                    team.Athlete_team = team_arr;
                    team.Home_team = team_hom;
                    team.Excel_path = excel_full_path;
                    return team;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);

            }
        }

        //salva modulo
        //salva campi
        private void saveModifiedFields()
        {
            try
            {
                if (team.Athlete_team == null || team.Home_team == null)
                {
                }
                team = parseTexboxToTeam();

                team.writeToFile(settings);
                string messageBoxText = "I dati sono stati salvati e verranno ripristinati alla prossima esecuzione";
                DialogBox.write(caption, messageBoxText, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                DialogBox.write(caption, ex.Message, MessageBoxImage.Error);

            }
        }
        //controlla se è presente un json con cui compilare i vari form
        void checkInput(string dateStr)
        {
            if (dateStr != null)
            {
                DateTime dt;
                bool test = DateTime.TryParse(dateStr, out dt);
                if (!test)
                {
                    if (!dateStr.Equals(""))
                        throw new Exception("La data inserita non è valida. Es. 20/03/2016");
                }
            }
        }

        private void validateInputTexBoxes()
        {

            for (int i = 0; i < tb00.Length; i++)
            {
                if (tb00[i].Text.Equals(""))
                {
                    throw new Exception("Bisogna compilare tutti i campi!");
                }
            }
            for (int i = 0; i < tb01.Length; i++)
            {
                if (tb01[i].Text.Equals(""))
                {
                    throw new Exception("Bisogna compilare tutti i campi!");
                }
            }
            checkInput(team_hom[1][3]);

        }
        private void createTexBoxes(StackPanel panel, TextBox[] tb)
        {
            if (tb == null)
                throw new Exception("non sono presenti textbox");
            for (int i = 0; i < tb.Length; i++)
            {
                tb[i] = new TextBox();
                tb[i].Width = 158;
                tb[i].Height = 23;
                tb[i].Margin = new Thickness(0, 8, 0, 0);
                tb[i].MaxLines = 1;
                panel.Children.Add(tb[i]);

            }
        }
        public MainWindow()
        {
            InitializeComponent();
            /*
            DateTime now = Program.getTime();
            int year = now.Year;
            int month = now.Month;
            if (year != 2016 && month > 2)
            {
                DialogBox.write(caption, "Il Periodo di prova è scaduto! 😟", MessageBoxImage.Stop);

                Environment.Exit(0);
            }
            */
            string jsonPath = settings;
            bool jsonExists = System.IO.File.Exists(jsonPath);
            bool updateFileExists = System.IO.File.Exists(update_path);
            try
            {
                createTexBoxes(panel0, tb0);
                createTexBoxes(panel1, tb1);
                createTexBoxes(panel2, tb2);
                createTexBoxes(panel3, tb3);
                tb00 = tb0.Concat(tb1).ToArray();
                tb01 = tb2.Concat(tb3).ToArray();
                team_arr[0] = new string[7] { "NOME_SOC", "CITTA_SOC", "INDIRIZZO_SOC", "MAIL_SOC", "PROV_SOC", "CAP_SOC", "TEL_SOC" };
                team_arr[1] = new string[7];
                team_hom[0] = new string[5] { "GAME_NAME", "GAME_HOME", "GAME_TEAM", "GAME_DATA", "TEAM_HOME" };
                team_hom[1] = new string[5];

                if (updateFileExists)
                {
                    Update test;
                    currentVersion = Update.readFiletoObject(update_path);
                    if ( (test = currentVersion.isRemoteUpdateAvaible(updateUrl) ) != null)
                    {
                        Window2 win2 = new Window2(test.Release, test.Remote_download_setup);
                        win2.Show();
                    }
                }
                if (jsonExists)
                {
                    team = Team.readFiletoObject(settings);

                    for (int i = 0; i < team.Athlete_team[0].Length; i++)
                    {
                        tb00[i].Text = team.Athlete_team[1][i];
                    }
                    for (int i = 0; i < team.Home_team[0].Length; i++)
                    {
                        tb01[i].Text = team.Home_team[1][i];
                    }

                    if (!team.Excel_path.Equals(""))
                    {
                        this.excel_full_path = team.Excel_path;
                        file_atleti.Text = team.Excel_path;
                        Output_path = team.Excel_path;
                        team.Athlete_list = (new openXML(this.excel_full_path, sheetName)).parseSpreadSheet(COLUMNS);
                        outputpath_label.Content = "I moduli verranno salvati nella cartella " + Output_path;
                    }

                }
                else
                {
                    team = new Team();

                }

            }
            catch (System.IO.FileNotFoundException ex)
            {
                DialogBox.write(caption, ex.Message + "\nDevi selezionare un File valido", MessageBoxImage.Information);
            }
            catch (System.IO.FileLoadException ex)
            {
                DialogBox.write(caption, ex.Message, MessageBoxImage.Information);
            }
            catch (NullReferenceException ex)
            {
                System.IO.File.Delete(jsonPath);
                Trace.WriteLine(ex.Message);
                Trace.WriteLine(ex.StackTrace);

            }
            catch (System.ArgumentException ex)
            {
                string message = "Non è stato possibile ripristinare il file degli atleti salvato in precedenza, devi selezionarne uno nuovo\n";
                DialogBox.write(caption, message + ex.Message, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                DialogBox.write(caption, ex.Message, MessageBoxImage.Information);
            }
        }

        //Salva modifiche
        private void button_Click(object sender, RoutedEventArgs e)
        {
            saveModifiedFields();
        }

        //seleziona una nuova lista atleti
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                fileSelect();
                outputpath_label.Content = "I moduli verranno salvati nella cartella " + Output_path;
            }
            catch (Exception ex)
            {

            }
        }

        //compila D
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            // fillModule();

            try
            {
                validateInputTexBoxes();
                if (team.Athlete_list == null)
                {
                    fileSelect();
                }
                openXML moduloD = new openXML(file_moduloD, file_moduloD2, Output_path);
                moduloD.fillD(team);
                DialogBox.write(caption, "Modulo D compilato.", MessageBoxImage.Information);
            }
            catch (OpenXmlPackageException ex)
            {
                DialogBox.write(caption, ex.Message + "\nER 01: modulo D non valido", MessageBoxImage.Error);
            }
            catch (System.IO.IOException ex)
            {
                DialogBox.write(caption, ex.Message + "\nER 02. Chiudere il Modulo D e riprovare.", MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                DialogBox.write(caption, ex.Message, MessageBoxImage.Error);
            }

        }

        //compila D2
        private void button4_Click(object sender, RoutedEventArgs e)
        {



            try
            {
                validateInputTexBoxes();
                if (team.Athlete_list == null)
                {
                    fileSelect();
                }
                int i;
                openXML moduloD2 = new openXML(file_moduloD, file_moduloD2, Output_path);
                for (i = 1; i < team.Athlete_list.Length; i++)
                {
                    moduloD2.fillD2(team, i);
                }
                string message;
                if (i == 1)
                {
                    message = "1 modulo D2 esportato";
                }
                else
                {
                    message = i + " moduli D2 esportati";

                }
                DialogBox.write(caption, message, MessageBoxImage.Information);

            }
            catch (OpenXmlPackageException ex)
            {
                throw new Exception("ER 03: modulo D2 non valido", ex);

            }
            catch (System.IO.IOException ex)
            {
                throw new Exception("ER 04. Chiudere il Modulo D2 e riprovare.", ex);
            }
            catch (Exception ex)
            {
                DialogBox.write(caption, ex.Message, MessageBoxImage.Error);
            }
        }
        //azzera tutti i campi
        private void button6_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < tb00.Length; i++)
            {
                tb00[i].Text = "";
            }
            for (int i = 0; i < tb01.Length; i++)
            {
                tb01[i].Text = "";
            }
            file_atleti.Text = "";
        }

        //controlla aggiornamenti
        private void button7_Click(object sender, RoutedEventArgs e)
        {
            //Update upd = Update.readRemoteFiletoObject(@url);
            //Update two = new Update(release, base_path + "update.txt", updateUrl, setupUrl);
            //two.writeToFile(two.Local_file);

        }

        //info
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            Window1 win1 = new Window1(release);
            win1.Show();
            //this.Close();
        }
    }
}
