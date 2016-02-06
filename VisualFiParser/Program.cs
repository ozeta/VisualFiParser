using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Sockets;
using System.Globalization;
using Newtonsoft.Json;
using System.Windows;

namespace VisualFiParser
{

    class Update
    {
        bool newUpdate;
        string local_file;
        string remote_file;

        public void writeToFile(string filename)
        {
            string output = JsonConvert.SerializeObject(this, Formatting.Indented);
            System.IO.File.WriteAllText(filename, output);
        }
        public Team readFiletoObject(string url)
        {
            StreamReader str = new StreamReader(@url);
            
            return JsonConvert.DeserializeObject<Team>(str.ToString());
        }
    }
    static class DialogBox
    {
        static public void write(string caption, string message, MessageBoxImage icon)
        {
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxResult result = MessageBox.Show(message, caption, button, icon);

        }
    }
    static class Program
    {
        private static bool enabledD = true;
        private static bool enabledD2 = true;

        static String file_iscritti = @"iscritti.CSV";
        static String file_team = @"team.CSV";
        static String file_moduloD = @"templateD.docx";
        static String file_moduloD2 = @"templateD2.docx";
        static String file_excel = @"Cartel1wl.xlsx";
        static String file_excel1 = @"Cartel1 -withoutlabel.xlsx";
        static String path_exe = @AppDomain.CurrentDomain.BaseDirectory;
        static String path_desktop = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        const int COLUMNS = 15;
        static string[] test = { "uno", "due", "tre", "quattro" };


        static public DateTime getTime()
        {
            DateTime localDateTime;
            try
            {
                var client = new TcpClient("time.nist.gov", 13);
                using (var streamReader = new StreamReader(client.GetStream()))
                {
                    var response = streamReader.ReadToEnd();
                    var utcDateTimeString = response.Substring(7, 17);
                    localDateTime = DateTime.ParseExact(utcDateTimeString, "yy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal);
                }
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                Console.Out.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
            }
            finally
            {
                localDateTime = DateTime.Now;

            }

            return localDateTime;
        }

        /*
        static void Main1(string[] args)
        {
            try
            {
                Console.WriteLine("FiParser v0.2 Demo!");
                //parsing cvs
                //Team team = new Team(file_team, file_iscritti, path_desktop);
                 //Settings sett = new Settings(new Team(path_desktop, file_team, file_excel, "FIPARSER", COLUMNS));
               // sett.writeToFile(path_desktop + @"\output.txt");

                //parsing excel
                Team team = new Team(path_desktop , file_excel, "FIPARSER", COLUMNS);
                
                DateTime now = getTime();
                int year = now.Year;
                int month = now.Month;
                if (year > 2016 && month >= 3)
                {
                    enabledD = enabledD & false;
                    enabledD2 = enabledD2 & false;
                    Console.WriteLine("Il Periodo di prova è scaduto! :(");
                }
                if (enabledD == true && team != null)
                {
                    try
                    {
                        openXML moduloD = new openXML("FIPARSER", file_moduloD, file_moduloD2, path_desktop);
                        moduloD.fillD(team);
                        Console.WriteLine("Modulo D compilato.");
                    }
                    catch (OpenXmlPackageException ex)
                    {
                        Console.WriteLine("modulo D non valido: {0}", ex.Message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        enabledD2 = false;
                    }
                }
                //creazione D2
                if (enabledD2 == true && team != null)
                {
                    try
                    {
                        openXML moduloD2 = new openXML("FIPARSER", file_moduloD, file_moduloD2, path_desktop);
                        int i;
                        for (i = 1; i < team.Athlete_list.Length; i++)
                        {
                            moduloD2.fillD2(team, i);
                        }
                        Console.WriteLine("Moduli D2 esportati: {0}", i - 1);
                    }
                    catch (OpenXmlPackageException ex)
                    {
                        Console.WriteLine("modulo D2 non valido: {0}", ex.Message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }

            }
            catch (ArrayBoundaryException ex)
            {
                Console.WriteLine("Parsing non eseguito. Errore nei file di importazione: {0}", ex.Message);
            }
            Console.WriteLine("");
            Console.WriteLine("Premere un tasto per uscire.");
            Console.ReadKey();
        }
        */
    }
}
