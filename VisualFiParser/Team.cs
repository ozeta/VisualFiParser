using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
namespace VisualFiParser
{
    class Team
    {
        private String[][] athlete_team;
        private String[][] home_team;
        private String[][] athlete_list;
        private String excel_path;
        private string web_home;
        public Team()
        {

        }
        /// <summary>
        /// nuovo oggetto team con squadra atleti e squadra casa impostati
        /// </summary>
        /// <param name="athlete_team"></param>
        /// <param name="home_team"></param>
        public Team(String[][] athlete_team, String[][] home_team)
        {
            this.Athlete_team = athlete_team;
            this.Home_team = home_team;
        }
        /// <summary>
        ///  nuovo oggetto team con squadra atleti e squadra casa, path per il file excel impostati
        /// </summary>
        /// <param name="athlete_team"></param>
        /// <param name="home_team"></param>
        /// <param name="excel_path"></param>
        public Team(String[][] athlete_team, String[][] home_team, string excel_path)
        {
            this.Athlete_team = athlete_team;
            this.Home_team = home_team;
            this.Excel_path = excel_path;
        }

        /// <summary>
        ///  nuovo oggetto creato solo con la lista degli atleti.
        /// </summary>
        /// <param name="home_path"></param>
        /// <param name="filename_team"></param>
        /// <param name="file_excel"></param>
        /// <param name="sheetName"></param>
        /// <param name="columns"></param>
        public Team(string home_path,String file_excel, string sheetName, int columns)
        {
            Athlete_team = null;
            Home_team = null;
            Athlete_list = null;
            try
            {
                athlete_list = (new openXML(home_path, file_excel, sheetName)).parseSpreadSheet(columns);

            }
            catch (System.IO.DirectoryNotFoundException ex)
            {

                Console.Out.WriteLine(ex.Message);
            }
        }

        public string[][] Athlete_team
        {
            get
            {
                return athlete_team;
            }

            set
            {
                athlete_team = value;
            }
        }

        public string[][] Athlete_list
        {
            get
            {
                return athlete_list;
            }

            set
            {
                athlete_list = value;
            }
        }

        public string[][] Home_team
        {
            get
            {
                return home_team;
            }

            set
            {
                home_team = value;
            }
        }

        public string Excel_path
        {
            get
            {
                return excel_path;
            }

            set
            {
                excel_path = value;
            }
        }

        public string Web_home
        {
            get
            {
                return web_home;
            }

            set
            {
                web_home = value;
            }
        }

        public void writeToFile(string filename)
        {
            string output = JsonConvert.SerializeObject(this, Formatting.Indented);
            System.IO.File.WriteAllText(filename, output);
        }
        static public Team readFiletoObject(string path)
        {
            string buffer = System.IO.File.ReadAllText(path);
            return JsonConvert.DeserializeObject<Team>(buffer);
        }

    }
}

