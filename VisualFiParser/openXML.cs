using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows;
using System.Diagnostics;

namespace VisualFiParser
{

    class openXML
    {

        private string path_moduloD_template;
        private string path_moduloD2_template;
        private string path_excel;
        private string sheetName;
        private string output_path;
        /// <summary>
        /// istanzia un nuovo oggetto openXML per la lettura di un file excel
        /// </summary>
        /// <param name="output_path">path in cui si troverà la cartella files e 
        /// le sottocartelle in e out</param>
        /// <param name="filename_excel">nome del file excel da leggere</param>
        /// <param name="sheetName">foglio del file excel che contiene i dati da leggere</param>
        /// 
        public openXML(String filename_excel, String sheetName)
        {
            this.path_excel = filename_excel;
            this.sheetName = sheetName;
        }
        /// <summary>
        /// crea un nuovo oggetto OPENXML per l'esportazione in word
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="path_moduloD">nome del file del modulo D</param>
        /// <param name="path_moduloD2">nome del file del modulo D2</param>
        /// <param name="output_path">path in cui si troverà la cartella files e 
        /// le sottocartelle in e out</param>
        public openXML(String path_moduloD, String path_moduloD2, String output_path)
        {
            this.path_moduloD_template = path_moduloD;
            this.path_moduloD2_template = path_moduloD2;
            this.output_path = output_path;

        }

        /// <summary>
        /// verifica la presenza della cartella out. se non è presente, la crea
        /// </summary>
        /// <param name="path">percorso da analizzare</param>


        private String getModuloD2PathBySurname(String surname)
        {
            return output_path + @"ModuloD2-" + surname + ".docx";
        }
        private String getModuloDOutputPath()
        {
            return output_path + @"ModuloD.docx";
        }
        private String getSpreadSheetPath()
        {
            return path_excel;
        }
        private String getModuloDtemplatePath()
        {

            return path_moduloD_template;
        }
        private String getModuloD2templatePath()
        {
            return path_moduloD2_template;
        }

        /// <summary>
        /// genera il file D2 per il singolo atleta
        /// </summary>
        /// <param name="team">oggetto contenente i dati di squadra ed atleti</param>
        /// <param name="i">indice dell'atleta selezionato</param>
        public void fillD2(Team team, int i)
        {
            String[] athleteDictionary = team.Athlete_list[0];      //indice atleta
            String[] athlete = team.Athlete_list[i];                //dati singolo atleta
            String[] teamDictionary = team.Athlete_team[0];         //indice squadra dell'atleta
            String[] teamArray = team.Athlete_team[1];              //dati squadra dell'atleta
            String[] homeTeamDictionary = team.Home_team[0];        //indice squadra di casa
            String[] homeTeamArray = team.Home_team[1];             //dati squadra di casa

            //crea il nome di output
            string var = getModuloD2PathBySurname(athlete[1]);

            //using: apre documento da parsare, crea il nuovo documento
            using (WordprocessingDocument mainDoc = WordprocessingDocument.Open(getModuloD2templatePath(), true))
            using (WordprocessingDocument resultDoc = WordprocessingDocument.Create(var, mainDoc.DocumentType))
            {
                //ricopia le parti del vecchio documento nel nuovo
                foreach (var part in mainDoc.Parts)
                    resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                string docText = null;
                //stringhifica il file
                using (StreamReader sr = new StreamReader(resultDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
                //sostituisce i dati dell'atleta, se esistenti
                docText = regFields(athlete, athleteDictionary, docText);
                //sostituisce i dati della squadra dell'atleta, se esistenti
                docText = regFields(teamArray, teamDictionary, docText);
                //sostituisce i dati della squadra di casa, se esistenti
                docText = regFields(homeTeamArray, homeTeamDictionary, docText);

                //scrive il nuovo file su disco
                using (StreamWriter sw = new StreamWriter(resultDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }



        /// <summary>
        ///cerca e sostituisce tutte le espressioni del dizionario, nel documento stringhificato.
        ///restituisce il documento modificato
        /// </summary>
        /// <param name="replace">array delle stringhe che verranno </param>
        /// <param name="strToReplace">array delle stringhe da cercare e sostituire</param>
        /// <param name="docText">documento che verrà letto, modificato e restituito</param>
        /// <returns></returns>
        private string regFields(String[] replace, String[] strToReplace, string docText)
        {
            //            foreach (string txt in strArray)
            int j = 0;
            for (int i = 0; i < strToReplace.Length; i++)
            {
                string pattrn = String.Format(@"\b{0}\b", strToReplace[j++]);
                try
                {
                    //in caso di stringa vuota
                    if (replace[i] == null || replace[i].Equals("") || replace[i].Equals("0"))
                    {
                        docText = Regex.Replace(docText, pattrn, "______________");

                    }
                    else
                    {
                        docText = Regex.Replace(docText, pattrn, replace[i]);

                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.Message);
                    Trace.WriteLine(ex.StackTrace);
                }

            }
            return docText;
        }
        /// <summary>
        /// riempie la tabella word con gli atleti, nel modulo D. usa la funzione getalignedarray
        /// per ottenere una lista di parametri già ordinati da inserire in modo sequenziale.
        /// La funzione richiede che l'oggetto Team sia stato creato tramite foglio excel
        /// </summary>
        /// <param name="t">Tabella word</param>
        /// <param name="team">team degli atleti</param>
        private void fillTable(DocumentFormat.OpenXml.Wordprocessing.Table t, Team team)
        {
            int j = 0;
            //array atleti. serve solo per snellire il codice
            String[][] athletes = team.Athlete_list;
            //sequenza info da estrapolare: 
            //2, 1, 4, CAT, SPEC, Tempo, 11, 12
            String[][] alignedMatrix = new String[athletes.Length - 1][];
            TableCell[][] cellArray = new TableCell[athletes.Length - 1][];

            for (int i = 1; i < athletes.Length; i++)
            {
                alignedMatrix[i - 1] = getAlignedArray(athletes[i]);
                cellArray[i - 1] = new TableCell[alignedMatrix[0].Length];
                for (j = 0; j < alignedMatrix[0].Length; j++)
                {
                    cellArray[i - 1][j] = new TableCell((new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(alignedMatrix[i - 1][j])))));
                }
                t.Append(new TableRow(cellArray[i - 1]));

            }


        }

        private string[] getAlignedArray(string[] athlete)
        {

            //numero di campi necessari
            int fieldNumber = 8;

            String[] res = new String[fieldNumber];
            res[0] = athlete[1];
            res[1] = athlete[0];
            res[2] = athlete[3];
            res[3] = athlete[12];
            res[4] = athlete[13];
            res[5] = athlete[14];
            res[6] = athlete[10];
            res[7] = athlete[11];


            return res;
        }

        /// <summary>
        /// riempie il modulo D ( elenco atleti ) usando il file excel
        /// </summary>
        /// <param name="team"></param>
        public void fillD(Team team)
        {
            //dati singolo atleta
            String[][] athletes = team.Athlete_list;
            //indice squadra dell'atleta
            String[] teamDictionary = team.Athlete_team[0];
            //dati squadra dell'atleta
            String[] teamArray = team.Athlete_team[1];
            //indice squadra di casa
            String[] homeTeamDictionary = team.Home_team[0];
            //dati squadra di casa
            String[] homeTeamArray = team.Home_team[1];
            string var = getModuloDOutputPath();
            string var1 = getModuloDtemplatePath();
            //using: apre documento da parsare, crea il nuovo documento
            using (WordprocessingDocument mainDoc = WordprocessingDocument.Open(var1, true))
            using (WordprocessingDocument resultDoc = WordprocessingDocument.Create(var, mainDoc.DocumentType))
            {
                //ricopia le parti del vecchio documento nel nuovo
                foreach (var part in mainDoc.Parts)
                    resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);

                //riempie la tabella di dati
                Body bod = resultDoc.MainDocumentPart.Document.Body;
                foreach (DocumentFormat.OpenXml.Wordprocessing.Table t in bod.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>())
                {
                    this.fillTable(t, team);
                }


            }
            //chiudo il flusso e lo riapro per salvare gli altri dati del modulo.
            //devo eseguire queste 2 procedure in sequenza perché la prima lavora sul body, la seconda
            //sullo stream xml
            using (WordprocessingDocument finalDoc = WordprocessingDocument.Open(var, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(finalDoc.MainDocumentPart.GetStream()))
                {
                    //leggo il flusso e lo converto in stringa
                    docText = sr.ReadToEnd();
                    //sostituisco i dati del team ospite e quello di casa
                    docText = regFields(teamArray, teamDictionary, docText);
                    docText = regFields(homeTeamArray, homeTeamDictionary, docText);
                }

                //salvo tutto su disco
                using (StreamWriter sw = new StreamWriter(finalDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }


            }

        }

        /// <summary>
        /// cerca il nome di una colonna e restituisce l'indice
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="test"></param>
        /// <returns></returns>
        private int getColumnIndex(string[][] matrix, string test)
        {
            int column = 0;
            string check = matrix[0][0];
            while (!check.Equals(test) && column < matrix[0].Length)
            {
                column++;
                if (column < matrix[0].Length)
                    check = matrix[0][column];

            }
            return column;

        }
        /// <summary>
        /// riceve la matrice che contiene i dati degli atleti. trova la colonna delle date
        /// di nascita e converte il numero letto dal foglio excel in formato leggibile
        /// </summary>
        /// <param name="matrix">matrice dei dati degli atleti, con riga delle intestazioni</param>
        /// <returns></returns>
        public bool parseBirthdays(String[][] matrix)
        {
            bool res = false;
            int row = 0;
            int column = 0;


            column = getColumnIndex(matrix, "DATA_NASCITA");
            if (column < matrix[0].Length)
            {
                res = true;

                row++;

                bool cont = true;
                while (row < matrix.Length && cont)
                {
                    try
                    {
                        DateTime birth = DateTime.FromOADate(Convert.ToInt64(matrix[row][column]));
                        matrix[row][column] = birth.Day + @"\" + birth.Month + @"\" + birth.Year;
                        //Console.Out.WriteLine(test[row][column]);
                        row++;
                    }
                    catch (FormatException ex)
                    {

                        Console.Out.WriteLine(ex.Message);
                        Console.Out.WriteLine("Controllare la data di nascita di {0} {1}",
                            matrix[row][0], matrix[row][1]);
                        cont = false;
                    }
                    catch (ArgumentException ex)
                    {
                        Console.Out.WriteLine(ex.Message);
                        Console.Out.WriteLine(ex.StackTrace);
                        cont = false;
                    }
                }
            }


            return res;
        }

        public int getRowsNumber(WorkbookPart wbPart)
        {
            int res = 0;
            string cell = "A" + res;

            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            if (theSheet == null)
                throw new ArgumentException("Foglio {0} non trovato", sheetName);
            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            //spreadsheet.GetCellValue(path_desktop + @"\files\in\" + file_excel, "FIPARSER", cell);
            try
            {
                string test = null;
                while ((test = GetCellValueNew(wsPart, wbPart, res, 0)) != null && !test.Equals(""))
                {
                    res++;
                    Console.Out.WriteLine(test);
                    cell = "A" + res;
                }
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine("getRowsNumber");
                Console.Out.WriteLine(ex.Message);
                Console.WriteLine("Premere un tasto per uscire.");
                Console.ReadKey();
                System.Environment.Exit(-1);
            }

            /**/
            return res;
        }

        /**
        legge un file excel e riempie la tabella degli atleti
         */
        public String[][] parseSpreadSheet(int columns)
        {
            string[][] res = null;


            string fileName = getSpreadSheetPath();
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    int rows = getRowsNumber(document.WorkbookPart);
                    Console.Out.WriteLine("Righe lette: {0}", rows);
                    if (rows > 0)
                        res = cycleFile(document, rows, columns);
                    else
                        throw new ArrayBoundaryException("numero di righe non valido: la prima riga è vuota?");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
                Trace.WriteLine(ex.StackTrace);
                throw;
            }
            return res;
        }

        /// <summary>
        /// cicla il foglio excel e ne estrae il contenuto in un array
        /// </summary>
        /// <param name="document"></param>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        public String[][] cycleFile(SpreadsheetDocument document, int rows, int columns)
        {
            string value = null;

            WorkbookPart wbPart = document.WorkbookPart;
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            if (theSheet == null)
                throw new ArgumentException("Foglio {0} non trovato", sheetName);
            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            string[][] res = new String[rows][];
            for (int i = 0; i < rows; i++)
            {
                res[i] = new String[columns];
                for (int j = 0; j < columns; j++)
                {
                    value = GetCellValueNew(wsPart, wbPart, i, j);
                    res[i][j] = value;

                }
            }


            if (parseBirthdays(res) == false)
                Console.Out.WriteLine("Colonna degli anni di nascita non trovata");
            return res;
        }
        /// <summary>
        /// restituisce il contenuto di una cella Excel
        /// </summary>
        /// <param name="wsPart">parametro del foglio excel</param>
        /// <param name="wbPart">parametro del foglio excel</param>
        /// <param name="i">riga</param>
        /// <param name="j">colonna</param>
        /// <returns></returns>
        public string GetCellValueNew(WorksheetPart wsPart, WorkbookPart wbPart, int i, int j)
        {
            string value = null;
            String addressName = Convert.ToString((char)('A' + j));
            addressName = (addressName + (i + 1));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();
            if (theCell != null)
            {
                string test = null;
                value = theCell.InnerText;
                if (theCell.DataType != null)
                {
                    test = theCell.DataType.Value.ToString();
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.String:
                            value = theCell.GetFirstChild<CellValue>()?.Text;
                            break;
                        case CellValues.SharedString:
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                    }
                }
                else {
                    value = theCell.GetFirstChild<CellValue>()?.Text;
                }
            }

            return value;
        }

        /**/
    }
}
