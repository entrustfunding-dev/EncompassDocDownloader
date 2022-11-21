using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.Client;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace EncompassDocDownloader
{
    class Program
    {
        private static string _clientId;
        private static string _username = "zmitchell";
        private static string _password = "Encompass21!";
        //private static Stopwatch _stopwatch;

        private static readonly string _ERRORTEXTFILENAME = "Errors.txt";

        static void Main(string[] args)
        {
            //_stopwatch = Stopwatch.StartNew();

            //  Make sure arguments are set properly
            if (args.Length != 2)
            {
                Console.WriteLine("ERROR: Must have two arguments passed in - filepath to a CSV document with loan GUIDS and the path of the folder to save loan files in (ending with a backslash)!");
                Console.ReadLine();
                return;
            }

            //  Start Encompass runtime services and run app
            Console.WriteLine("Initializing Encompass runtime services...");
            new EllieMae.Encompass.Runtime.RuntimeServices().Initialize();
            RunApplication(args);
        }

        private static void RunApplication(string[] args)
        {
            Session s = ConnectToServer();

            //  Make sure connection was successful
            if (s == null)
            {
                Console.WriteLine("ERROR: Connection could not be established to the Encompass server!");
                Console.ReadLine();
                return;
            }

            //  Make sure base directory path exists and is set up properly
            if (!Directory.Exists(args[1]))
                Directory.CreateDirectory(args[1]);

            string[] guids;

            //  Get all the loan guids from csv file
            using (FileStream fs = new FileStream(args[0], FileMode.Open))
            {
                StreamReader sr = new StreamReader(fs);
                string fileText = sr.ReadToEnd();
                guids = fileText.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                sr.Close();
            }

            foreach(string guid in guids)
            {
                //int attachmentsSaved = 0;

                //  Open loan and save docs to disk
                Console.WriteLine($"\nOpening loan {guid}...");
                Loan loan = s.Loans.Open(guid);

                //loan.Export(args[1] + "LoanExport", "", LoanExportFormat.FNMA32);

                //  Concat subfolder path for this loan and create it
                string subFilePath = args[1] + "\\" + loan.LoanNumber + "_" + loan.BorrowerPairs.Current.Borrower.LastName;
                if (!Directory.Exists(subFilePath))
                    Directory.CreateDirectory(subFilePath);

                //  Hashset to track already added docs
                Dictionary<string, ushort> filenameDict = new Dictionary<string, ushort>();

                //  Iterate backwards (most recent chronologically first)
                for(int i = loan.Attachments.Count - 1; i >= 0; --i)
                {
                    string cleanTitle = "";
                    Attachment att = loan.Attachments[i];
                    string startingTitle = att.Title;

                    //  Add pdf file type if not already included
                    if (!startingTitle.Contains(".pdf"))
                        startingTitle += ".pdf";

                    if (filenameDict.ContainsKey(att.Title))
                    {
                        //  If file title already exists, increment and save with that value as a marker
                        string fileVers = $"({++filenameDict[att.Title]})";
                        startingTitle = startingTitle.Insert(startingTitle.Length - 4, fileVers);
                    }
                    else
                    {
                        //  Otherwise add the title to the dictionary
                        filenameDict.Add(att.Title, 1);
                    }

                    try
                    {
                        //  Remove any characters not allowed in filenames
                        Regex regex = new Regex("[<>:\"/\\|?*]");
                        cleanTitle = regex.Replace(startingTitle, "");

                        //  Attempt to save to disk    
                        att.SaveToDisk(subFilePath + "\\" + cleanTitle);
                        Console.WriteLine("Saving file " + cleanTitle + " to disk...");
                        //++attachmentsSaved;
                    }
                    catch(Exception)
                    {
                        try
                        {
                            //  Catch on normal save file - try to save as original
                            att.SaveToDiskOriginal(subFilePath + "\\" + cleanTitle);
                            Console.WriteLine("Saving original file " + cleanTitle + " to disk...");
                            //++attachmentsSaved;
                        }
                        catch (Exception)
                        {
                            //  On further fail, write to error file
                            File.AppendAllText(args[1] + "\\" + guid + "\\" + _ERRORTEXTFILENAME, $"ERROR: Problem saving file {cleanTitle}\n");
                        }
                    }
                }

                //Console.WriteLine($"{attachmentsSaved} unique files attempted to be saved!");
                loan.Close();
            }

            Console.WriteLine("All loan files saved successfully!\nYou may now close this program!");
            //_stopwatch.Stop();
            //Console.WriteLine($"Total time taken: {_stopwatch.Elapsed}");
            Console.ReadLine();

            s.End();
            return;
        }

        private static Session ConnectToServer()
        {
#if DEBUG
            _clientId = "https://TEBE11214986.ea.elliemae.net$TEBE11214986";
#else
            _clientId = "https://BE11210494.ea.elliemae.net$BE11210494";
#endif

            //  Connect to Encompass server
            Console.WriteLine("Connecting to Encompass server...");
            Session s = new Session();
            s.Start(_clientId, _username, _password);
            return s;
        }
    }
}
