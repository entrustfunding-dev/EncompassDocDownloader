using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.Client;
using System.IO;

namespace EncompassDocDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            //  Make sure arguments are set properly
            if (args.Length != 2)
            {
                Console.WriteLine("ERROR: Must have two arguments passed in - loan GUID (including surrounding brackets) and the path of the folder to save loan files in (ending with a backslash)!");
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

            if (s == null)
            {
                Console.WriteLine("ERROR: Connection could not be established to the Encompass server!");
                Console.ReadLine();
                return;
            }

            try
            {
                //  Make sure directory path exists and is set up properly
                if (!Directory.Exists(args[1]))
                    Directory.CreateDirectory(args[1]);

                //  Open loan and save docs to disk
                Console.WriteLine("Opening loan...");
                Loan loan = s.Loans.Open(args[0]);

                //loan.Export(args[1] + "LoanExport", "", LoanExportFormat.FNMA32);

                //  Hashset to track already added docs
                HashSet<string> docHash = new HashSet<string>();

                //  Iterate backwards (most recent chronologically first)
                for(int i = loan.Attachments.Count - 1; i >= 0; --i)
                {
                    Attachment att = loan.Attachments[i];

                    //  Don't add non-most recent version of any doc
                    if (docHash.Contains(att.Title))
                        continue;

                    try
                    {
                        Console.WriteLine("Saving file " + att.Title + " to disk...");

                        //  Save depending on file
                        if (att.Title.Contains("FINDINGS") || att.Title.Contains("Credit Report"))
                            att.SaveToDiskOriginal(args[1] + att.Title + ".pdf");
                        else if (att.Title.Contains("CREDITPRINTFILE"))
                            att.SaveToDiskOriginal(args[1] + att.Title + ".txt");
                        else
                            att.SaveToDiskOriginal(args[1] + att.Title + ".html");

                        //  Add doc to hashset
                        docHash.Add(att.Title);
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine("ERROR SAVING FILE: " + att.Title + "!  " + ex.Message);
                    }
                }

                loan.Close();

                Console.WriteLine("Files saved successfully!");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
                Console.ReadLine();
                return;
            }

            s.End();
            return;
        }

        private static Session ConnectToServer()
        {
            //  Connect to Encompass server
            Console.WriteLine("Connecting to Encompass server...");
            Session s = new Session();
            s.Start("https://TEBE11214986.ea.elliemae.net$TEBE11214986", "zmitchell", "Encompass21!");
            return s;
        }
    }
}
