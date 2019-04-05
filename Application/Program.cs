using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace Application
{
    class Program
    {

        static void Main(string[] args)
        {
            var GetDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string[] dirs = Directory.GetDirectories(@"V:\TEMP\");
            StringBuilder sb = new StringBuilder();
            string titreLog = "************************************************************ logs " + DateTime.Now + " ****************************************************************";
            File.AppendAllText(GetDirectory + "/logsFile.csv", "" + Environment.NewLine);
            File.AppendAllText(GetDirectory + "/logsFile.csv", titreLog + Environment.NewLine);
            File.AppendAllText(GetDirectory + "/logsFile.csv", "" + Environment.NewLine);
            string chaineLog = "";



            foreach (string dir in dirs)
            {
                Boolean newLine = false;
                Console.ForegroundColor = ConsoleColor.Cyan;
                string user = System.IO.File.GetAccessControl(dir).GetOwner(typeof(System.Security.Principal.NTAccount)).ToString();
                string[] words = user.Split('\\');
                string[] words2 = dir.Split(':');
                string folderName = words2[1];
                Console.WriteLine(folderName);
                DateTime dateCreation = System.IO.File.GetCreationTime(dir);
                DateTime dt = Directory.GetLastWriteTime(dir);
                DateTime dt2 = dateCreation.AddDays(15);
                string dirMail = dir;

                double dif = DateTime.Now.Subtract(dateCreation).TotalDays;
                if (DateTime.Now.Subtract(dateCreation).TotalDays > 7)
                {

                    bool verif = dir.Contains("sera supprimé le");
                    if (verif)
                    {
                        if (DateTime.Now.Subtract(dt2).TotalDays > 0)
                        {
                            try
                            {
                                Directory.Delete(dirMail, true);
                                chaineLog = dirMail + "; date de création : " + dateCreation + "; Suppression OK";
                                File.AppendAllText("p:/logsFile.csv", chaineLog + Environment.NewLine, Encoding.Unicode);
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("Dossier supprimé avec succès!");
                            }
                            catch
                            {
                                chaineLog = dirMail + "; date de création : " + dateCreation + "; Suppression échouée";
                                File.AppendAllText("p:/logsFile.csv", chaineLog + Environment.NewLine, Encoding.Unicode);
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Suppression du dossier échouée!");
                            }
                            newLine = true;
                        }
                    }

                    if (!verif)
                    {
                        DateTime dateSuppression = dateCreation.AddDays(14);
                        string dateSuppressionS = dateSuppression.ToShortDateString();
                        string novName = dir + " - sera supprimé le " + dateSuppressionS.Replace("/", "-");
                        dirMail = novName;
                        try
                        {
                            Directory.Move(dir, novName);
                            chaineLog = dir + "; date de création : " + dateCreation + "; date de modification : " + dt + "; Renommage OK";
                            File.AppendAllText("p:/logsFile.csv", chaineLog + Environment.NewLine, Encoding.Unicode);
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Renommage réussi");
                        }
                        catch
                        {
                            chaineLog = dir + "; date de création : " + dateCreation + "; date de modification : " + dt + "; Renommage échoué ";
                            File.AppendAllText("p:/logsFile.csv", chaineLog + Environment.NewLine, Encoding.Unicode);
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Renommage échoué");
                        }
                        newLine = true;
                    }
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("Dernière modification était le {0} ", dt);
                    double result = (dt2 - dt).TotalDays;//comparer date de modification et date de suprission
                    if (result > 0 && result < 4)
                    {
                        Outlook._Application _app = new Outlook.Application();
                        Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);

                        mail.To = words[1] + "@ca-centreloire.fr";
                        mail.Subject = "Dossier sur V:\\TEMP bientôt supprimé";

                        mail.Body = "Bonjour,\n \nLe dossier \" " + folderName + " \" sera automatiquement supprimé selon les règles de gestion de V:\\TEMP, soit 15 jours après sa création. \n Veillez à prendre les mesures nécessaires si le contenu est toujours utile. \n " +
                            "Cordialement,\n\nJorge DUARTE";
                        mail.Importance = Outlook.OlImportance.olImportanceHigh;

                        try
                        {
                            ((Outlook._MailItem)mail).Send();
                            chaineLog = dir + "; suppression prévue le : " + dt2.ToShortDateString() + "; un mail est envoyé pour prévenir l'utilisateur " + words[1];
                        }
                        catch
                        {
                            chaineLog = dir + "; suppression prévue le : " + dt2.ToShortDateString() + "; envoi de mail échoué ";
                        }
                        newLine = true;
                        File.AppendAllText("p:/logsFile.csv", chaineLog + Environment.NewLine, Encoding.Unicode);
                        Console.WriteLine("Suppression prévue le {0}", dt2.ToShortDateString());
                    }
                }
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\n**********************************************************\n");
                if (newLine)
                {
                    File.AppendAllText("p:/logsFile.csv", "" + Environment.NewLine);
                }
            }
        }
    }
}
