using S7.Net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace RaportContorGazCuptorV1
{
    

    class Raport
    {

        int counter = 1;
        string pathFisierDeTrimis = "";
        // Variablile PLC-uri
        Plc cuptorPLC;
        //object locker = new object();

        // Variabile RullatriceVechePLC
        double indexConvertit;
        double tempIndexConvertit = 0;
        double indexNeconvertit;

        public Raport()
        {
            //MessageBox.Show("Am ajuns in constructor");
            cuptorPLC = new Plc(CpuType.S7300, "172.16.4.104", 0, 2);
            
        }

        // Prop Get Counter (not used)
        public int GetCounter()
        {
            return counter;
        }

        // Prop Set Counter (not used)
        public void SetCounter(int c=1)
        {
            counter = c;
        }


        // Functie Conectare PLC
        public void ConectarePlc()
        {
            //MessageBox.Show("Am ajuns in conectare plc");
            try
            {
                if (cuptorPLC.IsAvailable) {
                    cuptorPLC.Open();
                }
                else MessageBox.Show("Cuptor PLC is not available! Check connection! IP: 172.16.4.104");
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.Message);
            }
        }

        // Functie Deconectare Plc
        public void DeconectarePlc()
        {
            if (cuptorPLC.IsConnected)
            {
                cuptorPLC.Close();
            }
            MessageBox.Show("S-a inchis polc-ul");
        }

        // Functie Citire semnale
        public bool CitireSemnale()
        {
            try
            {
                if (cuptorPLC.IsConnected)
                {
                    var performTaskCitireVariabile = Task.Run(() =>
                    {
                        indexConvertit = ((uint)cuptorPLC.Read("MD114")).ConvertToFloat(); // MD114 Index convertit       
                        indexNeconvertit = ((uint)cuptorPLC.Read("MD52")).ConvertToFloat(); // MD52 Index neconvertit       

                    });

                    performTaskCitireVariabile.Wait(TimeSpan.FromSeconds(1)); // Asteapta Task sa fie complet in 1 sec  


                    // MessageBox.Show("S-a citit indexul: " + indexConvertit.ToString()); // proba
                    return true; // A avut loc citirea corecta semnalelor

                }
                else return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Remake the connection after the PLC come back
            if (cuptorPLC.IsAvailable)
                cuptorPLC.Open();
            else MessageBox.Show("Cuptor PLC is not available! Check connection! IP: 172.16.4.104");

            return false;
        }

        // Functie verificare comunicatie
        public bool VerificareComunicatie() {
            if (cuptorPLC.IsConnected)
                return true;
            return false;
        }

        public double GetIndexConvertit()
        {
            return indexConvertit;
        }


        // Creare raport zilnic index consum
        public bool CreareRaportZilnic(string oraRaport, string listaMailuiriDeTrimis, int minDelayReport)
        {

            // MessageBox.Show("Test");
            StringBuilder csvContent = new StringBuilder();
            // Setare nume fisier nou la ora si data setata
            string numeFisier = string.Format("Consum_Gaz_Cuptor_{0}", DateTime.Now.ToString("MMMM"));
            
            // Cand este ora setata se realizeaza raportul
            if (DateTime.Now.ToString("HH:mm:ss") == oraRaport) // setare ora pentru raport
            {
                uint consum = (uint)(indexConvertit - tempIndexConvertit);
                if (consum > 300000) consum = 0;  ///// DE INLOCUIT CONSUM SI FACUT AUTOMAT DE TRIMIS MAIL
                
                // MessageBox.Show("consum: " + consum.ToString());
                tempIndexConvertit = indexConvertit;
                //MessageBox.Show("Egal");
                counter++;
                // Setare cale fisier si denumire
                string csvPath = string.Format("{0}/{1}.csv", Raport.CreareFolderRaportare(), numeFisier);

                // Verificam daca exista fisier, pentru a crea cap tabel
                if (File.Exists(csvPath)) {

                    //Adaugare continut fisier din liste defecte
                    csvContent.AppendLine(string.Format("{0},{1},{2},{3}", DateTime.Now.ToString("dd.MM.yyyy"), indexConvertit, 
                        consum, DateTime.Now.ToString("hh:mm:ss")));

                    // Daca e ultima zi din luna retinem nume fisier
                    if (DateTime.Now.Month != DateTime.Now.AddDays(1).Month)
                    {
                        //csvContent.AppendLine(string.Format("{0}, =sum(C{1}:C{2}), {3}, {4}", "Total: ", 1, counter, 0, DateTime.Now.ToString("hh:mm:ss")));
                        pathFisierDeTrimis = csvPath;
                    }
                }
                else
                {
                    // Adaugare Cap tabel in continut fisier
                    csvContent.AppendLine("Ziua,Index Convertit, Consum gaz[m3], Ora");
                    csvContent.AppendLine(string.Format("{0},{1},{2},{3}", DateTime.Now.ToString("dd.MM.yyyy"), indexConvertit, consum, DateTime.Now.ToString("hh:mm:ss")));
                }

                if (DateTime.Now.Day == 1) // Trimitere mail pe data de 1
                {
                    StringBuilder newCsvContent = new StringBuilder(); ;
                    //Adaugare continut fisier din liste defecte
                    newCsvContent.AppendLine(string.Format("{0},{1},{2},{3}", DateTime.Now.ToString("dd.MM.yyyy"), indexConvertit,
                        consum, DateTime.Now.ToString("hh:mm:ss")));
                    

                    // resetare counter pe ale lunii
                    counter = 2;

                    //Functie trimitere mail (string adreseMailDeTrimis, string filePathDeTrimis, string subiect)
                    if (pathFisierDeTrimis != "") //trimitem mailul pe data de 1 pentru luna precedenta
                    {
                        // Adaugare text in fisiser
                        File.AppendAllText(pathFisierDeTrimis, newCsvContent.ToString(), Encoding.UTF8);
                        // Trimitere mail
                        Raport.TrimitereRaportMail(listaMailuiriDeTrimis, pathFisierDeTrimis, numeFisier);
                    }
                    pathFisierDeTrimis = "";
                }

                // Console.WriteLine(csvPath);

                // Adaugare text in fisiser
                File.AppendAllText(csvPath, csvContent.ToString(), Encoding.UTF8);

                //MessageBox.Show("S-a facut raport");
                // Trimitere mail zilnic daca avem consum gaz lui Mitran si Cernat
                // listaMailuiriDeTrimis = "a.cernat@beltrame-group.com, m.mitran@beltrame-group.com, v.moisei@beltrame-group.com, b.mitran@beltrame-group.com"
                if (consum > 100 && consum <250000)
                    Raport.TrimitereRaportMail(listaMailuiriDeTrimis, csvPath, numeFisier);

                // Delay 1 secunda pentru a evita sa faca mai multe rapoarte in aceeasi secunda
                System.Threading.Thread.Sleep(1000);

                return true;
            }

            return false;
        }

        public static string CreareFolderRaportare()
        {
            string path = string.Format(@"c:\Consum gaz Cuptor/{0}", DateTime.Now.ToString("yyyy"));
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    // Console.WriteLine("That path exists already.");
                    return path;
                }

                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                // Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));

                // Delete the directory.
                //di.Delete();
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (Exception ex)
            {
                // Console.WriteLine("The process failed: {0}", e.ToString());
                MessageBox.Show(ex.Message);
                throw ex;
            }

            return path;
        }

        // Functie trimitere mail
        public static void TrimitereRaportMail(string adreseMailDeTrimis, string filePathDeTrimis, string subiect)
        {
            try
            {
                // "don.rap.ajustaj@gmail.com", "v.moisei@beltrame-group.com, vladmoisei@yahoo.com"
                // Mail(emailFrom , emailTo)
                MailMessage mail = new MailMessage("don.rap.ajustaj@gmail.com", adreseMailDeTrimis);

                //mail.From = new MailAddress("don.rap.ajustaj@gmail.com");
                mail.Subject = "Consum gaz cuptor cu propulsie";
                string Body = string.Format("Buna dimineata. <br>Atasat gasiti consumul de gaz inregistrat de contor gaz " +
                    "cuptor cu propuslie pe luna {0}. <br>O zi buna.", DateTime.Now.AddMonths(-1).ToString("MMMM, yyyy"));
                mail.Body = Body;
                mail.IsBodyHtml = true;
                using (Attachment attachment = new Attachment(filePathDeTrimis)) { 
                    mail.Attachments.Add(attachment);

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com"; //Or Your SMTP Server Address
                    smtp.Credentials = new System.Net.NetworkCredential("don.rap.ajustaj@gmail.com", "Beltrame.1");
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.Send(mail);

                    mail = null;
                    smtp = null;
                }

                // Console.WriteLine("Mail Sent succesfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                // Console.WriteLine(ex.ToString());
                throw ex;
            }
        }

        


    }
}
