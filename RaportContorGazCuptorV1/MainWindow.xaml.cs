using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Diagnostics;
using System.Net.Mail;

namespace RaportContorGazCuptorV1
{



    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        bool executieProgram = true;
        string oraRaportare = "07:00:00";
        int indexRaport = 1;
        string adreseMailDeTrimis = "v.moisei@beltrame-group.com, m.chiru @beltrame-group.com, " +
                     "a.cernat@beltrame-group.com, m.mitran@beltrame-group.com, b.mitran@beltrame-group.com";
        public MainWindow()
        {
            InitializeComponent();


            //ProcessThreadCollection prt = Process.GetCurrentProcess().Threads;
            //foreach (ProcessThread item in prt)
            //{
            //    MessageBox.Show(item.ToString());
            //}

            Thread t = new Thread(new ThreadStart(ProgramBackground));
            t.Name = "PlcCommunication";
            t.Start();

            Stop_Comm.Click += StopCommButon_Click;
            Start_Comm.Click += StartCommButon_Click;
            Setare_Ora_Raport.Click += Setare_Ora_Raport_Click;
            AfisareOraRAport.Click += AfisareOraRAport_Click;
        }


        public void ProgramBackground()
        {
            try
            {
                //MessageBox.Show("A pornit thread!");
                double tempIndexRaport = 0;

                Raport gazRaport = new Raport();
                gazRaport.ConectarePlc();
                while (executieProgram)
                {
                    gazRaport.CitireSemnale();
                    
                    if (gazRaport.CreareRaportZilnic(oraRaportare, adreseMailDeTrimis, 1))
                    {
                        this.Dispatcher.Invoke(() => // Folosit pentru a modifica grafica hmi
                        {
                            indexText.Text = String.Format("Index citit: {0} la data: {1}", gazRaport.GetIndexConvertit(),
                                DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss"));
                        });

                    }

                    string timp = DateTime.Now.ToString("HH:mm:ss");
                    this.Dispatcher.Invoke(() =>
                    {
                        mesaje.Text = timp;
                    });


                    this.Dispatcher.Invoke(() =>
                    {

                        if (gazRaport.VerificareComunicatie())
                        {
                            eclipsa.Fill = Brushes.Green;
                        }
                        else eclipsa.Fill = Brushes.LightGray;
                    });


                }

                gazRaport.DeconectarePlc();

                this.Dispatcher.Invoke(() =>
                {
                    if (gazRaport.VerificareComunicatie())
                    {
                        eclipsa.Fill = Brushes.Green;
                    }
                    else eclipsa.Fill = Brushes.LightGray;
                });
            }
            catch (Exception ex)
            {

                try
                {
                    // "don.rap.ajustaj@gmail.com", "v.moisei@beltrame-group.com, vladmoisei@yahoo.com"
                    // Mail(emailFrom , emailTo)
                    MailMessage mail = new MailMessage("don.rap.ajustaj@gmail.com", "v.moisei@beltrame-group.com");

                    //mail.From = new MailAddress("don.rap.ajustaj@gmail.com");
                    mail.Subject = "Consum gaz cuptor cu propulsie";
                    string Body = string.Format("S-a oprit programul raportare gaz cu eroarea: <br>{0}", ex.Message);
                    mail.Body = Body;
                    mail.IsBodyHtml = true;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com"; //Or Your SMTP Server Address
                    smtp.Credentials = new System.Net.NetworkCredential("don.rap.ajustaj@gmail.com", "Beltrame.1");
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.Send(mail);

                    mail = null;
                    smtp = null;


                    // Console.WriteLine("Mail Sent succesfully");
                }
                catch (Exception excep)
                {
                    MessageBox.Show(excep.Message);
                    // Console.WriteLine(ex.ToString());                    
                }
                MessageBox.Show(ex.Message);
            }
            finally
            {
                executieProgram = false;
                MessageBox.Show("A iesit Thread");
            }
        }


        private void StartCommButon_Click(object sender, RoutedEventArgs e)
        {
            if (!executieProgram)
            {

                executieProgram = true;
                Thread t = new Thread(new ThreadStart(ProgramBackground));
                t.Name = "PlcCommunication";
                t.Start();

            }

            //MessageBox.Show(executieProgram.ToString());
        }

        private void StopCommButon_Click(object sender, RoutedEventArgs e)
        {
            if (executieProgram)
                executieProgram = false;
            //MessageBox.Show(executieProgram.ToString());
        }

        private void Setare_Ora_Raport_Click(object sender, RoutedEventArgs e)
        {
            oraRaportare = oraRaportTextBox.Text;
            adreseMailDeTrimis = AdreseMailDeTrimisTBox.Text;
        }

        private void AfisareOraRAport_Click(object sender, RoutedEventArgs e)
        {
            oraRaportTextBox.Text = oraRaportare;
            AdreseMailDeTrimisTBox.Text = adreseMailDeTrimis;
        }

        // Functie incheiere while loop cand inchzi fereastra
        void DataWindow_Closing(object sender, EventArgs e)
        {
            //MessageBox.Show("S-a inchis fereastra");
            executieProgram = false;
            //MessageBox.Show(executieProgram.ToString());
            //MessageBox.Show("S-a inchis fereastra!");
        }

    }
}
