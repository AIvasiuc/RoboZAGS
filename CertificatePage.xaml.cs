using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
using System.Windows.Threading;
using System.Net;
using System.Net.Mail;
using System.IO.Ports; //For work with COM-port (Arduino)
//using System.Drawing;

namespace WpfRoboZags
{
    /// <summary>
    /// Interaction logic for CertificatePage.xaml
    /// </summary>
    public partial class CertificatePage : BasePage
    {
        private DispatcherTimer timer;
        private MediaPlayer mediaPlayer = new MediaPlayer();
        FormattedText ftxt_P1;  //Input data of first person
        FormattedText ftxt_P2;  //Input data of second person

        SerialPort serial = new SerialPort(); //Create COM-port object (Arduino)

        public CertificatePage()
        {
            InitializeComponent();

            // --- COM-port settings (Arduino) ---
            serial.PortName = "COM4";            
            serial.BaudRate = 9600; 
            serial.Handshake = System.IO.Ports.Handshake.None;
            serial.Parity = Parity.None;
            serial.DataBits = 8;
            serial.StopBits = StopBits.One;
            serial.ReadTimeout = 200;
            serial.WriteTimeout = 50;

            serial.Open();     
        }
        // Timer for moving to next page if user doesn't answer the question
        private void start_Timer()
        {
            timer = new DispatcherTimer();
            timer.Tick += dispatcherTimer_Tick;
            timer.Interval = new TimeSpan(0, 0, 1);
            timer.Start();
        }

        const int timerTimeout = 10;
        private int resetCntr = timerTimeout;

        void dispatcherTimer_Tick(object sender, object e)
        {
            --resetCntr;
            if (resetCntr == 0)
            {
                timer.Stop();
                navigateNextFrame();
            }
        }
        // Navigate to next page
        private void navigateNextFrame()
        {
            this.NavigationService.Navigate(new Uri("FinalPage.xaml", UriKind.Relative));
        }       
        //Loading page
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            playSound("C:/RZAssets/CertificatePage.m4a");
            
            PrintDialog pd = new PrintDialog();
            BitmapImage certBg = new BitmapImage(new Uri("pack://application:,,,/Images/CertEmpty.png"));
            BitmapImage capture = new BitmapImage(new Uri("capture.png", UriKind.Relative));
            DrawingVisual visual = new DrawingVisual();

            PersonService persons = PersonService.Instance;
            Person p1 = persons.person1;
            Person p2 = persons.person2;

            using (DrawingContext dc = visual.RenderOpen())
            {
                dc.DrawImage(certBg, new Rect(0, 0, certBg.PixelWidth, certBg.PixelHeight));
                dc.DrawImage(capture, new Rect(283, 1463, 1184, 712));
                //// first person

                ftxt_P1 = new FormattedText(p1.firstName + "\n" + p1.middleName + " " + p1.lastName
                    , new CultureInfo("")
                    , FlowDirection.LeftToRight
                    , new Typeface("Verdana")
                    , 40
                    , Brushes.Black);
                ftxt_P1.TextAlignment = TextAlignment.Center;


                ftxt_P2 = new FormattedText(p2.firstName + "\n" + p2.middleName + " " + p2.lastName
                    , new CultureInfo("")
                    , FlowDirection.LeftToRight
                    , new Typeface("Verdana")
                    , 40
                    , Brushes.Black);
                ftxt_P2.TextAlignment = TextAlignment.Center;

                dc.DrawText(ftxt_P1, new Point(874, 815));
                dc.DrawText(ftxt_P2, new Point(874, 1134));
            }
            //Rendering photo
            RenderTargetBitmap target = new RenderTargetBitmap(certBg.PixelWidth, certBg.PixelHeight,
                                                       96, 96, PixelFormats.Default);
            target.Render(visual);

            // Saving final certificate
            JpegBitmapEncoder jpg = new JpegBitmapEncoder();
            jpg.Frames.Add(BitmapFrame.Create(target));
            using (Stream stm = File.Create("certificateSmall!.bmp"))
            {
                jpg.Save(stm);
            }
            
            Image img = new Image();
            img.Source = target;

            pd.PrintVisual(img, "RoboZags");
            
            //Writing into file all data
            StreamWriter write_text;  //Class for writing into file
            FileInfo file = new FileInfo("printlog.txt");
            write_text = file.AppendText(); //Attach info to file, create file if file doesn't exist
            DateTime localDate = DateTime.Now;
            string date = localDate.ToShortDateString();
            string time = localDate.ToShortTimeString();
            string line = date + "_" + time +  "_" + p1.lastName + "_" + p1.middleName + "_" + p1.firstName + "_" + p2.lastName + "_" + p2.middleName + "_" + p2.firstName;
            string nameOfFile = line + ".bmp";
            write_text.WriteLine(line); 
            write_text.Close(); // Close file


            // Sending certificate to e-mail
            // Sender - establish address and name
            MailAddress from = new MailAddress("robozags@balrobotov.ru", "Robozags");
            try
            {
                // Whom to send
                MailAddress toFirst = new MailAddress(p1.email, "Robozags");
                // Create message object
                MailMessage mToFirst = new MailMessage(from, toFirst);
                // Message's subject
                mToFirst.Subject = "Робозагс. Свидетельство о браке";
                // Text
                mToFirst.Body = "Любви вам!";

                mToFirst.IsBodyHtml = true;
                // Attach certificate
                mToFirst.Attachments.Add(new Attachment("certificateSmall!.bmp"));
                // SMTP server address and port
                SmtpClient smtp = new SmtpClient("smtp.mail.ru", 25);
                //SmtpClient smtp = new SmtpClient("smtp.gmail.com", 465);

                // Login and password
                smtp.Credentials = new NetworkCredential("robozags@balrobotov.ru", "&A8bmra4dPRI");
                smtp.EnableSsl = true;

                smtp.Send(mToFirst);
            }
            catch (Exception exc)
            {

            }
            // Same process for the second person
            try
            {
                MailAddress toSecond = new MailAddress(p2.email, "Robozags");
                MailMessage mToSecond = new MailMessage(from, toSecond);
                mToSecond.Subject = "Робозагс. Свидетельство о браке";
                mToSecond.Body = "Любви вам!";
                mToSecond.IsBodyHtml = true;
                mToSecond.Attachments.Add(new Attachment("certificateSmall!.bmp"));
                SmtpClient smtp = new SmtpClient("smtp.mail.ru", 25);
                smtp.Credentials = new NetworkCredential("robozags@balrobotov.ru", "&A8bmra4dPRI");
                smtp.EnableSsl = true;
                smtp.Send(mToSecond);
            }
            catch (Exception exc)
            {

            }
            Console.Read();

            // --- Sending data via COM-port (Arduino) ---
            if (serial.IsOpen)
            {
                try
                {
                    // Send the binary data out the port
                    //byte[] hexstring = Encoding.ASCII.GetBytes("Hello");
                    //foreach (byte hexval in hexstring)
                    //{
                    //byte[] _hexval = new byte[] { hexval };     // need to convert byte 
                    // to byte[] to write
                    serial.Write("111"); //Arduino could't respond sometimes, if we send "1" or "11". Don't know why:(
                    //Thread.Sleep(1);
                    //}
                }
                catch (Exception ex)
                {
                    //para.Inlines.Add("Failed to SEND" + data + "\n" + ex + "\n");
                    //mcFlowDoc.Blocks.Add(para);
                    //Commdata.Document = mcFlowDoc;
                }
            }

            serial.Close(); //Important
            // --- Sending data via COM-port (Arduino) ---
                   start_Timer();
        }        
    }
}
