using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using System.IO;

namespace Mail_Sender
{
    class Program
    {
        static void Main(string[] args)
        {
            //Sign in
            string username = "vuong.nc0582@gmail.com";
            string password = "Familyno1";
            string p;
            string u;

            login:;
            Console.WriteLine("Login");
            Console.Write("Username: "); u = Console.ReadLine();

            if (u != username)
            {
                Console.WriteLine("User not found, try again");
                goto login;
            }
            else
            {
                Console.Write("Password: "); p = Console.ReadLine();
                if (p != password)
                {
                    Console.WriteLine("Nahh, wrong password!");
                    goto login;
                }
            }

            //get files name from Folder Outlook
            string[] Attname = Directory.GetFiles(@"C:\Users\vuong\Desktop\OutLook", "*.pdf").Select(Path.GetFileName).ToArray();
            //get files name from Folder Outlook without name of file-type
            string[] Attname2 = Directory.GetFiles(@"C:\Users\vuong\Desktop\OutLook", "*.pdf").Select(Path.GetFileNameWithoutExtension).ToArray();
                       
            Console.WriteLine("\nSending mail");
            
            //Scan all file name to...
            for(int n =0; n < Attname.Length; n++)
            {
                //initializing Recipient.
                string recipient = "Unknown";

                //initializing Codename to know which file to send to which Recipient
                string codename = Attname[n].Substring(0,3);

                //testing
                Console.WriteLine(Attname[n]);
                Console.WriteLine(codename);

                //attachment with certain Codename will be sent to certain address
                switch(codename)
                {
                    case ("100"):
                        recipient = "vuong.nc0582@gmail.com";
                        break;
                    case ("101"):
                        recipient = "vuong.photo92@gmail.com";
                        break;
                    case ("102"):
                        recipient = "vuong.nc0582@gmail.com";
                        break;
                    case ("103"):
                        recipient = "vuong.nguyencong92@gmail.com";
                        break;
                    default:
                        Console.WriteLine("You have nothing to send");
                        break;

                }

                //testing
                Console.WriteLine(recipient);

                //Initializing Mail Message and SMTP to use method
                MailMessage mail = new MailMessage();
                SmtpClient server = new SmtpClient("smtp.gmail.com");
                // SMTP of Microsoft is smtp.live.com

                //Setup Email Address of Sender
                mail.From = new MailAddress("vuong.nc0582@gmail.com");

                //Setup Email Address of Reciever
                mail.To.Add(recipient);

                //Subject part
                mail.Subject = "Testing";

                //Body part
                mail.Body = "Testing";

                //Initializing Addin Attachments
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(@"C:\Users\vuong\Desktop\OutLook\"+ Attname[n]);
                mail.Attachments.Add(attachment);

                //Setup SMTP Port
                server.Port = 587;

                //setup account and send mail
                server.Credentials = new System.Net.NetworkCredential(u, p);
                server.EnableSsl = true;
                server.Send(mail);

                mail.Dispose();

                //Move file attachment which has been sent to another folder named Sent
                File.Copy(@"C:\Users\vuong\Desktop\OutLook\" + Attname[n], @"C:\Users\vuong\Desktop\OutLook\Sent\" + Attname2[n] + "-PROCESSED.pdf");
                File.Delete(@"C:\Users\vuong\Desktop\OutLook\" + Attname[n]);
            }
            Console.WriteLine("\nDone");
        }
    }
}
