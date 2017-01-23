using System;
using System.Linq;
using System.Net.Mail;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Mail_Sender
{ 
    class Mail_Sender
    {
        public static bool LogIn()
        {
            for(int i=0; i<=3; i++)
            {
                Console.WriteLine("_____LOGIN_____");
                Console.Write("USERNAME: ");
                string username = Console.ReadLine().Trim();
                Console.Write("PASSWORD: "); string password ="";
                //password masking
                while (true)
                {
                    var key = Console.ReadKey(true);
                    if (key.Key != ConsoleKey.Enter && key.Key != ConsoleKey.Backspace)
                    {
                        password += key.KeyChar;
                        Console.Write("*");
                    }
                    if (key.Key == ConsoleKey.Enter)
                        break;
                    if (key.Key == ConsoleKey.Backspace)
                    {
                        if (!string.IsNullOrEmpty(password))
                        {
                            password = password.Substring(0, password.Length - 1);
                            int pos = Console.CursorLeft;
                            Console.SetCursorPosition(pos - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        }
                    }
                }

                if (username == "vuong.nc0582@gmail.com" && password == "Familyno1")
                {
                    System.Threading.Thread.Sleep(400);
                    Console.Clear();
                    return true;
                }
                else
                {
                    Console.Write("\nWrong Password!");
                    System.Threading.Thread.Sleep(400);
                    Console.Clear();
                }
            }
            return false;
        }

        public static Dictionary<String, String[]> GetFilesArray()
        {
            Dictionary<String, String[]> dic = new Dictionary<string, string[]>();
            //Rewrite path here
            dic.Add("withExtension", Directory.GetFiles(@"C:\Users\vuong\Desktop\OutLook", "*.pdf").Select(Path.GetFileName).ToArray());
            dic.Add("withoutExtension", Directory.GetFiles(@"C:\Users\vuong\Desktop\OutLook", "*.pdf").Select(Path.GetFileNameWithoutExtension).ToArray());
            return dic;
        }

        public static Dictionary<String, String[]> GetFilesSent()
        {
            Dictionary<String, String[]> dic = new Dictionary<string, string[]>();
            //Rewrite path here
            dic.Add("withExtension", Directory.GetFiles(@"C:\Users\vuong\Desktop\OutLook\Sent", "*.pdf").Select(Path.GetFileName).ToArray());
            dic.Add("withoutExtension", Directory.GetFiles(@"C:\Users\vuong\Desktop\OutLook\Sent", "*.pdf").Select(Path.GetFileNameWithoutExtension).ToArray());
            return dic;
        }

        public static bool checkFileExist(string nameFile)
        {
            Dictionary<String, String[]> fileList = GetFilesSent();
            foreach(string file in fileList["withExtension"])
            {
                if(file == "Processed-" + nameFile)
                {
                    File.Delete(@"C:\Users\vuong\Desktop\OutLook\Sent\" + file);
                    //Console.WriteLine("Co file ton tai roi");
                    return false;
                }
            }
            //Console.WriteLine("Hem co gi het");
            return true;
        }

        public static bool sendMail(Dictionary<String, Object> input, String Sub, String Bod)
        {

            try
            {
                List<String> recipient = input["to"] as List<String>;
                String filePath = input["path"] as String;
                //Initializing Mail Message and SMTP to use method
                foreach(String address in recipient)
                {
                    MailMessage mail = new MailMessage();
                    SmtpClient server = new SmtpClient("smtp.gmail.com");
                    mail.From = new MailAddress("vuong.nc0582@gmail.com");
                    mail.To.Add(new MailAddress(address));

                    //naming the recipient
                    Console.WriteLine("Recipient: " + address);

                    Dictionary<String, String[]> fileList = GetFilesArray();
                    foreach(String file in fileList["withExtension"])
                    {
                        //Subject and Body for each Recipient
                        switch (file.Substring(0, 3))
                        {
                            case "218":
                                mail.Subject = Sub;
                                mail.Body = Bod;
                                break;
                            case "207":
                                mail.Subject = Sub;
                                mail.Body = Bod;
                                break;
                            case "227":
                                mail.Subject = Sub;
                                mail.Body = Bod;
                                break;
                            default:
                                continue;
                        }
                    }

                    //Rewrite this path over here
                    Attachment attachment = new Attachment(Path.Combine(@"C:\Users\vuong\Desktop\OutLook\") + filePath);

                    mail.Attachments.Add(attachment);
                    server.Port = 587;
                    server.Credentials = new System.Net.NetworkCredential("vuong.nc0582@gmail.com", "Familyno1");
                    server.EnableSsl = true;
                    server.Send(mail);
                    mail.Dispose();
                }

                //if file exists in Sent Folder, just Delete the existence of it
                if(!checkFileExist(filePath))
                {
                    File.Move(@"C:\Users\vuong\Desktop\OutLook\" + filePath, @"C:\Users\vuong\Desktop\OutLook\Sent\" + "Processed-" + filePath);
                }
                else
                {
                    File.Move(@"C:\Users\vuong\Desktop\OutLook\" + filePath, @"C:\Users\vuong\Desktop\OutLook\Sent\" + "Processed-" + filePath);
                }

                //File.Copy(@"C:\Users\vuong\Desktop\OutLook\" + filePath, @"C:\Users\vuong\Desktop\OutLook\Sent\" + "Processed-" + filePath, true);
                //File.Delete(@"C:\Users\vuong\Desktop\OutLook\" + filePath);

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e.ToString());
            }
            return false;
        }

        static void Main(string[] args)
        {
            bool haveFile = false;

            Console.OutputEncoding = Encoding.UTF8;

            //Login Part
            //if (!LogIn())
            //{
            //    return;
            //}

            Dictionary<String, String[]> fileList = GetFilesArray();

            //Making Color lmao!
            Console.Write("Scanning...");
            System.Threading.Thread.Sleep(800);
            Console.Write("Done!");
            System.Threading.Thread.Sleep(400);

            //Clear screen this time!
            Console.Clear();

            if (fileList["withoutExtension"].Length !=0)
            {
                foreach(string file in fileList["withExtension"])
                {
                    String Sub = "";
                    String Bod = "";
                    List<String> recipient = new List<string>();
                    try
                    {
                        switch (file.Substring(0, 3))
                        {
                            case "218":
                                Console.WriteLine("\n"+ file);
                                recipient.Add("vuong.nc0582@gmail.com");
                                recipient.Add("vuong.photo92@gmail.com");
                                Sub = "DKSHVN HEC - " + file.Remove(file.Length - 4);
                                Bod = "Kính gửi Quý khách hàng\n\nChúng tôi xin gửi đến Quý khách hóa đơn tài chính cho các đơn hàng đãđược đặt theo yêu cầu. \n\nQuý khách vui lòng kiểm tra và phản hồi cho chúng tôi ngay nếu như Quý hách phát hiện có sai sót trên hóa đơn. Theo quy định của nhà nước hóa đơn chỉ được phép hủy khi hàng hóa chưa bàn giao ký nhận và hai bên chưa kê khai thuế do đó DKSH chỉ chấp nhận hủy hóa đơn nếu như Quý khách hàng phản hồi cho chúng tôi trước hay ngay khi giao hàng. Các phản hồi sau khi giao hàng sẽ được xử lý bằng cách lập biên bản điều chỉnh và hóađơnđiều chỉnh. \n\nĐối với các đơn hàng theo hình thức thanh toán là chuyển khoản, Quý khách hàng vui lòng chuyển khoản cho chúng tôi theo số tài khoản như sau: \n\n-       Tên tài khoản : Công ty TNHH DKSH VIET NAM\n-        Số tài khoản : 6339402\n-       Tên ngân hàng : ANZ – CHI NHANH HCM\n\n Sau khi chuyển tiền, để thuận tiện cho việc theo dõi đơn hàng,Quý khách vui lòng Fax hoặc scan Ủy nhiệm chi cho chúng tôi theo số: 1800 54 54 20 (miễn cước) hoặc email customercare.vn@dksh.com\n\n Chúng tôi sẽ tiến hành chuyển hàng cho Quý khách theo tuyến giao hànggần nhất ngay khi nhận được thông tin xác nhận từ Ngân hàng. \n\n Lưu ý:Quý khách vui lòng thực hiện chuyển khoản trong vòng 4 ngày kể từ ngày nhận được thông báo này. Sau thời gian trên,nếu chúng tôi chưa nhậnđược xác nhận của Ngân hàng về việc thanh toán của Quý khách, chúng tôi buộc lòng hủy hóa đơn theo qui định mà không cần phải thông báo cho Quý khách. \n\n Nếu Quý khách hàng có bất cứ thắc mắc nào liên quan đến đơn hàng,hàng hóa, hóa đơn vui lòng liên hệ với chúng tôi qua số điện thoại tổng đài miễn cước 1800.54.54.02\n\nCông ty DKSH VN rất hân hạnh được phục vụ quý khách và mong tiếp tục nhận được đơn đặt hàng của Quý khách trong tương lai. \n\nChân trọng";
                                break;
                            case "207":
                                Console.WriteLine("\n" + file);
                                recipient.Add("vuong.nc0582@gmail.com");
                                Sub = "DKSHVN CG - " + file.Remove(file.Length - 4);
                                Bod = "Kính gửi Quý khách hàng\n\nChúng tôi xin gửi đến Quý khách hóa đơn tài chính cho các đơn hàng đã được đặt theo yêu cầu. \n\nQuý khách vui lòng kiểm tra và phản hồi cho chúng tôi ngay nếu như Quý khách phát hiện có sai sót trên hóa đơn. Theo quy định của nhà nước hóa đơn chỉ được phép hủy khi hàng hóa chưa bàn giao ký nhận và hai bên chưakê khai thuế do đó DKSH chỉ chấp nhận hủy hóa đơn nếu như Quý khách hàngphản hồi cho chúng tôi trước hay ngay khi giao hàng. Các phản hồi sau khi giao hàng sẽ được xử lý bằng cách lập biên bản điều chỉnh và hóa đơnđiều chỉnh. \n\n Đối với các đơn hàng theo hình thức thanh toán là chuyển khoản, Quý khách hàng vui lòng chuyển khoản cho chúng tôi theo số tài khoản như sau \n\n-       Tên tài khoản : Công ty TNHH DKSH VIET NAM chi nhánh tại Hà Nội\n-       Số tài khoản : 002 - 139707 - 001\n-       Tên ngân hàng : Ngan hang TNHH MTV HSBC Viet nam\n\nSau khi chuyển tiền, để thuận tiện cho việc theo dõi đơn hàng, Quý khách vui lòng Fax hoặc scan Ủy nhiệm chi cho chúng tôi theo số: 0650 3766 610 - 0650 3766636 hoặc email Customercare.Consumergoods @dksh.com\n\nChúng tôi sẽ tiến hành chuyển hàng cho Quý khách theo tuyến giao hànggần nhất ngay khi nhận được thông tin xác nhận từ Ngân hàng. \n\nLưu ý:Vui lòng chuyển tiền và gửi UNC cho DKSH trước 15h(đối với khu vực tỉnh) và 16h(đối với khu vực HCM, HN, DN) để được giao hàng vào ngàymai, nếu gửi sau thời gian này đơn hàng của quý khách sẽ được giữ lại và chuyển vào chuyến kế tiếp căn cứ theo lịch giao hàng\n\nNếu Quý khách hàng có bất cứ thắc mắc nào liên quan đến đơn hàng,hàng hóa, hóa đơn vui lòng liên hệ với chúng tôi qua số điện thoại tổng đài miễn cước 1800.54.54.05\n\nCông ty DKSH VN rất hân hạnh được phục vụ quý khách và mong tiếp tục nhận được đơn đặt hàng của Quý khách trong tương lai. \n\nChân trọng";
                                break;
                            case "227":
                                Console.WriteLine("\n" + file);
                                recipient.Add("vuong.nc0582@gmail.com");
                                Sub = "DKSHVN PM - " + file.Remove(file.Length - 4);
                                Bod = "Kính gửi Quý khách hàng\n\nChúng tôi xin gửi đến Quý khách hóa đơn tài chính cho các đơn hàng đãđược đặt theo yêu cầu. \n\nQuý khách vui lòng kiểm tra và phản hồi cho chúng tôi ngay nếu như Quý khách phát hiện có sai sót trên hóa đơn. Theo quy định của nhà nước hóa đơn chỉ được phép hủy khi hàng hóa chưa bàn giao ký nhận và hai bên chưakê khai thuế do đó DKSH chỉ chấp nhận hủy hóa đơn nếu như Quý khách hàng phản hồi cho chúng tôi trước hay ngay khi giao hàng. Các phản hồi sau khi giao hàng sẽ được xử lý bằng cách lập biên bản điều chỉnh và hóađơnđiều chỉnh. \n\n Đối với các đơn hàng theo hình thức thanh toán là chuyển khoản,Quý kháchhàng vui lòng chuyển khoản cho chúng tôi theo số tài khoản như sau:\n\n-       Tên tài khoản : Công ty TNHH DKSH VIET NAM chi nhánh tại Hà Nội\n-       Số tài khoản : 002 - 139707 - 001\n-       Tên ngân hàng : Ngan hang TNHH MTV HSBC Viet nam\n\n Sau khi chuyển tiền, để thuận tiện cho việc theo dõi đơn hàng, Quý khách vui lòng Fax hoặc scan Ủy nhiệm chi cho chúng tôi theo số: 08 3812 577 3 hoặc email pm_saleassistant @dksh.com\n\nChúng tôi sẽ tiến hành chuyển hàng cho Quý khách theo tuyến giao hàng gần nhất ngay khi nhận được thông tin xác nhận từ Ngân hàng. \n\nNếu Quý khách hàng có bất cứ thắc mắc nào liên quan đến đơn hàng ,hànghóa, hóa đơn vui lòng liên hệ với chúng tôi theo danh sách sau\n\n1.Nguyễn Thị Kim Yến: Nhóm hàng Diversey Foodcare\nPhone + 84 8 3812 5848 Ext. 203; Fax + 84 8 3812 5845\nMobile + 84 976 932432; Email: yen.thikim.nguyen @dksh.com\n\n 2.Lý Thị Mai Hương: Nhóm hàng Diversey Care\nPhone + 84 8 3812 5848 Ext. 725, Fax + 84 8 3812 5845\nMobile + 84 909 544 205; Email: huong.thimai.ly @dksh.com\n\n 3.Hồ Thị Thùy Dương: Nhóm hàng hóa chất Mỹ phẩm; hóa chất Thực phẩm và đồ uống\nPhone + 84 8 3812 5848 Ext. 429; Fax + 84 8 3812 5845\nMobile: Mobile + 84 908 783 090; Email: duong.thithuy.ho @dksh.com\n\n 4.Nguyễn Thị Minh Phúc: Nhóm hàng hóa chất Công nghiệp\nPhone + 84 8 3812 5848 Ext: 171; Fax + 84 8 3812 5845\nMobile: +84 976 101 671; Email: sci.support @dksh.com] \n\n 5.Bùi Thị Phương Thi: Trưởng nhóm -Phone + 84 8 3812 5848 Ext.162\nEmail: Thi.thiphuong.bui @dksh.com, \n\nCông ty DKSH VN rất hân hạnh được phục vụ quý khách và mong tiếp tục nhận được đơn đặt hàng của Quý khách trong tương lai. \n\n Chân trọng";
                                break;
                            default:
                                continue;
                        }

                        //check if there is anyone to send mail
                        if(recipient.Count > 0)
                        {
                            haveFile = true;
                        }

                    }
                    catch (Exception)
                    {
                        Console.Write("Something wrong with Recipient");
                        System.Threading.Thread.Sleep(800);
                    }

                    Dictionary<String, Object> info = new Dictionary<string, Object>();
                    info.Add("to", recipient);
                    info.Add("path", file);
                    sendMail(info, Sub,Bod);
                }

                if (haveFile)
                {
                    Console.WriteLine("Finished!");
                    System.Threading.Thread.Sleep(400);
                }
                else
                {
                    Console.Write("Nothing to be sent.");
                    System.Threading.Thread.Sleep(400);
                }
            }
        }
    }
}