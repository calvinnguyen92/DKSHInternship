using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace Mail_Sender
{
    class Mail_Sender
    {
        //setting up to hide Console APP
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        public static bool LogIn()
        {
            for (int i = 0; i <= 3; i++)
            {
                Console.WriteLine("_____LOGIN_____");
                Console.Write("USERNAME: ");
                string username = Console.ReadLine().Trim();
                Console.Write("PASSWORD: "); string password = "";
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

        public static Dictionary<String, String[]> GetFilesOutlook()
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
            foreach (string file in fileList["withExtension"])
            {
                if (file == "Processed-" + nameFile)
                {
                    File.Delete(@"C:\Users\vuong\Desktop\OutLook\Sent\" + file);
                    //Console.WriteLine("Co file ton tai roi");
                    return false;
                }
            }
            //Console.WriteLine("Hem co gi het");
            return true;
        }

        public static bool sendMail(Dictionary<String, Object> input, String Subj, String Body)
        {
            var handle = GetConsoleWindow();
            
            try
            {
                //ShowWindow(handle, SW_HIDE);
                List<String> recipient = input["to"] as List<String>;
                String filePath = input["path"] as String;
                
                foreach (String address in recipient)
                {
                    Microsoft.Office.Interop.Outlook.Application OutlookApp = new Microsoft.Office.Interop.Outlook.Application();
                    Microsoft.Office.Interop.Outlook.MailItem OutlookMail = (Microsoft.Office.Interop.Outlook.MailItem)OutlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    Microsoft.Office.Interop.Outlook.Recipients re = (Microsoft.Office.Interop.Outlook.Recipients)OutlookMail.Recipients;
                    
                    OutlookMail.Recipients.Add(address);

                    //naming the recipient
                    Console.WriteLine("Recipient: " + address);

                    Dictionary<String, String[]> fileList = GetFilesOutlook();
                    foreach (String file in fileList["withoutExtension"])
                    {
                        //Subject and Body for each Recipient
                        switch (file.Substring(0, 3))
                        {
                            case "218":
                                OutlookMail.Subject = Subj;
                                OutlookMail.HTMLBody = Body;
                                break;
                            case "207":
                                OutlookMail.Subject = Subj;
                                OutlookMail.HTMLBody = Body;
                                break;
                            case "227":
                                OutlookMail.Subject = Subj;
                                OutlookMail.HTMLBody = Body;
                                break;
                            default:
                                continue;
                        }
                    }

                    Microsoft.Office.Interop.Outlook.Attachment Attachment = OutlookMail.Attachments.Add(Path.Combine(@"C:\Users\vuong\Desktop\OutLook\") + filePath);
                    OutlookMail.Send();
                    OutlookMail = null;
                    OutlookApp = null;
                }

                //if file exists in Sent Folder, just Delete the existence of it
                if (!checkFileExist(filePath))
                {
                    File.Move(@"C:\Users\vuong\Desktop\OutLook\" + filePath, @"C:\Users\vuong\Desktop\OutLook\Sent\" + "Processed-" + filePath);
                }
                else
                {
                    File.Move(@"C:\Users\vuong\Desktop\OutLook\" + filePath, @"C:\Users\vuong\Desktop\OutLook\Sent\" + "Processed-" + filePath);
                }

                return true;
            }
            catch (System.Exception e)
            {
                ShowWindow(handle, SW_SHOW);
                Console.WriteLine("Error: {0}\n", e.ToString());
                Console.WriteLine("Cannot send mail!");
                Console.Clear();
            }
            return false;
        }

        static void Main(string[] args)
        {
            var handle = GetConsoleWindow();

            // Hide Console
            ShowWindow(handle, SW_HIDE);

            bool haveFile = false;

            //Login Part
            //if (!LogIn())
            //{
            //    return;
            //}

            Dictionary<String, String[]> fileList = GetFilesOutlook();

            //Making Color lmao!
            Console.Write("Scanning...");
            System.Threading.Thread.Sleep(1000);
            Console.Write("Done!");
            System.Threading.Thread.Sleep(500);

            //Clear screen this time!
            Console.Clear();

            if (fileList["withExtension"].Length != 0)
            {
                foreach (string file in fileList["withExtension"])
                {
                    String Sub = "";
                    String Bod = "";
                    List<String> recipient = new List<string>();
                    try
                    {
                        switch (file.Substring(0, 3))
                        {
                            case "218":
                                Console.WriteLine(file);
                                recipient.Add("vuong.nc0582@gmail.com");
                                recipient.Add("vuong.photo92@gmail.com");
                                Sub = "DKSHVN HEC - " + file.Remove(file.Length - 4);
                                Bod = "Kính gửi Quý khách hàng<p>Chúng tôi xin gửi đến Quý khách hóa đơn tài chính cho các đơn hàng đã được đặt theo yêu cầu.</p><p>Quý khách vui lòng kiểm tra và phản hồi cho chúng tôi ngay nếu như Quý khách phát hiện có sai sót trên hóa đơn. Theo quy định của nhà nước hóa đơn chỉ được phép hủy khi hàng hóa chưa bàn giao ký nhận và hai bên chưakê khai thuế do đó DKSH chỉ chấp nhận hủy hóa đơn nếu như Quý khách hàng phản hồi cho chúng tôi trước hay ngay khi giao hàng. Các phản hồi sau khi giao hàng sẽ được xử lý bằng cách lập biên bản điều chỉnh và hóa đơnđiều chỉnh.</p>Đối với các đơn hàng theo hình thức thanh toán là chuyển khoản, Quý khách hàng vui lòng chuyển khoản cho chúng tôi theo số tài khoản như sau</p><p>-          Tên tài khoản : Công ty TNHH DKSH VIET NAM<br>-          Số tài khoản : <b>6339402</b><br>-          Tên ngân hàng : <b>ANZ – CHI NHANH HCM</b><p>Sau khi chuyển tiền, để thuận tiện cho việc theo dõi đơn hàng, Quý khách vui lòng Fax hoặc scan Ủy nhiệm chi cho chúng tôi theo số: <b>1800 54 54 20 (miễn cước)</b> hoặc email  customercare.vn@dksh.com</p><p>Chúng tôi sẽ tiến hành chuyển hàng cho Quý khách theo tuyến giao hàng gần nhất ngay khi nhận được thông tin xác nhận từ Ngân hàng.</p><p>Lưu ý:Quý khách vui lòng thực hiện chuyển khoản trong vòng 4 ngày kể từ ngày nhận được thông báo này. Sau thời gian trên,nếu chúng tôi chưa nhậnđược xác nhận của Ngân hàng về việc thanh toán của Quý khách, chúng tôi buộc lòng hủy hóa đơn theo qui định mà không cần phải thông báo cho Quý khách.</p><p>Nếu Quý khách hàng có bất cứ thắc mắc nào liên quan đến đơn hàng,hàng hóa, hóa đơn vui lòng liên hệ với chúng tôi qua số điện thoại tổng đài miễn cước 1800.54.54.02</p><p>Công ty DKSH VN rất hân hạnh được phục vụ quý khách và mong tiếp tục nhận được đơn đặt hàng của Quý khách trong tương lai.</p><p>Chân trọng</p>";
                                break;
                            case "207":
                                Console.WriteLine(file);
                                recipient.Add("vuong.nc0582@gmail.com");
                                Sub = "DKSHVN CG - " + file.Remove(file.Length - 4);
                                Bod = "Kính gửi Quý khách hàng<p>Chúng tôi xin gửi đến Quý khách hóa đơn tài chính cho các đơn hàng đã được đặt theo yêu cầu.</p><p>Quý khách vui lòng kiểm tra và phản hồi cho chúng tôi ngay nếu như Quý khách phát hiện có sai sót trên hóa đơn. Theo quy định của nhà nước hóa đơn chỉ được phép hủy khi hàng hóa chưa bàn giao ký nhận và hai bên chưakê khai thuế do đó DKSH chỉ chấp nhận hủy hóa đơn nếu như Quý khách hàng phản hồi cho chúng tôi trước hay ngay khi giao hàng. Các phản hồi sau khi giao hàng sẽ được xử lý bằng cách lập biên bản điều chỉnh và hóa đơnđiều chỉnh.</p>Đối với các đơn hàng theo hình thức thanh toán là chuyển khoản, Quý khách hàng vui lòng chuyển khoản cho chúng tôi theo số tài khoản như sau</p><p>-          Tên tài khoản : Công ty TNHH DKSH VIET NAM chi nhánh tại Hà Nội<br>-          Số tài khoản : <b>002-139707-001</b><br>-          Tên ngân hàng : <b>Ngan hang TNHH MTV HSBC Viet nam</b><p>Sau khi chuyển tiền, để thuận tiện cho việc theo dõi đơn hàng, Quý khách vui lòng Fax hoặc scan Ủy nhiệm chi cho chúng tôi theo số: <b>0650 3766 610-0650 3766636</b> hoặc email Customercare.Consumergoods@dksh.com</p><p>Chúng tôi sẽ tiến hành chuyển hàng cho Quý khách theo tuyến giao hàng gần nhất ngay khi nhận được thông tin xác nhận từ Ngân hàng.</p><p>Lưu ý:Vui lòng chuyển tiền và gửi UNC cho DKSH trước 15h (đối với khu vực tỉnh) và 16h(đối với khu vực HCM, HN, DN) để được giao hàng vào ngàymai, nếu gửi sau thời gian này đơn hàng của quý khách sẽ được giữ lại và chuyển vào chuyến kế tiếp căn cứ theo lịch giao hàng</p><p>Nếu Quý khách hàng có bất cứ thắc mắc nào liên quan đến đơn hàng,hàng hóa, hóa đơn vui lòng liên hệ với chúng tôi qua số điện thoại tổng đài miễn cước 1800.54.54.05</p><p>Công ty DKSH VN rất hân hạnh được phục vụ quý khách và mong tiếp tục nhận được đơn đặt hàng của Quý khách trong tương lai.</p><p>Chân trọng</p>";
                                break;
                            case "227":
                                Console.WriteLine(file);
                                recipient.Add("vuong.nc0582@gmail.com");
                                Sub = "DKSHVN PM - " + file.Remove(file.Length - 4);
                                Bod = "Kính gửi Quý khách hàng<p>Chúng tôi xin gửi đến Quý khách hóa đơn tài chính cho các đơn hàng đã được đặt theo yêu cầu.</p><p>Quý khách vui lòng kiểm tra và phản hồi cho chúng tôi ngay nếu như Quý khách phát hiện có sai sót trên hóa đơn. Theo quy định của nhà nước hóa đơn chỉ được phép hủy khi hàng hóa chưa bàn giao ký nhận và hai bên chưakê khai thuế do đó DKSH chỉ chấp nhận hủy hóa đơn nếu như Quý khách hàng phản hồi cho chúng tôi trước hay ngay khi giao hàng. Các phản hồi sau khi giao hàng sẽ được xử lý bằng cách lập biên bản điều chỉnh và hóa đơnđiều chỉnh.</p>Đối với các đơn hàng theo hình thức thanh toán là chuyển khoản, Quý khách hàng vui lòng chuyển khoản cho chúng tôi theo số tài khoản như sau</p><p>-          Tên tài khoản : Công ty TNHH DKSH VIET NAM<br>-          Số tài khoản : <b>002-139707-001</b><br>-          Tên ngân hàng : <b>Ngan hang TNHH MTV HSBC Viet nam</b><p>Sau khi chuyển tiền, để thuận tiện cho việc theo dõi đơn hàng, Quý khách vui lòng Fax hoặc scan Ủy nhiệm chi cho chúng tôi theo số: <b>08 3812 577 3 </b> hoặc email: pm_saleassistant@dksh.com</p><p>Chúng tôi sẽ tiến hành chuyển hàng cho Quý khách theo tuyến giao hàng gần nhất ngay khi nhận được thông tin xác nhận từ Ngân hàng.</p><p><b>Nếu Quý khách hàng có bất cứ thắc mắc nào liên quan đến đơn hàng ,hànghóa, hóa đơn vui lòng liên hệ với chúng tôi theo danh sách sau</b></p><p><b>1. Nguyễn Thị Kim Yến: Nhóm hàng Diversey Foodcare<br>Phone +84 8 3812 5848 Ext. 203; Fax +84 8 3812 5845<br>Mobile +84 976 932432; Email: yen.thikim.nguyen@dksh.com<br>2. Lý Thị Mai Hương: Nhóm hàng Diversey Care<br>Phone +84 8 3812 5848 Ext. 725, Fax +84 8 3812 5845<br>Mobile +84 909 544 205; Email: huong.thimai.ly@dksh.com<br>3. Hồ Thị Thùy Dương: Nhóm hàng hóa chất Mỹ phẩm; hóa chất Thực phẩm và đồ uống<br>Phone +84 8 3812 5848 Ext. 429; Fax +84 8 3812 5845<br>Mobile: Mobile +84 908 783 090; Email: duong.thithuy.ho@dksh.com<br>4. Nguyễn Thị Minh Phúc: Nhóm hàng hóa chất Công nghiệp<br>Phone + 84 8 3812 5848 Ext: 171; Fax +84 8 3812 5845<br>Mobile: +84 976 101 671; Email: sci.support@dksh.com< br > 5.Bùi Thị Phương Thi: Trưởng nhóm -Phone + 84 8 3812 5848 Ext.162 < br > Email: Thi.thiphuong.bui @dksh.com,</ b ></ br >< p > Công ty DKSH VN rất hân hạnh được phục vụ quý khách và mong tiếp tục nhận được đơn đặt hàng của Quý khách trong tương lai.</ p >< p > Chân trọng </ p > ";
                                break;
                            default:
                                continue;
                        }

                        if (recipient.Count > 0)
                        {
                            haveFile = true;
                        }

                    }
                    catch (System.Exception)
                    {
                        //show Console
                        ShowWindow(handle, SW_SHOW);
                        Console.Write("Something wrong with Recipient");
                        System.Threading.Thread.Sleep(800);
                    }

                    Dictionary<String, Object> info = new Dictionary<string, Object>();
                    info.Add("to", recipient);
                    info.Add("path", file);
                    sendMail(info, Sub, Bod);
                }
                if (haveFile)
                {
                    //ShowWindow(handle, SW_SHOW);
                    Console.Clear();
                    Console.WriteLine("Finished!");
                    System.Threading.Thread.Sleep(1000);
                }
                else
                {
                    ShowWindow(handle, SW_SHOW);
                    Console.Write("Nothing to be sent.");
                    System.Threading.Thread.Sleep(1000);
                }
            }
        }
    }
}