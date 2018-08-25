using MySendEmail.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace MySendEmail
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        //发件人邮箱
        public static string MailFrom = Config.GetValue("MailFrom");
        public static string MailSshPwd = Config.GetValue("MailSshPwd");
        //收件人地址，抄送地址（多个地址使用英文分号分割）
        public static string MailToStr = Config.GetValue("MailToStr");
        public static string MailToCcStr = Config.GetValue("MailToCcStr"); 
        //邮件主题及内容
        public static string MailSubject = Config.GetValue("MailSubject"); 
        public static string MailBody = Config.GetValue("MailBody");
        //邮件附件所在文件夹路径
        public static string MailAttachmentsPath = Config.GetValue("MailAttachmentsPath");
        //定时发送时间
        public static string MailSendTime = Config.GetValue("MailSendTime");

        /// <summary>
        /// 带发送全部附件的路径（一个文件夹内的所有Excel文件）
        /// </summary>
        List<string> MailAttachmentsList = new List<string>();

        private void Form1_Load(object sender, EventArgs e)
        {
            Runtime.ServerLog = this.ServerLog;
            Runtime.m_IsRunning = false;
            Runtime.ShowLog("初始化软件");
            Config.log.Info("初始化软件");
            txtBoxEmailAddress.Text = MailToStr;
            Runtime.ShowLog("发件人地址:" + MailFrom);
            Runtime.ShowLog("收件人地址:" + MailToStr);
            Runtime.ShowLog("抄送地址:" + MailToCcStr);
            textBoxSendTime.Text = MailSendTime;
            Runtime.ShowLog("定时发送时间为:" + MailSendTime);



        }

        public void SendMailBy163()
        {
            //163-- > qq   成功
            /*
                        SmtpClient client = new SmtpClient("smtp.163.com", 25);
                        client.EnableSsl = true;
                        client.Credentials = new NetworkCredential("edgarwen@163.com", "edgarwen1215");
                        MailMessage mail = new MailMessage();
                        mail.From = new MailAddress("edgarwen@163.com");
                        mail.To.Add(new MailAddress("522191646@qq.com"));
                        mail.IsBodyHtml = true;
                        mail.Priority = MailPriority.Normal;
                        mail.Subject = "这是主题";
                        mail.Body = "这是内容";
                        //电子邮件正文的编码
                        mail.BodyEncoding = Encoding.Default;

                        //在有附件的情况下添加附件
                        try
                        {
                            if (attachmentsPath != null && attachmentsPath.Length > 0)
                            {
                                Attachment attachFile = null;
                                foreach (string path in attachmentsPath)
                                {
                                    attachFile = new Attachment(path);
                                    mail.Attachments.Add(attachFile);
                                }
                            }
                        }
                        catch (Exception err)
                        {
                            throw new Exception("在添加附件时有错误:" + err);
                        }
                        client.Send(mail);
                        */
        }


        public void SendMailBy163MailService(List<string> attachFileList)
        {
            try
            {

                Email myEmail = new Email();
                myEmail.host = "smtp.163.com";
                myEmail.mailSshPwd = MailSshPwd;
                myEmail.mailFrom = MailFrom;
                myEmail.mailToArray = MailToStr.Split(';'); 
                myEmail.mailCcArray = MailToCcStr.Split(';');
                myEmail.mailSubject = MailSubject;
                myEmail.mailBody = MailBody;
                myEmail.attachmentsPath = attachFileList.ToArray();
                if (myEmail.Send())
                {
                    Config.log.Info("Email 发送 成功!");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now.ToLongTimeString() + "Email 发送 失败! 详细：" + ex.Message);
                Config.log.Info(DateTime.Now.ToLongTimeString() + "Email 发送 失败! 详细：" + ex.Message);
            }

        }

        public void RunSendMailLoop()
        {
            while (Runtime.m_IsRunning)
            {
                try
                {
                    //获取一个文件夹内全部Excel文件
                    DirectoryInfo folder = new DirectoryInfo(MailAttachmentsPath);
                    MailAttachmentsList.Clear();
                    foreach (FileInfo file in folder.GetFiles("*.xls"))
                    {
                        MailAttachmentsList.Add(file.FullName);
                        Runtime.ShowLog("当前路径：" + folder + "  中 的文件为：" + file.FullName);
                    }

                    Email myEmail = new Email();
                    myEmail.host = "smtp.163.com";
                    myEmail.mailSshPwd = MailSshPwd;
                    myEmail.mailFrom = MailFrom;
                    myEmail.mailToArray = MailToStr.Split(';');
                    myEmail.mailCcArray = MailToCcStr.Split(';');
                    myEmail.mailSubject = MailSubject;
                    myEmail.mailBody = MailBody;
                    myEmail.attachmentsPath = MailAttachmentsList.ToArray();
                    if (myEmail.Send())
                    {
                        Config.log.Info("Email 发送 成功!");
                        Runtime.ShowLog("Email 发送 成功!");
                    }

                    System.Threading.Thread.Sleep(60000*60);
                    continue;
                }
                catch (Exception ex)
                {
                    Runtime.ShowLog("！！！ 发送邮件失败 ！！！  详细：" + ex.Message);
                   Config.log.Error("！！！  发送邮件失败！！！  详细：" + ex.Message);
                    System.Threading.Thread.Sleep(60000*30);
                    continue;
                }
            }

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                Runtime.m_IsRunning = true;
                //获取一个文件夹内全部Excel文件
                DirectoryInfo folder = new DirectoryInfo(MailAttachmentsPath);
                MailAttachmentsList.Clear();
                foreach (FileInfo file in folder.GetFiles("*.xls"))
                {
                    MailAttachmentsList.Add(file.FullName);
                    Runtime.ShowLog("当前路径："+ folder+"  中 的文件为：" + file.FullName);
                }

                Thread StartSendEmailThread = new Thread(RunSendMailLoop);
                StartSendEmailThread.Start();

            }
            catch (Exception ex)
            {
                Runtime.ShowLog("！！！ 启动服务失败！！！  详细：" + ex.Message);
                Config.log.Error("！！！ 启动服务失败！！！  详细：" + ex.Message);
            }
        }



        private void btnStopServer_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            Runtime.m_IsRunning = false;

        }

        private void btnTestSendMail_Click(object sender, EventArgs e)
        {
            //获取一个文件夹内全部Excel文件
            DirectoryInfo folder = new DirectoryInfo(MailAttachmentsPath);
            MailAttachmentsList.Clear();
            foreach (FileInfo file in folder.GetFiles("*.xls"))
            {
                MailAttachmentsList.Add(file.FullName);
                Runtime.ShowLog("当前路径：" + folder + "  中 的文件为：" + file.FullName);
            }
            SendMailBy163MailService(MailAttachmentsList);
        }
    }
}
