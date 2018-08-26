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

        public static DES dd = new DES();
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

            DateTime sendTime = Convert.ToDateTime(MailSendTime);
            DateTime nowTime = DateTime.Now;

            Runtime.ShowLog("定时发送时间为:" + sendTime);
            Runtime.ShowLog("当前的时间为:" + nowTime);

        }

        //测试的发送邮件
        public void SendMailBy163MailService(List<string> attachFileList)
        {
            try
            {
                DateTime nowTime = DateTime.Now;
                Email myEmail = new Email();
                myEmail.host = "smtp.163.com";
                myEmail.mailSshPwd = MailSshPwd;
                myEmail.mailFrom = MailFrom;
                myEmail.mailToArray = MailToStr.Split(';');
                myEmail.mailCcArray = MailToCcStr.Split(';');
                myEmail.mailSubject = MailSubject + "(" + nowTime.ToString("yyyy-MM-dd HH:mm:ss") + ") 测试成功";
                myEmail.mailBody = MailBody + "(" + nowTime.ToString("yyyy-MM-dd HH:mm:ss") + ") 测试成功"; ;
                myEmail.attachmentsPath = attachFileList.ToArray();
                if (myEmail.Send())
                {
                    Config.log.Info("Email 发送 成功!");
                    Runtime.ShowLog("Email 发送 成功!");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now.ToLongTimeString() + "Email 发送 失败! 详细：" + ex.Message);
                Config.log.Info(DateTime.Now.ToLongTimeString() + "Email 发送 失败! 详细：" + ex.Message);
                Runtime.ShowLog("Email 发送 失败! 详细：" + ex.Message);
            }

        }
        //定时发送邮件
        public void RunSendMailLoop()
        {
            while (Runtime.m_IsRunning)
            {
                try
                {
                    DateTime sendTime = Convert.ToDateTime(MailSendTime);
                    DateTime nowTime = DateTime.Now;

                    if (nowTime.Hour.Equals(sendTime.Hour) && nowTime.Minute.Equals(sendTime.Minute))
                    {

                        MailAttachmentsList.Clear();

                        string[] MailPathArry = MailAttachmentsPath.Split(';');
                        if (MailPathArry != null)
                        {
                            for (int i = 0; i < MailPathArry.Length; i++)
                            {
                                //获取一个文件夹内全部Excel文件
                                List<string> mailAttachmentsLastOne = new List<string>();
                                DirectoryInfo folder = new DirectoryInfo(MailPathArry[i].ToString());
                                foreach (FileInfo file in folder.GetFiles("*.xls"))
                                {
                                    mailAttachmentsLastOne.Add(file.FullName);
                                    Config.log.Info("当前路径：" + folder + "  中 的文件为：" + file.FullName);
                                    Runtime.ShowLog("当前路径：" + folder + "  中 的文件为：" + file.FullName);
                                }
                                //获取最新的一个文件
                                mailAttachmentsLastOne.Sort();
                                if (mailAttachmentsLastOne != null && mailAttachmentsLastOne.Count > 0)
                                {
                                    MailAttachmentsList.Add(mailAttachmentsLastOne.Last());
                                    Runtime.ShowLog("当前路径：" + folder + "  中 的最新一个文件为：" + MailAttachmentsList.Last());
                                    Config.log.Info("当前路径：" + folder + "  中 的最新一个文件为：" + MailAttachmentsList.Last());
                                }
                            }
                        }

                        Email myEmail = new Email();
                        myEmail.host = "smtp.163.com";
                        myEmail.mailSshPwd = MailSshPwd;
                        myEmail.mailFrom = MailFrom;
                        myEmail.mailToArray = MailToStr.Split(';');
                        myEmail.mailCcArray = MailToCcStr.Split(';');
                        myEmail.mailSubject = MailSubject + "(" + nowTime.ToString("yyyy-MM-dd") + ")";
                        myEmail.mailBody = MailBody + "(" + nowTime.ToString("yyyy-MM-dd") + ")";
                        myEmail.attachmentsPath = MailAttachmentsList.ToArray();
                        if (myEmail.Send())
                        {
                            Config.log.Info("Email 发送 给：" + MailToStr + "成功");
                            Config.log.Info("Email 抄送 给：" + MailToCcStr + "成功");
                            //Config.log.Info("Email 附件为："+ MailAttachmentsLastOne.ToString());
                            Runtime.ShowLog("Email 发送 给：" + MailToStr + "成功");
                            //Runtime.ShowLog("Email 附件为："+ MailAttachmentsLastOne.ToString());
                        }
                    }
                    System.Threading.Thread.Sleep(60000 * 1);
                    continue;
                }
                catch (Exception ex)
                {
                    Runtime.ShowLog("！！！ 发送邮件失败 ！！！  详细：" + ex.Message);
                    Config.log.Error("！！！  发送邮件失败！！！  详细：" + ex.Message);
                    System.Threading.Thread.Sleep(60000 * 30);
                    continue;
                }
            }

        }

        //文件夹中按时间排序最新的文件读取
        public class DirectoryLastTimeComparer : IComparer<DirectoryInfo>
        {
            #region IComparer<DirectoryInfo> 成员

            public int Compare(DirectoryInfo x, DirectoryInfo y)
            {
                return x.LastWriteTime.CompareTo(y.LastWriteTime);
                //依名称排序
                //return x.FullName.CompareTo(y.FullName);//递增
                //return y.FullName.CompareTo(x.FullName);//递减

                //依修改日期排序
                //return x.LastWriteTime.CompareTo(y.LastWriteTime);//递增
                //return y.LastWriteTime.CompareTo(x.LastWriteTime);//递减
            }

            #endregion
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                Runtime.m_IsRunning = true;

                btnStart.Enabled = false;
                btnStart.BackColor = Color.Lime;

                btnStop.Enabled = true;
                btnStop.BackColor = Color.GhostWhite;
                btnTestSendMail.Enabled = false;
                //获取每个区域最新的文件
                MailAttachmentsList.Clear();
                string[] MailPathArry = MailAttachmentsPath.Split(';');
                if (MailPathArry != null)
                {
                    for (int i = 0; i < MailPathArry.Length; i++)
                    {
                        //获取一个文件夹内全部Excel文件
                        List<string> mailAttachmentsLastOne = new List<string>();
                        DirectoryInfo folder = new DirectoryInfo(MailPathArry[i].ToString());
                        foreach (FileInfo file in folder.GetFiles("*.xls"))
                        {
                            mailAttachmentsLastOne.Add(file.FullName);
                            Config.log.Info("当前路径：" + folder + "  中 的文件为：" + file.FullName);
                            Runtime.ShowLog("当前路径：" + folder + "  中 的文件为：" + file.FullName);
                        }
                        //获取最新的一个文件
                        mailAttachmentsLastOne.Sort();
                        if (mailAttachmentsLastOne != null && mailAttachmentsLastOne.Count > 0)
                        {
                            MailAttachmentsList.Add(mailAttachmentsLastOne.Last());
                            Runtime.ShowLog("当前路径：" + folder + "  中 的最新一个文件为：" + MailAttachmentsList.Last());
                            Config.log.Info("当前路径：" + folder + "  中 的最新一个文件为：" + MailAttachmentsList.Last());
                        }
                    }
                }
                //启动定时发送服务
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
            btnStart.BackColor = Color.GhostWhite;

            btnStop.Enabled = false;
            btnStop.BackColor = Color.Red;
            btnTestSendMail.Enabled = true;

            Runtime.m_IsRunning = false;

        }

        private void btnTestSendMail_Click(object sender, EventArgs e)
        {
            MailAttachmentsList.Clear();

            string[] MailPathArry = MailAttachmentsPath.Split(';');
            if (MailPathArry != null)
            {
                for (int i = 0; i < MailPathArry.Length; i++)
                {
                    //获取一个文件夹内全部Excel文件
                    List<string> mailAttachmentsLastOne = new List<string>();
                    DirectoryInfo folder = new DirectoryInfo(MailPathArry[i].ToString());
                    foreach (FileInfo file in folder.GetFiles("*.xls"))
                    {
                        mailAttachmentsLastOne.Add(file.FullName);
                        Config.log.Info("当前路径：" + folder + "  中 的文件为：" + file.FullName);
                        Runtime.ShowLog("当前路径：" + folder + "  中 的文件为：" + file.FullName);
                    }
                    //获取最新的一个文件
                    mailAttachmentsLastOne.Sort();
                    if (mailAttachmentsLastOne != null && mailAttachmentsLastOne.Count > 0)
                    {
                        MailAttachmentsList.Add(mailAttachmentsLastOne.Last());
                        Runtime.ShowLog("当前路径：" + folder + "  中 的最新一个文件为：" + MailAttachmentsList.Last());
                        Config.log.Info("当前路径：" + folder + "  中 的最新一个文件为：" + MailAttachmentsList.Last());
                    }
                }
            }

            SendMailBy163MailService(MailAttachmentsList);
        }

        //------------------窗体最小化，不退出软件--------------------------------------------------------
        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            //窗体关闭原因为单击"关闭"按钮或Alt+F4  
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;           //取消关闭操作 表现为不关闭窗体  
                this.Hide();               //隐藏窗体  
            }
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            //点击鼠标"左键"发生  
            if (e.Button == MouseButtons.Left)
            {
                this.Visible = true;                        //窗体可见  
                this.WindowState = FormWindowState.Normal;  //窗体默认大小  
                this.notifyIcon1.Visible = true;            //设置图标可见  
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //点击鼠标"左键"发生  
            if (e.Button == MouseButtons.Left)
            {
                this.Visible = true;                        //窗体可见  
                this.WindowState = FormWindowState.Normal;  //窗体默认大小  
                this.notifyIcon1.Visible = true;            //设置图标可见  
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Show();                                //窗体显示  
            this.WindowState = FormWindowState.Normal;  //窗体状态默认大小  
            this.Activate();                            //激活窗体给予焦点 
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            //点击"是(YES)"退出程序  
            if (MessageBox.Show("确定要退出程序?", "重要提示",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Warning)
                == System.Windows.Forms.DialogResult.Yes)
            {
                notifyIcon1.Visible = false;   //设置图标不可见  
                this.Close();                  //关闭窗体  
                this.Dispose();                //释放资源  
                Application.Exit();            //关闭应用程序窗体  
            }
        }
    }
}
