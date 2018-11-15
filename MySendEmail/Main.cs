using MySendEmail.Common;
using MySendEmail.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
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
        //Thread StartSendEmailThread = new Thread(RunSendMailLoop);

        ManualResetEvent _shutdownEvent = new ManualResetEvent(false);
        ManualResetEvent _pauseEvent = new ManualResetEvent(true);

        Thread _thread;

        public static int CheckTime = Convert.ToInt16(Config.GetValue("CheckTime"));
        public static DES dd = new DES();
        //发件人邮箱
        public static string MailFrom = Config.GetValue("MailFrom");
        public static string MailSshPwd = Config.GetValue("MailSshPwd");
        //收件人地址，抄送地址（多个地址使用英文分号分割）
        public static string MailToStr = Config.GetValue("MailToStr");
        public static string MailToCcStr = Config.GetValue("MailToCcStr");
        //接收软件运行异常告警邮箱地址
        public static string ReceiverAlarm = Config.GetValue("ReceiverAlarm");

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
        static List<string> MailAttachmentsList = new List<string>();

        private void Form1_Load(object sender, EventArgs e)
        {
            Runtime.ServerLog = this.ServerLog;

            Runtime.m_IsRunning = false;
            btnStart.Text = "启动服务";
            btnStop.Text = "服务已停止";

            Runtime.ShowLog("初始化软件");
            Config.log.Info("初始化软件");
            txtBoxEmailAddress.Text = MailToStr;
            textBoxSendTime.Text = MailSendTime;

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
                myEmail.mailSubject = MailSubject + "(" + nowTime.ToString("yyyy-MM-dd HH:mm:ss") + ") ";

                //判断附件是否为空
                if (MailAttachmentsList.Count==0 || MailAttachmentsList == null)
                {
                    myEmail.mailBody = "内容：尊敬的用户，昨日" + "(" + nowTime.AddDays(-1).ToString("yyyy-MM-dd") + ")"
                        + "的用能报表尚未生成，请检查系统是否运行正常";
                }
                else
                {
                    myEmail.mailBody = MailBody + "(" + nowTime.AddDays(-1).ToString("yyyy-MM-dd") + ")";
                }
                //myEmail.mailBody = MailBody + "(" + nowTime.ToString("yyyy-MM-dd HH:mm:ss") + ")"; ;
                myEmail.attachmentsPath = attachFileList.ToArray();

                if (myEmail.Send())
                {
                    Config.log.Info("Email 发送 成功!");
                    Runtime.ShowLog("Email 发送 成功!");
                    string emailAttachmnets = null;
                    foreach (var item in attachFileList)
                    {
                        emailAttachmnets += item;
                    }
                    // 将发送成功的邮件 存入数据库
                    string sql = @"INSERT INTO SendEmailResult ( Sender,Receiver,SendTime,State,MailSubject,MailBody,Attachments)
                                                                VALUES(@Sender,@Receiver,@SendTime,@State,@MailSubject,@MailBody,@Attachments); ";

                    SQLiteParameter[] parameters =  {
                                new SQLiteParameter("@Sender", myEmail.mailFrom),
                                new SQLiteParameter("@Receiver",MailToStr),
                                new SQLiteParameter("@SendTime",DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                                new SQLiteParameter("@State",1),
                                new SQLiteParameter("@MailSubject",myEmail.mailSubject),
                                new SQLiteParameter("@MailBody",myEmail.mailBody),
                                new SQLiteParameter("@Attachments", emailAttachmnets)
                             };

                    int sqlResult = SqliteHelper.ExecuteNonQuery(sql, parameters);
                    if (sqlResult >= 1)
                    {
                        Config.log.Info("Email 用能日报 写入数据库完成");
                        Runtime.ShowLog("Email 用能日报 写入数据库完成");
                    }
                }

            }
            catch (Exception ex)
            {
                Config.log.Info("Email 发送 失败! 详细：" + ex.Message);
                Runtime.ShowLog("Email 发送 失败! 详细：" + ex.Message);
            }

        }

        //定时发送邮件
        public static void RunSendMailLoop()
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
                                foreach (FileInfo file in folder.GetFiles("*.csv"))
                                {
                                    mailAttachmentsLastOne.Add(file.FullName);
                                    Config.log.Info("当前路径：" + folder + "  中 的文件为：" + file.Name);
                                    Runtime.ShowLog("当前路径：" + folder + "  中 的文件为：" + file.Name);
                                }
                                //获取最新的一个文件
                                mailAttachmentsLastOne.Sort();
                                if (mailAttachmentsLastOne != null && mailAttachmentsLastOne.Count > 0)
                                {
                                    string sql = @"SELECT COUNT(1) FROM SendEmailResult WHERE Attachments LIKE '%" + mailAttachmentsLastOne.Last() + "%';";
                                    
                                    int sqlResult = Convert.ToInt16(SqliteHelper.ExecuteScalar(sql, null));
                                    if (sqlResult == 0)
                                    {
                                        MailAttachmentsList.Add(mailAttachmentsLastOne.Last());
                                        Runtime.ShowLog("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 未发送! 已添加到附件中");
                                        Config.log.Info("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 未发送! 已添加到附件中");
                                    }
                                    else
                                    {
                                        Runtime.ShowLog("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 已发送");
                                        Config.log.Info("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 已发送");
                                    }
                                }
                            }
                        }

                        
                        Email myEmail = new Email();
                        myEmail.host = "smtp.163.com";
                        myEmail.mailSshPwd = MailSshPwd;
                        myEmail.mailFrom = MailFrom;
                        myEmail.mailToArray = MailToStr.Split(';');
                        myEmail.mailCcArray = MailToCcStr.Split(';');
                        myEmail.mailSubject = MailSubject + "(" + nowTime.AddDays(-1).ToString("yyyy-MM-dd") + ")";
                        //判断附件是否为空
                        if (MailAttachmentsList.Count == 0 || MailAttachmentsList == null)
                        {
                            myEmail.mailBody = "内容：尊敬的用户，昨日" + "(" + nowTime.AddDays(-1).ToString("yyyy-MM-dd") + ")"
                                +"的用能报表尚未生成，请检查系统是否运行正常";
                        }
                        else {
                            myEmail.mailBody = MailBody + "(" + nowTime.AddDays(-1).ToString("yyyy-MM-dd") + ")";
                        }
                       
                        myEmail.attachmentsPath = MailAttachmentsList.ToArray();
                        if (myEmail.Send())
                        {
                            Config.log.Info("Email 发送 给：" + MailToStr + "成功");
                            Config.log.Info("Email 抄送 给：" + MailToCcStr + "成功");

                            string emailAttachmnets = null;
                            foreach (var item in MailAttachmentsList)
                            {
                                emailAttachmnets += item;
                            }
                            // 将发送成功的邮件 存入数据库
                            EmailModel eModel = new EmailModel();
                            eModel.Sender = myEmail.mailFrom;
                            eModel.Receiver = MailToStr;
                            eModel.CarbonCopy = MailToCcStr;
                            eModel.SendTime = nowTime.ToString("yyyy-MM-dd HH:mm:ss");
                            eModel.SendState = 1;
                            eModel.Subject = myEmail.mailSubject;
                            eModel.Body = myEmail.mailBody;
                            eModel.Attachment = emailAttachmnets;

                            int sqlResult = SqliteHelper.SetEmailToDB(eModel);
                            if (sqlResult == 1)
                            {
                                Config.log.Info("Email 用能日报表 写入数据库 完成！");
                                Runtime.ShowLog("Email 用能日报表 写入数据库 完成！");
                            }
                            else
                            {
                                Config.log.Info("Email 用能日报表 写入数据库 失败！");
                                Runtime.ShowLog("Email 用能日报表 写入数据库 失败！");
                            }

                            //Config.log.Info("Email 附件为："+ MailAttachmentsLastOne.ToString());
                            Runtime.ShowLog("Email 发送 给：" + MailToStr + "成功");
                            //Runtime.ShowLog("Email 附件为："+ MailAttachmentsLastOne.ToString());
                        }
                    }

                    //定时监测软件是否正常运行
                    if (nowTime.Minute % CheckTime == 0)
                    {
                        ProcessState processState = new ProcessState();

                        List<ProcessState> processStates = processState.GetProcessState();
                        int processStopCount = 0;

                        foreach (var ps in processStates)
                        {
                            if (ps.State == 0)
                                processStopCount++;

                            Runtime.ShowLog("当前软件运行状态：" + " 软件名称：" + ps.ProcessName + "; 运行状态： " + ps.State + "; 检查时间： " + ps.UpdateTime);
                            Config.log.Info("当前软件运行状态：" + " 软件名称：" + ps.ProcessName + "; 运行状态： " + ps.State + "; 检查时间： " + ps.UpdateTime);
                        }
                        //出现2个以上的软件没有运行则发送邮件提示
                        if (processStopCount >= 2)
                        {
                            SendProcesssReport(processStates);
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




        /// <summary>
        /// 启动发送服务
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
               
                textBoxSendTime.Text = MailSendTime;
                Runtime.ShowLog("发件人地址:" + MailFrom);
                Config.log.Info("发件人地址:" + MailFrom);

                Runtime.ShowLog("收件人地址:" + MailToStr);
                Config.log.Info("收件人地址:" + MailToStr);

                Runtime.ShowLog("抄送地址:" + MailToCcStr);
                Config.log.Info("抄送地址:" + MailToCcStr);

                Runtime.ShowLog("定时发送时间为:" + MailSendTime);
                Config.log.Info("定时发送时间为:" + MailSendTime);

                Runtime.ShowLog("当前的时间为:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                Config.log.Info("当前的时间为:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

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
                //Thread StartSendEmailThread = new Thread(RunSendMailLoop);
                //StartSendEmailThread.Start();
                if (_thread == null)
                {
                    _thread = new Thread(RunSendMailLoop);
                    _thread.IsBackground = true;
                    _thread.Start();
                    Runtime.ShowLog("服务启动：" + _thread.ManagedThreadId + "  " + _thread.Name + ": " + _thread.ThreadState);

                }
                else
                {
                    Resume();
                    Runtime.ShowLog("恢复服务：" + _thread.ManagedThreadId + "  " + _thread.Name + ": " + _thread.ThreadState);
                }
                //启动成功后，设置按钮显示
                Runtime.m_IsRunning = true;

                btnStart.Enabled = false;
                btnStart.Text = "服务已运行";
                btnStart.BackColor = Color.Lime;

                btnStop.Enabled = true;
                btnStop.Text = "停止服务";
                btnStop.BackColor = Color.GhostWhite;
                btnTestSendMail.Enabled = false;


            }
            catch (Exception ex)
            {
                Runtime.ShowLog("！！！ 启动服务失败！！！  详细：" + ex.Message);
                Config.log.Error("！！！ 启动服务失败！！！  详细：" + ex.Message);
            }
        }

        private void btnStopServer_Click(object sender, EventArgs e)
        {
            try
            {
                btnStart.Enabled = true;
                btnStart.Text = "启动服务";
                btnStart.BackColor = Color.GhostWhite;

                btnStop.Enabled = false;
                btnStop.Text = "服务已停止";
                btnStop.BackColor = Color.Red;
                btnTestSendMail.Enabled = true;
                //暂停线程
                Pause();
                //Stop();
                Runtime.ShowLog("停止服务");
                Runtime.m_IsRunning = false;
            }
            catch (Exception ex)
            {

                Runtime.ShowLog("停止服务 失败: "+ex.Message);
            }
            
        }

        /// <summary>
        /// 测试按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                    foreach (FileInfo file in folder.GetFiles("*.csv"))
                    {
                        mailAttachmentsLastOne.Add(file.FullName);
                        Config.log.Info("当前路径：" + folder + "  中 的文件为：" + file.Name);
                        Runtime.ShowLog("当前路径：" + folder + "  中 的文件为：" + file.Name);
                    }
                    //获取最新的一个文件
                    mailAttachmentsLastOne.Sort();
                    if (mailAttachmentsLastOne != null && mailAttachmentsLastOne.Count > 0)
                    {
                        string sql = @"SELECT COUNT(1) FROM SendEmailResult WHERE Attachments LIKE '%"+ mailAttachmentsLastOne.Last() +"%';";
                       
                        int sqlResult = Convert.ToInt16( SqliteHelper.ExecuteScalar(sql, null));
                        if (sqlResult == 0)
                        {
                            MailAttachmentsList.Add(mailAttachmentsLastOne.Last());
                            Runtime.ShowLog("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 未发送! 已添加到附件中");
                            Config.log.Info("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 未发送! 已添加到附件中");
                        }
                        else {
                            Runtime.ShowLog("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 已发送");
                            Config.log.Info("当前路径：" + folder + "  中 的最新一个文件为：" + mailAttachmentsLastOne.Last() + " 已发送");
                        }
                        
                    }
                }
            }
            ProcessState processState = new ProcessState();

            List<ProcessState> processStates = processState.GetProcessState();
            int processStopCount = 0;

            foreach (var ps in processStates)
            {
                if (ps.State == 0)
                    processStopCount++;

                Runtime.ShowLog("当前软件运行状态：" + " 软件名称：" + ps.ProcessName + "; 运行状态： " + ps.State + "; 检查时间： " + ps.UpdateTime);
            }
            if (processStopCount >= 2)
            {
                //SendProcesssReport(processStates);
            }
            //发送附件
            SendMailBy163MailService(MailAttachmentsList);
        }

        public static void SendProcesssReport(List<ProcessState> processStates)
        {
            try
            {
                string emailBody = "**********  监测到软件运行异常，可能影响系统正常运行; 请及时检查电力监控系统是否运行正常？ ********* \n  当前所监测软件运行状态如下：\n";
                foreach (var ps in processStates)
                {
                    if (ps.State == 1)
                        emailBody += "  软件名称: " + ps.ProcessName + ";  状态 : 运行  \n";
                    else
                        emailBody += "  软件名称: " + ps.ProcessName + ";  状态 : 停止（或异常） \n";
                }

                DateTime nowTime = DateTime.Now;
                Email myEmail = new Email();
                myEmail.host = "smtp.163.com";
                myEmail.mailSshPwd = MailSshPwd;
                myEmail.mailFrom = MailFrom;
                myEmail.mailToArray = ReceiverAlarm.Split(';');
                myEmail.mailCcArray = MailToCcStr.Split(';');
                myEmail.mailSubject = "监测软件运行异常提示 " + "(" + nowTime.ToString("yyyy-MM-dd HH:mm") + ") ";
                myEmail.mailBody = emailBody + "******************************************************************************************************"
                    + " \n  本次监测时间：" + nowTime.ToString("yyyy-MM-dd HH:mm:ss");
                //myEmail.attachmentsPath = attachFileList.ToArray();

                if (myEmail.Send())
                {
                    Config.log.Info("Email: 软件监测告警 发送 成功!");
                    Runtime.ShowLog("Email: 软件监测告警 发送 成功!");

                    // 将发送成功的邮件 存入数据库
                    EmailModel eModel = new EmailModel();
                    eModel.Sender = myEmail.mailFrom;
                    eModel.Receiver = ReceiverAlarm;
                    eModel.CarbonCopy = MailToCcStr;
                    eModel.SendTime = nowTime.ToString("yyyy-MM-dd HH:mm:ss");
                    eModel.SendState = 1;
                    eModel.Subject = myEmail.mailSubject;
                    eModel.Body = myEmail.mailBody;

                    int sqlResult = SqliteHelper.SetEmailToDB(eModel);
                    if (sqlResult == 1)
                    {
                        Config.log.Info("Email 软件监测告警 写入数据库完成！");
                        Runtime.ShowLog("Email 软件监测告警 写入数据库完成！");
                    }
                    else
                    {
                        Config.log.Info("Email 软件监测告警 写入数据库 失败！");
                        Runtime.ShowLog("Email 软件监测告警 写入数据库 失败！");
                    }
                }

            }
            catch (Exception ex)
            {
                Config.log.Info(DateTime.Now.ToLongTimeString() + "Email 软件监测告警 发送 失败! 详细：" + ex.Message);
                Runtime.ShowLog("Email 软件监测告警 发送 失败! 详细：" + ex.Message);

            }
        }



        public void Start(ThreadStart DoWork)
        {
            _thread = new Thread(DoWork);
            _thread.Start();
        }

        public void Pause()
        {
            _pauseEvent.Reset();
        }

        public void Resume()
        {
            _pauseEvent.Set();
        }

        public void Stop()
        {
            // Signal the shutdown event
            _shutdownEvent.Set();

            // Make sure to resume any paused threads
            _pauseEvent.Set();

            // Wait for the thread to exit
            _thread.Join();
            _thread = null;
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

        //------------------窗体最小化，不退出软件----------------------------
        #region 窗体最小化，不退出软件
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
        #endregion 窗体最小化，不退出软件

    }
}
