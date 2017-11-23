using Air5DurationTool.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Collections;
using Infomedia.EB.Plugins;
using Infomedia.EB.Common.Models.Player;
using Infomedia.EB.Common.Enums;

namespace Air5DurationTool
{
    public partial class FrmMain : Form
    {
        private Dart.PowerTCP.Ftp.Ftp ftp1;
        private Dart.PowerTCP.Ftp.Ftp ftp2;
        bool isStart;
        public bool IsStart
        {
            get
            {
                return isStart;
            }

            set
            {
                isStart = value;
                isRepeat = true;
                if (isStart)
                {
                    this.btnStart.Image = global::Air5DurationTool.Properties.Resources.stop;

                    this.btnStart.Text = "停止";
                }
                else
                {
                    this.btnStart.Image = global::Air5DurationTool.Properties.Resources.Start;

                    this.btnStart.Text = "启动";
                }
            }
        }
        Thread ExcThread;
        public FrmMain()
        {
            InitializeComponent();
        }

        private void Init()
        {
            try
            {

                AudioPlayerEx.Instance.Init(AudioPlayerMode.Dll);

                if (!Directory.Exists(Application.StartupPath + "\\Temp"))
                {
                    Directory.CreateDirectory(Application.StartupPath + "\\Temp");
                }
                string sourceFile = @"C:\Users\sharp\Desktop\歌库时长校正\A\BACK ROUND.s48";
                AudioPlayInfo info = new AudioPlayInfo();
                info.FileName = sourceFile;
                info.Title = Path.GetFileName(sourceFile);
                info.Protocal = AudioPlayProtocal.Local;
                AudioPlayerEx.Instance.Request(info, false);
                AudioPlayerEx.Instance.Stop();
                long Duration_File = AudioPlayerEx.Instance.GetPlayLength(sourceFile);

            }
            catch (Exception ex)
            {

                LogService.WriteErr("系统错误，方法：Init() 信息：" + ex.Message);
            }
        }
        /// <summary>
        /// 日志输出
        /// </summary>
        /// <param name="pContent"></param>
        private void AppendMessageLine(string pContent)
        {
            this.txtInfo.AppendText(DateTime.Now + ":  " + pContent + "\n");
        }




        #region //任务执行

        private void TaskFunc()
        {
            try
            {
                this.Invoke(new Tasks(this.Task));

            }
            catch { }
        }

        private delegate void Tasks();
        private bool isRepeat = true;
        private void Task()
        {
            int countError = 0;
            try
            {
                if (txtInfo.Text.Length > txtInfo.MaxLength)
                    txtInfo.Text = "";
                Application.DoEvents();
                AppendMessageLine("开始查询数据！");
                LogService.Write("开始查询数据！");
                DataTable taskDataTable = new DataTable();
                using (SqlConnection conn = new SqlConnection(Globals.Air5ConnectionStr))
                {
                    //string sql = " select p.ItemID,(select Top 1 CategoryName from Categories c,Stacks s where c.CategoryID=s.StackID and s.ItemID=p.ItemID and c.CategoryType=p.CategoryType) AS CategoryName,p.Title ,p.Duration,p.DateAdded ,p.SoundStorageID ,p.SoundFileName from Items p where  p.Deleted=0 and  p.CategoryType=1 and p.AudioExists=1   order by CategoryName,DateAdded";
                    string sql = " select p.ItemID,(select Top 1 CategoryName from Categories c,Stacks s where c.CategoryID=s.StackID and s.ItemID=p.ItemID and c.CategoryType=p.CategoryType) AS CategoryName,p.Title ,p.Duration,p.DateAdded ,p.SoundStorageID ,p.SoundFileName from Items p where  p.Deleted=0 and  p.CategoryType=1 and p.AudioExists=1   and p.ItemID in('b427f772-a696-4f27-8629-64c2b56826e9', 'e05442e3-d816-436d-971c-02c0406101c6','8ccae3c5-48a4-4255-a946-277547d4b3b7','b977bacc-35f5-477a-9885-7cf958817a86','eb172285-9a99-4e17-aadb-265b45c5bfb0','5f8da943-9177-4224-a0a1-583fb6ce2637','804f513c-5085-4626-9a2b-720242580439','101c1b66-3574-4c29-9330-6cb0bd2c9126','0555e041-de7e-40a4-856f-4962b5be10fe','ab7e9612-c306-48ee-916a-ad7b86ea6233','f9e52e1f-1878-41ec-9f98-e404886a3d61','2ea578ba-5829-4691-b3ff-9bd2d9bcc153','d89b26e3-8780-404e-abd2-e1d297cdc0f6','026d67cd-c9e0-41a2-98c5-bf456190df19','c1cf2f3f-2635-4d2c-8448-b19ac1428ef6','e3eaf30d-82fb-4eb5-985a-1d1bd1675d6e','7cb93036-2dc5-4480-b944-88a22f16bb34','f124f6d0-7765-4738-b2ad-cc1735e85fbe','2729d4f7-9f1a-4fdc-9e22-8d1b78f9a166','547003c2-fada-42af-8f49-de95976021ba','4f65e092-f9cb-405a-90c9-df29a8513b5d','f6d8c6ec-db5d-41bd-8b77-948f7b6053d2','8f516592-fb3f-47c0-b476-d1ac9caac7af','f96b826c-e404-4803-a9af-5c60198f5ff9','6c4b9405-9e3e-42a5-a186-f420bd7c3199','abfa56b7-a0e8-40f0-a471-b7bc7130bfcc','763d86a5-e896-45f8-b7f2-c61b34b86437','e0ad22f5-90b7-4ec6-9e49-dcbd7cd525fa','0a0b3a27-8c59-4202-9f8b-d68bdf95bdf5','7fc5cdeb-02a4-4c33-bfc6-8870e15c9e04','32514566-8148-4b72-afd7-87ebdbcdad79','2c0a9177-165e-481e-b36b-e18a35c6f3ef','61a9f866-ae69-4a5a-8c2f-8ad3b9803ef6','25a1c394-40c5-4ab8-9e7a-c0539a0340de') order by CategoryName,DateAdded";

                    using (SqlDataAdapter sd = new SqlDataAdapter(sql, conn))
                    {
                        sd.Fill(taskDataTable);
                    }

                    if (taskDataTable == null || taskDataTable.Rows.Count <= 0)
                    {
                        if (isRepeat)
                        {
                            AppendMessageLine("未获取到新任务！");
                            LogService.Write("未获取到新任务！");
                        }
                        isRepeat = false;
                        return;
                    }
                    isRepeat = true;
                    AppendMessageLine("共获取到：" + taskDataTable.Rows.Count + "条数据");
                    int count = 0;

                    foreach (DataRow item in taskDataTable.Rows)
                    {
                        if (!isStart)
                        {
                            AppendMessageLine("任务停止！");
                            LogService.Write("任务停止！");
                            return;
                        }
                        count++;
                        LogService.Write(count.ToString());
                        AppendMessageLine(count.ToString());
                        string sourceFile = Application.StartupPath + "\\temp\\" + item["ItemID"].ToString() + ".s48";
                        try
                        {

                            Application.DoEvents();
                            string SQLItem = "select * from StorageZone where  StorageID='" + item["SoundStorageID"] + "' and ZoneID=" + Globals.Zone + " ";
                            using (SqlDataAdapter sd = new SqlDataAdapter(SQLItem, conn))
                            {
                                DataTable dtFTP = new DataTable();
                                sd.Fill(dtFTP);
                                if (dtFTP == null || dtFTP.Rows.Count <= 0)
                                {
                                    continue;
                                }
                                DataRow row = dtFTP.Rows[0];

                                if (File.Exists(sourceFile))
                                {
                                    File.Delete(sourceFile);
                                }
                                LogService.Write("开始下载：" + item["Title"]);
                                AppendMessageLine("开始下载：" + item["Title"]);
                                #region //下载音频

                                try
                                {
                                    using (this.ftp1 = new Dart.PowerTCP.Ftp.Ftp())
                                    {
                                        this.ftp1.Server = row["FTPServer1"].ToString();
                                        this.ftp1.ServerPort = int.Parse(row["FTPPort1"].ToString());
                                        this.ftp1.Username = row["FTPUser1"].ToString();
                                        this.ftp1.Password = row["FTPPassword1"].ToString();

                                        this.ftp1.Get(item["SoundFileName"].ToString().Replace("\\", "/"), sourceFile);

                                    }
                                }
                                catch
                                {

                                    try
                                    {
                                        using (this.ftp1 = new Dart.PowerTCP.Ftp.Ftp())
                                        {
                                            this.ftp1.Server = row["FTPServer2"].ToString();
                                            this.ftp1.ServerPort = int.Parse(row["FTPPort2"].ToString());
                                            this.ftp1.Username = row["FTPUser2"].ToString();
                                            this.ftp1.Password = row["FTPPassword2"].ToString();
                                            this.ftp1.Get(item["SoundFileName"].ToString().Replace("\\", "/"), sourceFile);

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        AppendMessageLine("音频文件下载失败！ItemID:" + item["ItemID"].ToString());
                                        LogService.Write("音频文件下载失败！ItemID:" + item["ItemID"].ToString());
                                        continue;
                                    }
                                }
                                #endregion
                                LogService.Write("开始计算时长：" + item["Title"]);
                                AppendMessageLine("开始计算时长：" + item["Title"]);
                                #region 比较时长 记录信息
                                long Duration_DB = Convert.ToInt64(item["Duration"].ToString());
                                AudioPlayInfo info = new AudioPlayInfo();
                                info.FileName = sourceFile;
                                info.Title = Path.GetFileName(sourceFile);
                                info.Protocal = AudioPlayProtocal.Local;
                                AudioPlayerEx.Instance.Request(info, false);
                                AudioPlayerEx.Instance.Stop();
                                long Duration_File = AudioPlayerEx.Instance.GetPlayLength(sourceFile);
                                LogService.Write("Duration_DB：" + Duration_DB);
                                AppendMessageLine("Duration_File：" + Duration_File);
                                if (Duration_DB != Duration_File)
                                {
                                    countError++;
                                    string Info = item["CategoryName"].ToString().Replace(',', '_') + "," + item["Title"].ToString().Replace(',', '_') + "," + item["Duration"] + "," + item["DateAdded"] + "," + Duration_File + "," + item["ItemID"];
                                    LogService.WriteInfo(Info);
                                    #region 更新时长
                                    using (SqlCommand command = new SqlCommand())
                                    {
                                        try
                                        {
                                            conn.Open();
                                            command.Connection = conn;
                                            command.CommandText = "Update Items set  Duration=@Duration  where ItemID=@ItemID";
                                            command.CommandType = CommandType.Text;
                                            command.Parameters.Clear();
                                            command.Parameters.AddWithValue("@Duration", Duration_File);
                                            command.Parameters.AddWithValue("@ItemID", item["ItemID"].ToString());
                                            int isOK = command.ExecuteNonQuery();
                                            if (isOK > 0)
                                            {
                                                LogService.WriteInfoUpdateDB("OK" + "," + item["ItemID"].ToString() + "," + Duration_File + "," + Duration_DB);
                                            }
                                            else
                                            {
                                                LogService.WriteInfoUpdateDB("FAIL" + "," + item["ItemID"].ToString() + "," + Duration_File + "," + Duration_DB);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            LogService.WriteErr("更新数据库时长失败 节目ID：" + item["ItemID"].ToString() + " 错误信息：" + ex.Message);
                                            continue;
                                        }
                                        finally
                                        {
                                            if (conn.State == ConnectionState.Open)
                                                conn.Close();
                                        }
                                    }

                                    #endregion
                                }

                                #endregion
                            }

                            if (!isStart)
                            {
                                AppendMessageLine("任务停止！");
                                LogService.Write("任务停止！");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {

                            AppendMessageLine("对比歌曲的时长错误，ID：" + item["ItemID"].ToString() + " 错误信息：" + ex.Message);
                            LogService.WriteErr("导出任务失败 节目ID：" + item["ItemID"].ToString() + " 错误信息：" + ex.Message);
                            continue;
                        }
                        finally
                        {
                            if (File.Exists(sourceFile))
                            {
                                File.Delete(sourceFile);
                            }
                        }

                    }
                }


            }
            catch (Exception ex)
            {
                AppendMessageLine("程序错误，方法：Task 错误信息：" + ex.Message);
                LogService.WriteErr("程序错误，方法：Task 错误信息：" + ex.Message);
            }
            finally
            {
                AppendMessageLine("共有：" + countError + "条数据时长不匹配");
                LogService.Write("共有：" + countError + "条数据时长不匹配");
                Globals.ExecuteStopDate = DateTime.Now;
                Common.FlushMemory();
            }
        }

        #endregion

        #region//窗体事件
        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (ExcThread != null && ExcThread.IsAlive)
            {
                ExcThread.Abort();
            }

        }
        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //判断是否已经最小化于托盘 
            if (WindowState == FormWindowState.Minimized)
            {
                //还原窗体显示 
                WindowState = FormWindowState.Normal;
                //激活窗体并给予它焦点 
                this.Activate();
                //任务栏区显示图标 
                this.ShowInTaskbar = true;
                //托盘区图标隐藏 
                notifyIcon.Visible = false;
            }
        }
        private void FrmMain_SizeChanged(object sender, EventArgs e)
        {
            //判断是否选择的是最小化按钮 
            if (WindowState == FormWindowState.Minimized)
            {
                //托盘显示图标等于托盘图标对象 
                //注意notifyIcon1是控件的名字而不是对象的名字 

                //隐藏任务栏区图标 
                this.ShowInTaskbar = false;
                //图标显示在托盘区 
                notifyIcon.Visible = true;
            }
        }

        private void btnShowForm_Click(object sender, EventArgs e)
        {
            notifyIcon_MouseDoubleClick(sender, null);
        }
        private void btnStart_Click(object sender, EventArgs e)
        {

            if (ExcThread == null)
            {
                ExcThread = new Thread(this.TaskFunc);
                ExcThread.Start();
            }

            IsStart = !IsStart;


        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            Test();
            if (IsStart)
                MessageBox.Show("程序运行中，无法设置！", "提示！", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void btnCloseForm_Click(object sender, EventArgs e)
        {
            if (IsStart)
            {
                notifyIcon_MouseDoubleClick(sender, null);
                MessageBox.Show("程序运行中，请先停止！", "提示！", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                Close();
            }
        }
        #endregion

        private void FrmMain_Load(object sender, EventArgs e)
        {
            Init();
        }

        private void Test()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.Air5ConnectionStr))
                {
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.CommandTimeout = 2000;
                        command.Connection = conn;
                        conn.Open();
                        string ID = GetDictIDByName(command, "12345678-1234-1234-0002-123456789010", 66, "节目组9").ToString();
                        this.txtInfo.AppendText("ID:" + ID);
                    }
                    catch (Exception ex)
                    {

                        this.txtInfo.AppendText(ex.Message);
                    }
                    finally
                    { conn.Close(); }

                }
            }
            catch (Exception ex)
            {
                this.txtInfo.AppendText(ex.Message);

            }
        }

        protected Guid GetDictIDByName(SqlCommand command, string workgroupID, int kind, string name)
        {
            command.CommandText = string.Format("select DictID from ItemDict where WorkgroupID='{0}' and Name='{1}' and kind='{2}'", workgroupID, name.Trim(), kind);
            command.Parameters.Clear();
            command.CommandType = CommandType.Text;
            object result = command.ExecuteScalar();
            if (result != null)
            {
                return new Guid(result.ToString());
            }
            else
            {
                string  code="";
                if (kind == 121 || kind == 122 || kind == 129)
                    code="";
                else
                {
                    code = GetItemDictCode(command, kind);
                }
                string sqlInsert = "if not exists (select DictID from ItemDict where WorkgroupID=@WorkgroupID and Kind=@Kind and Name=@Name)" +
                              " insert into ItemDict(DictID,WorkgroupID,Kind,Name,Code) values (@DictID,@WorkgroupID,@Kind,@Name,@Code)";
                command.CommandText = sqlInsert;
                command.CommandType = CommandType.Text;
                command.Parameters.Clear();
                Guid dictID = Guid.NewGuid();
                command.Parameters.AddWithValue("@DictID", dictID);
                command.Parameters.AddWithValue("@WorkgroupID", workgroupID);
                command.Parameters.AddWithValue("@Kind", kind);
                command.Parameters.AddWithValue("@Name", name);
                command.Parameters.AddWithValue("@Code", code);

               int i= command.ExecuteNonQuery();
               if (i > 0)
               {
                   return dictID;
               }
               else
               {
                   return new Guid(result.ToString());
                  
               }
            }
        }
        public string GetItemDictCode(SqlCommand command, int kind)
        {
            string value = "";

            string str = string.Format("select MAX(CAST(Code as int)+1) as Code from ItemDict where Kind='{0}' and Code not like '%[^0-9]%'", kind);
            command.CommandText = str;
            command.CommandType = CommandType.Text;
            //command.Parameters.Clear();
            object result = command.ExecuteScalar();
            if (result == null || result == DBNull.Value)
            {
                return "01";
            }

            if (result.ToString().Length == 1)
            {
                value = "0" + result;
            }
            return value;
        }


    }
}
