using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;

namespace XMLExportTool
{
    public partial class Form4 : Form
    {
        //必要变量定义
        private DbOperation dboperation;
        private string fileFolder = "C:\\";
        private string dateString = DateTime.Now.ToString("yyyy-MM-dd");
        private string serverIP;
        private string dbName;
        private string userName;
        private string password;
        private string filePath;
        private string currentSql;      //记录当前表格的sql语句
        private FileStream fs;
        private StreamWriter sw;
        private SqlDataReader reader;
        public static string chargeName = "钢铁研究总院";
        public static string chargeAddress = "北京市海淀区学院南路76号14信箱";
        public static string chargeCode = "100081";
        public static string chargePhone = "(010)62188863";
        public static string chargeEmail = "ceas@analysis.org.cn";
        public static string visitLimit = "1";
        public static string classifyName = "中国应急分析测试平台栏目分类标准";
        public static string classifyVersion = "3.0";
        public Form4(string serverIP, string dbName, string userName, string password)
        {
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
            dboperation = new DbOperation();
            this.serverIP = serverIP;
            this.dbName = dbName;
            this.userName = userName;
            this.password = password;
            dboperation.Connect(serverIP, dbName, userName, password);
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            reader = dboperation.GetDataReader("select ChannelID,ChannelName from PE_Channel where ItemCount !=0");
            DbOperation mydboperation = new DbOperation();
            mydboperation.Connect(serverIP, dbName, userName, password);
            SqlDataReader myreader;
            while (reader.Read())
            {
                checkedListBox1.Items.Add(reader.GetString(1));
                TreeNode node = new TreeNode();
                node.Text = reader.GetString(1);
                treeView1.Nodes.Add(node);
                myreader = mydboperation.GetDataReader("select ClassName from PE_Class where ChannelID =" + reader.GetInt32(0));
                while (myreader.Read())
                {
                    TreeNode childNode = new TreeNode();
                    childNode.Text = myreader.GetString(0);
                    node.Nodes.Add(childNode);
                }
                myreader.Close();
            }
            mydboperation.Close();
            //默认全部勾选
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
            reader.Close();
        }

        //实现查询
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                reader = dboperation.GetDataReader(currentSql+" and "+textBox1.Text);
                if (reader == null)
                {
                    MessageBox.Show("请指定正确的查询条件！");
                    return;
                }
                if (!reader.HasRows)
                {
                    MessageBox.Show("没有查询结果！");
                    reader.Close();
                    return;
                }
                BindingSource Bs = new BindingSource();
                Bs.DataSource = reader;
                dataGridView1.DataSource = Bs;
            }
            catch (Exception)
            {
            }
            reader.Close();
        }

        //"高级设置"事件
        private void button7_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
        }


        //导出功能，最核心的部分
        private void button1_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(exportXmlProcess));
            thread.Start();
        }

        //浏览文件夹功能实现
        private void button6_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderbrowserdialog = new FolderBrowserDialog();
            if (folderbrowserdialog.ShowDialog() == DialogResult.OK)
            {
                fileFolder = folderbrowserdialog.SelectedPath.ToString();
                label1.Text = fileFolder;
            }
        }

        //注销功能实现
        private void button5_Click(object sender, EventArgs e)
        {
            foreach (Form f in Application.OpenForms)
            {
                if (f.Name == "Form1")
                {
                    f.Show();
                    this.Dispose();
                    break;
                }
            }
        }

        //打开文件浏览器
        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Explorer.exe", fileFolder);
        }

        //重置功能实现
        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
        }

        //全选功能实现
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        //关闭Form4窗体操作
        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Form f in Application.OpenForms)
            {
                if (f.Name != "Form4")
                {
                    f.Dispose();
                }
            }
            dboperation.Close();
        }

        //导出线程函数
        private void exportXmlProcess()
        {
            //导出channel数据库表并写入xml文件
            progressBar1.Minimum = 0;
            progressBar1.Maximum = checkedListBox1.CheckedItems.Count + 2;
            filePath = fileFolder + "\\ceas_channel_" + dateString + ".xml";
            fs = new FileStream(filePath, FileMode.Create);
            sw = new StreamWriter(fs, Encoding.GetEncoding("GB2312"));
            sw.WriteLine("<?xml version=\"1.0\" encoding=\"gb2312\"?>\r\n<results>");
            reader = dboperation.GetDataReader("select ChannelID,ChannelName from PE_Channel where ItemCount !=0");
            while (reader.Read())
            {
                sw.WriteLine("<result>\r\n<资源库ID>{0}</资源库ID>\r\n<资源库名称>{1}</资源库名称>\r\n</result>", reader.GetInt32(0), reader.GetString(1));
            }
            sw.WriteLine("</results>");
            reader.Close();
            sw.Close();
            fs.Close();
            progressBar1.Value++;

            //导出class数据库表并写入xml文件
            filePath = fileFolder + "\\ceas_class_" + dateString + ".xml";
            fs = new FileStream(filePath, FileMode.Create);
            sw = new StreamWriter(fs, Encoding.GetEncoding("GB2312"));
            sw.WriteLine("<?xml version=\"1.0\" encoding=\"gb2312\"?>\r\n<results>");
            reader = dboperation.GetDataReader("select ClassID,ChannelID,ClassName,ParentPath from PE_Class");
            while (reader.Read())
            {
                string parentPath = reader.GetString(3);
                string categoryCode = "";
                if (parentPath != "0")
                {
                    categoryCode = parentPath.Substring(2) + ";";
                }
                sw.WriteLine("<result>\r\n<类目ID>{0}</类目ID>\r\n<资源库ID>{1}</资源库ID>\r\n<类目名称>{2}</类目名称>\r\n<类目编码>{3}</类目编码>\r\n</result>", reader.GetInt32(0), reader.GetInt32(1), reader.GetString(2), categoryCode + reader.GetInt32(0));
            }
            sw.WriteLine("</results>");
            reader.Close();
            sw.Close();
            fs.Close();
            progressBar1.Value++;

            //循环需要导出的文件
            DbOperation mydboperation = new DbOperation();
            mydboperation.Connect(serverIP, dbName, userName, password);
            SqlDataReader myreader;
            int ChannelID = 0;
            string ChannelDir = "";
            string ChannelShortName = "";
            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                myreader = mydboperation.GetDataReader("select ChannelID,ChannelDir,ChannelShortName from PE_Channel where ChannelName ='" + checkedListBox1.CheckedItems[i].ToString()+"'");
                if (myreader.Read())
                {
                    ChannelID = myreader.GetInt32(0);
                    ChannelDir = myreader.GetString(1);
                    ChannelShortName = myreader.GetString(2);
                }
                else
                {
                    continue;
                }
                myreader.Close();

                filePath = fileFolder + "\\ceas_" + checkedListBox1.CheckedItems[i].ToString() + "_" + dateString + ".xml";
                fs = new FileStream(filePath, FileMode.Create);
                sw = new StreamWriter(fs, Encoding.GetEncoding("GB2312"));
                sw.WriteLine("<?xml version=\"1.0\" encoding=\"gb2312\"?>\r\n<results>");
                exportXml("select ArticleID,Title,Keyword,Intro,UpdateTime,ClassID from PE_Article where ChannelID=" + ChannelID + " order by UpdateTime desc", ChannelDir,ChannelShortName, "Article");
                exportXml("select SoftID,SoftName,Keyword,SoftIntro,UpdateTime,ClassID from PE_Soft where ChannelID=" + ChannelID + " order by UpdateTime desc", ChannelDir, ChannelShortName, "Soft");
                exportXml("select PhotoID,PhotoName,Keyword,PhotoIntro,UpdateTime,ClassID from PE_Photo where ChannelID=" + ChannelID + " order by UpdateTime desc", ChannelDir, ChannelShortName, "Photo");
                sw.WriteLine("</results>");
                sw.Close();
                fs.Close();
                progressBar1.Value++;
            }
            mydboperation.Close();
            if (MessageBox.Show("导出完成！") == DialogResult.OK)
            {
                progressBar1.Value = 0;
            }
        }

        //导出每类的xml文件
        private bool exportXml(string sqlstr, string ChannelDir, string ChannelShortName, string tableName)
        {
            try
            {
                
                DbOperation mydboperation = new DbOperation();
                mydboperation.Connect(serverIP, dbName, userName, password);
                SqlDataReader myreader;
                myreader = dboperation.GetDataReader(sqlstr);
                while (myreader.Read())
                {
                    sw.WriteLine("<result>");
                    sw.WriteLine("<标识符>{0}</标识符>", myreader.GetInt32(0));
                    sw.WriteLine("<名称>{0}</名称>", myreader.GetString(1));
                    sw.WriteLine("<最近提交日期>{0}</最近提交日期>", myreader.GetDateTime(4));
                    sw.WriteLine("<描述>{0}</描述>", RemoveSpaceHtmlTag(myreader.GetString(3)));
                    sw.WriteLine("<负责单位名称>{0}</负责单位名称>", chargeName);
                    sw.WriteLine("<负责单位通讯地址>{0}</负责单位通讯地址>", chargeAddress);
                    sw.WriteLine("<负责单位邮政编码>{0}</负责单位邮政编码>", chargeCode);
                    sw.WriteLine("<负责单位联系电话>{0}</负责单位联系电话>", chargePhone);
                    sw.WriteLine("<负责单位电子邮件地址>{0}</负责单位电子邮件地址>", chargeEmail);
                    string keyword = myreader.GetString(2);
                    if (keyword.Length >= 2)
                    {
                        keyword = keyword.Substring(1, keyword.Length - 2).Replace('|', ';');
                    }
                    sw.WriteLine("<关键词>应急;{0};{1}</关键词>", keyword,ChannelShortName);
                    sw.WriteLine("<访问限制>{0}</访问限制>", visitLimit);
                    reader = mydboperation.GetDataReader("select ClassName,ParentPath from PE_Class where ClassID=" + myreader.GetInt32(5));
                    if (reader.Read())
                    {
                        string parentPath = reader.GetString(1);
                        string categoryCode = "";
                        if (parentPath != "0")
                        {
                            categoryCode = parentPath.Substring(2) + ",";
                        }
                        sw.WriteLine("<类目名称>{0}</类目名称>", reader.GetString(0));
                        sw.WriteLine("<类目编码>{0}</类目编码>", categoryCode + myreader.GetInt32(5));
                    }
                    reader.Close();
                    sw.WriteLine("<分类标准名称>{0}</分类标准名称>", classifyName);
                    sw.WriteLine("<分类标准版本号>{0}</分类标准版本号>", classifyVersion);
                    sw.WriteLine("<资源信息链接地址>http://www.ceas.org.cn/{0}/Show{1}.asp?{1}ID={2}</资源信息链接地址>", ChannelDir, tableName, myreader.GetInt32(0));
                    sw.WriteLine("</result>");
                }
                mydboperation.Close();
                myreader.Close();
            }
            catch (Exception)
            {
                if (reader != null)
                {
                    reader.Close();
                }
                return false;
            }
            return true;
        }

        //去除Html标签
        private string RemoveSpaceHtmlTag(string Input)
        {
            string input = Input;
            try
            {
                //去html标签
                input = Regex.Replace(input, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"-->", "", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"<!--.*", "", RegexOptions.IgnoreCase);

                input = Regex.Replace(input, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(nbsp|#160);", " ", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
                input = Regex.Replace(input, @"&#(\d+);", "", RegexOptions.IgnoreCase);

                input.Replace("<", "");
                input.Replace(">", "");
                input.Replace("\r\n", "");
                //去两端空格，中间多余空格
                input = Regex.Replace(input.Trim(), "\\s+", " ");
            }
            catch (Exception)
            {
            }
            return input;
        }

        //选中叶节点事件
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Level == 1)
            {
                int ClassID = 0;
                textBox1.Text = "";
                //查询PE_Class表
                reader = dboperation.GetDataReader("select ClassID from PE_Class where ClassName = '"+e.Node.Text+"'");
                if (reader.Read())
                {
                    ClassID = reader.GetInt32(0);
                }
                reader.Close();
                //查询PE_Article表
                currentSql = "select * from PE_Article where ClassID=" + ClassID;
                reader = dboperation.GetDataReader(currentSql);
                if (reader != null && reader.HasRows)
                {
                    BindingSource Bs = new BindingSource();
                    Bs.DataSource = reader;
                    dataGridView1.DataSource = Bs;
                    reader.Close();
                    return;
                }
                reader.Close();
                //查询PE_Photo表
                currentSql = "select * from PE_Photo where ClassID=" + ClassID;
                reader = dboperation.GetDataReader(currentSql);
                if (reader != null && reader.HasRows)
                {
                    BindingSource Bs = new BindingSource();
                    Bs.DataSource = reader;
                    dataGridView1.DataSource = Bs;
                    reader.Close();
                    return;
                }
                reader.Close();
                //查询PE_Soft表
                currentSql = "select * from PE_Soft where ClassID=" + ClassID;
                reader = dboperation.GetDataReader(currentSql);
                if (reader != null && reader.HasRows)
                {
                    BindingSource Bs = new BindingSource();
                    Bs.DataSource = reader;
                    dataGridView1.DataSource = Bs;
                    reader.Close();
                    return;
                }
                reader.Close();
            }
        }

        //"高级设置"事件
        private void button7_Click_1(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
        }

        //返回
        private void button9_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            reader = dboperation.GetDataReader(currentSql);
            if (reader == null)
            {
                MessageBox.Show("请指定正确的查询条件！");
                return;
            }
            if (!reader.HasRows)
            {
                MessageBox.Show("没有查询结果！");
                reader.Close();
                return;
            }
            try
            {
                BindingSource Bs = new BindingSource();
                Bs.DataSource = reader;
                dataGridView1.DataSource = Bs;
                reader.Close();
            }
            catch (Exception)
            {
            }
        }
    }
}
