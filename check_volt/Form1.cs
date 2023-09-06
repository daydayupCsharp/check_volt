using LumenWorks.Framework.IO.Csv;
using MySql.Data;
using MySql.Data.MySqlClient;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace check_volt
{
    public partial class Form1 : Form
    {
        //SerialNumber,UpdateOn
        DataTable dt1 = new DataTable();
        string addUser = "";
        string SerialNumber = "";
        string UpdateOn = "";
        string bodymodel = "";
        string gradeCode = "";
        string capacity = "";
        string remarks = "";
        string addtime = "";
        string connstr = "Database=mes;Data Source=10.52.67.28;port=3306;User Id=wwb;Password=123@mes";
        double row;
        public Form1()
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
        }
        private IWorkbook workbook = null;

        public void txt_write()
        {

            string writeBody = "";
            writeBody = DateTime.Now.ToString() + "  " + textBox1.Text + "   " + richTextBox3.Text + "-----" + richTextBox2.Text + "-----";
            var pFileName = DateTime.Now.ToString("yyyy-MM-dd"); //文件名
            string pSaveLogPath = @"" + Directory.GetCurrentDirectory().Substring(0, 3) + "";//路径
            string pSaveLogName = pSaveLogPath + "\\" + pFileName + "Log.txt";//完整LOG文件路径
            //判断当日日志文件是否存在
            if (File.Exists(pSaveLogName))
            {
                //已存在，从文件结束处添加新的日志
                FileStream fileStream = new FileStream(pSaveLogName, FileMode.Append, FileAccess.Write);
                StreamWriter LogWriter = new StreamWriter(fileStream);


                LogWriter.WriteLine(writeBody);//文件写入

                LogWriter.Flush(); //清除缓存
                LogWriter.Close(); //关闭文件流

            }
            else
            {
                //不存在，创建新的TXT文件

                //FileAccess.ReadWrite  对文件的读写访问权限。数据可以写入和读取文件。
                //FileMode.OpenOrCreate 指定操作系统应打开文件（如果存在）；否则应创建一个新文件。
                //FileStream errorLogFile = new System.IO.FileStream(pSaveLogName, FileMode.OpenOrCreate, FileAccess.ReadWrite);//创建文件流
                FileStream errorLogFile = new System.IO.FileStream(pSaveLogName, FileMode.OpenOrCreate, FileAccess.ReadWrite);//创建文件流

                StreamWriter LogWriter = new StreamWriter(errorLogFile);

                LogWriter.WriteLine(writeBody);//文件写入

                LogWriter.Flush(); //清除缓存
                LogWriter.Close(); //关闭文件流

            }
            writeBody = "";

        }

        /// <summary>
        /// 读取CSV文件
        /// </summary>
        /// <param name="fileName">文件路径</param>
        public static DataTable ReadCSV(string fileName)
        {
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs, Encoding.UTF8);

            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] arrayLine = null;
            //分隔符
            string[] separators = { "," };
            //判断，若是第一次，建立表头
            bool isFirst = true;

            //逐行读取CSV文件
            while ((strLine = sr.ReadLine()) != "")
            {
                strLine = strLine.Trim();//去除头尾空格
                arrayLine = strLine.Split(separators, StringSplitOptions.RemoveEmptyEntries);//分隔字符串，返回数组
                int dtColumns = arrayLine.Length;//列的个数

                if (isFirst)  //建立表头
                {
                    for (int i = 0; i < dtColumns; i++)
                    {
                        dt.Columns.Add(arrayLine[i]);//每一列名称
                    }
                }
                else   //表内容
                {
                    DataRow dataRow = dt.NewRow();//新建一行
                    for (int j = 0; j < dtColumns; j++)
                    {
                        dataRow[j] = arrayLine[j];
                    }
                    dt.Rows.Add(dataRow);//添加一行
                }
            }
            sr.Close();
            fs.Close();

            return dt;
        }
        public DataTable CsvToTable(Stream stream)
        {
            return CsvToTable(stream, 1);
        }

        public DataTable CsvToTable(Stream stream, int titleCount)
        {
            //如果遇到中文乱码情况可以设置一下编码字符集
            //Encoding _encode = Encoding.GetEncoding("GB2312");
            using (stream)
            {
                using (StreamReader input = new StreamReader(stream))
                {
                    using (CsvReader csv = new CsvReader(input, false))
                    {
                        DataTable dt = new DataTable();
                        int columnCount = csv.FieldCount;
                        for (int i = 0; i < columnCount; i++)
                        {
                            dt.Columns.Add("col" + i.ToString());
                        }
                        for (int i = 0; i < titleCount; i++)
                        {
                            csv.ReadNextRecord();
                        }
                        while (csv.ReadNextRecord())
                        {
                            DataRow dr = dt.NewRow();
                            for (int i = 0; i < columnCount; i++)
                            {
                                if (!string.IsNullOrWhiteSpace(csv[i]))
                                {
                                    dr[i] = csv[i];
                                }
                            }
                            dt.Rows.Add(dr);
                        }
                        csv.Dispose();
                        input.Close();
                        return dt;
                    }

                }

            }
        }
        public void excel_check(IWorkbook workbook, RichTextBox richTextBox1, string sheet_name)
        {
            var sheet = workbook.GetSheet(sheet_name/*ofd.SafeFileName.Substring(0, 31)*/);
            if (sheet != null)
            {
                for (int i = 1; i < sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row != null)
                    {
                        for (int j = 3; j < row.LastCellNum; j++)
                        {
                            var cell = row.GetCell(j);
                            if (cell != null)
                            {

                                //Console.WriteLine("{0} {1} {2}", cell.ToString(),
                                //cell.NumericCellValue, cell.StringCellValue);
                                if (cell.NumericCellValue >= 3300)
                                {

                                    richTextBox1.SelectionColor = Color.Red;
                                    richTextBox1.AppendText(i + 1 + "行--" + (j + 1) + "列--" + cell.NumericCellValue + "--电压值大于等于3300\r\n");
                                }
                            }
                            else
                            {
                                richTextBox1.SelectionColor = Color.Red;
                                richTextBox1.AppendText("单元格数据为空，请检查数据表\r\n");
                            }
                        }
                    }
                    else
                    {
                        richTextBox1.SelectionColor = Color.Red;
                        richTextBox1.AppendText("行数据为空，请检查数据表\r\n");
                    }
                }
            }
            else
            {
                richTextBox1.SelectionColor = Color.Red;
                richTextBox1.AppendText("数据表数据为空，请检查数据表\r\n");
            }
        }
        public static int MaxIndex(List<double> row)
        {
            int i_Pos = 0;
            int j_Pos = 0;
            double value_max = row[0];
            double value_max_2 = row[0];
            for (int i = 1; i < row.Count; ++i)
            {

                if (row[i] > value_max && row[i] > 0)
                {
                    value_max = row[i];
                    i_Pos = i;
                }
            }
            for (int j = 1; j < row.Count; ++j)
            {

                if (row[j] >= value_max_2 && row[j] > 0)
                {
                    value_max_2 = row[j];
                    j_Pos = j;
                }
            }
            if (i_Pos == j_Pos)
            { return i_Pos; }
            else { return 99999; }

        }
        public static int MinIndex(List<double> row)
        {
            int m_Pos = 0;
            int n_Pos = 0;
            double value_min = row[0];
            double value_min_2 = row[0];
            for (int m = 1; m < row.Count; ++m)
            {

                if (row[m] < value_min && row[m] > 0)
                {
                    value_min = row[m];
                    m_Pos = m;
                }
            }
            for (int n = 1; n < row.Count; ++n)
            {

                if (row[n] <= value_min_2 && row[n] > 0)
                {
                    value_min_2 = row[n];
                    n_Pos = n;
                }
            }
            if (m_Pos == n_Pos)
            { return m_Pos; }
            else { return 99999; }
        }

        //        private void button1_Click(object sender, EventArgs e)
        //        {
        //            richTextBox1.Clear();
        //            richTextBox2.Clear();
        //            richTextBox3.Clear();
        //            double time = Convert.ToDouble(DateTime.Now.ToString("yyyyMMddHHmmss"));
        //            string path = "";
        //            MySqlConnection conn = new MySqlConnection(connstr);
        //            conn.Open();
        //            MySqlDataReader test2 = new MySqlCommand(@"select a.gradeCode,min(a.capacity) from tdcbasecellinfo a 
        //left join tdcproducemoduleassembleitem b on a.cellLabelCode =b.cellLabelCode 
        //left join tdcproducepackassembleitem c on b.moduleLabelCode =c.moduleLabelCode 
        //where c.packLabelCode ='" + textBox1.Text + "'", conn).ExecuteReader();
        //            while (test2.Read())
        //            {
        //                gradeCode = test2[0].ToString();
        //                capacity = test2[1].ToString();
        //            }
        //            test2.Close();
        //            richTextBox3.SelectionColor = Color.Blue;
        //            richTextBox3.AppendText("包体码：" + textBox1.Text + "\r\n");
        //            richTextBox3.AppendText("档位：" + gradeCode + "      包内电芯最小容量：" + capacity + "\r\n");
        //            gradeCode = null;
        //            capacity = null;
        //            conn.Close();

        //            string csv_path = @"C:\Users\wang.hui110\Desktop\新建文件夹\P_" + textBox1.Text.Substring(5, 5) + "_JINGTAI1_" + textBox1.Text + "_" + time + "_Bv.csv";
        //            for (int i = 0; i <= 40000; i++)
        //            {
        //                if (!File.Exists(csv_path))
        //                {
        //                    time--;
        //                    csv_path = @"C:\Users\wang.hui110\Desktop\新建文件夹\P_" + textBox1.Text.Substring(5, 5) + "_JINGTAI1_" + textBox1.Text + "_" + time + "_Bv.csv";
        //                }
        //                else
        //                {
        //                    path = @"C:\Users\wang.hui110\Desktop\新建文件夹\P_" + textBox1.Text.Substring(5, 5) + "_JINGTAI1_" + textBox1.Text + "_" + time + "_Bv.csv";
        //                    break;
        //                }
        //            }

        //            //OpenFileDialog ofd = new OpenFileDialog();
        //            //ofd.Title = "请选择需要校验的文件";
        //            //ofd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        //            //ofd.Filter = "(*.csv)|*.csv|(*.xlsx)|*.xlsx|(*.xls)|*.xls|(*.*)|*.*";
        //            //ofd.ShowDialog();

        //            //var path = csv_path; /*ofd.FileName;*/
        //            string ofd = path.Substring(0, 66);
        //            if (File.Exists(path))
        //            {
        //                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        //                {

        //                    if (path.IndexOf(".xlsx") > 0) // 2007版本
        //                    {
        //                        workbook = new XSSFWorkbook(fs);
        //                        fs.Close();
        //                        excel_check(workbook, richTextBox1, ofd);
        //                    }
        //                    else if (path.IndexOf(".xls") > 0) // 2003版本
        //                    {
        //                        workbook = new HSSFWorkbook(fs);
        //                        fs.Close();
        //                        excel_check(workbook, richTextBox1, ofd);
        //                    }
        //                    else if (path.IndexOf(".csv") > 0) // 2003版本
        //                    {


        //                        DataTable dt = CsvToTable(fs);
        //                        fs.Close();

        //                        richTextBox1.AppendText("数据表共" + dt.Rows.Count + "行--" + dt.Columns.Count + "列--\r\n");
        //                        bool b = false;
        //                        bool b3 = true;
        //                        bool b2 = true;
        //                        Barrier barrier = new Barrier(2, it =>
        //                        {
        //                            if (b2 == false | b3 == false)
        //                            {
        //                                richTextBox3.SelectionColor = Color.Red;
        //                                richTextBox3.AppendText("行的电压最大值和最小值是否相邻判定：" + b2 + "----列的电压数值结果判定：" + b3 + "整包判定不合格\r\n");
        //                            }
        //                            else
        //                            {
        //                                //richTextBox3.SelectionColor = Color.Green;
        //                                //richTextBox3.AppendText("--此包电压测试合格\r\n");
        //                            }
        //                            txt_write();
        //                        });
        //                        Task.Run(() =>
        //                        {
        //                            for (int j = 3; j < dt.Columns.Count - 1; j++)
        //                            {
        //                                b = false;
        //                                for (int i = 0; i <= dt.Rows.Count - 1; i++)
        //                                {
        //                                    if (Convert.ToInt32(dt.Rows[i][j].ToString()) != 3300 && Convert.ToInt32(dt.Rows[i][j].ToString()) > 0)
        //                                    {
        //                                        //richTextBox1.SelectionColor = Color.Red;
        //                                        //richTextBox1.AppendText((i + 1) + "行--" + (j + 1) + "列--" + dt.Rows[i][j].ToString() + "--电压值大于等于3300\r\n");
        //                                        b = true;
        //                                    };
        //                                }
        //                                if (b == false)
        //                                {
        //                                    richTextBox1.SelectionColor = Color.Red;
        //                                    richTextBox1.AppendText((j + 1) + "列--" + "--测试电压数值均等于3300\r\n");
        //                                    b3 = false;
        //                                }
        //                                else
        //                                {
        //                                    //richTextBox1.SelectionColor = Color.Green;
        //                                    //richTextBox1.AppendText((j + 1) + "列--" + "--测试电压数值结果合格\r\n");
        //                                }
        //                            }
        //                            barrier.SignalAndWait();
        //                        }

        //              );
        //                        List<int> row = new List<int>();
        //                        int max;
        //                        int min;

        //                        Task.Run(() =>
        //                        {
        //                            for (int i = 0; i <= dt.Rows.Count - 1; i++)
        //                            {

        //                                for (int j = 3; j < dt.Columns.Count - 1; j++)
        //                                {
        //                                    row.Add(Convert.ToInt32(dt.Rows[i][j].ToString()));
        //                                }
        //                                max = MaxIndex(row);
        //                                min = MinIndex(row);
        //                                if (max - min == 1 | min - max == 1)
        //                                {
        //                                    richTextBox2.SelectionColor = Color.Red;
        //                                    richTextBox2.AppendText(i + 1 + "行--" + "--测试电压最大值和最小值相邻\r\n");
        //                                    b2 = false;
        //                                }
        //                                else
        //                                {
        //                                    //richTextBox2.SelectionColor = Color.Green;
        //                                    //richTextBox2.AppendText(i + 1 + "行--" + "--测试电压最大值和最小值不相邻\r\n");
        //                                }
        //                                row.Clear();
        //                            }
        //                            barrier.SignalAndWait();
        //                        }

        //                                    );

        //                    }
        //                }
        //            }
        //            else
        //            {
        //                richTextBox1.SelectionColor = Color.Red;
        //                richTextBox1.AppendText("无文件，请检查文件\r\n");
        //            }
        //        }
        private void textpack_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {

                richTextBox1.Clear();
                richTextBox2.Clear();
                richTextBox3.Clear();
                DateTime overtrue = DateTime.Now;
                double time = Convert.ToDouble(overtrue.ToString("yyyyMMddHHmmss"));
                string m_time = overtrue.ToString("yyyy-MM-dd-HH-mm-ss");
                string path = "";
                MySqlConnection conn = new MySqlConnection(connstr);
                MySqlConnection conn2 = new MySqlConnection(connstr);
                MySqlConnection conn3 = new MySqlConnection(connstr);




                Barrier barrier_allpro = new Barrier(4, it =>
                {
                    conn.Close();
                    conn2.Close();
                    conn3.Close();
                    txt_write();
                    textBox1.Clear();
                    GC.Collect();
                });


                Task.Run(() =>
                {
                    conn.Open();
                    if (checkBox1.Checked == true)
                    {
                        MySqlCommand msc2 = new MySqlCommand(@"select a.gradeCode,min(a.capacity) from tdcbasecellinfo a 
left join tdcproducemoduleassembleitem b on a.cellLabelCode =b.cellLabelCode 
left join tdcproducepackassembleitem c on b.moduleLabelCode =c.moduleLabelCode 
where c.packLabelCode ='" + textBox1.Text + "'", conn);
                        MySqlDataReader test2 = msc2.ExecuteReader();
                        try
                        {
                            while (test2.Read())
                            {
                                gradeCode = test2[0].ToString();
                                capacity = test2[1].ToString();
                            }
                        }
                        catch
                        {
                            richTextBox3.SelectionColor = Color.Red;
                            richTextBox3.AppendText("无包体码信息\r\n");
                        }
                        test2.Close();
                        msc2.Dispose();
                        richTextBox3.SelectionColor = Color.Blue;
                        richTextBox3.AppendText("包体码：" + textBox1.Text + "\r\n档位：" + gradeCode + "      包内电芯最小容量：" + capacity + "\r\n");
                        // richTextBox3.AppendText("档位：" + gradeCode + "      包内电芯最小容量：" + capacity + "\r\n");
                        gradeCode = "";
                        capacity = "";
                    }
                    else if (checkBox5.Checked == true)
                    {
                        MySqlCommand msc2 = new MySqlCommand(@"select a.gradeCode,min(a.capacity) from tdcbasecellinfo a 
left join tdcproducemoduleassembleitem b on a.cellLabelCode =b.cellLabelCode 
where b.moduleLabelCode ='" + textBox1.Text + "'", conn);
                        MySqlDataReader test2 = msc2.ExecuteReader();
                        try
                        {
                            while (test2.Read())
                            {
                                gradeCode = test2[0].ToString();
                                capacity = test2[1].ToString();
                            }
                        }
                        catch
                        {
                            richTextBox3.SelectionColor = Color.Red;
                            richTextBox3.AppendText("无模组码信息\r\n");
                        }
                        test2.Close();
                        msc2.Dispose();
                        richTextBox3.SelectionColor = Color.Blue;
                        richTextBox3.AppendText("模组码：" + textBox1.Text + "\r\n档位：" + gradeCode + "      模组内电芯最小容量：" + capacity + "\r\n");
                        // richTextBox3.AppendText("档位：" + gradeCode + "      包内电芯最小容量：" + capacity + "\r\n");
                        gradeCode = "";
                        capacity = "";
                    }
                    barrier_allpro.SignalAndWait();
                });

                Task.Run(() =>
                {
                    conn2.Open();
                    if (checkBox2.Checked == true)
                    {
                        MySqlCommand msc4 = new MySqlCommand(@"select oldLabelCode  from tdcpack_module_cellinfo where packLabelCode ='" + textBox1.Text + "'", conn2);
                        MySqlDataReader test4 = msc4.ExecuteReader();
                        while (test4.Read())
                        {
                            bodymodel += "'" + test4[0].ToString() + "',";//listModel.Add(test4[0].ToString());
                        }
                        test4.Close();
                        msc4.Dispose();
                        MySqlCommand msc5 = new MySqlCommand(@"select SerialNumber,UpdateOn from tdc_qm_defect
where SerialNumber in (" + bodymodel.Remove(bodymodel.Length - 1, 1) + @")
and DefectStatus='0'
and ActionID='3'", conn);
                        MySqlDataReader test5 = msc5.ExecuteReader();
                        while (test5.Read())
                        {
                            richTextBox3.SelectionColor = Color.Red;
                            richTextBox3.AppendText("此包的隔离电芯：" + test5[0].ToString() + "--隔离时间：" + test5[1].ToString() + "\r\n");
                            SerialNumber += test5[0].ToString();
                            UpdateOn += test5[1].ToString();

                        }
                        if (SerialNumber == "" | UpdateOn == "")
                        {
                            richTextBox3.SelectionColor = Color.Green;
                            richTextBox3.AppendText("此包内电芯未隔离\r\n");
                        }
                        bodymodel = "";
                        SerialNumber = "";
                        UpdateOn = "";
                        test5.Close();
                        msc5.Dispose();

                    }
                    else if (checkBox6.Checked == true)
                    {
                        MySqlCommand msc4 = new MySqlCommand(@"select oldLabelCode  from tdcpack_module_cellinfo where moduleLabelCode ='" + textBox1.Text + "'", conn2);
                        MySqlDataReader test4 = msc4.ExecuteReader();
                        while (test4.Read())
                        {
                            bodymodel += "'" + test4[0].ToString() + "',";//listModel.Add(test4[0].ToString());
                        }
                        test4.Close();
                        msc4.Dispose();
                        MySqlCommand msc5 = new MySqlCommand(@"select SerialNumber,UpdateOn from tdc_qm_defect
where SerialNumber in (" + bodymodel.Remove(bodymodel.Length - 1, 1) + @")
and DefectStatus='0'
and ActionID='3'", conn);
                        MySqlDataReader test5 = msc5.ExecuteReader();
                        while (test5.Read())
                        {
                            richTextBox3.SelectionColor = Color.Red;
                            richTextBox3.AppendText("此模组的隔离电芯：" + test5[0].ToString() + "--隔离时间：" + test5[1].ToString() + "\r\n");
                            SerialNumber += test5[0].ToString();
                            UpdateOn += test5[1].ToString();

                        }
                        if (SerialNumber == "" | UpdateOn == "")
                        {
                            richTextBox3.SelectionColor = Color.Green;
                            richTextBox3.AppendText("此模组内电芯未隔离\r\n");
                        }
                        bodymodel = "";
                        SerialNumber = "";
                        UpdateOn = "";
                        test5.Close();
                        msc5.Dispose();
                    }
                    barrier_allpro.SignalAndWait();
                });


                Task.Run(() =>
                {
                    conn3.Open();
                    if (checkBox3.Checked == true)
                    {

                        //                        MySqlDataReader test3 = new MySqlCommand(@"select  remarks,addUser from tdcqualitypackbadentryresult
                        //where badLabelCode='" + textBox1.Text+"'", conn3).ExecuteReader();

                        //                        while(test3.Read())
                        //                        {
                        //                            richTextBox3.SelectionColor = Color.Red;
                        //                            richTextBox3.AppendText("此包被隔离。隔离原因：" + test3[0].ToString() + "--隔离人工号：" + test3[1].ToString() + "\r\n");
                        //                            remarks += test3[0].ToString();
                        //                            addUser += test3[1].ToString();
                        //                        }
                        //                        if (remarks == "" | addUser == "")
                        //                        {
                        //                            richTextBox3.SelectionColor = Color.Green;
                        //                            richTextBox3.AppendText("此包未隔离\r\n");
                        //                        }

                        //                        test3.Close();
                        //                        addUser = "";
                        //                        remarks = "";


                        MySqlCommand msc = new MySqlCommand(@"select  remarks,addUser from tdcqualitypackbadentryresult
where badLabelCode='" + textBox1.Text + "'", conn3);
                        MySqlDataAdapter msd = new MySqlDataAdapter(msc);
                        msd.Fill(dt1);
                        if (dt1 == null || dt1.Rows.Count == 0)
                        {
                            richTextBox3.SelectionColor = Color.Green;
                            richTextBox3.AppendText("此包未隔离\r\n");
                        }
                        else
                        {
                            richTextBox3.SelectionColor = Color.Red;
                            richTextBox3.AppendText("此包被隔离。隔离原因：" + dt1.Rows[0][0] + "--隔离人工号：" + dt1.Rows[0][1] + "\r\n");
                        }
                        dt1.Clear();
                        msc.Dispose();
                        msd.Dispose();
                    }
                    barrier_allpro.SignalAndWait();
                });


                Task.Run(() =>
                {

                    if (checkBox4.Checked == true)
                    {

                        string csv_path = textBox2.Text + @"\P_" + textBox1.Text.Substring(5, 5) + "_JINGTAI1_" + textBox1.Text + "_" + time + "_Bv.csv";
                        for (int i = 0; i <= 40000; i++)
                        {
                            if (!File.Exists(csv_path))
                            {
                                time--;
                                csv_path = textBox2.Text + @"\P_" + textBox1.Text.Substring(5, 5) + "_JINGTAI1_" + textBox1.Text + "_" + time + "_Bv.csv";
                            }
                            else
                            {
                                path = textBox2.Text + @"\P_" + textBox1.Text.Substring(5, 5) + "_JINGTAI1_" + textBox1.Text + "_" + time + "_Bv.csv";
                                break;
                            }
                        }

                        //OpenFileDialog ofd = new OpenFileDialog();
                        //ofd.Title = "请选择需要校验的文件";
                        //ofd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        //ofd.Filter = "(*.csv)|*.csv|(*.xlsx)|*.xlsx|(*.xls)|*.xls|(*.*)|*.*";
                        //ofd.ShowDialog();

                        //var path = csv_path; /*ofd.FileName;*/
                        string ofd = path.Substring(0, 66);
                        if (File.Exists(path))
                        {
                            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                            {

                                if (path.IndexOf(".xlsx") > 0) // 2007版本
                                {
                                    workbook = new XSSFWorkbook(fs);
                                    fs.Close();
                                    excel_check(workbook, richTextBox1, ofd);
                                }
                                else if (path.IndexOf(".xls") > 0) // 2003版本
                                {
                                    workbook = new HSSFWorkbook(fs);
                                    fs.Close();
                                    excel_check(workbook, richTextBox1, ofd);
                                }
                                else if (path.IndexOf(".csv") > 0) // 2003版本
                                {


                                    DataTable dt = CsvToTable(fs);
                                    fs.Close();

                                    richTextBox3.AppendText("数据表共" + dt.Rows.Count + "行--" + dt.Columns.Count + "列--\r\n");
                                    bool b = false;
                                    bool b3 = true;
                                    bool b2 = true;
                                    Barrier barrier = new Barrier(2, it =>
                                    {
                                        if (b2 == false | b3 == false)
                                        {
                                            richTextBox3.SelectionColor = Color.Red;
                                            richTextBox3.AppendText("行的电压最大值和最小值是否相邻判定：" + b2 + "----列的电压数值结果判定：" + b3 + "整包判定不合格\r\n");
                                        }
                                        else
                                        {
                                            richTextBox3.SelectionColor = Color.Green;
                                            richTextBox3.AppendText("--此包电压测试合格\r\n");
                                        }

                                    });
                                    Task.Run(() =>
                                    {
                                        for (int j = 3; j < dt.Columns.Count - 1; j++)
                                        {
                                            b = false;
                                            for (int i = 0; i <= dt.Rows.Count - 1; i++)
                                            {
                                                if (Convert.ToInt32(dt.Rows[i][j].ToString()) != 3300 && Convert.ToInt32(dt.Rows[i][j].ToString()) > 0)
                                                {
                                                    //richTextBox1.SelectionColor = Color.Red;
                                                    //richTextBox1.AppendText((i + 1) + "行--" + (j + 1) + "列--" + dt.Rows[i][j].ToString() + "--电压值大于等于3300\r\n");
                                                    b = true;
                                                };
                                            }
                                            if (b == false)
                                            {
                                                richTextBox1.SelectionColor = Color.Red;
                                                richTextBox1.AppendText((j + 1) + "列--" + "--测试电压数值均等于3300\r\n");
                                                b3 = false;
                                            }
                                            else
                                            {
                                                //richTextBox1.SelectionColor = Color.Green;
                                                //richTextBox1.AppendText((j + 1) + "列--" + "--测试电压数值结果合格\r\n");
                                            }
                                        }
                                        barrier.SignalAndWait();
                                    }

                          );
                                    List<double> row = new List<double>();
                                    int max;
                                    int min;

                                    Task.Run(() =>
                                    {
                                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                                        {

                                            for (int j = 3; j < dt.Columns.Count - 1; j++)
                                            {
                                                row.Add(Convert.ToDouble(dt.Rows[i][j].ToString()));
                                            }
                                            if (textBox1.Text.Substring(5, 3) == "PE4")
                                            {
                                                row.RemoveAt(30);
                                                row.RemoveAt(95);
                                            }

                                            else if (textBox1.Text.Substring(5, 3) == "PE5" | textBox1.Text.Substring(5, 3) == "PE6")
                                            {
                                                row.RemoveAt(42);
                                                row.RemoveAt(125);
                                            }
                                            max = MaxIndex(row);
                                            min = MinIndex(row);
                                            if (max - min == 1 | min - max == 1)
                                            {
                                                richTextBox2.SelectionColor = Color.Red;
                                                richTextBox2.AppendText(i + 1 + "行--" + "--测试电压最大值和最小值相邻\r\n");
                                                b2 = false;
                                            }
                                            else
                                            {
                                                //richTextBox2.SelectionColor = Color.Green;
                                                //richTextBox2.AppendText(i + 1 + "行--" + "--测试电压最大值和最小值不相邻\r\n");
                                            }
                                            row.Clear();
                                        }
                                        barrier.SignalAndWait();
                                    });

                                }
                            }
                        }
                        else
                        {
                            richTextBox1.SelectionColor = Color.Red;
                            richTextBox1.AppendText("无文件，请检查文件\r\n");
                        }
                    }
                    else if (checkBox7.Checked == true)
                    {
                        if (textBox1.Text.Substring(3, 1) == "M")
                        {
                            string csv_path = textBox2.Text + @"\M_" + textBox1.Text.Substring(5, 3) + "_MZCY_" + textBox1.Text + "_" + m_time + "_Bv.csv";
                            for (int i = 0; i <= 40000; i++)
                            {
                                if (!File.Exists(csv_path))
                                {

                                    m_time = overtrue.AddSeconds(-i - 1).ToString("yyyy-MM-dd-HH-mm-ss");
                                    csv_path = textBox2.Text + @"\M_" + textBox1.Text.Substring(5, 3) + "_MZCY_" + textBox1.Text + "_" + m_time + "_Bv.csv";
                                }
                                else
                                {
                                    path = textBox2.Text + @"\M_" + textBox1.Text.Substring(5, 3) + "_MZCY_" + textBox1.Text + "_" + m_time + "_Bv.csv";
                                    break;
                                }
                            }
                        }
                        else if (textBox1.Text.Substring(3, 1) == "P")
                        {
                            string csv_path = textBox2.Text + @"\P_" + textBox1.Text.Substring(5, 3) + "_BTCY_" + textBox1.Text + "_" + m_time + "_Bv.csv";
                            for (int i = 0; i <= 40000; i++)
                            {
                                if (!File.Exists(csv_path))
                                {

                                    m_time = overtrue.AddSeconds(-i - 1).ToString("yyyy-MM-dd-HH-mm-ss");
                                    csv_path = textBox2.Text + @"\P_" + textBox1.Text.Substring(5, 3) + "_BTCY_" + textBox1.Text + "_" + m_time + "_Bv.csv";
                                }
                                else
                                {
                                    path = textBox2.Text + @"\P_" + textBox1.Text.Substring(5, 3) + "_BTCY_" + textBox1.Text + "_" + m_time + "_Bv.csv";
                                    break;
                                }
                            }
                        }
                        //OpenFileDialog ofd = new OpenFileDialog();
                        //ofd.Title = "请选择需要校验的文件";
                        //ofd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        //ofd.Filter = "(*.csv)|*.csv|(*.xlsx)|*.xlsx|(*.xls)|*.xls|(*.*)|*.*";
                        //ofd.ShowDialog();

                        //var path = csv_path; /*ofd.FileName;*/
                        string ofd = path.Substring(0, 66);
                        if (File.Exists(path))
                        {
                            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                            {

                                if (path.IndexOf(".xlsx") > 0) // 2007版本
                                {
                                    workbook = new XSSFWorkbook(fs);
                                    fs.Close();
                                    excel_check(workbook, richTextBox1, ofd);
                                }
                                else if (path.IndexOf(".xls") > 0) // 2003版本
                                {
                                    workbook = new HSSFWorkbook(fs);
                                    fs.Close();
                                    excel_check(workbook, richTextBox1, ofd);
                                }
                                else if (path.IndexOf(".csv") > 0) // 2003版本
                                {


                                    DataTable dt = CsvToTable(fs);
                                    fs.Close();
                                    List<double> doubles = new List<double>();
                                    for (int j = 1; j < dt.Columns.Count - 1; j++)
                                    {
                                        for (int i = 0; i <= dt.Rows.Count - 1; i++)
                                        {
                                            row = Convert.ToDouble(dt.Rows[i][j]);
                                            if (row != 0)
                                            {
                                                doubles.Add(row);
                                            }
                                        }
                                    }
                                    if (textBox1.Text.Substring(3, 1) == "M")
                                    {
                                        if (textBox1.Text.Substring(5, 5) == "PE4M1" | textBox1.Text.Substring(5, 5) == "PE4M3" | textBox1.Text.Substring(5, 5) == "PE4M5" | textBox1.Text.Substring(5, 5) == "PE4M7")
                                        {
                                            doubles.RemoveAt(32);
                                        }
                                        else if (textBox1.Text.Substring(5, 5) == "PE4M2" | textBox1.Text.Substring(5, 5) == "PE4M4" | textBox1.Text.Substring(5, 5) == "PE4M6" | textBox1.Text.Substring(5, 5) == "PE4M8")
                                        {
                                            doubles.RemoveAt(30);
                                        }
                                        else if (textBox1.Text.Substring(5, 3) == "PE5" | textBox1.Text.Substring(5, 3) == "PE6")
                                        {
                                            doubles.RemoveAt(42);
                                        }
                                    }
                                    else if (textBox1.Text.Substring(3, 1) == "P")
                                    {
                                        if (textBox1.Text.Substring(5, 3) == "PE4")
                                        {
                                            doubles.RemoveAt(30);
                                            doubles.RemoveAt(95);
                                        }

                                        else if (textBox1.Text.Substring(5, 3) == "PE5" | textBox1.Text.Substring(5, 3) == "PE6")
                                        {
                                            doubles.RemoveAt(42);
                                            doubles.RemoveAt(125);
                                        }
                                    }
                                    richTextBox3.AppendText("条码为：" + textBox1.Text + "数据表共" + dt.Rows.Count + "行--" + dt.Columns.Count + "列--\r\n");
                                    bool b = false;
                                    bool b3 = true;
                                    bool b2 = true;
                                    bool b4 = true;
                                    Barrier barrier = new Barrier(3, it =>
                                    {
                                        if (b2 == false | b3 == false | b4 == false)
                                        {
                                            richTextBox3.SelectionColor = Color.Red;
                                            richTextBox3.AppendText("行的电压最大值和最小值是否相邻判定：" + b2 + "----列的电压数值结果判定：" + b3 + "整包判定不合格\r\n");
                                        }
                                        else
                                        {
                                            richTextBox3.SelectionColor = Color.Green;
                                            richTextBox3.AppendText("--此包电压测试合格\r\n");
                                        }

                                    });
                                    Task.Run(() =>
                                    {
                                        for (int j = 1; j < dt.Columns.Count - 1; j++)
                                        {
                                            b = false;
                                            for (int i = 0; i <= dt.Rows.Count - 1; i++)
                                            {
                                                if (Convert.ToInt32(dt.Rows[i][j].ToString()) != 3300 /*&& Convert.ToInt32(dt.Rows[i][j].ToString()) > 0*/)
                                                {
                                                    //richTextBox1.SelectionColor = Color.Red;
                                                    //richTextBox1.AppendText((i + 1) + "行--" + (j + 1) + "列--" + dt.Rows[i][j].ToString() + "--电压值大于等于3300\r\n");
                                                    b = true;
                                                };
                                            }
                                            if (b == false)
                                            {
                                                richTextBox1.SelectionColor = Color.Red;
                                                richTextBox1.AppendText((j + 1) + "列--" + "--测试电压数值均等于3300\r\n");
                                                b3 = false;
                                            }
                                            else
                                            {
                                                //richTextBox1.SelectionColor = Color.Green;
                                                //richTextBox1.AppendText((j + 1) + "列--" + "--测试电压数值结果合格\r\n");
                                            }
                                        }
                                        barrier.SignalAndWait();
                                    }

                          );

                                    int max;
                                    int min;

                                    Task.Run(() =>
                                    {

                                        max = MaxIndex(doubles);
                                        min = MinIndex(doubles);
                                        if (max - min == 1 | min - max == 1)
                                        {
                                            richTextBox2.SelectionColor = Color.Red;
                                            richTextBox2.AppendText("--测试电压最大值和最小值相邻\r\n");
                                            b2 = false;
                                        }
                                        else
                                        {
                                            //richTextBox2.SelectionColor = Color.Green;
                                            //richTextBox2.AppendText(i + 1 + "行--" + "--测试电压最大值和最小值不相邻\r\n");
                                        }


                                        barrier.SignalAndWait();
                                    });
                                    Task.Run(() =>
                                    {
                                        double m_max = doubles.Max();
                                        double m_min = doubles.Min();
                                        double m_average = doubles.Average();
                                        double m_sqrt = Math.Sqrt(doubles.Sum(x => Math.Pow(x - m_average, 2)) / (doubles.Count() - 1));
                                        decimal result1 = Convert.ToDecimal(m_average) + Convert.ToDecimal(4.5) * Convert.ToDecimal(m_sqrt);
                                        decimal result2 = Convert.ToDecimal(m_average) - Convert.ToDecimal(4.5) * Convert.ToDecimal(m_sqrt);
                                        if (m_max > Convert.ToDouble(result1) | m_min < Convert.ToDouble(result2))
                                        {
                                            b = false;
                                            richTextBox3.SelectionColor = Color.Red;
                                            richTextBox3.AppendText("--电压最大值：" + m_max + "--电压最小值：" + m_min + "\r\n--电压平均值：" + m_average + "--电压标准差：" + m_sqrt + "此包判定NG\r\n");
                                        }
                                        else
                                        {
                                            richTextBox3.SelectionColor = Color.Green;
                                            richTextBox3.AppendText("--电压最大值：" + m_max + "--电压最小值：" + m_min + "\r\n--电压平均值：" + m_average + "--电压标准差：" + m_sqrt + "此包判定OK\r\n");

                                        }
                                        barrier.SignalAndWait();
                                    });

                                }
                            }
                        }
                        else
                        {
                            richTextBox1.SelectionColor = Color.Red;
                            richTextBox1.AppendText("无文件，请检查文件\r\n");
                        }
                    }
                    barrier_allpro.SignalAndWait();
                });
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Visible == false && textBox1.Text == "5029756")
            {
                textBox2.Visible = true;
                checkBox1.Visible = true;
                checkBox2.Visible = true;
                checkBox3.Visible = true;
                checkBox4.Visible = true;
                checkBox5.Visible = true;
                checkBox6.Visible = true;
                checkBox7.Visible = true;

            }
            else
            {
                textBox2.Visible = false;
                checkBox1.Visible = false;
                checkBox2.Visible = false;
                checkBox3.Visible = false;
                checkBox4.Visible = false;
                checkBox5.Visible = false;
                checkBox6.Visible = false;
                checkBox7.Visible = false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<CheckBox> list = new List<CheckBox>();
            list.Add(checkBox1);
            list.Add(checkBox2);
            list.Add(checkBox3);
            list.Add(checkBox4);
            list.Add(checkBox5);
            list.Add(checkBox6);
            list.Add(checkBox7);
            string[] config = System.IO.File.ReadAllLines(Directory.GetCurrentDirectory() + @"\config.txt");
            textBox2.Text = config[0];
            for (int i = 1; i < list.Count + 1; i++)
            {
                if (config[i].Contains("true")) { list[i - 1].Checked = true; }
                else if (config[i].Contains("false")) { list[i - 1].Checked = false; }
            }
        }
    }
}
