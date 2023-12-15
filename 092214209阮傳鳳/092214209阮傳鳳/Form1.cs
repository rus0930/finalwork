using System;
using System.Windows.Forms;
using System.Data.SQLite;
using static _092214209阮傳鳳.Form1;
using static System.Windows.Forms.AxHost;
using System.Xml.Linq;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq.Expressions;
using Excel = Microsoft.Office.Interop.Excel;



namespace _092214209阮傳鳳
{
    public partial class Form1 : Form
    {
        // 紀錄清單編號
        int item_index = 0;
        int log_index = 0;

        // 資料庫設定
        public class DBConfig
        {
            //log.db要放在【bin\Debug底下】      
            public static string dbFile = Application.StartupPath + @"\myC#.db";

            public static string dbPath = "Data source=" + dbFile;

            public static SQLiteConnection sqlite_connect;
            public static SQLiteCommand sqlite_cmd;
            public static SQLiteDataReader sqlite_datareader;
        }

        // 載入資料庫
        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open
        }

        // 刷新商品選項清單
        private void Show_DB_item()
        {
            // 清空表格
            this.dataGridView4.Rows.Clear();
            item_comboBox.Items.Clear();

            // SQL查詢語法
            string sql = @"SELECT * from item;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();
                        
            if (DBConfig.sqlite_datareader.HasRows) // 若讀取結果有數據
            {
                while (DBConfig.sqlite_datareader.Read()) // 逐行讀取
                {
                    // 讀取資料庫欄位
                    string _itemNo = Convert.ToString(DBConfig.sqlite_datareader["item_no"]);
                    string _itemName = Convert.ToString(DBConfig.sqlite_datareader["item_name"]);

                    // 加入表格
                    dataGridView4.Rows.Add(
                        new Object[]
                        {
                            _itemNo, _itemName
                        }
                    );

                    item_comboBox.Items.Add("[" + _itemNo + "] " + _itemName); // 下拉式選單新增項目
                    item_index = int.Parse(_itemNo.Substring(2)); // 紀錄編號
                }
                DBConfig.sqlite_datareader.Close(); // 關閉資料庫讀取
            }

        }

        // 刷新紀錄清單
        private void Show_DB_log()
        {
            // 清空表格
            this.dataGridView1.Rows.Clear();
            this.dataGridView3.Rows.Clear();

            // SQL查詢語法
            string sql = @"SELECT * from log;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows) // 若讀取結果有數據
            {
                while (DBConfig.sqlite_datareader.Read()) // 逐行讀取
                {
                    // 讀取資料庫欄位
                    string _no = Convert.ToString(DBConfig.sqlite_datareader["no"]);
                    string _date = Convert.ToString(DBConfig.sqlite_datareader["date"]);
                    string _type = Convert.ToString(DBConfig.sqlite_datareader["type"]);
                    string _itemno = Convert.ToString(DBConfig.sqlite_datareader["itemno"]);
                    string _item = Convert.ToString(DBConfig.sqlite_datareader["item"]);
                    string _unit = Convert.ToString(DBConfig.sqlite_datareader["unit"]);
                    string _num = Convert.ToString(DBConfig.sqlite_datareader["num"]);
                    string _total = Convert.ToString(DBConfig.sqlite_datareader["total"]);
                    string _action = Convert.ToString(DBConfig.sqlite_datareader["action"]);

                    // 若為新增才加入表格(紀錄清單)
                    if (_action == "新增")
                    {
                        dataGridView1.Rows.Add(
                            new Object[]
                            {
                                _no, _date, _type, _itemno, _item, _unit, _num, _total
                            }
                        );
                    }
                    // 全部加入表格(操作紀錄)
                    dataGridView3.Rows.Add(
                        new Object[]
                        {
                            _no, _date, _type, _itemno, _item, _unit, _num, _total, _action
                        }
                    );

                    log_index = int.Parse(_no.Substring(2)); // 紀錄編號
                }
                DBConfig.sqlite_datareader.Close(); // 關閉資料庫讀取
            }

            // 設定表格外觀
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["type"].Value != null &&
                    row.Cells["type"].Value.ToString() == "出貨")
                {
                    row.Cells["type"].Style.ForeColor = Color.Green; // 字體顏色
                    row.Cells["type"].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold); // 粗體
                }
                if (row.Cells["type"].Value != null &&
                    row.Cells["type"].Value.ToString() == "進貨")
                {
                    row.Cells["type"].Style.ForeColor = Color.Red; // 字體顏色
                    row.Cells["type"].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold); // 粗體
                }
            }

            // 設定表格外觀
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row.Cells["action"].Value != null &&
                    row.Cells["action"].Value.ToString() == "新增")
                {
                    row.DefaultCellStyle.BackColor = Color.WhiteSmoke; // 背景顏色
                }
                if (row.Cells["action"].Value != null &&
                    row.Cells["action"].Value.ToString() == "取消")
                {
                    row.DefaultCellStyle.BackColor = Color.MistyRose; // 背景顏色
                }

                if (row.Cells["type3"].Value != null &&
                    row.Cells["type3"].Value.ToString() == "出貨")
                {
                    row.Cells["type3"].Style.ForeColor = Color.Green; // 字體顏色
                    row.Cells["type3"].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold); // 粗體
                }
                if (row.Cells["type3"].Value != null &&
                    row.Cells["type3"].Value.ToString() == "進貨")
                {
                    row.Cells["type3"].Style.ForeColor = Color.Red; // 字體顏色
                    row.Cells["type3"].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold); // 粗體
                }
            }
        }

        // 刷新庫存清單
        private void Show_DB_stock()
        {
            // 清空表格
            this.dataGridView2.Rows.Clear();

            // SQL查詢語法
            string sql = @"SELECT itemno, item, sum(CASE WHEN type = '進貨' THEN num ELSE -num END) AS sum " +
                "FROM log WHERE action = '新增' " +
                "GROUP BY item HAVING sum != 0 ORDER BY itemno;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows) // 若讀取結果有數據
            {
                while (DBConfig.sqlite_datareader.Read()) // 逐行讀取
                {
                    // 讀取資料庫欄位
                    string _itemno = Convert.ToString(DBConfig.sqlite_datareader["itemno"]);
                    string _item = Convert.ToString(DBConfig.sqlite_datareader["item"]);
                    string _sum = Convert.ToString(DBConfig.sqlite_datareader["sum"]);
                    // 加入表格
                    dataGridView2.Rows.Add(
                            new Object[]
                            {
                               "[" + _itemno + "]", _item, _sum
                            }
                        );
                }
                DBConfig.sqlite_datareader.Close(); // 關閉資料庫讀取
            }
        }

        public Form1()
        {
            InitializeComponent(); // 初始化
            Load_DB(); // 載入資料庫
            Show_DB_item(); // 刷新商品選項清單
            Show_DB_log(); // 刷新紀錄清單
            Show_DB_stock(); // 刷新庫存清單

            // 欄位初始為0
            unit_textBox.Text = "0";
            num_textBox.Text = "0";
        }


        /* 第一頁 - 【進出貨紀錄】*/
        double unitValue, numValue; // 單價及數量


        // 單價改變事件
        private void unit_textBox_TextChanged(object sender, EventArgs e)
        {
            // 若能轉成double則計算總價，否則總價顯示-
            if (double.TryParse(unit_textBox.Text, out unitValue) && double.TryParse(num_textBox.Text, out numValue))
            {
                total_label.Text = (unitValue * numValue).ToString();
            }
            else
            {
                total_label.Text = "-";
            }

        }

        // 數量改變事件
        private void num_textBox_TextChanged(object sender, EventArgs e)
        {
            // 若能轉成double則計算總價，否則總價顯示-
            if (double.TryParse(unit_textBox.Text, out unitValue) && double.TryParse(num_textBox.Text, out numValue))
            {
                total_label.Text = (unitValue * numValue).ToString();
            }
            else
            {
                total_label.Text = "-";
            }

        }

        // 新增紀錄按鈕
        private void addlog_btn_Click(object sender, EventArgs e)
        {
            // 判斷radio按鈕為"進貨"或"出貨"
            String radiotext = "";
            Boolean addBoolean = true;
            if (radioButton1.Checked) { radiotext = "進貨"; addBoolean = true; }
            // 若為"出貨"則檢查庫存是否能出貨
            if (radioButton2.Checked)
            {
                radiotext = "出貨";

                // SQL查詢語法
                string sql_stock = @"SELECT itemno, item, sum(CASE WHEN type = '進貨' THEN num ELSE -num END) AS sum " +
                "FROM log WHERE action = '新增' " +
                "GROUP BY item;";
                DBConfig.sqlite_cmd = new SQLiteCommand(sql_stock, DBConfig.sqlite_connect);
                DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

                if (DBConfig.sqlite_datareader.HasRows) // 若讀取結果有數據
                {
                    while (DBConfig.sqlite_datareader.Read()) // 逐行讀取
                    {
                        // 讀取資料庫欄位
                        string _itemno = Convert.ToString(DBConfig.sqlite_datareader["itemno"]);
                        string _sum = Convert.ToString(DBConfig.sqlite_datareader["sum"]);

                        // 若ID一樣且庫存不足，則布林值為false
                        if (item_comboBox.Text.Substring(1, 4) == _itemno &&
                            int.Parse(_sum) < int.Parse(num_textBox.Text))
                        {
                            addBoolean = false;
                            break;
                        }
                    }
                    DBConfig.sqlite_datareader.Close(); // 關閉資料庫讀取
                }
            }

            // 布林為true則可以新增
            if (addBoolean)
            {
                // SQL查詢語法
                string sql = @"INSERT INTO log (no, date, type, itemno, item, unit, num, total, action)
                VALUES( "
                 + " 'L" + (log_index + 1).ToString("D3") + "' , "
                 + " '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "' , "
                 + " '" + radiotext + "' , "
                 + " '" + item_comboBox.Text.Substring(1, 4) + "' , "
                 + " '" + item_comboBox.Text.Substring(7) + "' , "
                 + " '" + unit_textBox.Text + "' , "
                 + " '" + num_textBox.Text + "' , "
                 + " '" + total_label.Text + "' , "
                 + " '新增'   "
                 + ");";
                DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
                DBConfig.sqlite_cmd.ExecuteNonQuery();

                Show_DB_log(); // 刷新紀錄清單

                // 欄位數值清空或歸0
                item_comboBox.Text = "";
                unit_textBox.Text = "0";
                num_textBox.Text = "0";
            }
            else
            {
                MessageBox.Show(
                    "庫存不足，無法出貨，請重新檢查！",
                    "警告",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                    );
            }
            Show_DB_stock(); // 刷新庫存清單
        }

        // 紀錄清單刪除按鈕
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 若被點選的是"取消"按鈕
            var colname = dataGridView1.Columns[e.ColumnIndex].Name;
            if (colname.ToLower() == "cancel")
            {
                int rowIndex = e.RowIndex; // 行數

                if (rowIndex >= 0 && rowIndex < dataGridView1.Rows.Count) // 行數正常沒溢出
                {
                    // 抓取表格的對應欄位
                    string id = dataGridView1.Rows[rowIndex].Cells["id"].Value.ToString();
                    string createDate = dataGridView1.Rows[rowIndex].Cells["createDate"].Value.ToString();
                    string type = dataGridView1.Rows[rowIndex].Cells["type"].Value.ToString();
                    string itemId = dataGridView1.Rows[rowIndex].Cells["itemId"].Value.ToString();
                    string item = dataGridView1.Rows[rowIndex].Cells["item"].Value.ToString();
                    string cost = dataGridView1.Rows[rowIndex].Cells["cost"].Value.ToString();
                    string num = dataGridView1.Rows[rowIndex].Cells["num"].Value.ToString();
                    string total = dataGridView1.Rows[rowIndex].Cells["total"].Value.ToString();

                    DialogResult result = MessageBox.Show(
                        $"確定要取消此項紀錄？\n* 注意！取消某些紀錄會導致庫存數量異常，請謹慎操作！\n\n單號：[{id}]\n時間：{createDate}\n【{type}】[{itemId}] {item}\n單價：{cost}\n數量：{num}\n總計：{total}",
                        "警告",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning
                        );

                    if (result == DialogResult.OK)
                    {
                        // SQL查詢語法
                        string sql = @"UPDATE log " +
                            "SET action = '取消' " +
                            "where no = '" + id + "';";

                        DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
                        DBConfig.sqlite_cmd.ExecuteNonQuery();
                        Show_DB_log(); // 刷新紀錄清單
                        Show_DB_stock(); // 刷新庫存清單
                    }
                }
            }
        }


        /* 第二頁 - 【操作紀錄&選項管理】*/


        // 新增商品選項按鈕
        private void additem_btn_Click(object sender, EventArgs e)
            {
                // 輸入框不為空
                if (itemName_textBox.Text != "")
                {
                    // SQL查詢語法
                    string sql = @"INSERT INTO item (item_no, item_name)
                    VALUES( "
                     + " 'M" + (item_index + 1).ToString("D3") + "' , "
                     + " '" + itemName_textBox.Text + "'   "
                     + ");";
                    DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
                    DBConfig.sqlite_cmd.ExecuteNonQuery();

                    Show_DB_item(); // 刷新商品選項清單

                    itemName_textBox.Text = "";
                }
                else
                {
                    MessageBox.Show(
                        "請輸入新增的商品選項",
                        "警告",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                        );
                }
            }

        // 商品選項刪除按鈕
        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 若被點選的是"刪除"按鈕
            var colname = dataGridView4.Columns[e.ColumnIndex].Name;
            if (colname.ToLower() == "delete")
            {
                int rowIndex = e.RowIndex; // 行數

                if (rowIndex >= 0 && rowIndex < dataGridView4.Rows.Count) // 行數正常沒溢出
                {
                    // 抓取表格的對應欄位
                    string itemNo = dataGridView4.Rows[rowIndex].Cells["itemNo"].Value.ToString();
                    string itemName = dataGridView4.Rows[rowIndex].Cells["itemName"].Value.ToString();

                    DialogResult result = MessageBox.Show(
                        $"確定要刪除此項商品？：\n商品編號：{itemNo}\n商品名稱：{itemName}",
                        "警告",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning
                        );

                    if (result == DialogResult.OK)
                    {
                        // SQL查詢語法
                        string sql = @"DELETE from item " +
                            "where item_no = '" + itemNo + "';";

                        DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
                        DBConfig.sqlite_cmd.ExecuteNonQuery();
                        Show_DB_item(); // 刷新商品選項清單

                    }
                }
            }
        }


        /* 第三頁 - 【統計圖表】*/
        String imgname;

        // 繪圖 - 庫存
        private void chart1_btn_Click(object sender, EventArgs e)
        {
            imgname = "庫存圖表";
            // 清空圖表數據
            this.chart1.Series.Clear();

            // 創建一個新的 Series 物件
            Series newSeries = new Series("stocks");

            // 將圖表類型設置為長條圖
            newSeries.ChartType = SeriesChartType.Column;

            // SQL查詢語法
            string sql = @"SELECT itemno, item, sum(CASE WHEN type = '進貨' THEN num ELSE -num END) AS sum " +
                "FROM log WHERE action = '新增' " +
                "GROUP BY item HAVING sum != 0 ORDER BY itemno;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows) // 若讀取結果有數據
            {
                while (DBConfig.sqlite_datareader.Read()) // 逐行讀取
                {
                    // 讀取資料庫欄位
                    string _itemno = Convert.ToString(DBConfig.sqlite_datareader["itemno"]);
                    string _item = Convert.ToString(DBConfig.sqlite_datareader["item"]);
                    string _sum = Convert.ToString(DBConfig.sqlite_datareader["sum"]);

                    // 放入 Series
                    DataPoint point1 = new DataPoint();
                    point1.SetValueY(_sum);
                    point1.Label = _item + "\n" + _sum;
                    point1.LabelBackColor = Color.LightGray;
                    point1.LabelForeColor = Color.DarkBlue;
                    newSeries.Points.Add(point1);
                }
                DBConfig.sqlite_datareader.Close(); // 關閉資料庫讀取
            }

            // 將 Series 放入圖表
            this.chart1.Series.Add(newSeries);
        }

        // 繪圖 - 進貨出貨比例
        private void chart2_btn_Click(object sender, EventArgs e)
        {
            imgname = "進貨出貨比例圖表";
            // 清空圖表數據
            this.chart1.Series.Clear();

            // 創建一個新的 Series 物件
            Series newSeries = new Series("stocks");

            // 將圖表類型設置為圓餅圖
            newSeries.ChartType = SeriesChartType.Pie;

            // 設定進貨出貨變數
            int in_item_num = 0;
            int out_item_num = 0;

            // SQL查詢語法
            string sql = @"SELECT * from log;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows) // 若讀取結果有數據
            {
                while (DBConfig.sqlite_datareader.Read()) // 逐行讀取
                {
                    // 讀取資料庫欄位
                    string _type = Convert.ToString(DBConfig.sqlite_datareader["type"]);
                    string _num = Convert.ToString(DBConfig.sqlite_datareader["num"]);
                    string _action = Convert.ToString(DBConfig.sqlite_datareader["action"]);

                    // 若為新增才加入計算進貨出貨
                    if (_action == "新增")
                    {
                        switch (_type)
                        {
                            case "進貨":
                                in_item_num += int.Parse(_num);
                                break;
                            case "出貨":
                                out_item_num += int.Parse(_num);
                                break;
                        }
                    }
                }

                DBConfig.sqlite_datareader.Close(); // 關閉資料庫讀取
            }

            // 放入 Series
            DataPoint point1 = new DataPoint();
            point1.SetValueY(in_item_num);
            point1.Label = "進貨\n#VAL\n(#PERCENT)";
            newSeries.Points.Add(point1);

            DataPoint point2 = new DataPoint();
            point2.SetValueY(out_item_num);
            point2.Label = "出貨\n#VAL\n(#PERCENT)";
            newSeries.Points.Add(point2);

            // 將 Series 放入圖表
            this.chart1.Series.Add(newSeries);
        }

        // 匯出圖片
        private void output_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = imgname;
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            chart1.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

    }
}