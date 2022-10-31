using ClosedXML.Excel;

namespace BalanceCompute
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "Excel File (*.xlsx)|*.xls;*.xlsx"
            };

            var fileResult = openFileDialog1.ShowDialog();

            if (fileResult == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "Excel File (*.xlsx)|*.xls;*.xlsx"
            };

            var fileResult = openFileDialog2.ShowDialog();

            if (fileResult == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = openFileDialog2.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var balanceData = LoadBalanceData(textBox1.Text);

            var systemData = LoadSysytemData(textBox2.Text);

            var filePath = ExportResult(balanceData, systemData);

            textBox3.Text = textBox3.Text + Environment.NewLine + string.Format("產生完成 匯出檔案: {0}", filePath);
        }


        private static IEnumerable<SystemData> LoadSysytemData(string filePath)
        {
            List<SystemData> data = new List<SystemData>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(2);

                var lastRow = ws.LastRowUsed().RowNumber();

                for (int i = 1; i <= lastRow; i++)
                {
                    string rowData = ws.Cell(i, 1).Value.ToString() ?? string.Empty;

                    if (rowData.IndexOf("門市代號") > -1)
                    {
                        SystemData temp = new SystemData();

                        var idxSpec = rowData.LastIndexOf(" ");

                        temp.Store = rowData.Substring(idxSpec, rowData.Length - idxSpec).Trim().Replace("圓環", "店");

                        i += 3;

                        decimal.TryParse(ws.Cell(i, 2).Value.ToString() ?? string.Empty, out decimal amount);

                        temp.Cash = amount;

                        data.Add(temp);
                    }
                }
            }

            return data;
        }

        private static IEnumerable<BalanceData> LoadBalanceData(string filePath)
        {
            List<BalanceData> data = new List<BalanceData>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                var lastRow = ws.LastRowUsed().RowNumber();

                for (int i = 2; i <= lastRow; i++)
                {


                    string rowData = ws.Cell(i, 1).Value.ToString() ?? string.Empty;

                    BalanceData temp = new BalanceData();


                    temp.Store = rowData;

                    if (temp.Store.IndexOf("合計") > -1)
                    {
                        break;
                    }

                    decimal.TryParse(ws.Cell(i, 2).Value.ToString() ?? string.Empty, out decimal amount);

                    temp.LastBalance = amount;

                    data.Add(temp);
                }
            }

            return data;
        }

        private static string ExportResult(IEnumerable<BalanceData> _Balance, IEnumerable<SystemData> _System)
        {
            string fileName = AppDomain.CurrentDomain.BaseDirectory + "result.xlsx";

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("1030");

                int i = 1;

                int j = 1;

                ws.Cell(i, j++).SetValue("門市");
                ws.Cell(i, j++).SetValue("10/29餘額");
                ws.Cell(i, j++).SetValue("10/30現金收入");             
                ws.Cell(i, j++).SetValue("10/30餘額");


                foreach (var item in _Balance)
                {
                    var system = _System.FirstOrDefault(x => x.Store.IndexOf(item.Store) > -1);

                    item.Cash = system.Cash;

                    j = 1;

                    ws.Cell(++i, j++).SetValue(item.Store);

                    ws.Cell(i, j++).SetValue(item.LastBalance);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0.00";

                    ws.Cell(i, j++).SetValue(item.Cash);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0.00";

                    ws.Cell(i, j++).SetValue(item.NowBalance);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0.00";
                }

                wb.SaveAs(fileName);
            }

            return fileName;
        }


    }
}