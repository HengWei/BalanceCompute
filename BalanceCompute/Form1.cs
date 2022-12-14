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
            var systemData = LoadSysytemData(textBox2.Text, out DateTime? date, out string sysytemmessage);

            if (!string.IsNullOrEmpty(sysytemmessage))
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + sysytemmessage;
                return;
            }

            var balanceData = LoadBalanceData(textBox1.Text, out string title, out string balanceMessage);
            

            if (!string.IsNullOrEmpty(balanceMessage))
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + balanceMessage;
                return;
            }

            var filePath = ExportResult(balanceData, systemData, title, date.Value);

            textBox3.Text = textBox3.Text + Environment.NewLine + string.Format("???????? ???X????: {0}", filePath);
        }


        public static IEnumerable<SystemData> LoadSysytemData(string filePath, out DateTime? date, out string message)
        {
            message = string.Empty;

            date = null;

            List<SystemData> data = new List<SystemData>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("?? 1 ??");

                var lastRow = ws.LastRowUsed().RowNumber();

                for (int i = 1; i <= lastRow; i++)
                {
                    string rowData = ws.Cell(i, 1).Value.ToString() ?? string.Empty;

                    if (rowData.IndexOf("?????N??") > -1)
                    {
                        SystemData temp = new SystemData();

                        var idxSpec = rowData.LastIndexOf(" ");

                        temp.Store = rowData.Substring(idxSpec, rowData.Length - idxSpec).Trim().Replace("????", "??");

                        i += 3;

                        decimal.TryParse(ws.Cell(i, 2).Value.ToString() ?? string.Empty, out decimal amount);

                        temp.Cash = amount;

                        if(!date.HasValue)
                        {
                            var dateArray = (ws.Cell(i, 1).Value.ToString()??String.Empty).Split("/");

                            date = new DateTime(int.Parse(dateArray[0])+1911, int.Parse(dateArray[1]), int.Parse(dateArray[2]));
                        }

                        data.Add(temp);
                    }
                }
            }

            if(data.Count()==0)
            {
                message = "???s????????";
            }

            return data;
        }

        private static IEnumerable<BalanceData> LoadBalanceData(string filePath, out string title, out string message)
        {
            List<BalanceData> data = new List<BalanceData>();

            message = string.Empty;
            title = string.Empty;

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                var lastRow = ws.LastRowUsed().RowNumber();

                int j = 0;

                while (true)
                {
                    var tempTitle = ws.Cell(1, ++j).Value.ToString() ?? string.Empty;

                    if(string.IsNullOrEmpty(tempTitle))
                    {
                        break;
                    }
                    else
                    {
                        title = tempTitle;
                    }
                }
                

                for (int i = 2; i <= lastRow; i++)
                {
                    string rowData = ws.Cell(i, 1).Value.ToString() ?? string.Empty;

                    BalanceData temp = new BalanceData();


                    temp.Store = rowData;

                    if (temp.Store.IndexOf("?X?p") > -1)
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

        private static string ExportResult(IEnumerable<BalanceData> _Balance, IEnumerable<SystemData> _System, string title, DateTime date)
        {
            string fileName = AppDomain.CurrentDomain.BaseDirectory + String.Format("{0}.xlsx", date.ToString("MMdd"));

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet(date.ToString("MMdd"));

                int i = 1;

                int j = 1;

                ws.Cell(i, j++).SetValue("????");
                ws.Cell(i, j++).SetValue(title);
                ws.Cell(i, j++).SetValue(string.Format("{0}?{?????J", date.ToString("MM/dd")));             
                ws.Cell(i, j++).SetValue(string.Format("{0}?l?B", date.ToString("MM/dd")));


                foreach (var item in _Balance)
                {
                    var system = _System.FirstOrDefault(x => x.Store.IndexOf(item.Store) > -1);

                    item.Cash = system.Cash;

                    j = 1;

                    ws.Cell(++i, j).SetValue(item.Store);

                    ws.Cell(i, ++j).SetValue(item.LastBalance);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                    ws.Cell(i, ++j).SetValue(item.Cash);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                    ws.Cell(i, ++j).SetValue(item.NowBalance);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";
                }


                j = 1;
                ws.Cell(++i, j).SetValue("?X?p");
                ws.Range(i,1,i,4).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                ws.Cell(i, ++j).SetValue(_Balance.Sum(x=>x.LastBalance));
                ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                ws.Cell(i, ++j).SetValue(_Balance.Sum(x => x.Cash));
                ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                ws.Cell(i, ++j).SetValue(_Balance.Sum(x => x.NowBalance));
                ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";


                ws.Columns().AdjustToContents();

                wb.SaveAs(fileName);
            }

            return fileName;
        }


    }
}