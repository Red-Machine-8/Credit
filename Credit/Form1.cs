using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace Credit
{

    public partial class Form1 : Form
    {
        String filePath = "";
        DataTable dt = new DataTable();
        NpgsqlDataAdapter adapter;
        String updateRow = "Update reports SET x2 = '{0}', x3 = '{1}', x4 = '{2}', x5 = '{3}', x6 = '{4}', x7 = '{5}', x8 = '{6}', x9 = '{7}', x10 = '{8}', x11 = '{9}', x12 = '{10}', x13 = '{11}', x14 = '{12}' WHERE x1a = '{13}' AND x1b = '{14}' AND x1c = '{15}' AND x1d = '{16}'";
        String insertRow = "INSERT INTO reports VALUES ('{0}', '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16})";
        NpgsqlConnection con = new NpgsqlConnection("Server=127.0.0.1;Port=5432;User Id=postgres;Password=1234;Database=credits;");
        public Form1()
        {
            InitializeComponent();
            Connection();

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Reload()
        {
            dt.Clear();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
            //dt = data.GetChanges();
        }

        private void Connection()
        {

            try
            {
                con.Open();
                adapter = new NpgsqlDataAdapter("select * from reports", con);
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;

                Console.WriteLine("Соединение установлено!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка соединения");
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private Boolean Comparison(String[] data, int i)
        {
            if (dt.Rows.Count >= i)
            {
                if ((data[0] == dt.Rows[i][0].ToString()) && (data[1] == dt.Rows[i][1].ToString()) && (data[2] == dt.Rows[i][2].ToString()) && (data[3] == dt.Rows[i][3].ToString()))
                {
                    return true;
                }
                else
                {
                    i += 1;
                    return Comparison(data, i);
                }
            }
            else
            {
                return false;
            }
        }

        private void загрузитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string[] rowValue = { };
            OpenFileDialog file = new OpenFileDialog();
            file.InitialDirectory = "c:\\";

            if (file.ShowDialog() == DialogResult.OK)
            {
                filePath = file.FileName;

                var Stream = file.OpenFile();
                StreamReader reader = new StreamReader(Stream, System.Text.Encoding.Default);
                var row = reader.ReadLine();

                while (row != null)
                {
                    if ((row[0] != '*') && (row[0] != '#') && (row != "ТБ=01"))
                    {
                        row = row.Replace(",", ".");
                        rowValue = row.Split('|');
                        rowValue = rowValue.Take(rowValue.Count() - 1).ToArray();
                        if (Comparison(rowValue, 0))
                        {
                            string up = string.Format(updateRow, rowValue[4], rowValue[5], rowValue[6], rowValue[7], rowValue[8], rowValue[9], rowValue[10], rowValue[11], rowValue[12], rowValue[13], rowValue[14], rowValue[15], rowValue[16], rowValue[0], rowValue[1], rowValue[2], rowValue[3]);
                            var cmd = new NpgsqlCommand(up, con);
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            string ins = string.Format(insertRow, rowValue[0], rowValue[1], rowValue[2], rowValue[3], rowValue[4], rowValue[5], rowValue[6], rowValue[7], rowValue[8], rowValue[9], rowValue[10], rowValue[11], rowValue[12], rowValue[13], rowValue[14], rowValue[15], rowValue[16]);
                            var cmd = new NpgsqlCommand(ins, con);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    row = reader.ReadLine();
                }
                reader.Close();
                Reload();
            }
        }

        private void xMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XML files (*.xml)|*.xml";
            if (sfd.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            string path = sfd.FileName;
            ds.Tables.Add(dt);
            ds.WriteXml(path, XmlWriteMode.IgnoreSchema);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sqlcmd;
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "")
            {
                sqlcmd = string.Format("Select * From reports Where x1a = '{0}' and x1b = '{1}' and x1c = '{2}' and x1d = '{3}'", textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else if (textBox1.Text == "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "")
            {
                sqlcmd = string.Format("Select * From reports Where x1b = '{0}' and x1c = '{1}' and x1d = '{2}'", textBox2.Text, textBox3.Text, textBox4.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text != "" && textBox4.Text != "")
            {
                sqlcmd = string.Format("Select * From reports Where x1c = '{0}' and x1d = '{1}'", textBox3.Text, textBox4.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text != "")
            {
                sqlcmd = string.Format("Select * From reports Where x1d = '{0}'", textBox4.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text != "" && textBox4.Text == "")
            {
                sqlcmd = string.Format("Select * From reports Where x1c = '{0}'", textBox3.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else if (textBox1.Text == "" && textBox2.Text != "" && textBox3.Text == "" && textBox4.Text == "")
            {
                sqlcmd = string.Format("Select * From reports Where x1b = '{0}'", textBox2.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else if (textBox1.Text != "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "")
            {
                sqlcmd = string.Format("Select * From reports Where x1a = '{0}'", textBox1.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }else if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text == "" && textBox4.Text == "")
            {
                sqlcmd = string.Format("Select * From reports Where x1a = '{0}' and x1b = '{1}'", textBox1.Text, textBox2.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }else if (textBox1.Text != "" && textBox2.Text == "" && textBox3.Text != "" && textBox4.Text == "")
            {
                sqlcmd = string.Format("Select * From reports Where x1a = '{0}' and x1c = '{1}'", textBox1.Text, textBox3.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }else if (textBox1.Text != "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text != "")
            {
                sqlcmd = string.Format("Select * From reports Where x1a = '{0}' and x1d = '{1}'", textBox1.Text, textBox4.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }else if (textBox1.Text == "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text == "")
            {
                sqlcmd = string.Format("Select * From reports Where x1b = '{0}' and x1c = '{1}'", textBox2.Text, textBox3.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }else if (textBox1.Text == "" && textBox2.Text != "" && textBox3.Text == "" && textBox4.Text != "")
            {
                sqlcmd = string.Format("Select * From reports Where x1b = '{0}' and x1d = '{1}'", textBox2.Text, textBox4.Text);
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else
            {
                sqlcmd = "Select * From reports for xml auto";
                adapter = new NpgsqlDataAdapter(sqlcmd, con);
                dt.Clear();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach(DataGridViewRow row in dataGridView1.SelectedRows)
            {
                    string cmd = String.Format("Delete From reports Where x1a = '{0}' and x1b = '{1}' and x1c = '{2}' and x1d = '{3}'", row.Cells[0].Value.ToString(), row.Cells[1].Value.ToString(), row.Cells[2].Value.ToString(), row.Cells[3].Value.ToString());
                    var del = new NpgsqlCommand(cmd, con);
                    dataGridView1.Rows.RemoveAt(row.Index);
                    del.ExecuteNonQuery();
            }
            Reload();
        }

        private string klei(int i)
        {
            string str = dataGridView1[0, i].Value.ToString()+ " " + dataGridView1[1, i].Value.ToString() + " " + dataGridView1[2, i].Value.ToString() + " " + dataGridView1[3, i].Value.ToString();
            return str;
        }

        private void Table(Excel.Worksheet wsh)
        {
            Excel.Range r1 = wsh.Range["A1", "A4"];
            r1.Merge();
            r1.Value = "Номер (код) счета бюджетного учета";
            r1.EntireColumn.AutoFit();
            r1.Borders.Color = ColorTranslator.ToOle(Color.Black);
            r1 = wsh.Range["B1", "N1"];
            r1.Merge();
            r1.Value = "Сумма задолженности, руб";
            r1.EntireColumn.AutoFit();
            r1.HorizontalAlignment = HorizontalAlignment.Center;
            r1 = wsh.Range["B2", "D2"];
            r1.Merge();
            r1.Value = "на начало года";
            r1.EntireColumn.AutoFit();
            r1.HorizontalAlignment = HorizontalAlignment.Center;
            r1 = wsh.Range["B3", "B4"];
            r1.Merge();
            r1.Value = "всего";
            r1.EntireColumn.AutoFit();
            r1.HorizontalAlignment = HorizontalAlignment.Center;
            r1 = wsh.Range["C3", "D3"];
            r1.Merge();
            r1.Value = "из них:";
            r1.EntireColumn.AutoFit();
            r1.HorizontalAlignment = HorizontalAlignment.Center;
        }

        private void xLSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();

            exApp.Workbooks.Open(Application.StartupPath.ToString() + "\\Шаблон.xlsx");
            Excel.Worksheet wsh = exApp.Worksheets["Бюджет"];
            //Table(wsh);
            int i, j;
            for (i = 0; i <= dataGridView1.RowCount - 2; i++)
            {
                for (j = 4; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    Console.WriteLine(klei(i));
                    wsh.Cells[i + 6, 1] = klei(i);
                    wsh.Cells[i + 6, 1].Borders.Color = ColorTranslator.ToOle(Color.Black);
                    if(Convert.ToInt32(dataGridView1[j, i].Value) == 0)
                    {
                        wsh.Cells[i + 6, j - 2] = "-";
                        wsh.Cells[i + 6, j - 2].Borders.Color = ColorTranslator.ToOle(Color.Black);
                    }
                    else
                    {
                        wsh.Cells[i + 6, j - 2] = dataGridView1[j, i].Value;
                        wsh.Cells[i + 6, j - 2].Borders.Color = ColorTranslator.ToOle(Color.Black);
                    }                    
                }
            }
            exApp.Visible = true;
        }
    }
}
