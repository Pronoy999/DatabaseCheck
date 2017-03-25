using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace DatabaseCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =E:\Documents\NITMAS\Project\Database1.accdb");
            String sel = "SELECT * FROM SHEET2";
            OleDbCommand cmd = new OleDbCommand(sel, connection);
            OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
            connection.Open();
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet, "Sheet2");
            connection.Close();
            dataGridView1.DataSource = dataSet;
            dataGridView1.DataMember = "Sheet2";
            DataTable dataTable = dataSet.Tables["SHEET2"];
            HashSet<string> set1 = new HashSet<string>();
            foreach(DataRow row in dataTable.Rows)
            {
                set1.Add(row[1].ToString());
               
            }
            HashSet<int> numbers1 = new HashSet<int>();
            HashSet<int> numbers2 = new HashSet<int>();
            foreach(string ele in set1)            {
                
                foreach (DataRow row in dataTable.Rows)
                {
                    if (ele.Equals("D"))
                    {
                        if (ele.Equals(row[1].ToString()))
                            numbers1.Add((int)row[0]);
                    }
                    else if (ele.Equals("E"))
                    {
                        if (ele.Equals(row[1].ToString()))
                            numbers2.Add((int)row[0]);
                    }
                }
            }
            foreach(int a in numbers1)
            {
                comboBox1.Items.Add(a);
            }
            foreach (int a in numbers2)
            {
                comboBox2.Items.Add(a);
            }
            HashSet<HashSet<int>> numbers = new HashSet<HashSet<int>>();
            numbers.Add(numbers1);
            numbers.Add(numbers2);
            label1.Text = numbers.Count.ToString();
        }
    }
}
