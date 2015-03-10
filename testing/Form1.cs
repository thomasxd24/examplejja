using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;

namespace testing
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            webBrowser1.Navigate("http://easy-rea.com/products/index");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.ShowDialog();
                string path = openFileDialog1.FileName;
                DataTable dt = exceldata(path);
                richTextBox1.Text = "// ca commence..." + Environment.NewLine;
                foreach (DataRow row in dt.Rows)
                {
                    string text = row[0].ToString();
                    string textcap = text.ToUpper();
                    richTextBox1.Text += "setTimeout(function(){updateProduct('" + text + "', '0', '/products/update', 1, 110, 'divAlertQteDispoMessage', 0, 900, 'divAlertNbRowsMessage'); }, 3000);" + Environment.NewLine;
                }
                button2.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Vous ete en train de utiliser le file excel, il faut que vous fermez. Oringinale erreur:" + ex);
            }
        }

        public static DataTable exceldata(string filePath)
        {     
            DataTable dtexcel = new DataTable();
                string HDR = "No";
                string strConn;
                if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                else
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                //Looping Total Sheet of Xl File
                /*foreach (DataRow schemaRow in schemaTable.Rows)
                {
                }*/
                //Looping a first Sheet of Xl File
                DataRow schemaRow = schemaTable.Rows[0];
                string sheet = schemaRow["TABLE_NAME"].ToString();
                if (!sheet.EndsWith("_"))
                {
                    string query = "SELECT  * FROM [sheet1$]";
                    OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                    daexcel.Fill(dtexcel);
                }
            
            conn.Close();
            return dtexcel;
 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string path = openFileDialog1.FileName;
            DataTable dt = exceldata(path);
            richTextBox1.Text = "var myElem = 'Les articles:'" + Environment.NewLine;
            foreach (DataRow row in dt.Rows)
            {
                string text = row[0].ToString();
                string textcap = text.ToUpper();
                richTextBox1.Text += "if (document.getElementById('product" + textcap + "') == null) { myElem += '" + text + ",'; };" + Environment.NewLine;

            }
            richTextBox1.Text += "document.getElementById('footer').innerHTML = myElem + '. Ne peuvent pas etre commandé'";
        }

        private void button3_Click(object sender, EventArgs e)
        {

            MyInvokeScript("updateProduct", "102550","4","/products/update",1,110,"divAlertQteDispoMessage",0,900,"divAlertNbRowsMessage");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
HtmlDocument doc = webBrowser1.Document;
HtmlElement head = doc.GetElementsByTagName("head")[0];
HtmlElement s = doc.CreateElement("script");
s.SetAttribute("text","function sayHello() {updateProduct('709663341','0','/products/update',1,110,'divAlertQteDispoMessage',0,900,'divAlertNbRowsMessage'); }");
head.AppendChild(s);
        }

        private object MyInvokeScript(string name, params object[] args)
        {
            return webBrowser1.Document.InvokeScript(name, args);
        }
    }
}
