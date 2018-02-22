using SAPbobsCOM;
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

namespace csv2activities
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            textBox1.Text = Properties.connection.Default.server;
            textBox2.Text = Properties.connection.Default.port;
            comboBox1.Text = Properties.connection.Default.type;
            textBox4.Text = Properties.connection.Default.dbUser;
            textBox5.Text = Properties.connection.Default.dbPassword;

            textBox6.Text = Properties.connection.Default.sapUser;
            textBox7.Text = Properties.connection.Default.sapPassword;
            textBox8.Text = Properties.connection.Default.sapCompany;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "CSV files | *.csv"; // file types, that will be allowed to upload
            dialog.Filter += "| TXT files | *.txt"; 
            dialog.Multiselect = false; // allow/deny user to upload more than one file at a time
            if (dialog.ShowDialog() == DialogResult.OK) // if user clicked OK
            {
                Application.DoEvents();
                Cursor = Cursors.WaitCursor; 
                var myCompany = connectSap();

                string path = dialog.FileName; // get name of file
                bool check = SAPHelper.insActFromCSV(myCompany,path,',');

                if (check)
                {
                    MessageBox.Show("All records are imported in SAP B1", "Import from CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else {
                    MessageBox.Show("Somethings goes wrong, check the CSV format", "Import from CSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                SAPHelper.destroyCom(myCompany);
                Cursor = Cursors.Default;
            }
         
        }

        public Company connectSap(bool destroy=false, bool popup = false) {

            string server = textBox1.Text;
            string port = textBox2.Text;

            string dbUser = textBox4.Text;
            string dbPassword = textBox5.Text;

            string b1User = textBox6.Text;
            string b1Pass = textBox7.Text;
            string b1Company = textBox8.Text;

            var serverType = comboBox1.Text;

            SAPHelper.tipoServer tipo;
            switch (serverType)
            {
                case "hana":
                    tipo = SAPHelper.tipoServer.hana;
                    break;

                case "mssql2012":
                    tipo = SAPHelper.tipoServer.mssql2012;
                    break;

                case "mssql2014":
                    tipo = SAPHelper.tipoServer.mssql2014;
                    break;

                default:
                    tipo = SAPHelper.tipoServer.hana;
                    break;
            }
            
            //connetti a DB
            
            Company myCompany = SAPHelper.getSocieta(server, port, tipo, dbUser, dbPassword, b1User, b1Pass, b1Company);
            if (popup)
            {
                if (myCompany != null)
                {
                    MessageBox.Show("Correct configuration", "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Some data is missing or wrong", "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }


            //SAPHelper.getAct(myCompany);
            if (destroy)
            {
                SAPHelper.destroyCom(myCompany);
                return null;
            }

            return myCompany;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            Cursor = Cursors.WaitCursor;
            connectSap(true, true);
            Cursor = Cursors.Default;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            Cursor = Cursors.WaitCursor;
            var myCompany = connectSap();
            SAPHelper.getActToXML(myCompany);
            SAPHelper.destroyCom(myCompany);
            Cursor = Cursors.Default;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "CSV files | *.csv"; // file types, that will be allowed to upload
            dialog.Filter += "| TXT files | *.txt";
            dialog.Multiselect = false; // allow/deny user to upload more than one file at a time
            if (dialog.ShowDialog() == DialogResult.OK) // if user clicked OK
            {
                Application.DoEvents();
                Cursor = Cursors.WaitCursor;
                var myCompany = connectSap();

                string path = dialog.FileName; // get name of file
                bool check = SAPHelper.updActFromCSV(myCompany, path, ',');

                if (check)
                {
                    MessageBox.Show("All records updated correctly", "Update from CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Somethings goes wrong, check the CSV format", "Update from CSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                SAPHelper.destroyCom(myCompany);
                Cursor = Cursors.Default;

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Properties.connection.Default.server = textBox1.Text;
            Properties.connection.Default.port = textBox2.Text;
            Properties.connection.Default.type = comboBox1.Text;
            Properties.connection.Default.dbUser = textBox4.Text;
            Properties.connection.Default.dbPassword = textBox5.Text;
            Properties.connection.Default.sapUser = textBox6.Text;
            Properties.connection.Default.sapPassword = textBox7.Text;
            Properties.connection.Default.sapCompany = textBox8.Text;
            
            Properties.connection.Default.Save();

            MessageBox.Show("Configuration saved", "Configuration", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
    }
}
