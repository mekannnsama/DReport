using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using DReportUtil.DBControl;
using DReportUtil.Utils;

namespace DReport
{
    public partial class Form1 : XtraForm
    {
        DBConnecter dbconn = new DBConnecter();
        public Form1()
        {
            InitializeComponent();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Escape))
            {
                Application.Exit();
                return true;
            }
            

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                if(txtUserName.Text.Length != 0)
                {
                    DataTable dt = dbconn.getTable(DBQry.login(txtUserName.Text));
                    if(dt.Rows.Count == 1)
                    {
                        string pass = dt.Rows[0]["USER_PASSWORD"].ToString();
                        string uName = dt.Rows[0]["USER_FIRST_NAME"].ToString();
                        string permi = dt.Rows[0]["DIVISION"].ToString();
                        if (EncDecFunction.Decrypt(pass) == txtPassword.Text)
                        {
                            this.Hide();
                            MainReportViewer report = new MainReportViewer(permi);
                            report.Show();
                        }
                        else
                        {
                            MessageBox.Show("Пассворд буруу байна");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Нэр буруу байна");
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }


    }
}
