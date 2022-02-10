using System;
using System.Net;
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
using System.IO;
using System.IO.Compression;

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
                        string uName = dt.Rows[0]["USER_LAST_NAME"].ToString();
                        string division = dt.Rows[0]["DIVISION"].ToString();
                        string email = dt.Rows[0]["USER_EMAIL"].ToString();
                        string uId = dt.Rows[0]["USER_ID"].ToString();
                        string uFName = dt.Rows[0]["USER_LAST_NAME"].ToString();
                        string uLName = dt.Rows[0]["USER_FIRST_NAME"].ToString();
                        string uMobile = dt.Rows[0]["USER_PHONE"].ToString();
                        ReportSession.UserId = uId;
                        ReportSession.Lname = uLName;
                        ReportSession.Fname = uFName;
                        ReportSession.Mobilenumber = uMobile;
                        ReportSession.Email = email;
                        ReportSession.Division = division;
                        if (EncDecFunction.Decrypt(pass) == txtPassword.Text)
                        {
                            DataTable viewablereports_dt = dbconn.getTable(DBQry.permission_reports(division));
                            List<string> viewablereports_list = new List<string>();
                            foreach (DataRow dtrow in viewablereports_dt.Rows)
                            {
                                viewablereports_list.Add(dtrow["REPORTS"].ToString());
                            }
                            this.Hide();
                            MainReportViewer report = new MainReportViewer(uId , uName , viewablereports_list);
                            report.Show();
                        }
                        else
                        {
                            MessageBox.Show("Пассворд буруу байна");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Нэвтрэх нэр буруу байна. IT хэлтэстэй холбогдож шинэ хэрэглэгч үүсгүүлнэ үү.");
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Forgotpassbutton_Click(object sender, EventArgs e)
        {
            if (txtUserName.Text.Length == 0)
            {
                XtraMessageBox.Show("Утасны дугаараа оруулна уу?");
            }
            else
            {
                DataTable dt = dbconn.getTable(DBQry.login(txtUserName.Text));
                if (dt.Rows.Count == 1)
                {
                    string email = dt.Rows[0]["USER_EMAIL"].ToString();
                    string pass = dt.Rows[0]["USER_PASSWORD"].ToString();
                    string uLName = dt.Rows[0]["USER_LAST_NAME"].ToString();
                    string decripted_pass = EncDecFunction.Decrypt(pass);
                    string content = "Сайн уу, " + uLName + "</br></br></br> Таны DReport -руу нэвтрэх нууц үг : <b>" + decripted_pass + "</b><br/> " + "/Энэхүү мэйл нь программаас очиж байгаа тул хариу бичихгүй байна уу/";
                    if(Email.emailSender(email , "noreply@ddishtv.mn", "DREPORT PASSWORD" , content, out string message))
                    {
                        string prefix = "Таны нууц үг " + email  + " хаяг руу очлоо. Та mail хаягаа шалгаарай.";
                        MessageBox.Show(prefix);
                    }
                    else
                    {
                        string prefix = "Амжилтгүй боллоо. " + message;
                        MessageBox.Show(prefix);
                    }
                }
                else
                {
                    MessageBox.Show("Утасны дугаар буруу байна. IT-тай холбогдож шинэ хэрэглэгч үүсгүүлээрэй.");
                }
            }
        }
    }
}
