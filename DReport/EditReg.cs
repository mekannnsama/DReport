using DevExpress.XtraEditors;
using DReportUtil.DBControl;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DReport
{
    public partial class EditReg : DevExpress.XtraEditors.XtraForm
    {
        DBConnecter dbconn = new DBConnecter();
        public EditReg( string _userId , string _fname , string _lname , string _mnumber , string _email , string _division)
        {
            InitializeComponent();
            textEdit1.EditValue = _fname;
            textEdit2.EditValue = _lname;
            textEdit3.EditValue = _mnumber;
            textEdit4.EditValue = _email;
            textEdit5.EditValue = _division;
            textEdit5.Enabled = false;
            textEdit3.Enabled = false;
            textEdit4.Enabled = false;
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            //Овог update хийх //
   /*         XtraMessageBox.Show(ReportSession.UserId);*/
            string updatesquery = string.Format(@"UPDATE DREPORT_LOGIN SET USER_FIRST_NAME = '{0}' WHERE USER_ID = {1} ", textEdit2.EditValue , ReportSession.UserId);
            string result = dbconn.iDBCommand(updatesquery);
            if ( result.Substring( result.Length - 3 , 3 ) == "000")
            {
                MessageBox.Show(result);
            }
            else
            {
                MessageBox.Show(result);
            }
        }
        private void simpleButton3_Click(object sender, EventArgs e)
        {
            //Нэр update хийх //
            string updatesquery = string.Format(@"UPDATE DREPORT_LOGIN SET USER_LAST_NAME = '{0}' WHERE USER_ID = {1}", textEdit1.EditValue, ReportSession.UserId);
            string result = dbconn.iDBCommand(updatesquery);
            if (result.Substring(result.Length - 3, 3) == "000")
            {
                MessageBox.Show(result);
            }
            else
            {
                MessageBox.Show(result);
            }
        }
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }


    }
}