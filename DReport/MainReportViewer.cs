using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.AccessControl;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.DocumentView.Controls;
using DevExpress.XtraBars.Alerter;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraTab;
using DReportUtil.DBControl;
using DReportUtil.Utils;

namespace DReport
{
    public partial class MainReportViewer : XtraForm
    {
        
        DBConnecter dbconn = new DBConnecter();
        private string dbres = string.Empty;
        public MainReportViewer(string userid , string username , List<string> viewreports)
        {
            InitializeComponent( );
            user_greetings.Text = "Сайн байна уу, " + username;
            ActivateTabPage(viewreports);
            user_time.Text = DateTime.Now.ToString("G") ;
        }
        private void ActivateTabPage( List<string> accessiblereports )
        {
            foreach (XtraTabPage tabPage in xtratabcontrol_report.TabPages)
            {
/*                label1.Text = tabPage.Name;*/
                
                if (!accessiblereports.Contains(tabPage.Name))
                {
                    tabPage.PageVisible = false;
                }
            }
        }
        private void MainReportViewer_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = XtraMessageBox.Show("ТА СИСТЕМЭЭС ГАРАХДАА ИТГЭЛТЭЙ БАЙНА УУ ?", "ПРОГРАМ ХААХ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    Application.Exit();
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

       

        
        DataTable dt_rpt_28 = new DataTable();
        DataTable dt_rpt_25 = new DataTable();
        private object dt_rpt_26;
        private DataTable dt_rpt_27;
        private DataTable dt_rpt_24;
        private DataTable dt_rpt_1;
        private DataTable dt_rpt_2;
        private DataTable dt_rpt_3;
        private object dt_rpt_4;
        private object dt_rpt_5;
        private object dt_rpt_6;
        private object dt_rpt_7;
        private object dt_rpt_8;
        private object dt_rpt_9;
        private object dt_rpt_10;
        private object dt_rpt_12;
        private object dt_rpt_13;
        private object dt_rpt_14;
        private object dt_rpt_15;
        private object dt_rpt_16;
        private object dt_rpt_17;
        private object dt_rpt_18;
        private object dt_rpt_19;
        private object dt_rpt_20;
        private object dt_rpt_21;
        private object dt_rpt_22;
        private object dt_rpt_23;

        
        

        private void rpt_1()
        {
            try
            {
                string beginDate = dtBeginRpt_1.Text.Replace("-", "");
                string endDate = dtEndRpt_1.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_1(beginDate, endDate));
                    dt_rpt_1 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT1");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_2()
        {
            try
            {
                string beginDate = dtBeginRpt_2.Text.Replace("-", "");
                string endDate = dtEndRpt_2.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_2(beginDate, endDate));
                    dt_rpt_2 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT2");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_3()
        {
            try
            {
                string beginDate = dtBeginRpt_3.Text.Replace("-", "");
                string endDate = dtEndRpt_3.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_3(beginDate, endDate));
                    dt_rpt_3 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT3");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_4()
        {
            try
            {
                string beginDate = dtBeginRpt_4.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_4(beginDate));
                    dt_rpt_4 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT4");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_5()
        {
            try
            {
                string beginDate = dtBeginRpt_5.Text.Replace("-", "");
                string endDate = dtEndRpt_5.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_5(beginDate, endDate));
                    dt_rpt_5 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT5");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_6()
        {
            try
            {
                string beginDate = dtBeginRpt_6.Text.Replace("-", "");
                string endDate = dtEndRpt_6.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_6(beginDate, endDate));
                    dt_rpt_6 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT6");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_7()
        {
            try
            {
                string beginDate = dtBeginRpt_7.Text.Replace("-", "");
                string endDate = dtEndRpt_7.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_7(beginDate, endDate));
                    dt_rpt_7 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT13");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_8()
        {
            try
            {
                string beginDate = dtBeginRpt_8.Text.Replace("-", "");
                string endDate = dtEndRpt_8.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_8(beginDate, endDate));
                    dt_rpt_8 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT8");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_9()
        {
            try
            {
                string beginDate = dtBeginRpt_9.Text.Replace("-", "");
                string endDate = dtEndRpt_9.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_9(beginDate, endDate));
                    dt_rpt_9 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT9");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_10()
        {
            try
            {
                string beginDate = dtBeginRpt_10.Text.Replace("-", "");
                string endDate = dtEndRpt_10.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_10(beginDate, endDate));
                    dt_rpt_10 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT10");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_12()
        {
            try
            {
                string beginDate = dtBeginRpt_12.Text.Replace("-", "");
                string endDate = dtEndRpt_12.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_12(beginDate, endDate));
                    dt_rpt_12 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT12");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_13()
        {
            try
            {
                string beginDate = dtBeginRpt_13.Text.Replace("-", "");
                string endDate = dtEndRpt_13.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_13(beginDate, endDate));
                    dt_rpt_13 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT13");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_14()
        {
            try
            {
                string beginDate = dtBeginRpt_14.Text.Replace("-", "");
                string endDate = dtEndRpt_14.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_14(beginDate, endDate));
                    dt_rpt_14 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT14");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_15()
        {
            try
            {
                string beginDate = dtBeginRpt_15.Text.Replace("-", "");
                string endDate = dtEndRpt_15.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_15(beginDate, endDate));
                    dt_rpt_15 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT15");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_16()
        {
            try
            {
                string beginDate = dtBeginRpt_16.Text.Replace("-", "");
                //string endDate = dtEndRpt21.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_16(beginDate));
                    dt_rpt_16 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT16");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_17()
        {
            try
            {
                string beginDate = dtBeginRpt_17.Text.Replace("-", "");
                string endDate = dtEndRpt_17.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_17(beginDate,endDate));
                    dt_rpt_17 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT17");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_18()
        {
            try
            {
                string beginDate = dtBeginRpt_18.Text.Replace("-", "");
                string endDate = dtEndRpt_18.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_18(beginDate, endDate));
                    dt_rpt_18 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT18");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_19()
        {
            try
            {
                string beginDate = dtBeginRpt_19.Text.Replace("-", "");
                string endDate = dtEndRpt_19.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_19(beginDate, endDate));
                    dt_rpt_19 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT19");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_20()
        {
            try
            {
                string beginDate = dtBeginRpt_20.Text.Replace("-", "");
                string endDate = dtEndRpt_20.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_20(beginDate, endDate));
                    dt_rpt_20 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT20");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_21()
        {
            try
            {
                string beginDate = dtBeginRpt_21.Text.Replace("-", "");
                string endDate = dtEndRpt_21.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_21(beginDate, endDate));
                    dt_rpt_21 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT21");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_22()
        {
            try
            {
                string beginDate = dtBeginRpt_22.Text.Replace("-", "");
                string endDate = dtEndRpt_22.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_22(beginDate, endDate));
                    dt_rpt_22 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT22");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_23()
        {
            try
            {
                string beginDate = dtBeginRpt_23.Text.Replace("-", "");
                string endDate = dtEndRpt_23.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_23(beginDate, endDate));
                    dt_rpt_23 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT27");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_24()
        {
            try
            {
                string beginDate = dtBeginRpt_24.Text.Replace("-", "");
                string endDate = dtEndRpt_24.Text.Replace("-", "");
                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_24(beginDate, endDate));
                    dt_rpt_24 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }

            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT24");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_25()
        {
            try
            {
                string beginDate = dtBeginRpt_25.Text.Replace("-", "");
                string endDate = dtEndRpt_25.Text.Replace("-", "");
                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_25(beginDate, endDate));
                    dt_rpt_25 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }

            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT25");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt_26()
        {
            try
            {
                string beginDate = dtBeginRpt_26.Text.Replace("-", "");
                string endDate = dtEndRpt_26.Text.Replace("-", "");
                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_26(beginDate, endDate));
                    dt_rpt_26 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT26");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_27()
        {
            try
            {
                string beginDate = dtBeginRpt_27.Text.Replace("-", "");
                string endDate = dtEndRpt_27.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_27(beginDate, endDate));
                    dt_rpt_27 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT27");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt_28()
        {
            try
            {
                string beginDate = dtBeginRpt_28.Text.Replace("-", "");
                string endDate = dtEndRpt_28.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt_28(beginDate, endDate));
                    dt_rpt_28 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT28");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void expdata_24(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_24.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_25(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_25.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_26(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_26.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_27(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_27.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_28(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_28.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_1(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_1.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_2(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_2.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_3(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_3.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_4(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_4.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_5(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_5.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_6(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_6.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_7(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_7.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_8(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));
                gridControl_8.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_9(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_9.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_10(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));
                gridControl_10.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_12(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_12.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_13(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_13.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_14(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_14.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata_15(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_15.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_16(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_16.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_17(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_17.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_18(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_18.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_19(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_19.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_20(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_20.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_21(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_21.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_22(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_22.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata_23(string filename)
        {
            try
            {
                var logfolder = AppDomain.CurrentDomain.BaseDirectory + "\\EXPDATA";
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(logfolder);
                FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
                DirectorySecurity ds = null;
                if (!di.Exists)
                {
                    Directory.CreateDirectory(logfolder);
                }
                ds = di.GetAccessControl();
                ds.AddAccessRule(fsar);
                string fname = string.Format(@"{0}\\{1}{2}.xlsx", logfolder, filename, DateTime.Today.ToString("yyyyMMdd"));//logfolder + "\\" + DateTime.Today.ToString("yyyyMMdd") + "productlist.xlsx";
                gridControl_23.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnSearchRpt_24_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_24))
            {
                frm.ShowDialog(this);
            }
            gridControl_24.DataSource = dt_rpt_24;
        }

        private void btnSearchRpt_25_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_25))
            {
                frm.ShowDialog(this);
            }
            gridControl_25.DataSource = dt_rpt_25;
        }

        private void btnSearchRpt_28_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_28))
            {
                frm.ShowDialog(this);
            }
            gridControl_28.DataSource = dt_rpt_28;
        }

        private void btnSearchRpt_26_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_26))
            {
                frm.ShowDialog(this);
            }
            gridControl_26.DataSource = dt_rpt_26;
        }

        private void btnSearchRpt_27_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_27))
            {
                frm.ShowDialog(this);
            }
            gridControl_27.DataSource = dt_rpt_27;
        }


        private void btnSearchRpt_1_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_1))
            {
                frm.ShowDialog(this);
            }
            gridControl_1.DataSource = dt_rpt_1;
        }

        private void btnSearchRpt_2_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_2))
            {
                frm.ShowDialog(this);
            }
            gridControl_2.DataSource = dt_rpt_2;
        }

        private void btnSearchRpt_3_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_3))
            {
                frm.ShowDialog(this);
            }
            gridControl_3.DataSource = dt_rpt_3;
        }

        private void btnSearchRpt_4_Click(object sender, EventArgs e)
        {
/*            MessageBox.Show("Засвартай байна. Лизинг үлдэгдэл гээд жижгээр бичсэн рүү орно уу !!!");
*/            using (frmWaitForm frm = new frmWaitForm(rpt_4))
            {
                frm.ShowDialog(this);
            }
            gridControl_4.DataSource = dt_rpt_4;
        }


        private void btnSearchRpt_5_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_5))
            {
                frm.ShowDialog(this);
            }
            gridControl_5.DataSource = dt_rpt_5;
        }
        private void btnSearchRpt_6_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_6))
            {
                frm.ShowDialog(this);
            }
            gridControl_6.DataSource = dt_rpt_6;
        }
        private void btnSearchRpt_7_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_7))
            {
                frm.ShowDialog(this);
            }
            gridControl_7.DataSource = dt_rpt_7;
        }
        private void btnSearchRpt_8_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_8))
            {
                frm.ShowDialog(this);
            }
            gridControl_8.DataSource = dt_rpt_8;
        }
        private void btnSearchRpt_9_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_9))
            {
                frm.ShowDialog(this);
            }
            gridControl_9.DataSource = dt_rpt_9;
        }
        private void btnSearchRpt_10_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_10))
            {
                frm.ShowDialog(this);
            }
            gridControl_10.DataSource = dt_rpt_10;
        }

        private void btnSearchRpt_12_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_12))
            {
                frm.ShowDialog(this);
            }
            gridControl_12.DataSource = dt_rpt_12;
        }

        private void btnSearchRpt_13_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_13))
            {
                frm.ShowDialog(this);
            }
            gridControl_13.DataSource = dt_rpt_13;
        }

        private void btnSearchRpt_14_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_14))
            {
                frm.ShowDialog(this);
            }
            gridControl_14.DataSource = dt_rpt_14;
        }

        private void btnSearchRpt_15_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_15))
            {
                frm.ShowDialog(this);
            }
            gridControl_15.DataSource = dt_rpt_15;
        }
        private void btnSearchRpt_16_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_16))
            {
                frm.ShowDialog(this);
            }
            gridControl_16.DataSource = dt_rpt_16;
        }

        private void btnSearchRpt_17_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_17))
            {
                frm.ShowDialog(this);
            }
            gridControl_17.DataSource = dt_rpt_17;

        }
        private void btnSearchRpt_18_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_18))
            {
                frm.ShowDialog(this);
            }
            gridControl_18.DataSource = dt_rpt_18;

        }

        private void btnSearchRpt_19_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_19))
            {
                frm.ShowDialog(this);
            }
            gridControl_19.DataSource = dt_rpt_19;

        }

        private void btnSearchRpt_20_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_20))
            {
                frm.ShowDialog(this);
            }
            gridControl_20.DataSource = dt_rpt_20;

        }
        private void btnSearchRpt_21_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_21))
            {
                frm.ShowDialog(this);
            }
            gridControl_21.DataSource = dt_rpt_21;
        }
        private void btnSearchRpt_22_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_22))
            {
                frm.ShowDialog(this);
            }
            gridControl_22.DataSource = dt_rpt_22;
        }
        private void btnSearchRpt_23_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt_23))
            {
                frm.ShowDialog(this);
            }
            gridControl_23.DataSource = dt_rpt_23;
        }


        private void btnToXLSX_24_Click(object sender, EventArgs e)
        {
            expdata_24("CashTopup_");
        }
        private void btnToXLSX_25_Click(object sender, EventArgs e)
        {
            expdata_25("Recharge_");
        }

        private void btnToXLSX_26_Click(object sender, EventArgs e)
        {
            expdata_26("Virtualacc_topup_");
        }
        private void btnToXLSX_27_Click(object sender, EventArgs e)
        {
            expdata_27("Virtualacc_movie_");
        }

        private void btnToXLSX_28_Click(object sender, EventArgs e)
        {
            expdata_28("Free_days_");
        }

        private void btnToXLSX_1_Click(object sender, EventArgs e)
        {
            expdata_1("Online_leasing_");
        }

        private void btnToXLSX_2_Click(object sender, EventArgs e)
        {
            expdata_2("Openfreeproduct_");
        }

        private void btnToXLSX_3_Click(object sender, EventArgs e)
        {
            expdata_3("Upoint_");
        }

        private void btnToXLSX_4_Click(object sender, EventArgs e)
        {
            expdata_4("VA_balance_");
        }

        private void btnToXLSX_5_Click(object sender, EventArgs e)
        {
            expdata_5("Darkhan_3installer_");
        }

        private void btnToXLSX_6_Click(object sender, EventArgs e)
        {
            expdata_6("Darkhan_promo_");
        }

        private void btnToXLSX_7_Click(object sender, EventArgs e)
        {
            expdata_7("Uni_unit_");
        }

        private void btnToXLSX_8_Click(object sender, EventArgs e)
        {
            expdata_8("Promo_usage_");
        }

        private void btnToXLSX_9_Click(object sender, EventArgs e)
        {
            expdata_9("Installerpromo_");
        }

        private void btnToXLSX_10_Click(object sender, EventArgs e)
        {
            expdata_10("Leasing_creation_report");
        }

        private void btnToXLSX_12_Click(object sender, EventArgs e)
        {
            expdata_12("Leasing_report_");
        }
        private void btnToXLSX_13_Click(object sender, EventArgs e)
        {
            expdata_13("Gerduuren_");
        }

        private void btnToXLSX_14_Click(object sender, EventArgs e)
        {
            expdata_14("Gerduuren_bank_");
        }

        private void btnToXLSX_15_Click(object sender, EventArgs e)
        {
            expdata_15("Gerduuren_guilgee_");
        }

        private void btnToXLSX_16_Click(object sender, EventArgs e)
        {
            expdata_16("Gerduuren_INVOICE_");
        }
        private void btnToXLSX_17_Click(object sender, EventArgs e)
        {
            expdata_17("FlashLeasing_");
        }

        private void btnToXLSX_18_Click(object sender, EventArgs e)
        {
            expdata_18("Installerreport_");
        }

        private void btnToXLSX_19_Click(object sender, EventArgs e)
        {
            expdata_19("Mobileapp_report_");
        }

        private void btnToXLSX_20_Click(object sender, EventArgs e)
        {
            expdata_20("Upsell_");
        }
        private void btnToXLSX_21_Click(object sender, EventArgs e)
        {
            expdata_21("InstallerGeneralReport_");
        }

        private void btnToXLSX_22_Click(object sender, EventArgs e)
        {
            expdata_22("InfoUpdatePromo_");
        }
        private void btnToXLSX_23_Click(object sender, EventArgs e)
        {
            expdata_23("MTFEEreport_");
        }

        public void label3_Click(object sender, EventArgs e )
        {
            //MessageBox.Show("Засвартай байна. Орох боломжгүй байна.");
            EditReg editreg = new EditReg(ReportSession.UserId , ReportSession.Fname , ReportSession.Lname , ReportSession.Mobilenumber , ReportSession.Email , ReportSession.Division);
            editreg.Show();
        }
    }
}