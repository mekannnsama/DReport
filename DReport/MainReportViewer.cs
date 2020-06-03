using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DReportUtil.DBControl;
using DReportUtil.Utils;

namespace DReport
{
    public partial class MainReportViewer : XtraForm
    {
        DBConnecter dbconn = new DBConnecter();
        private string dbres = string.Empty;
        public MainReportViewer()
        {
            InitializeComponent();
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

        DataTable dt_rpt2 = new DataTable();

        private void rpt1()
        {
            try
            {
                string beginDate = dtBeginRpt1.Text.Replace("-", "");
                string endDate = dtEndRpt1.Text.Replace("-", "");
                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt1(beginDate, endDate));
                    dt_rpt1 = dt;
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

        private void rpt2()
         {
            try
            {
                string beginDate = dtBeginRpt2.Text.Replace("-", "");
                string endDate = dtEndRpt2.Text.Replace("-", "");
                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt2(beginDate, endDate));
                    dt_rpt2 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }

            }
            catch(Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT2");
                XtraMessageBox.Show(ex.Message);
            }
         }

        internal void show()
        {
            throw new NotImplementedException();
        }

        private void rpt3()
        {
            try
            {
                string beginDate = dtBeginRpt3.Text.Replace("-", "");
                string endDate = dtEndRpt3.Text.Replace("-", "");
                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt3(beginDate, endDate));
                    dt_rpt3 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch(Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT3");
                XtraMessageBox.Show(ex.Message);
            }
        }

        DataTable dt_rpt6 = new DataTable();
        private object dt_rpt3;
        private DataTable dt_rpt4;
        private object dt_expdata1;
        private DataTable dt_rpt1;
        private DataTable dt_rpt7;
        private DataTable dt_rpt8;
        private DataTable dt_rpt9;
        private object dt_rpt10;
        private object dt_rpt5;
        private object dt_rpt12;
        private object dt_rpt13;
        private object dt_rpt14;
        private DataTable dt_rpt15;
        private DataTable dt_rpt16;

        private void rpt6()
        {
            try
            {
                string beginDate = dtBeginRpt6.Text.Replace("-", "");
                string endDate = dtEndRpt6.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt6(beginDate, endDate));
                    dt_rpt6 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch(Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT6");
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void rpt4()
        {
            try
            {
                string beginDate = dtBeginRpt4.Text.Replace("-", "");
                string endDate = dtEndRpt4.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt4(beginDate, endDate));
                    dt_rpt4 = dt;
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
        
        private void rpt7()
        {
            try
            {
                string beginDate = dtBeginRpt7.Text.Replace("-", "");
                string endDate = dtEndRpt7.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt7(beginDate, endDate));
                    dt_rpt7 = dt;
                }
                else
                {
                    XtraMessageBox.Show(dbres);
                }
            }
            catch (Exception ex)
            {
                exceptionManager.ManageException(ex, "RPT7");
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void rpt8()
        {
            try
            {
                string beginDate = dtBeginRpt8.Text.Replace("-", "");
                string endDate = dtEndRpt8.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt8(beginDate, endDate));
                    dt_rpt8 = dt;
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

        private void rpt9()
        {
            try
            {
                string beginDate = dtBeginRpt9.Text.Replace("-", "");
                string endDate = dtEndRpt9.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt9(beginDate, endDate));
                    dt_rpt9 = dt;
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

        private void rpt10()
        {
            try
            {
                string beginDate = dtBeginRpt10.Text.Replace("-", "");
                string endDate = dtEndRpt10.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt10(beginDate, endDate));
                    dt_rpt10 = dt;
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

        private void rpt5()
        {
            try
            {
                string beginDate = dtBeginRpt5.Text.Replace("-", "");
                string endDate = dtEndRpt5.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt5(beginDate, endDate));
                    dt_rpt5 = dt;
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
        private void rpt12()
        {
            try
            {
                string beginDate = dtBeginRpt12.Text.Replace("-", "");
                string endDate = dtEndRpt12.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt12(beginDate, endDate));
                    dt_rpt12 = dt;
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
        private void rpt13()
        {
            try
            {
                string beginDate = dtBeginRpt13.Text.Replace("-", "");
                string endDate = dtEndRpt13.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt13(beginDate, endDate));
                    dt_rpt13 = dt;
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

        private void rpt14()
        {
            try
            {
                string beginDate = dtBeginRpt14.Text.Replace("-", "");
                string endDate = dtEndRpt14.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt14(beginDate, endDate));
                    dt_rpt14 = dt;
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
        private void rpt15()
        {
            try
            {
                string beginDate = dtBeginRpt15.Text.Replace("-", "");
                string endDate = dtEndRpt15.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt15(beginDate, endDate));
                    dt_rpt15 = dt;
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

        private void rpt16()
        {
            try
            {
                string beginDate = dtBeginRpt16.Text.Replace("-", "");
                string endDate = dtEndRpt16.Text.Replace("-", "");

                if (dbconn.idbCheck(out dbres))
                {
                    DataTable dt = dbconn.getTable(DBQry.rpt16(beginDate, endDate));
                    dt_rpt16 = dt;
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

        private void expdata1(string filename)
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
                gridControl1.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata2(string filename)
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
                gridControl2.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata3(string filename)
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
                gridControl7.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata4(string filename)
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
                gridControl8.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata6(string filename)
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
                gridControl10.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata7(string filename)
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
                gridControl11.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata8(string filename)
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
                gridControl12.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata9(string filename)
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
                gridControl13.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata10(string filename)
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
                gridControl5.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void expdata5(string filename)
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
                gridControl9.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata12(string filename)
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
                gridControl14.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata13(string filename)
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
                gridControl15.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata14(string filename)
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
                gridControl16.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata15(string filename)
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
                gridControl17.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void expdata16(string filename)
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
                gridControl18.ExportToXlsx(fname);
                XtraMessageBox.Show("Амжилттай хуулж дууслаа.", "Өгөгдөл хуулах", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(logfolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void btnSearchRpt1_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt1))
            {
                frm.ShowDialog(this);
            }
            gridControl1.DataSource = dt_rpt1;
        }

        private void btnSearchRpt2_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt2))
            {
                frm.ShowDialog(this);
            }
            gridControl2.DataSource = dt_rpt2;
        }

        private void btnSearchRpt6_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt6))
            {
                frm.ShowDialog(this);
            }
            gridControl10.DataSource = dt_rpt6;
        }

        private void btnSearchRpt3_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt3))
            {
                frm.ShowDialog(this);
            }
            gridControl7.DataSource = dt_rpt3;
        }

        private void btnSearchRpt4_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt4))
            {
                frm.ShowDialog(this);
            }
            gridControl8.DataSource = dt_rpt4;
        }


        private void btnSearchRpt7_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt7))
            {
                frm.ShowDialog(this);
            }
            gridControl11.DataSource = dt_rpt7;
        }

        private void btnSearchRpt8_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt8))
            {
                frm.ShowDialog(this);
            }
            gridControl12.DataSource = dt_rpt8;
        }

        private void btnSearchRpt9_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt9))
            {
                frm.ShowDialog(this);
            }
            gridControl13.DataSource = dt_rpt9;
        }

        private void btnSearchRpt10_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt10))
            {
                frm.ShowDialog(this);
            }
            gridControl5.DataSource = dt_rpt10;
        }


        private void btnSearchRpt11_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt5))
            {
                frm.ShowDialog(this);
            }
            gridControl9.DataSource = dt_rpt5;
        }
        private void btnSearchRpt12_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt12))
            {
                frm.ShowDialog(this);
            }
            gridControl14.DataSource = dt_rpt12;
        }
        private void btnSearchRpt13_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt13))
            {
                frm.ShowDialog(this);
            }
            gridControl15.DataSource = dt_rpt13;
        }
        private void btnSearchRpt14_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt14))
            {
                frm.ShowDialog(this);
            }
            gridControl16.DataSource = dt_rpt14;
        }
        private void btnSearchRpt15_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt15))
            {
                frm.ShowDialog(this);
            }
            gridControl17.DataSource = dt_rpt15;
        }
        private void btnSearchRpt16_Click(object sender, EventArgs e)
        {
            using (frmWaitForm frm = new frmWaitForm(rpt16))
            {
                frm.ShowDialog(this);
            }
            gridControl18.DataSource = dt_rpt16;
        }

        private void btnToXLSX2_Click(object sender, EventArgs e)
        {
            expdata2("rpt2");
        }

        private void btnToXLSX3_Click(object sender, EventArgs e)
        {
            expdata3("rpt3");
        }

        private void btnToXLSX1_Click(object sender, EventArgs e)
        {
            expdata1("rpt1");
        }

        private void btnToXLSX4_Click(object sender, EventArgs e)
        {
            expdata4("rpt4");
        }

        private void btnToXLSX6_Click(object sender, EventArgs e)
        {
            expdata6("rpt6");
        }

        private void btnToXLSX7_Click(object sender, EventArgs e)
        {
            expdata7("rpt7");
        }

        private void btnToXLSX8_Click(object sender, EventArgs e)
        {
            expdata8("rpt8");
        }

        private void btnToXLSX9_Click(object sender, EventArgs e)
        {
            expdata9("rpt9");
        }

        private void btnToXLSX10_Click(object sender, EventArgs e)
        {
            expdata10("rpt10");
        }


        private void btnToXLSX11_Click(object sender, EventArgs e)
        {
            expdata5("rpt5");
        }

        private void btnToXLSX12_Click(object sender, EventArgs e)
        {
            expdata12("rpt12");
        }

        private void btnToXLSX13_Click(object sender, EventArgs e)
        {
            expdata13("rpt13");
        }

        private void btnToXLSX14_Click(object sender, EventArgs e)
        {
            expdata14("rpt14");
        }

        private void btnToXLSX15_Click(object sender, EventArgs e)
        {
            expdata15("rpt15");
        }

        private void btnToXLSX16_Click(object sender, EventArgs e)
        {
            expdata16("rpt16");
        }
    }
}
