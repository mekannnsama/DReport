using DreportLauncher.LauncherUtils;
using Microsoft.Extensions.Configuration;
using Renci.SshNet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Windows.Forms;
using static DreportLauncher.DreportModels;

namespace DreportLauncher
{
    public partial class Form1 : Form
    {
        BackgroundWorker worker = new BackgroundWorker();

        public Form1()
        {
            InitializeComponent();
            worker.WorkerSupportsCancellation = true;
            worker.WorkerReportsProgress = true;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerAsync();
        }
        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try 
            { 
                DownloadFiles("", "DreportFiles.json");
                var path = AppDomain.CurrentDomain.BaseDirectory;
                var builder = new ConfigurationBuilder().AddJsonFile(path + "DreportFiles.json", true, true).AddEnvironmentVariables();
                IConfiguration ConfigRoot = builder.Build();

                DreportConfigs appConfig = new DreportConfigs();
                appConfig.Appversion = ConfigRoot.GetSection("Appversion").Value;
                appConfig.Files = ConfigRoot.GetSection("Files").Get<List<FileNames>>();
                
                if (Constants.VERSION != appConfig.Appversion)
                {
                    MessageBox.Show("Update хийнэ үү!");
                    int loopcount = appConfig.Files.Count();

                    for (int i = 0; i < loopcount; i++)
                    {
                        string DownloadFile = appConfig.Files[i].file_name;
                        DownloadFiles("Dreportinstall/", DownloadFile);
                        
                        List<string> Userstateinfo = new List<string>();
                        Userstateinfo.Add(DownloadFile);
                        Userstateinfo.Add((i + 1).ToString() + " of " + loopcount.ToString() + " files are downloaded...");
                        worker.ReportProgress((int)((i + 1) * 100 / loopcount) , (List<string>)(Userstateinfo));
                    }
                }
                Process.Start(AppDomain.CurrentDomain.BaseDirectory + "DReport.exe");
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Application.Exit();
            }
        }
        public void Tsebo()
        {
            
        }

        public void DownloadFiles(string address, string fileName)
        {
            var xmlFolder = AppDomain.CurrentDomain.BaseDirectory;
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(xmlFolder);
            FileSystemAccessRule fsar = new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessControlType.Allow);
            DirectorySecurity ds = null;

            if (!di.Exists)
            {
                Directory.CreateDirectory(xmlFolder);
            }
            ds = di.GetAccessControl();
            ds.AddAccessRule(fsar);


            SftpClient sftpClient = new SftpClient(Constants.FTP_SERVER, Convert.ToInt32(Constants.FTP_SERVER_PORT), Constants.FTP_SERVER_NAME, Constants.FTP_SERVER_PASS);
            try
            {
                sftpClient.Connect();
                using (Stream rmFile = File.OpenWrite(xmlFolder + fileName))
                {
                    if ("DreportFiles.json" != fileName)
                    {
                        /*                        string server_file_hash = GetMD5HashFromFile(Constants.FTP_FILE_PATH + address + fileName + fileName);
                                                string local_file_hash = GetMD5HashFromFile(AppDomain.CurrentDomain.BaseDirectory + fileName);
                                                if (local_file_hash != server_file_hash)
                                                {*/
                        sftpClient.DownloadFile(Constants.FTP_FILE_PATH + address + fileName, rmFile);
                        /*                      }*/
                    }
                    else
                    {
                        sftpClient.DownloadFile(Constants.FTP_FILE_PATH + address + fileName, rmFile);
                    }
                }
                sftpClient.Disconnect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Application.Exit();
            }   
            ;
           
        }
        private static string GetMD5HashFromFile(string fileName)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(fileName))
                {
                    return BitConverter.ToString(md5.ComputeHash(stream)).Replace("-", string.Empty);
                }
            }
        }
        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try 
            { 
                Downloading.Value = e.ProgressPercentage;
                IList userinfo = (IList)e.UserState;
                label1.Text = "Downloading files from server , please wait ...";
                label2.Text = "Now downloading " + userinfo[0] + " ..." ;
                label3.Text = userinfo[1].ToString() ;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                Application.Exit();
            }
        }
    }
}
