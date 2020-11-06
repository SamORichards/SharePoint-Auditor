using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1 {

	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window {
		public MainWindow() {
			InitializeComponent();
            string[] workingExe = System.Reflection.Assembly.GetEntryAssembly().Location.Split('\\');
            string workingFolder = "";
            for (int i = 0; i < workingExe.Length - 1; i++) {
                workingFolder += workingExe[i] + (i == workingExe.Length - 2 ? "" : "/");
            }
            if (File.Exists( workingFolder + "/UserInfo.sav")) {
                using (StreamReader reader = new StreamReader(workingFolder + "/UserInfo.sav")) {
                    TenantTextBox.Text = reader.ReadLine();
                    AdminEmailTextBox.Text = reader.ReadLine();
                }
            }
        }
        bool pathSelected = false;
        private async void LocationButton_Click(object sender, RoutedEventArgs e) {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            DialogResult result = folderDlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK) {
                ChooseLocationLabel.Text = folderDlg.SelectedPath;
                pathSelected = true;
            } else {
                ChooseLocationLabel.Text = "Operation Cancelled. Choose Again.";
                pathSelected = false;
            }
        }

        private void RunButton_Click(object sender, RoutedEventArgs e) {
            string tenant = TenantTextBox.Text.ToString();
            string email = AdminEmailTextBox.Text.ToString();
            bool logs = (bool)LogsCheck.IsChecked;
            int throttle = int.Parse(ThrottleLimitValueTextBox.Text);
            if (tenant != "" && email != "" && !logs.Equals(null) && !throttle.Equals(null) && pathSelected) {
                string reportPath = ChooseLocationLabel.Text.ToString();
                var proc1 = new ProcessStartInfo();
                string[] workingExe = System.Reflection.Assembly.GetEntryAssembly().Location.Split('\\');
                string workingFolder = "";
                for (int i = 0; i < workingExe.Length - 1; i++) {
                    workingFolder += workingExe[i] + (i == workingExe.Length - 2 ? "" : "/");
                }
                reportPath = reportPath.Replace('\\', '/');
                SaveInfo();
                var command = "cd " + workingFolder + " && powershell -ExecutionPolicy Bypass -Command \".\\SharePointPermissionsAuditor.ps1 " + tenant + " " + email +  " " + (logs ? "1" : "0") + " " + throttle + " " + reportPath + "\" -Verb RunAs";
                proc1.UseShellExecute = true;
                proc1.WorkingDirectory = System.Reflection.Assembly.GetEntryAssembly().Location;
                proc1.FileName = @"C:\Windows\System32\cmd.exe";
                proc1.Verb = "runas";
                proc1.Arguments = "/c " + command;
                proc1.WindowStyle = ProcessWindowStyle.Maximized;
                Process.Start(proc1);
            }
        }

        void SaveInfo() {
            string[] workingExe = System.Reflection.Assembly.GetEntryAssembly().Location.Split('\\');
            string workingFolder = "";
            for (int i = 0; i < workingExe.Length - 1; i++) {
                workingFolder += workingExe[i] + (i == workingExe.Length - 2 ? "" : "/");
            }
            using (StreamWriter Save = new StreamWriter(workingFolder + "/UserInfo.sav")) {
                Save.WriteLine(TenantTextBox.Text);
                Save.WriteLine(AdminEmailTextBox.Text);
            }
        }
    }
}
