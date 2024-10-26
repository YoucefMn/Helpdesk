using System.Data;
using System.Diagnostics;
using System.Text;
using Microsoft.Win32;

namespace Helpdesk
{
    public partial class FLogin : Form
    {
        public static string Sid = "";
        public static string User = "";
        public static string Appsfolder = "";
        public FLogin()
        {
            InitializeComponent();
            var dragManager = new Drag(this);
            panel1.MouseDown += dragManager.Panel_MouseDown;
            panel1.MouseMove += dragManager.Panel_MouseMove;
            panel1.MouseUp += dragManager.Panel_MouseUp;

        }
        private void Close(object sender, EventArgs e) => Close();
        private void Minimize(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            
        }
        private void BFolder(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }
        private void BDone(object sender, EventArgs e)
        {
            Appsfolder = folderBrowserDialog1.SelectedPath;
            Hide();
            FMain f2 = new();
            f2.FormClosed += (s, args) => Show();
            if (!string.IsNullOrEmpty(Sid)) f2.Show();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            User = comboBox1.Text;
            label4.Visible = false;
            button2.Visible = false;
            Sid = GetUserSID(comboBox1.Text);
            try
            {
                using var _k = Registry.Users.OpenSubKey(Sid);
                if (_k != null) button2.Visible = true;
                else label4.Visible = true;
            }
            catch
            {
                label4.Visible = true;
            }


        }
        private void FLogin_Load(object sender, EventArgs e) =>
            GetUsersList()?.ForEach(user => comboBox1.Items.Add(user));

        private List<string>? GetUsersList() => Directory.GetDirectories(@"C:\Users")?
         .Where(dir => !new HashSet<string> { "Default", "Default User", "Public", "All Users" }.Contains(Path.GetFileName(dir)))
         .Select(dir => Path.GetFileName(dir))
         .ToList();
        private string GetUserSID(string user)
        {
            string x;
            x = Command($@"$userNTAccount = New-Object System.Security.Principal.NTAccount('{user}');
             $userSID = $userNTAccount.Translate([System.Security.Principal.SecurityIdentifier]);
             $userSID.Value");
            if (!string.IsNullOrEmpty(x)) return x;
            return "";
        }
        private string Command(string command)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{command}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            var output = new StringBuilder();

            using var process = Process.Start(psi);
            if (process != null)
            {
                Action<StreamReader> appendOutput = reader => output.AppendLine(reader.ReadToEnd());
                appendOutput(process.StandardOutput);
                appendOutput(process.StandardError);
                process.WaitForExit();
            }


            return output.ToString().Trim();
        }

        private void pictureBox2_DoubleClick(object sender, EventArgs e)
        {
            Appsfolder = folderBrowserDialog1.SelectedPath;
            Hide();
            FMain f2 = new();
            f2.FormClosed += (s, args) => Show();
            f2.Show();
        }
    }
}
