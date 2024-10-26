using System.Diagnostics;
using System.Text;

namespace Helpdesk.Forms;
public partial class FKaspersky : Form
{
    public FKaspersky()
    {
        InitializeComponent();
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;
    }
    public async void Local(object sender, EventArgs e)
    {


        var k = Path.Combine(FLogin.Appsfolder, "KES local.exe");
        var k2 = Path.Combine(FLogin.Appsfolder, "Agent - local.exe");

        if (!File.Exists(k) && !File.Exists(k2))
        {
            button1.ForeColor = Color.Red;
            button1.Enabled = false;
            return;
        }
        try
        {
            this.Visible = false;
            FMain.setkaspersky(false);
            var process1 = new Process
            {
                StartInfo = new ProcessStartInfo { FileName = k }
            };
            var process2 = new Process
            {
                StartInfo = new ProcessStartInfo { FileName = k2, }
            };
            await Task.Run(() =>
            {

                Task.Run(() => MessageBox.Show("Installations de KES en cours...", "Kaspersky", MessageBoxButtons.OK, MessageBoxIcon.Information));
                process1.Start();
                process1.WaitForExit();
                Task.Run(() => MessageBox.Show("KES Installé avec succees \nInstallation de l'agent en cours...", "Kaspersky", MessageBoxButtons.OK, MessageBoxIcon.Information));
                process2.Start();
                process2.WaitForExit();
                Task.Run(() => MessageBox.Show("Agent Installé avec succees", "Kaspersky", MessageBoxButtons.OK, MessageBoxIcon.Information));

                Invoke(() =>
                {
                    FMain.setkaspersky(true);
                    Close();
                });
            });
        }
        catch (Exception)
        {
            throw;
        }

    }

    public async void Cloud(object sender, EventArgs e)
    {
        var k = Path.Combine(FLogin.Appsfolder, "KES local.exe");
        var k2 = Path.Combine(FLogin.Appsfolder, "KesCloud.exe");

        if (!File.Exists(k) && !File.Exists(k2))
        {
            button2.ForeColor = Color.Red;
            button2.Enabled = false;
            return;
        }
        try
        {
            FMain.setkaspersky(false);
            this.Hide();
            await Task.Run(() =>
            {
                var process1 = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = k,
                        Arguments = "/s",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                var process2 = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = k2,
                        Arguments = "/s",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };
                Task.Run(() => MessageBox.Show("Installations de KES en cours...", "Kaspersky", MessageBoxButtons.OK, MessageBoxIcon.Information));
                process1.Start();
                process1.WaitForExit();
                process2.Start();
                process2.WaitForExit();
                Close();
                FMain.setkaspersky(true);
            });
        }
        catch (Exception)
        {
            throw;
        }

    }
    public void Exit(object sender, EventArgs e)
    {
        FMain.setkaspersky(true);
        Close();

    }
    public string Command(string command)
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

 
}
