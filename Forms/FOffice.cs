using System.Diagnostics;
using System.Text;

namespace Helpdesk.Forms;
public partial class FOffice : Form
{
    public FOffice()
    {
        InitializeComponent();
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;
    }

    public async void Pro(object sender, EventArgs e)
    {


        var office = Path.Combine(FLogin.Appsfolder, @"Office_Professional_Plus_2016_64Bit\setup.exe");
        var adminFilePath = Path.Combine(FLogin.Appsfolder, @"Office_Professional_Plus_2016_64Bit\config.MSP");

        if (!File.Exists(office))
        {
            button1.ForeColor = Color.Red;
            button1.Enabled = false;
            return;
        }
        try
        {
            FMain.setoffice(false);
            Hide();
            var p = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = office,
                    Arguments = $"/adminfile {adminFilePath}",
                }
            };
            Task.Run(() => { MessageBox.Show("Installations office en cours", "OFFICE 2016 PRO", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            await Task.Run(async () =>
            {
                p.Start();

                while (!p.HasExited)
                {
                    for (int i = 0; i <= 100; i++)
                    {
                        await Task.Delay(100);

                        if (p.HasExited)
                        {
                            break;
                        }
                    }
                }

                MessageBox.Show("office installé avec succeés", "OFFICE 2016 PRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
                FMain.setoffice(true);
            });

        }
        catch (Exception)
        {

            throw;
        }

    }
    public async void _365(object sender, EventArgs e)
    {
        var s = Path.Combine(FLogin.Appsfolder, "2016");
        if (Directory.Exists(s))
        {
            FMain.setoffice(false);
            Hide();
            await Task.Run(() => CopyFolder(s, @"C:\2016"));
            await Task.Run(() => Command($@"start-process -FilePath 'C:\2016\setup.exe' -ArgumentList '/configure C:\2016\config64.xml' "));
            Close();
            FMain.setoffice(true);
        }
        else
        {
            button2.ForeColor = Color.Red;
        }
    }
    public void Exit(object sender, EventArgs e) => Hide();
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
    private void CopyFolder(string source, string dest)
    {

        Directory.CreateDirectory(dest);
        Directory.EnumerateFiles(source).ToList().ForEach(file =>
        {
            File.Copy(file, Path.Combine(dest, Path.GetFileName(file)), true);
        });
        Directory.EnumerateDirectories(source).ToList().ForEach(subdir =>
        {
            CopyFolder(subdir, Path.Combine(dest, Path.GetFileName(subdir)));
        });
    }


}
