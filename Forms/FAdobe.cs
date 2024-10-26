using System.Diagnostics;
using System.Text;

namespace Helpdesk.Forms;
public partial class FAdobe : Form
{
    public FAdobe()
    {
InitializeComponent();
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;
    }

    public void Reaader(object sender, EventArgs e)
    {
        var k = Path.Combine(FLogin.Appsfolder, @"AdbeRdr930_fr_FR.exe");
        if (File.Exists(k))
        {
            Task.Run(() =>
            {
                Command($@"start-process '{k}' -wait");
                Close();
            });
        }
        else
        {
            button1.ForeColor = Color.Red;
        }

    }
    public async void XI(object sender, EventArgs e)
    {
        var x = Path.Combine(FLogin.Appsfolder, @"Adobe Acrobat XI Pro\Adobe Acrobat XI\Setup.exe");
        var p = Path.Combine(FLogin.Appsfolder, @"Adobe Acrobat XI Pro\amtemu.v0.9.1-painter.exe");
        if (File.Exists(x))
        {
            Clipboard.SetText(@"C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\amtlib.dll");
            await Task.Run(() => Command($@"start-process -FilePath '{x}' -wait "));
            await Task.Run(() => Command($@"start-process -FilePath '{p}' -wait "));
        }
        else
        {
            button2.ForeColor = Color.Red;
        }
    }
    public void Exit(object sender, EventArgs e) => Close();
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

    private void FAdobe_Load(object sender, EventArgs e)
    {
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;
    }
}
