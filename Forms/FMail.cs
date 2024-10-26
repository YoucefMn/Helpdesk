using System.Diagnostics;
using System.Text;

namespace Helpdesk.Forms;
public partial class FMail : Form
{
    public FMail()
    {
        InitializeComponent();
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;
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
    public string TargetPc;
    private void button1_Click(object sender, EventArgs e)
    {
        var TargetPath = $@"\\{TargetPc}\c$\Mail.txt";
        //var TargetPath = $@"c:\Mail.txt";
        try
        {
            File.WriteAllText(TargetPath, textBox1.Text);
            MessageBox.Show($"Email enregisté dans : {TargetPath}");
            Close();
        }
        catch (Exception z)
        {
            MessageBox.Show(z.Message);
            Close();
        }

    }


 
}
