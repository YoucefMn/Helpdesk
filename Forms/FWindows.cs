using System.Diagnostics;

namespace Helpdesk.Forms;
public partial class FWindows : Form
{
    public FWindows()
    {
        InitializeComponent();
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;
    }

    public void Pro(object sender, EventArgs e)
    {

        var p = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = @"slmgr.vbs //b /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX",
            CreateNoWindow = true
        };
        var pp = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = @"slmgr.vbs //b /skms srvkms001.condor.com",
            CreateNoWindow = true
        };
        using var process1 = Process.Start(p);
        using var process2 = Process.Start(pp);
        Close();
    }
    public void Enterprise(object sender, EventArgs e)
    {
        var p = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = @"slmgr.vbs //b /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43",
            CreateNoWindow = true
        };
        var pp = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = @"slmgr.vbs //b /skms srvkms001.condor.com",
            CreateNoWindow = true
        };
        using var process1 = Process.Start(p);
        using var process2 = Process.Start(pp);
        Close();
    }
    public void Exit(object sender, EventArgs e) => Close();


}
