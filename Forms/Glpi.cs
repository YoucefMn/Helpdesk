using System.Diagnostics;
using System.Text;

namespace Helpdesk;
public partial class FGlpi : Form
{
    private System.Windows.Forms.Timer debounceTimer;
    List<string> listOriginal = new List<string>();
    List<string> listNew = new List<string>();
    public FGlpi()
    {
        InitializeComponent();
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;

        listOriginal = new List<string> {
            "BU Climatisation Centralisée", "BU Climatiseurs et Machines à Laver", "BU Condor Security Systems",
            "BU Cuissons & Transformation Métallique", "BU Polysterene", "BU Refrigérateurs", "BU Transformation Plastique",
            "Direction Commerciale", "Direction de l'Administration Générale", "Direction des Achats",
            "Direction des Ressources Humaines", "Direction des Systèmes d'Information", "Direction Développement International",
            "Direction Finance et Comptabilité", "Direction Générale", "Direction QHSE", "Direction Régionale Est",
            "Direction Régionale Ouest", "OT- MES", "Alver", "Argilor", "Batigec", "Bordj Steel", "Candy",
            "Condor Academy", "Condor Dasan", "Condor Engineering", "Condor Immo", "Condor Logistics",
            "Condor Multimedia", "Condor Tunisie", "Confidance Plast", "Convia, Aima", "Enicab", "FECO",
            "GB Pharma", "Gerbior, Extra", "Hodna Metal", "Hôtel Beni HAMAD", "Khadamaty", "Polyben",
            "Proxima", "Travocovia", "Travoshop", "Zentech"
        };
        cb.Items.AddRange(listOriginal.ToArray());

        listNew = new List<string>();

        // Initialize the debounce timer
        debounceTimer = new System.Windows.Forms.Timer();
        debounceTimer.Interval = 200;
        debounceTimer.Tick += DebounceTimer_Tick;

    }
    public void Exit(object sender, EventArgs e) { Close(); }
    public void Install(object sender, EventArgs e)
    {
        var selected = cb.Text;
        switch (selected)
        {
            case "BU Climatisation Centralisée":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_BUCC.msi"));
                break;
            case "BU Climatiseurs et Machines à Laver":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_BUCML.msi"));
                break;
            case "BU Condor Security Systems":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_BUCSS.msi"));
                break;
            case "BU Cuissons & Transformation Métallique":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_BUCTM.msi"));
                break;
            case "BU Polysterene":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_BUPol.msi"));
                break;
            case "BU Refrigérateurs":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_BURef.msi"));
                break;
            case "BU Transformation Plastique":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_BUTP.msi"));
                break;
            case "Direction de l'Administration Générale":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DAC.msi"));
                break;
            case "Direction des Achats":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_Achats.msi"));
                break;
            case "Direction des Ressources Humaines":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DRH.msi"));
                break;
            case "Direction des Systèmes d'Information":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DSI.msi"));
                break;
            case "Direction Développement International":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DDI.msi"));
                break;
            case "Direction Finance et Comptabilité":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DFC.msi"));
                break;
            case "Direction Générale":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DG.msi"));
                break;
            case "Direction QHSE":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_QHSE.msi"));
                break;
            case "Direction Régionale Est":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DRE.msi"));
                break;
            case "Direction Régionale Ouest":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DRO.msi"));
                break;
            case "Condor Academy":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_ACADEMY.msi"));
                break;
            case "Condor Dasan":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_DASAN.msi"));
                break;
            case "Condor Engineering":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_ENGINEERING.msi"));
                break;
            case "Condor Immo":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_IMMO.msi"));
                break;
            case "Condor Logistics":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_LOGISTICS.msi"));
                break;
            case "Condor Multimedia":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_MULTIMEDIA.msi"));
                break;
            case "Condor Tunisie":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_TUNISIE.msi"));
                break;
            case "Proxima":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_PROXIMA.msi"));
                break;
            case "Travocovia":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_TRAVOCOVIA.msi"));
                break;
            case "Khadamaty":
                Execute(Path.Combine(FLogin.Appsfolder, @"Parc IT\GLPI-Agent-1.10_KHADAMATY.msi"));
                break;
            case "OT- MES":
                Execute(Path.Combine(FLogin.Appsfolder, @"")); //empty
                break;
            case "Alver":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Argilor":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Batigec":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Bordj Steel":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Candy":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Direction Commerciale":
                Execute(Path.Combine(FLogin.Appsfolder, @"")); //empty
                break;
            case "Confidance Plast":
                Execute(Path.Combine(FLogin.Appsfolder, @"")); //empty
                break;
            case "Convia, Aima":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Enicab":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "FECO":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "GB Pharma":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Gerbior, Extra":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Hodna Metal":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Hôtel Beni HAMAD":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Polyben":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Travoshop":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;
            case "Zentech":
                Execute(Path.Combine(FLogin.Appsfolder, @""));
                break;

        }
    }
    public void Install2(object sender, EventArgs e)
    {
        Clipboard.SetText(@"https://itam.condor.dz/Assets/plugins/fusioninventory/");
        var k = Path.Combine(FLogin.Appsfolder, @"glpi64.exe");

        if (File.Exists(k))
        {
            try
            {
                var process1 = new Process { StartInfo = new ProcessStartInfo { FileName = k } };
                process1.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        else { button1.ForeColor = Color.Red; }
    }
    private void DebounceTimer_Tick(object sender, EventArgs e)
    {
        debounceTimer.Stop();
        FilterItems();
    }
    private void FilterItems()
    {
        string inputText = cb.Text.ToLower();

        if (string.IsNullOrEmpty(inputText))
        {
            cb.Items.Clear();
            cb.Items.AddRange(listOriginal.ToArray());
        }
        else
        {
            listNew = listOriginal.Where(item => item.ToLower().Contains(inputText)).ToList();
            cb.Items.Clear();
            cb.Items.AddRange(listNew.ToArray());
        }

        cb.SelectionStart = cb.Text.Length;
        Cursor = Cursors.Default;
        cb.DroppedDown = true;
    }
    private void Execute(string program)
    {
        try
        {
            Command($@"start-process '{program}' /passive -wait ");
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
    private void cb_TextUpdate(object sender, EventArgs e)
    {

        debounceTimer.Stop();
        debounceTimer.Start();
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
