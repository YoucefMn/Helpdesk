using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Helpdesk.Forms;
using Helpdesk.Properties;
using Microsoft.Win32;

namespace Helpdesk;

public partial class FMain : Form
{
    private static FMain _instance;

    public FMain()
    {
        InitializeComponent();
        _instance = this;
        var dragManager = new Drag(this);
        panel1.MouseDown += dragManager.Panel_MouseDown;
        panel1.MouseMove += dragManager.Panel_MouseMove;
        panel1.MouseUp += dragManager.Panel_MouseUp;
    }


    private void Form1_Load(object sender, EventArgs e)
    {

        Task2(); // check installed apps
        Task3(); // check language
        Task4();// check windows
        Task5(); // check rdp
        Task6(); // check parefeu
        Task7(); // check services

    }
    private void Close(object sender, EventArgs e) => Close();
    private void Minimize(object sender, EventArgs e) => WindowState = FormWindowState.Minimized;
    // Office
    public void Word(object sender, EventArgs e) => CopyFile(Programmes, Desktop, "Wor*.lnk", p_word);
    public void Excel(object sender, EventArgs e) => CopyFile(Programmes, Desktop, "Exc*.lnk", p_excel);
    public void Outlook(object sender, EventArgs e) => CopyFile(Programmes, Desktop, "Outl*.lnk", p_outlook);
    public void Access(object sender, EventArgs e) => CopyFile(Programmes, Desktop, "Acces*.lnk", p_access);
    public void Teams(object sender, EventArgs e) => CopyFile(Programmes, Desktop, @"Microsoft Dynamics NAV 2009 R2 Classic with Microsoft SQL Server.lnk", p_teams);

    // Desktop
    public void Thispc(object sender, EventArgs e) => DesktopReg(@"{20D04FE0-3AEA-1069-A2D8-08002B30309D}", p_thispc);
    public void Userfolder(object sender, EventArgs e) => DesktopReg(@"{59031a47-3f72-44a7-89c5-5595fe6b30ee}", p_userfolder);
    public void Pannel(object sender, EventArgs e) => DesktopReg(@"{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}", p_pannel);
    public void Network(object sender, EventArgs e) => DesktopReg(@"{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", p_network);
    public void Bin(object sender, EventArgs e) => DesktopReg(@"{645FF040-5081-101B-9F08-00AA002F954E}", p_bin);

    // Services
    public void Windows(object sender, EventArgs e)
    {
        var k = new FWindows();
        k.Location = MousePosition;
        k.FormClosed += (s, args) => { Task4(); };
        k.ShowDialog();
    }
    public void Parefeu(object sender, EventArgs e)
    {
        p_parefeu.Image = null;
        Task.Run(() =>
        {
            string r = Command(@"netsh advfirewall set allprofiles state off");
            p_parefeu.Image = r switch
            {
                var res when res.ToLower().Contains("ok") => Resources.act,
                var res when res.ToLower().Contains("accord") => Resources.act,
                _ => Resources.error
            };
        });
    }
    public void Rdp(object sender, EventArgs e)
    {
        p_rdp.Image = null;
        Task.Run(() =>
        {

            string r = @"SYSTEM\CurrentControlSet\Control\Terminal Server";
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(r, true);
                if (key != null)
                {
                    key.SetValue("fDenyTSConnections", 0, RegistryValueKind.DWord);
                    var users = new HashSet<string> { @"condor.com\t-admin", @"condor.com\helpdesk-ad" };
                    users.ToList()
                        .ForEach(user =>
                            Command($@"Add-LocalGroupMember -Group 'Utilisateurs du Bureau à distance' -Member {user} ")
                        );
                    p_rdp.Image = Resources.act;
                }
                else p_rdp.Image = Resources.error;
            }
            catch
            {
                p_rdp.Image = Resources.error;
            }
        });
    }
    public async void Services(object sender, EventArgs e)
    {
        p_services.Image = null;
        var t1 = Task.Run(() => Command(@"Set-Service RemoteRegistry -StartupType Manual; Start-Service RemoteRegistry"));
        var t2 = Task.Run(() => Command(@"$service1 = (Get-Service -Name UserDataSvc_*).Name; Set-ItemProperty -Path ""HKLM:\SYSTEM\CurrentControlSet\Services\$service1"" -Name Start -Value 3;Start-Service -Name $service1"));
        var t3 = Task.Run(() => Command(@"Set-Service WinRM -StartupType Automatic; Start-Service WinRM"));
        var t4 = Task.Run(() => Command(@"winrm qc -force;
                Enable-PSRemoting -Force;
                Set-Item -Path WSMan:\localhost\Service\AllowUnencrypted -Value $true;
                Set-Item -Path WSMan:\localhost\Service\Auth\Basic -Value $true;
                Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value ""*"" -Force;
                "));

        var results = await Task.WhenAll(t1, t2, t3, t4);
        if (string.IsNullOrEmpty(results[0]) && string.IsNullOrEmpty(results[1]) && string.IsNullOrEmpty(results[2]))
        {
            p_services.Image = Resources.act;
        }
        else
        {
            p_services.Image = Resources.error;
        }
    }

    // Configuration
    public void Sapconfig(object sender, EventArgs e)
    {
        p_sapconfig.Image = null;
        var s = Path.Combine(FLogin.Appsfolder, "SAP");
        var d = Path.Combine(Roaming, "SAP");

        if (Directory.Exists(s))
        {
            CopyFolder(s, d);
            try
            {
                using var a = new StreamWriter(@"C:\Windows\System32\drivers\etc\services", append: true);
                a.WriteLine(@"sapmsCPH 3600/tcp");
                p_sapconfig.Image = Resources.done;
            }
            catch { p_sapconfig.Image = Resources.error; }
        }
        else p_sapconfig.Image = Resources.error;
    }
    //public void Mbamconfig(object sender, EventArgs e)
    //{
    //    p_mbam.Image = null;

    //    var k = @"SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement";
    //    try
    //    {
    //        using var key = Registry.LocalMachine.OpenSubKey(k, true);
    //        bt_mbamconfig.Enabled = false;
    //        key.SetValue("ClientWakeupFrequency", 1, RegistryValueKind.DWord);
    //        key.SetValue("StatusReportingFrequency", 1, RegistryValueKind.DWord);
    //        key.Close();
    //        var p = new Process
    //        {
    //            StartInfo = new ProcessStartInfo
    //            {
    //                FileName = "powershell.exe",
    //                Arguments = "gpupdate /force",
    //                CreateNoWindow = true,
    //                UseShellExecute = false
    //            }
    //        };
    //        Task.Run(() =>
    //        {
    //            p.Start();
    //            p.WaitForExit();

    //            Invoke((Delegate)(() =>
    //            {
    //                using var key2 = Registry.LocalMachine.OpenSubKey(k, true);
    //                key2.SetValue("ClientWakeupFrequency", 1, RegistryValueKind.DWord);
    //                key2.SetValue("StatusReportingFrequency", 1, RegistryValueKind.DWord);
    //                key2.Close();
    //                p_mbam.Image = Resources.done;
    //                bt_mbamconfig.Enabled = true;
    //            }));

    //        });

    //    }
    //    catch
    //    {
    //        bt_mbamconfig.Enabled = true;
    //        p_mbam.Image = Resources.error;
    //    }
    //}
    public async void Mbamconfig(object sender, EventArgs e)
    {
        p_mbam.Image = null;
        p_mbam.Visible = false;


        var k = @"SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement";
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(k, true);
            bt_mbamconfig.Enabled = false;
            key.SetValue("ClientWakeupFrequency", 1, RegistryValueKind.DWord);
            key.SetValue("StatusReportingFrequency", 1, RegistryValueKind.DWord);
            key.Close();

            progressBar0.Minimum = 0;
            progressBar0.Maximum = 100;
            progressBar0.Value = 0;
            progressBar0.Visible = true;

            var p = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "gpupdate /force",
                    CreateNoWindow = true,
                    UseShellExecute = false
                }
            };
            int totalEstimatedTime = 5000; // 5 seconds

            await Task.Run(async () =>
            {
                p.Start();

                // Start a timer to update the progress bar
                var startTime = DateTime.Now;
                bool taskCompleted = false;

                while (!taskCompleted)
                {
                    await Task.Delay(100); // Update progress every 100 milliseconds
                    var elapsedTime = (DateTime.Now - startTime).TotalMilliseconds;
                    int progress = (int)((elapsedTime / totalEstimatedTime) * 100);
                    if (progress > 100) progress = 100;

                    Invoke((Delegate)(() =>
                    {
                        progressBar0.Value = progress;
                    }));

                    if (p.HasExited)
                    {
                        taskCompleted = true;
                    }
                }

                p.WaitForExit();

                Invoke((Delegate)(() =>
            {
                using var key2 = Registry.LocalMachine.OpenSubKey(k, true);
                key2.SetValue("ClientWakeupFrequency", 1, RegistryValueKind.DWord);
                key2.SetValue("StatusReportingFrequency", 1, RegistryValueKind.DWord);
                key2.Close();
                p_mbam.Image = Resources.done;
                p_mbam.Visible = true;
                bt_mbamconfig.Enabled = true;

                progressBar0.Value = 0;
                progressBar0.Visible = false;
            }));

            });

        }
        catch
        {
            bt_mbamconfig.Enabled = true;
            p_mbam.Visible = true;
            p_mbam.Image = Resources.error;
            progressBar0.Value = 0;
            progressBar0.Visible = false;
        }
    }
    public void Glpisync(object sender, EventArgs e)
    {
        var s = @"C:\Program Files\FusionInventory-Agent\fusioninventory-agent.bat";
        var s2 = @"C:\Program Files\GLPI-Agent\glpi-agent.bat";

        if (File.Exists(s))
        {
            progressBar2.Visible = true;
            bt_glpiconfig.Enabled = false;
            p_glpisync.Image = Resources.done;
            Task.Run(async () =>
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = s,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true // Prevents cmd window from appearing
                    }
                };
                process.Start();

                while (!process.HasExited)
                {
                    for (int i = 0; i <= 100; i++)
                    {
                        if (process.HasExited)
                            break;

                        Invoke(new Action(() => progressBar2.Value = i));
                        await Task.Delay(100);
                    }
                }

                process.WaitForExit();
                Invoke(new Action(() => progressBar2.Visible = false));
            });

        }
        else if (File.Exists(s2))
        {
            progressBar2.Visible = true;
            p_glpisync.Image = Resources.done;
            Task.Run(async () =>
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = s2,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true // Prevents cmd window from appearing
                    }
                };
                process.Start();

                while (!process.HasExited)
                {
                    for (int i = 0; i <= 100; i++)
                    {
                        if (process.HasExited)
                            break;

                        Invoke(new Action(() => progressBar2.Value = i));
                        await Task.Delay(100);
                    }
                }

                process.WaitForExit();
                Invoke(new Action(() => progressBar2.Visible = false));
            });
            bt_glpiconfig.Enabled = true;

        }
        else
        {
            p_glpisync.Image = Resources.error;
        }
    }
    public void Gpupdate(object sender, EventArgs e)
    {
        progressBar3.Visible = true;
        p_gpupdate.Image = Resources.done;
        Task.Run(async () =>
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "gpupdate",
                    Arguments = "/force",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            process.Start();
            bt_gpupdate.Enabled = false;
            while (!process.HasExited)
            {
                for (int i = 0; i <= 100; i++)
                {
                    if (process.HasExited)
                        break;

                    Invoke(new Action(() => progressBar3.Value = i));
                    await Task.Delay(100);
                }
            }

            process.WaitForExit();
            Invoke(new Action(() => bt_gpupdate.Enabled = true));
            Invoke(new Action(() => progressBar3.Visible = false));
        });

    }
    public void Mail(object sender, EventArgs e)
    {
        string m = @"C:\Program Files\Microsoft Office\Office16\MLCFG32.CPL";
        string m2 = @"C:\Program Files\Microsoft Office\root\Office16\MLCFG32.CPL";
        if (File.Exists(m) || File.Exists(m2))
        {
            string filePath = File.Exists(m) ? m : m2;
            try
            {
                var p = new Process { StartInfo = new ProcessStartInfo { FileName = filePath, UseShellExecute = true } };
                p.Start();
                Thread.Sleep(1000);
                SendKeys.SendWait("{TAB}");
                SendKeys.SendWait("{ENTER}");
                SendKeys.SendWait("Test");
                SendKeys.SendWait("{ENTER}");
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{ENTER}");
                SendKeys.SendWait("{DOWN}");
                SendKeys.SendWait("{ENTER}");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        else
        {
            MessageBox.Show("MLCFG32.CPL Not found");
        }

    }

    // Clavier
    public void Arab(object sender, EventArgs e) => Addlaguages("ar-DZ", p_ar);
    public void Frensh(object sender, EventArgs e) => Addlaguages("fr-FR", p_fr);
    public void English(object sender, EventArgs e) => Addlaguages("en-US", p_en);
    public void RArab(object sender, EventArgs e) => Removelanguages("ar-*", p_ar);
    public void RFrensh(object sender, EventArgs e) => Removelanguages("fr-*", p_fr);
    public void REnglish(object sender, EventArgs e) => Removelanguages("en-*", p_en);
    private void Addlaguages(string langue, PictureBox p)
    {
        Task.Run(() =>
        {
            string script = Command($@"
        $cl = Get-WinUserLanguageList
        if (-not ($cl | Where-Object {{$_.LanguageTag -eq '{langue}' }})) {{
            $cl.Add('{langue}')
            Set-WinUserLanguageList $cl -Force
        }} else {{
            Write-output 'error'
        }}");
            if (script.Contains("error")) p.Image = Resources._0;
            else p.Image = Resources._1;
        });
    }
    private void Removelanguages(string langue, PictureBox p)
    {
        Task.Run(() =>
        {
            Command($@"
        $LanguageList = Get-WinUserLanguageList;
       $LanguageList.Remove(($LanguageList | Where-Object LanguageTag -like '{langue}'));
        Set-WinUserLanguageList $LanguageList -Force
        ");
            p.Image = Resources._0;
        });
    }

    #region Chiffrement percentage
    public void Task1(object sender, EventArgs e) // chiffrement percentage
    {
        var x = string.Empty;
        Task.Run(() =>
        {
            x = Command(@"
                $x = manage-bde -status;
                $result = New-Object System.Text.StringBuilder
                foreach ($line in $x)
                {
                    if ($line -match 'Volume\s+([A-Z]):') {
                        $drive = $matches[1]
                    }
                    elseif ($line -match 'Pourcentage chiffré\s*:\s*(.+)'){
                        $null = $result.AppendLine($drive + ': = ' + $matches[1].Trim())
                    }
                }
                $result.ToString().Trim()
                 ");
            Invoke(() => { richTextBox1.Text = x; });
        });
    }
    #endregion
    #region Check installed Apps
    public async void Task2()
    {
        try
        {
            await Task.Run(async () =>
            {
                var results = await Task.WhenAll(
                    Checkapp("Kaspersky.*Endpoint"),
                    Checkapp("office\\s*2016"),
                    Checkapp("Office.*365"),
                    Checkapp("Microsoft.*365"),
                    Checkapp("Glpi"),
                    Checkapp("FusionInventory"),
                    Checkapp("sap"),
                    Checkapp("password sol"),
                    Checkapp("mbam"),
                    Checkapp("Adobe\\s*Reader"),
                    Checkapp("Acrobat.*XI"),
                    Checkapp("Acrobat.*DC"),
                    Checkapp("Dynamics"),
                    Checkapp("Microsoft.*NAV"),
                    Checkapp("chrome"),
                    Checkapp("firefox"),
                    Checkapp("silverlight"),
                    Checkapp("classic shell")
                );

                this.Invoke((MethodInvoker)delegate
                {
                    p_kaspersky.Image = results[0] ? Resources.done : null;

                    // Handle Office 2016 and Office 365
                    if (results[2] || results[3])
                    {
                        p_office.Image = Resources._365;
                    }
                    else if (results[1])
                    {
                        p_office.Image = Resources._2016;
                    }
                    else
                    {
                        p_office.Image = null;
                    }

                    p_glpiinstall.Visible = results[4] ? true : false;
                    p_glpiinstall2.Visible = results[5] ? true : false;
                    p_sapinstall.Image = results[6] ? Resources.done : null;
                    p_lapsinstall.Image = results[7] ? Resources.done : null;
                    p_bitlockerinstall.Image = results[8] ? Resources.done : null;
                    p_reader.Visible = results[9] ? true : false;
                    if (results[10] || results[11])
                    {
                        p_xi.Visible = true;
                    }
                    else
                    {
                        p_xi.Visible = false;
                    }

                    p_navision.Image = results[12] || results[13] ? Resources.done : null;
                    p_chrome.Image = results[14] ? Resources.done : null;
                    p_firefox.Image = results[15] ? Resources.done : null;
                    p_silverlight.Image = results[16] ? Resources.done : null;
                    p_classicshell.Image = results[17] ? Resources.done : null;

                    // Optionally, you can hide the PictureBox if the Image is null
                    p_kaspersky.Visible = p_kaspersky.Image != null;
                    p_office.Visible = p_office.Image != null;
                    //p_glpiinstall.Visible = p_glpiinstall.Image != null;
                    p_sapinstall.Visible = p_sapinstall.Image != null;
                    p_lapsinstall.Visible = p_lapsinstall.Image != null;
                    p_bitlockerinstall.Visible = p_bitlockerinstall.Image != null;
                    p_navision.Visible = p_navision.Image != null;
                    p_chrome.Visible = p_chrome.Image != null;
                    p_firefox.Visible = p_firefox.Image != null;
                    p_silverlight.Visible = p_silverlight.Image != null;
                    p_classicshell.Visible = p_classicshell.Image != null;
                });
            });
        }
        catch (Exception ex)
        {
            // Log the exception or handle it appropriately
            throw;
        }
    }
    #endregion
    #region Check keyboard languages
    public void Task3() // check language
    {
        var x = Command(@"Get-WinUserLanguageList");
        if (x.ToLower().Contains("fr-")) p_fr.Image = Resources._1;
        if (x.ToLower().Contains("ar-")) p_ar.Image = Resources._1;
        if (x.ToLower().Contains("en-")) p_en.Image = Resources._1;
    }
    #endregion
    #region Check windows
    public async void Task4()
    {
        p_windows.Image = null;
        string r = await Task.Run(() => Command(@"(Get-CimInstance -Query 'SELECT PartialProductKey FROM SoftwareLicensingProduct WHERE LicenseStatus = 1').PartialProductKey"));
        if (r.Contains("T83GX")) p_windows.Image = Resources.Pro;
        else if (r.Contains("2YT43")) p_windows.Image = Resources.enterprise;
        else p_windows.Image = Resources.error;
    }


    #endregion
    #region Check rdp
    public void Task5()
    {
        string r = @"SYSTEM\CurrentControlSet\Control\Terminal Server";
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(r, true);
            string x = key!.GetValue("fDenyTSConnections")!.ToString()!;
            if (x == "0") p_rdp.Image = Resources.act;
        }
        catch { }

    }
    #endregion
    #region Check parefeu
    public async void Task6()
    {
        var r = Task.Run(() => Command(@"(Get-NetFirewallProfile -Name Public,Domain,Private).enabled"));
        string e = await r;
        if (!e.ToString().Contains("True")) p_parefeu.Image = Resources.des;
    }
    #endregion
    #region Check services
    public async void Task7()
    {
        var t1 = Task.Run(() => Command(@"(Get-Service -Name UserDataSvc_*).status"));
        var t2 = Task.Run(() => Command(@"(Get-Service -Name UserDataSvc_*).status"));
        var b = string.Join(',', await Task.WhenAll(t1, t2));
        if (!b.Contains("Stopped")) p_services.Image = Resources.act;
    }
    #endregion
    #region Close
    private void FMain_FormClosed(object sender, FormClosedEventArgs e)
    {
        FLogin.Appsfolder = string.Empty;

    }
    #endregion


    //Installation
    //private CancellationTokenSource _cts;
    //public void Kaspersky(object sender, EventArgs e)
    //{
    //        ProgressKas.Value = 0;
    //        ProgressKas.Visible = true;
    //        p_kaspersky.Visible = false;

    //        // Create a cancellation token
    //        _cts = new CancellationTokenSource();
    //        var token = _cts.Token;

    //        // Loop to update the progress bar until the form is closed
    //        var progressTask = Task.Run((Func<Task?>)(async () =>
    //        {
    //            try
    //            {
    //                while (!token.IsCancellationRequested)
    //                {
    //                    for (int i = 0; i <= 100; i++)
    //                    {
    //                        this.ProgressKas.Invoke((Action)(() => this.ProgressKas.Value = i));
    //                        await Task.Delay(20, token); // Respect the cancellation token
    //                    }
    //                }
    //            }
    //            catch (TaskCanceledException)
    //            {
    //                // Task was canceled, do nothing to avoid unhandled exceptions
    //            }
    //        }), token);

    //        // Show the FKaspersky form
    //        var k = new FKaspersky();
    //        k.Location = MousePosition;

    //        k.Show();

    //        // When the FKaspersky form is closed, stop the progress bar loop
    //        k.FormClosed += async (s, args) =>
    //        {
    //            // Cancel the progress task
    //            _cts.Cancel();

    //            // Wait for the task to complete
    //            try
    //            {
    //                await progressTask;
    //            }
    //            catch (TaskCanceledException)
    //            {
    //                // Ignore cancellation
    //            }

    //            // Reset visibility and enable other controls
    //            ProgressKas.Visible = false;
    //            p_kaspersky.Visible = true;
    //            bt_kaspersky.Enabled = true;
    //            Task2();
    //        };
    //    }
    public void Kaspersky(object sender, EventArgs e)
    {
        bt_kaspersky.Enabled = false;
        var k = new FKaspersky();
        k.Location = MousePosition;
        k.ShowDialog();
        k.FormClosed += (s, args) =>
        {
            //p_kaspersky.Visible = true;
            //bt_kaspersky.Enabled = true;
            Task2();
        };
    }


    public void Office(object sender, EventArgs e)
    {
        var k = new FOffice();
        k.Location = MousePosition;
        k.FormClosed += (s, args) => bt_office.Enabled = true; Task2();
        k.ShowDialog();
    }
    public void Glpiagent(object sender, EventArgs e)
    {
        var k = new FGlpi();
        k.StartPosition = FormStartPosition.CenterScreen;
        k.FormClosed += (s, args) => Task2();
        k.ShowDialog();
        //var g = Path.Combine(FLogin.Appsfolder, @"glpi64.exe");
        //if (File.Exists(g))
        //{
        //    var image = Resources.glpi;
        //    string tempPath = Path.Combine(Path.GetTempPath(), "tempImage.png");
        //    image.Save(tempPath);
        //    try { Process.Start(new ProcessStartInfo(tempPath) { UseShellExecute = true }); } catch { }
        //    Clipboard.SetText("https://itam.condor.dz/Assets/plugins/fusioninventory/");
        //    await Task.Run(() => Command($@"start-process '{g}' -wait"));
        //    Task2();
        //}
        //else p_glpiinstall.Image = Resources.error;

    }
    public async void Laps(object sender, EventArgs e)
    {
        var g = Path.Combine(FLogin.Appsfolder, @"LAPSx64.msi");
        if (!File.Exists(g))
        {
            var p = new PictureBox();
            p.Parent = bt_laps;
            p.Image = Resources._0;
            p.Height = bt_laps.Height - 10;
            p.Width = bt_laps.Width - 10;
            p.Location = new Point(5, 5);
            p.BackColor = Color.Transparent;
            p.SizeMode = PictureBoxSizeMode.Zoom;
            p.BringToFront();
            return;
        }
        try
        {
            var p = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "msiexec.exe",
                    Arguments = $"/i \"{g}\" /qn",
                }
            };
            await Task.Run(() =>
            {
                Task.Run(() => { MessageBox.Show("Installation Laps est en cours ...", "LAPS", MessageBoxButtons.OK, MessageBoxIcon.Information); });
                p.Start();
                p.WaitForExit();
                Task.Run(() => { MessageBox.Show("Installation Laps est terminée", "LAPS", MessageBoxButtons.OK, MessageBoxIcon.Information); });

            });
            Task2();
        }
        catch (Exception)
        {
            throw;
        }
    }
    public async void Mbaminstall(object sender, EventArgs e)
    {
        var p = Path.Combine(FLogin.Appsfolder, @"MbamClientSetup-2.5.1100.0x64.msi");
        var c = Path.Combine(FLogin.Appsfolder, @"MBAM2.5_Client_x64_KB4586232.msp");
        var a = Path.Combine(FLogin.Appsfolder, @"MBAM2.5_Client_x64_KB4586232.exe");
        if (!File.Exists(p) && !File.Exists(c) && !File.Exists(a))
        {
            var pp = new PictureBox();
            pp.Parent = bt_mbam;
            pp.Image = Resources._0;
            pp.Height = bt_mbam.Height - 10;
            pp.Width = bt_mbam.Width - 10;
            pp.Location = new Point(5, 5);
            pp.BackColor = Color.Transparent;
            pp.SizeMode = PictureBoxSizeMode.Zoom;
            pp.BringToFront();
            return;
        }
        var package = new Process { StartInfo = new ProcessStartInfo { FileName = "msiexec.exe", Arguments = $"/i \"{p}\"/passive" } };
        var correctif = new Process { StartInfo = new ProcessStartInfo { FileName = "msiexec.exe", Arguments = $"/update \"{c}\" " } };
        var app = new Process { StartInfo = new ProcessStartInfo { FileName = a } };
        Task.Run(() => { MessageBox.Show("Installations MBAM est en cours ...", "MBAM", MessageBoxButtons.OK, MessageBoxIcon.Information); });
        await Task.Run(() =>
        {
            package.Start();
            package.WaitForExit();
            correctif.Start();
            correctif.WaitForExit();
            app.Start();
            app.WaitForExit();
            Task.Run(() => { MessageBox.Show("Installations MBAM terminée", "MBAM", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            Task2();
        });
    }
    public void Adobe(object sender, EventArgs e)
    {
        var k = new FAdobe();
        k.Location = MousePosition;
        k.FormClosed += (s, args) => Task2();
        k.ShowDialog();
    }
    public async void Chrome(object sender, EventArgs e)
    {
        var a = Path.Combine(FLogin.Appsfolder, @"ChromeStandaloneSetup64.exe");
        if (!File.Exists(a))
        {
            var p = new PictureBox();
            p.Parent = bt_chrome;
            p.Image = Resources._0;
            p.Height = bt_chrome.Height - 10;
            p.Width = bt_chrome.Width - 10;
            p.Location = new Point(5, 5);
            p.BackColor = Color.Transparent;
            p.SizeMode = PictureBoxSizeMode.Zoom;
            p.BringToFront();
            return;
        }
        try
        {
            var proc = new Process { StartInfo = new ProcessStartInfo { FileName = a, Arguments = "/silent", CreateNoWindow = true, UseShellExecute = false } };
            bt_chrome.Enabled = false;
            Task.Run(() => { MessageBox.Show("Installation de CHROME est en cours ...", "CHROME", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            await Task.Run(() =>
            {
                proc.Start();
                proc.WaitForExit();
                Task.Run(() => { MessageBox.Show("Installation de CHROME est terminée", "CHROME", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            });
            bt_chrome.Enabled = true;

            Task2();
        }
        catch (Exception)
        {
            throw;
        }
    }
    public async void Firefox(object sender, EventArgs e)
    {
        var a = Path.Combine(FLogin.Appsfolder, @"Firefox Setup 128.0.3.exe");
        if (!File.Exists(a))
        {
            var p = new PictureBox();
            p.Parent = bt_firefox;
            p.Image = Resources._0;
            p.Height = bt_firefox.Height - 10;
            p.Width = bt_firefox.Width - 10;
            p.Location = new Point(5, 5);
            p.BackColor = Color.Transparent;
            p.SizeMode = PictureBoxSizeMode.Zoom;
            p.BringToFront();
            return;
        }
        try
        {
            var proc = new Process { StartInfo = new ProcessStartInfo { FileName = a, Arguments = "/silent", CreateNoWindow = true, UseShellExecute = false } };
            bt_firefox.Enabled = false;
            Task.Run(() => { MessageBox.Show("Installation de FIREFOX est en cours ...", "FIREFOX", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            await Task.Run(() =>
            {
                proc.Start();
                proc.WaitForExit();
                Task.Run(() => { MessageBox.Show("Installation de FIREFOX est terminée", "FIREFOX", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            });
            bt_firefox.Enabled = true;

            Task2();
        }
        catch (Exception)
        {
            throw;
        }
    }
    public async void Silver(object sender, EventArgs e)
    {
        var a = Path.Combine(FLogin.Appsfolder, @"silverlight.exe");
        if (!File.Exists(a))
        {
            var p = new PictureBox();
            p.Parent = bt_silver;
            p.Image = Resources._0;
            p.Height = bt_silver.Height - 10;
            p.Width = bt_silver.Width - 10;
            p.Location = new Point(5, 5);
            p.BackColor = Color.Transparent;
            p.SizeMode = PictureBoxSizeMode.Zoom;
            p.BringToFront();
            return;
        }
        try
        {
            var proc = new Process { StartInfo = new ProcessStartInfo { FileName = a, Arguments = "/q", CreateNoWindow = true, UseShellExecute = false } };
            bt_silver.Enabled = false;
            Task.Run(() => { MessageBox.Show("Installation de Silverlight est en cours ...", "Silverlight", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            await Task.Run(() =>
            {
                proc.Start();
                proc.WaitForExit();
                Task.Run(() => { MessageBox.Show("Installation de Silverlight est terminée", "Silverlight", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            });
            bt_silver.Enabled = true;

            Task2();
        }
        catch (Exception)
        {
            throw;
        }
    }
    public async void Classicshell(object sender, EventArgs e)
    {
        var a = Path.Combine(FLogin.Appsfolder, @"ClassicShellSetup_4_3_1-fr.exe");
        if (!File.Exists(a))
        {
            var p = new PictureBox();
            p.Parent = bt_classicshell;
            p.Image = Resources._0;
            p.Height = bt_classicshell.Height - 10;
            p.Width = bt_classicshell.Width - 10;
            p.Location = new Point(5, 5);
            p.BackColor = Color.Transparent;
            p.SizeMode = PictureBoxSizeMode.Zoom;
            p.BringToFront();
            return;
        }
        try
        {
            var proc = new Process { StartInfo = new ProcessStartInfo { FileName = a, Arguments = "/q", CreateNoWindow = true, UseShellExecute = false } };
            bt_classicshell.Enabled = false;
            Task.Run(() => { MessageBox.Show("Installation de Classic Shell est en cours ...", "Classic Shell", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            await Task.Run(() =>
            {
                proc.Start();
                proc.WaitForExit();
                Task.Run(() => { MessageBox.Show("Installation de Classic Shell est terminée", "Classic Shell", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            });
            bt_classicshell.Enabled = true;

            Task2();
        }
        catch (Exception)
        {
            throw;
        }
    }
    public async void Navition(object sender, EventArgs e)
    {
        var a = Path.Combine(FLogin.Appsfolder, @"Dynamics.NAV60R2.FR.1097337\setup.exe");
        if (File.Exists(a))
        {
            await Task.Run(() => Command($@"start-process '{a}' -wait"));
            Task2();
        }
        else
        {
            var pp = new PictureBox();
            pp.Parent = bt_nav;
            pp.Image = Resources._0;
            pp.Height = bt_nav.Height - 10;
            pp.Width = bt_nav.Width - 10;
            pp.Location = new Point(5, 5);
            pp.BackColor = Color.Transparent;
            pp.SizeMode = PictureBoxSizeMode.Zoom;
            pp.BringToFront();
        }
    }
    public async void Sapinstall(object sender, EventArgs e)
    {
        var a = Path.Combine(FLogin.Appsfolder, @"SAPi\SetupAll.exe");
        if (!File.Exists(a))
        {
            var pp = new PictureBox();
            pp.Parent = bt_sap;
            pp.Image = Resources._0;
            pp.Height = bt_sap.Height - 10;
            pp.Width = bt_sap.Width - 10;
            pp.Location = new Point(5, 5);
            pp.BackColor = Color.Transparent;
            pp.SizeMode = PictureBoxSizeMode.Zoom;
            pp.BringToFront();
            return;
        }
        try
        {
            var proc = new Process { StartInfo = new ProcessStartInfo { FileName = a, Arguments = "/silent", CreateNoWindow = true, UseShellExecute = false } };
            bt_sap.Enabled = false;
            Task.Run(() => { MessageBox.Show("Installation de SAP est en cours ...", "SAP", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            await Task.Run(() =>
            {
                proc.Start();
                proc.WaitForExit();
                Task.Run(() => { MessageBox.Show("Installation de SAP est terminée", "SAP", MessageBoxButtons.OK, MessageBoxIcon.Information); });
            });
            bt_sap.Enabled = true;

            Task2();
        }
        catch (Exception)
        {
            throw;
        }
    }

    // Autre
    public void Info(object sender, EventArgs e)
    {
        var k = new FInfo();
        k.StartPosition = FormStartPosition.CenterScreen;
        k.Show();
    }
    public void Wifi(object sender, EventArgs e)
    {
        Task.Run(() =>
        {
            try
            {
                Invoke(() => { button4.Enabled = false; });
                string tempFilePath = Path.GetTempFileName();
                File.WriteAllText(tempFilePath, Resources.DIR_INFRA);
                bool profileAdded = false;
                var s = Command(@"netsh wlan show profiles");
                if (s.Contains("DIR-INFRA")) profileAdded = true;
                if (!profileAdded) Command($@" netsh wlan add profile filename=""{tempFilePath}"" interface=""Wi-Fi*"" ");
                Thread.Sleep(2000);
                int retryCount = 5;
                bool connected = false;

                for (int i = 0; i < retryCount; i++)
                {
                    var d = Command($@"netsh wlan connect name=""DIR-INFRA"" ssid=""DIR-INFRA"" interface=""Wi-Fi*""");
                    if (d.Contains("La demande de connexion a"))
                    {
                        connected = true;
                        MessageBox.Show("Connected to Wi-Fi successfully.");
                        Invoke(() => { button4.Enabled = true; });
                        break;
                    }
                    else
                    {
                        Thread.Sleep(1000);
                    }
                }

                if (!connected)
                {
                    MessageBox.Show($"Failed to connect to Wi-Fi after {retryCount} attempts.");
                    Invoke(() => { button4.Enabled = true; });
                }

                File.Delete(tempFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Invoke(() =>
                {
                    button4.Enabled = true;
                });
            }
        });
    }
    public void Diskmgmt(object sender, EventArgs e) => Task.Run(() => Command("diskmgmt.msc"));
    public void Snappy(object sender, EventArgs e)
    {
        var s = Path.Combine(FLogin.Appsfolder, @"SDI_RUS_2023_05\s.lnk");
        if (File.Exists(s))
        {
            Task.Run(() => Command($@"start-process '{s}'"));
        }
        else
        {
            var p = new PictureBox();
            p.Parent = bt_snappy;
            p.Image = Resources._0;
            p.Height = bt_snappy.Height - 10;
            p.Width = bt_snappy.Width - 10;
            p.Location = new Point(5, 5);
            p.BackColor = Color.Transparent;
            p.SizeMode = PictureBoxSizeMode.Zoom;
            p.BringToFront();
        };
    }
    public void Appwiz(object sender, EventArgs e)
    {
        [DllImport("shell32.dll", SetLastError = true)]
        static extern IntPtr ShellExecute(IntPtr hwnd, string lpOperation, string lpFile, string? lpParameters, string? lpDirectory, int nShowCmd);
        ShellExecute(IntPtr.Zero, "open", "ms-settings:appsfeatures", null, null, 1);
        Command("appwiz.cpl");
    }
    public void Renamme(object sender, EventArgs e)
    {
        Task.Run(() =>
        {
            var x = Process.Start("SystemPropertiesComputerName.exe");
            x.WaitForInputIdle();
            Task.Delay(1000).Wait();
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{TAB}");
            SendKeys.SendWait("{enter}");
        });
    }
    public void Time(object sender, EventArgs e)
    {
        Task.Run(() =>
        {
            Command(@"
            Set-Timezone -Name 'Afr. centrale Ouest';
            w32tm.exe /resync
        ");
        });
    }
    private Task<bool> Checkapp(string pattern)
    {
        return Task.Run(() =>
        {
            // Check in both Local Machine and Current User registry paths
            var registryPaths = new string[]
            {
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",  // For 32-bit apps on 64-bit Windows
            @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            // Create a regular expression object from the pattern
            var regex = new Regex(pattern, RegexOptions.IgnoreCase);

            foreach (var path in registryPaths)
            {
                using var key = Registry.LocalMachine.OpenSubKey(path);
                if (key != null)
                {
                    foreach (string subkeyName in key.GetSubKeyNames())
                    {
                        using var subkey = key.OpenSubKey(subkeyName);
                        if (subkey?.GetValue("DisplayName") is string displayName && regex.IsMatch(displayName))
                        {
                            //MessageBox.Show(subkey?.GetValue("DisplayName") as string);
                            return true; // App found
                        }
                    }
                }
            }

            return false; // App not found
        });
    }
    private void CopyFile(string source, string dest, string pattern, PictureBox p)
    {
        try
        {
            var x = Directory.EnumerateFiles(source, pattern);
            if (x.Any())
            {
                Directory.EnumerateFiles(source, pattern)
                  .ToList()
                  .ForEach(file =>
                      File.Copy(file, Path.Combine(dest, Path.GetFileName(file)), true)
                  );
                p.Image = Resources._1;
            }
            else p.Image = Resources._0;
        }
        catch (Exception)
        {
            p.Image = Resources._0;
        }

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
    private void DesktopReg(string key, PictureBox p)
    {
        var r = $@"{FLogin.Sid}\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel";
        using var k = Registry.Users.OpenSubKey(r, true);
        if (k != null) { k.SetValue(key, 0, RegistryValueKind.DWord); p.Image = Resources._1; RefDesktop(); }
        else p.Image = Resources._0;
    }
    private void RefDesktop()
    {
        const uint SHCNE_ASSOCCHANGED = 0x08000000;
        const uint SHCNF_IDLIST = 0x0000;

        [DllImport("Shell32.dll")]
        static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
    }

    public void Regedit(object sender, EventArgs e)
    {
        try
        {
            Task.Run(() => new Process { StartInfo = new ProcessStartInfo { FileName = "Regedit.exe" } }.Start());
            Clipboard.SetText(@"Ordinateur\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\FVE\MDOPBitLockerManagement");
        }
        catch (Exception)
        {
            throw;
        }
    }
    public void Pshell(object sender, EventArgs e)
    {
        Task.Run(() => new Process { StartInfo = new ProcessStartInfo { FileName = "Powershell.exe" } }.Start());
    }
    // Paths
    public string Desktop => $@"C:\Users\{FLogin.User}\Desktop\";
    public string Roaming => $@"C:\Users\{FLogin.User}\AppData\Roaming";
    public string Programmes = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs";

    private void checkBox1_CheckedChanged(object sender, EventArgs e)
    {
        if (checkBox1.Checked) { timer1.Enabled = true; }
        else { timer1.Enabled = false; richTextBox1.Text = ""; }
    }

    private void button6_Click(object sender, EventArgs e)
    {
        Task2();
    }

    public static void setkaspersky(bool status)
    {
        _instance.bt_kaspersky.Enabled = status;

    }
    public static void setoffice(bool status)
    {
        _instance.bt_office.Enabled = status;
    }

    private void guna2vSeparator1_Click(object sender, EventArgs e)
    {

    }

    private void groupBox7_Enter(object sender, EventArgs e)
    {

    }

    private void groupBox6_Enter(object sender, EventArgs e)
    {

    }
}
