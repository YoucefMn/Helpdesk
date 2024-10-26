using System.Diagnostics;
using System.Management;
using System.Text;
using Helpdesk.Forms;
using Microsoft.Data.SqlClient;

namespace Helpdesk
{
    public partial class FInfo : Form
    {

        public FInfo()
        {
            InitializeComponent();
            var dragManager = new Drag(this);
            panel1.MouseDown += dragManager.Panel_MouseDown;
            panel1.MouseMove += dragManager.Panel_MouseMove;
            panel1.MouseUp += dragManager.Panel_MouseUp;
            saveFileDialog1.Filter = "Text Files | *.txt";
            saveFileDialog1.DefaultExt = "txt";


        }
        private void Close(object sender, EventArgs e) => Close();
        private void Minimize(object sender, EventArgs e) => WindowState = FormWindowState.Minimized;
        //public async void Inf(string Computer = "")
        //{
        //    var hostname = Command("Hostname");
        //    if (string.IsNullOrEmpty(Computer)) tlaps.Text = Command($@"(Get-LapsADPassword {hostname} -AsPlainText).password");
        //    else tlaps.Text = Command($@"(Get-LapsADPassword {Computer} -AsPlainText).password");
        //    tnom.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{Hostname}}"));
        //    tdom.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{(Get-WmiObject Win32_ComputerSystem).Domain}}"));
        //    tmodel.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{(Get-WmiObject Win32_ComputerSystem).Model}}"));
        //    tcpu.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{(Get-WmiObject Win32_Processor).Name}}"));
        //    tram.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{Write-Output ""$( (Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB ) GB""}}"));
        //    tbios.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{(Get-WmiObject Win32_BIOS).SerialNumber}}"));
        //    tcm.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{(Get-WmiObject Win32_BaseBoard).SerialNumber}}"));
        //    tip.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{Get-NetIPConfiguration | Where-Object {{$_.IPv4DefaultGateway -ne $null -and $_.NetAdapter.Status -ne 'Disconnected'}}| ForEach-Object {{""$($_.InterfaceAlias) + ' | ' + $($_.IPv4Address.IPAddress)""}}}}"));
        //    tmac.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{Get-NetAdapter | Select-Object Name, MacAddress | ForEach-Object {{""$($_.MacAddress) + ' | ' + $($_.Name)""}}}}"));
        //    tdisk.Text = await Task.Run(() => Command($@"Invoke-Command {Computer} {{Get-WmiObject Win32_DiskDrive |ForEach-Object {{""$($_.Model) + ' | ' + $($_.SerialNumber) + ' | ' + $([math]::Round($_.Size / 1GB)) + 'GB' ""}}}}"));
        //}
        //public string CommandF(string command)
        //{
        //    var psi = new ProcessStartInfo
        //    {
        //        FileName = "powershell.exe",
        //        Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{command}\"",
        //        RedirectStandardOutput = true,
        //        RedirectStandardError = true,
        //        UseShellExecute = false,
        //        CreateNoWindow = true
        //    };
        //    var output = new StringBuilder();

        //    using var process = Process.Start(psi);
        //    if (process != null)
        //    {
        //        Action<StreamReader> appendOutput = reader => output.AppendLine(reader.ReadToEnd());
        //        appendOutput(process.StandardOutput);
        //        appendOutput(process.StandardError);
        //        process.WaitForExit();
        //    }


        //    return output.ToString().Trim();
        //}
        //public string Command(string command)
        //{
        //    var psi = new ProcessStartInfo
        //    {
        //        FileName = "powershell.exe",
        //        Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{command}\"",
        //        RedirectStandardOutput = true,
        //        RedirectStandardError = true,
        //        UseShellExecute = false,
        //        CreateNoWindow = true
        //    };
        //    var output = new StringBuilder();

        //    using var process = Process.Start(psi);
        //    if (process != null)
        //    {
        //        Action<StreamReader> appendOutput = reader => output.AppendLine(reader.ReadToEnd());
        //        appendOutput(process.StandardOutput);
        //        appendOutput(process.StandardError);
        //        process.WaitForExit();
        //    }


        //    return output.ToString().Trim();
        //}

        //private void FInfo_Load(object sender, EventArgs e) =>Inf();

        //private void BClick(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        Inf(textBox1.Text);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    string tempFilePath = Path.Combine(Path.GetTempPath(), "inf.ps1");
        //    File.WriteAllText(tempFilePath, Resources.info);
        //    var p = saveFileDialog1.ShowDialog();
        //    if (p == DialogResult.OK)
        //    {
        //        var f = saveFileDialog1.OpenFile();
        //        var x = CommandF(tempFilePath);
        //        var ff = new StreamWriter(f);
        //        ff.Write(x);
        //        ff.Flush();
        //        ff.Close();
        //    }
        //}
        private async void buttonGetInfo_Click(object sender, EventArgs e)
        {
            // Disable the button while loading
            textBoxTargetPC.Enabled = false;
            //buttonGetInfo.Enabled = false;
            textBoxName.Clear();
            textBoxDomain.Clear();
            textBoxModel.Clear();
            textBoxCpuName.Clear();
            textBoxRamSize.Clear();
            textBoxBiosSerial.Clear();
            textBoxMotherboardSerial.Clear();
            richTextBoxNetworkAdapters.Clear();
            richTextBoxDiskInfo.Clear();
            tlaps.Clear();

            string targetPC = textBoxTargetPC.Text.Trim();
            if (!string.IsNullOrEmpty(targetPC))
            {
                BMstsc.Enabled = true;
                bGoverlan.Enabled = true;
            }

            try
            {
                await Task.Run(() => GetPCInfo(targetPC));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                // Re-enable the button after loading
                textBoxTargetPC.Enabled = true;
            }
        }
        private void GetPCInfo(string targetPC)
        {
            string name, domain, model, cpuName, ramSize, biosSerial, motherboardSerial, diskInfo, networkAdapters, LapsPassword;

            try
            {
                if (string.IsNullOrEmpty(targetPC))
                {
                    name = Environment.MachineName;
                    domain = GetDomainName();
                    model = GetModel();
                    cpuName = GetCpuName();
                    ramSize = GetRamSizeGB();
                    biosSerial = GetBiosSerial();
                    motherboardSerial = GetMotherboardSerial();
                }
                else
                {
                    name = GetRemoteProperty(targetPC, "Name", "Win32_ComputerSystem");
                    domain = GetRemoteProperty(targetPC, "Domain", "Win32_ComputerSystem");
                    model = GetRemoteProperty(targetPC, "Model", "Win32_ComputerSystem");
                    cpuName = GetRemoteProperty(targetPC, "Name", "Win32_Processor");
                    ramSize = GetRemoteRamSizeGB(targetPC);
                    biosSerial = GetRemoteProperty(targetPC, "SerialNumber", "Win32_BIOS");
                    motherboardSerial = GetRemoteProperty(targetPC, "SerialNumber", "Win32_BaseBoard");

                    BMstsc.Enabled = true;
                    bGoverlan.Enabled = true;
                }

                diskInfo = GetDiskInfo(targetPC);
                networkAdapters = GetNetworkAdapters(targetPC);


                // Update the UI from the main thread
                this.Invoke((MethodInvoker)delegate
                {
                    textBoxName.Text = name;
                    textBoxDomain.Text = domain;
                    textBoxModel.Text = model;
                    textBoxCpuName.Text = cpuName;
                    textBoxRamSize.Text = ramSize;
                    textBoxBiosSerial.Text = biosSerial;
                    textBoxMotherboardSerial.Text = motherboardSerial;
                    richTextBoxDiskInfo.Text = diskInfo;
                    richTextBoxNetworkAdapters.Text = networkAdapters;
                    LapsPassword = GetLapsPassword(textBoxName.Text);
                    tlaps.Text = LapsPassword;
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }
        private string GetDomainName()
        {
            try
            {
                return new ManagementObjectSearcher("SELECT Domain FROM Win32_ComputerSystem").Get().Cast<ManagementObject>().FirstOrDefault()?["Domain"]?.ToString() ?? "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetModel()
        {
            try
            {
                return new ManagementObjectSearcher("SELECT Model FROM Win32_ComputerSystem").Get().Cast<ManagementObject>().FirstOrDefault()?["Model"]?.ToString() ?? "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetCpuName()
        {
            try
            {
                return new ManagementObjectSearcher("SELECT Name FROM Win32_Processor").Get().Cast<ManagementObject>().FirstOrDefault()?["Name"]?.ToString() ?? "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetRamSizeGB()
        {
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
                var obj = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
                if (obj != null)
                {
                    ulong totalMemoryBytes = (ulong)obj["TotalPhysicalMemory"];
                    double totalMemoryGB = totalMemoryBytes / (1024.0 * 1024.0 * 1024.0);
                    return Math.Ceiling(totalMemoryGB) + " GB";
                }
                return "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetBiosSerial()
        {
            try
            {
                return new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS").Get().Cast<ManagementObject>().FirstOrDefault()?["SerialNumber"]?.ToString() ?? "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetMotherboardSerial()
        {
            try
            {
                return new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard").Get().Cast<ManagementObject>().FirstOrDefault()?["SerialNumber"]?.ToString() ?? "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        //private string GetDiskInfo(string targetPC)
        //{
        //    try
        //    {
        //        var scope = string.IsNullOrEmpty(targetPC) ? new ManagementScope() : new ManagementScope($@"\\{targetPC}\root\cimv2", new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate });
        //        var query = new ObjectQuery("SELECT Caption, SerialNumber, Size FROM Win32_DiskDrive");
        //        var searcher = new ManagementObjectSearcher(scope, query);
        //        var disks = searcher.Get();

        //        string info = "";
        //        foreach (var disk in disks)
        //        {
        //            string caption = disk["Caption"]?.ToString() ?? "N/A";

        //            if (caption.Contains("USB", StringComparison.OrdinalIgnoreCase))
        //            {
        //                continue;
        //            }
        //            string serialNumber = disk["SerialNumber"]?.ToString() ?? "N/A";
        //            ulong sizeBytes = (ulong)(disk["Size"] ?? 0);
        //            double sizeGB = sizeBytes / (1000 * 1000 * 1000); // Use base-10 units
        //            info += $"{caption} | {serialNumber} | {sizeGB} GB\n";
        //        }
        //        return info;
        //    }
        //    catch
        //    {
        //        return "N/A";
        //    }
        //}
        //private string GetNetworkAdapters(string targetPC) // ip and name
        //{
        //    try
        //    {
        //        var scope = string.IsNullOrEmpty(targetPC) ? new ManagementScope() : new ManagementScope($@"\\{targetPC}\root\cimv2", new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate });
        //        var query = new ObjectQuery("SELECT Name, MACAddress FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True");
        //        var searcher = new ManagementObjectSearcher(scope, query);
        //        var adapters = searcher.Get();

        //        string info = "";
        //        foreach (var adapter in adapters)
        //        {
        //            string name = adapter["Name"]?.ToString() ?? "N/A";
        //            string macAddress = adapter["MACAddress"]?.ToString() ?? "N/A";

        //            // Skip Bluetooth adapters
        //            if (name.Contains("Bluetooth"))
        //            {
        //                continue;
        //            }

        //            // Skip virtual adapters
        //            if (name.Contains("Virtual") || name.Contains("VPN") || name.Contains("Hyper-V") || name.Contains("PPPoP"))
        //            {
        //                continue;
        //            }

        //            // Get IP addresses for the current adapter
        //            var networkInterfaces = NetworkInterface.GetAllNetworkInterfaces();
        //            bool isConnected = false;
        //            string ipAddress = null;

        //            foreach (var networkInterface in networkInterfaces)
        //            {
        //                if (networkInterface.Description == name)
        //                {
        //                    var ipProperties = networkInterface.GetIPProperties();
        //                    foreach (var ip in ipProperties.UnicastAddresses)
        //                    {
        //                        if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork && !IPAddress.IsLoopback(ip.Address) && !ip.Address.ToString().StartsWith("169.254."))
        //                        {
        //                            ipAddress = ip.Address.ToString();
        //                            isConnected = true;
        //                            break;
        //                        }
        //                    }
        //                    if (isConnected) break;
        //                }
        //            }

        //            info += $"- {name} | {macAddress} | {ipAddress ?? "N/A"} \n";
        //        }
        //        return info;
        //    }
        //    catch
        //    {
        //        return "N/A";
        //    }
        //} // avec ip and name//
        private string GetDiskInfo(string targetPC)
        {
            try
            {
                var scope = string.IsNullOrEmpty(targetPC) ? new ManagementScope() : new ManagementScope($@"\\{targetPC}\root\cimv2", new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate });
                var query = new ObjectQuery("SELECT Caption, SerialNumber FROM Win32_DiskDrive");
                var searcher = new ManagementObjectSearcher(scope, query);
                var disks = searcher.Get();

                List<string> diskNames = new List<string>();
                List<string> diskSerials = new List<string>();

                foreach (var disk in disks)
                {
                    string caption = disk["Caption"]?.ToString() ?? "N/A";

                    if (caption.Contains("USB", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    string serialNumber = disk["SerialNumber"]?.ToString() ?? "N/A";
                    diskNames.Add(caption);
                    diskSerials.Add(serialNumber);
                }

                string namesLine = string.Join(" | ", diskNames);
                string serialsLine = string.Join(" | ", diskSerials);

                return $"{namesLine}{Environment.NewLine}{serialsLine}";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetNetworkAdapters(string targetPC)
        {
            try
            {
                var scope = string.IsNullOrEmpty(targetPC) ? new ManagementScope() : new ManagementScope($@"\\{targetPC}\root\cimv2", new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate });
                var query = new ObjectQuery("SELECT Name, MACAddress FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True");
                var searcher = new ManagementObjectSearcher(scope, query);
                var adapters = searcher.Get();

                string macAddresses = "";
                foreach (var adapter in adapters)
                {
                    string name = adapter["Name"]?.ToString() ?? "N/A";
                    string macAddress = adapter["MACAddress"]?.ToString() ?? "N/A";

                    // Skip Bluetooth adapters
                    if (name.Contains("Bluetooth"))
                    {
                        continue;
                    }

                    // Skip virtual adapters
                    if (name.Contains("Virtual") || name.Contains("VPN") || name.Contains("Hyper-V") || name.Contains("PPPoP"))
                    {
                        continue;
                    }

                    // Append MAC address to the result string
                    if (macAddresses.Length > 0)
                    {
                        macAddresses += " | ";
                    }
                    macAddresses += macAddress;
                }
                return macAddresses;
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetRemoteProperty(string targetPC, string propertyName, string wmiClass)
        {
            try
            {
                var scope = new ManagementScope($@"\\{targetPC}\root\cimv2", new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate });
                var query = new ObjectQuery($"SELECT {propertyName} FROM {wmiClass}");
                var searcher = new ManagementObjectSearcher(scope, query);
                var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
                return result?[propertyName]?.ToString() ?? "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetRemoteRamSizeGB(string targetPC)
        {
            try
            {
                var scope = new ManagementScope($@"\\{targetPC}\root\cimv2", new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate });
                var query = new ObjectQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
                var searcher = new ManagementObjectSearcher(scope, query);
                var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
                if (result != null)
                {
                    ulong totalMemoryBytes = (ulong)result["TotalPhysicalMemory"];
                    double totalMemoryGB = totalMemoryBytes / (1024.0 * 1024.0 * 1024.0);
                    return Math.Ceiling(totalMemoryGB) + " GB";
                }
                return "N/A";
            }
            catch
            {
                return "N/A";
            }
        }
        private string GetLapsPassword(string targetPC)
        {
            string lapsPassword = "Not Found";

            try
            {
                string computerName = string.IsNullOrEmpty(targetPC) ? Environment.MachineName : targetPC;

                // Create a new process start info
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = "powershell.exe";
                psi.Arguments = $"-Command \"Get-LapsAdPassword {computerName} -AsPlainText | Select-Object -ExpandProperty Password\"";
                psi.RedirectStandardOutput = true;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;

                // Start the process
                using (Process process = Process.Start(psi))
                {
                    // Read the output
                    using (StreamReader reader = process.StandardOutput)
                    {
                        lapsPassword = reader.ReadToEnd().Trim();
                    }
                    if (lapsPassword.Contains("Current machine is not AD domain-joined"))
                    {
                        lapsPassword = "Current machine is not AD domain-joined";
                    }
                }
            }
            catch (Exception ex)
            {
                lapsPassword = $"Error: {ex.Message}";
            }

            return lapsPassword;
        }
        private void buttonSaveInfo_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = $"{textBoxName.Text}.txt";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog1.FileName;

                try
                {
                    using (StreamWriter writer = new StreamWriter(fileName))
                    {
                        writer.WriteLine($"PC Nom: {textBoxName.Text}");
                        writer.WriteLine($"Domain: {textBoxDomain.Text}");
                        writer.WriteLine($"Model: {textBoxModel.Text}");
                        writer.WriteLine($"CPU Name: {textBoxCpuName.Text}");
                        writer.WriteLine($"RAM : {textBoxRamSize.Text}");
                        writer.WriteLine($"BIOS Serial: {textBoxBiosSerial.Text}");
                        writer.WriteLine($"Motherboard Serial: {textBoxMotherboardSerial.Text}");
                        writer.WriteLine($"_________ Disk Info _______________\n{richTextBoxDiskInfo.Text}");
                        writer.WriteLine($"_____________________________");
                        writer.WriteLine($"Mac: {richTextBoxNetworkAdapters.Text}");
                        writer.WriteLine($"LAPS Password: {tlaps.Text}");
                    }

                    MessageBox.Show("Information saved successfully!", "Save Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving information: {ex.Message}", "Save Info", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void FInfo_Load(object sender, EventArgs e)
        {
            GetPCInfo("");
            BMstsc.Enabled = false;
            bGoverlan.Enabled = false;
        }
        private async void buttonOpenMstsc_Click(object sender, EventArgs e)
        {
            string targetPC = textBoxTargetPC.Text; // Assuming you have a TextBox named textBoxTargetPC to input the target PC name or IP address
            string targetUser = $@"{targetPC}\helpdesk-admin";
            string targetPass = tlaps.Text;
            if (!string.IsNullOrEmpty(targetPC))
            {
                try
                {
                    await Task.Run(() => Command($@"cmdkey /generic:{targetPC} /user:{targetUser} /pass: '{targetPass}'"));
                    await Task.Run(() => Command($@"mstsc /v:{targetPC}"));

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error opening MSTSC: {ex.Message}", "MSTSC Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                BMstsc.Enabled = false;
            }
        }
        private void Goverlan_Click(object sender, EventArgs e)
        {
            string Gover = @"C:\Program Files (x86)\GoverLAN v7\GoverRMC.exe";
            if (!string.IsNullOrEmpty(textBoxTargetPC.Text))
            {
                if (File.Exists(Gover))
                {
                    var process = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = Gover,
                            Arguments = textBoxTargetPC.Text,
                            UseShellExecute = true,
                        }
                    };
                    process.Start();
                    Task.Delay(1000).Wait();
                    SendKeys.SendWait("{ENTER}");
                    SendKeys.SendWait("{TAB}");
                }
                else
                {
                    bGoverlan.Enabled = false;
                }
            }
            else
            {
                bGoverlan.Enabled = false;
            }
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
        private void textbox_MouseEnter(object sender, EventArgs e)
        {
            // Cast the sender to a TextBox
            TextBox textBox = sender as TextBox;

            if (textBox != null)
            {
                // Copy the text of the TextBox to the clipboard
                Clipboard.SetText(textBox.Text);

            }
        }

        private void InsertDataIntoDatabase(object sender, EventArgs e)
        {
            //string connectionString = "Server=DZCEAO0072;Database=gpi;User Id=root;Password=123456";
            string connectionString = "Server=CONDOR-S2\\SQLEXPRESS;Database=test;User Id=sa;Password=1234;TrustServerCertificate=True";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    InsertOrUpdateAllInfo(connection);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw;
            }
        }

        private void InsertOrUpdateAllInfo(SqlConnection connection)
        {
            try
            {
                string[] Fields = { "nom", "Model", "cpu", "ram", "bios", "cm", "mac" };
                string[] Values = { textBoxName.Text, textBoxModel.Text, textBoxCpuName.Text,
            textBoxRamSize.Text, textBoxBiosSerial.Text, textBoxMotherboardSerial.Text,
            richTextBoxNetworkAdapters.Text
        };

                string[] diskInfoLines = richTextBoxDiskInfo.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                List<string> diskNames = new List<string>();
                List<string> serialNumbers = new List<string>();

                if (diskInfoLines.Length == 2)
                {
                    string diskName = diskInfoLines[0].Trim();
                    string serialNumber = diskInfoLines[1].Trim();

                    if (diskName != "N/A" && serialNumber != "N/A")
                    {
                        diskNames.Add(diskName);
                        serialNumbers.Add(serialNumber);
                    }
                }

                string diskNamesConcat = string.Join(" | ", diskNames);
                string serialNumbersConcat = string.Join(" | ", serialNumbers);

                // Check if the computer name exists
                string checkQuery = "SELECT COUNT(*) FROM PC WHERE nom = @nom";
                using (SqlCommand checkCommand = new SqlCommand(checkQuery, connection))
                {
                    checkCommand.Parameters.AddWithValue("@nom", textBoxName.Text);
                    int count = (int)checkCommand.ExecuteScalar();

                    if (count > 0)
                    {
                        // Update the existing record
                        string updateQuery = "UPDATE PC SET model = @model, cpu = @cpu, ram = @ram, bios = @bios, cm = @cm, mac = @mac, disknom = @disknom, diskserial = @diskserial WHERE nom = @nom";
                        using (SqlCommand updateCommand = new SqlCommand(updateQuery, connection))
                        {
                            for (int i = 1; i < Fields.Length; i++) // Skip "nom" as it's used in the WHERE clause
                            {
                                if (!string.IsNullOrEmpty(Values[i]) && Values[i] != "N/A")
                                {
                                    updateCommand.Parameters.AddWithValue($"@{Fields[i]}", Values[i]);
                                }
                                else
                                {
                                    updateCommand.Parameters.AddWithValue($"@{Fields[i]}", DBNull.Value);
                                }
                            }

                            updateCommand.Parameters.AddWithValue("@disknom", diskNamesConcat);
                            updateCommand.Parameters.AddWithValue("@diskserial", serialNumbersConcat);
                            updateCommand.Parameters.AddWithValue("@nom", textBoxName.Text);
                            updateCommand.ExecuteNonQuery();
                        }

                        MessageBox.Show("Data updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Insert a new record
                        string insertQuery = "INSERT INTO PC (nom, model, cpu, ram, bios, cm, mac, disknom, diskserial) VALUES (@nom, @model, @cpu, @ram, @bios, @cm, @mac, @disknom, @diskserial)";
                        using (SqlCommand insertCommand = new SqlCommand(insertQuery, connection))
                        {
                            for (int i = 0; i < Fields.Length; i++)
                            {
                                if (!string.IsNullOrEmpty(Values[i]) && Values[i] != "N/A")
                                {
                                    insertCommand.Parameters.AddWithValue($"@{Fields[i]}", Values[i]);
                                }
                                else
                                {
                                    insertCommand.Parameters.AddWithValue($"@{Fields[i]}", DBNull.Value);
                                }
                            }

                            insertCommand.Parameters.AddWithValue("@disknom", diskNamesConcat);
                            insertCommand.Parameters.AddWithValue("@diskserial", serialNumbersConcat);
                            insertCommand.ExecuteNonQuery();
                        }

                        MessageBox.Show("Data inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to insert or update all information: {ex.Message}");
            }
        }
        private void TextBox_MouseUp(object sender, MouseEventArgs e)
        {
            var textBox = sender as Guna.UI2.WinForms.Guna2TextBox;
            if (textBox != null && !string.IsNullOrEmpty(textBox.SelectedText))
            {
                Clipboard.SetText(textBox.SelectedText);
            }
        }
        private void Mail(object sender, EventArgs e)
        {
            var mail = new FMail();
            mail.Location = MousePosition;
            if (!string.IsNullOrEmpty(textBoxName.Text) && textBoxName.Text != "N/A")
            {
                if (string.IsNullOrEmpty(textBoxTargetPC.Text))
                {
                    mail.TargetPc = textBoxName.Text;
                mail.Show();
                }
                else
                {
                mail.TargetPc = textBoxTargetPC.Text;
                mail.Show();
                }
            }
        }
    }
}