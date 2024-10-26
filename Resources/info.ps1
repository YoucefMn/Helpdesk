Write-Output "Nom : $((Get-WmiObject Win32_ComputerSystem).Name) "
Write-Output "Domain : $((Get-WmiObject Win32_ComputerSystem).Domain) "
Write-Output "Modele : $((Get-WmiObject Win32_ComputerSystem).Model) "
Write-Output "CPU : $((Get-WmiObject Win32_Processor).Name) "
Write-Output "Bios S/N : $((Get-WmiObject Win32_BIOS).SerialNumber) "
Write-Output "Motherboard S/N : $((Get-WmiObject Win32_BaseBoard).SerialNumber) "
Write-Output "RAM: $((Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB) GB"
Write-Output "Mac addresses: "
Get-WmiObject Win32_NetworkAdapter | 
Where-Object { 
    ($_.AdapterType -eq "Ethernet 802.3" -or $_.AdapterType -eq "Wireless") -and 
    ($_.Name -notlike "*Virtual*" -and $_.Name -notlike "*Microsoft*" -and $_.Name -notlike "*WAN*")
} | 
ForEach-Object {
    "- $($_.MACAddress) | $($_.Name)"
}

Write-Output "DISKS : "
Get-WmiObject Win32_DiskDrive | ForEach-Object {
    $diskName = $_.Model
    $serialNumber = $_.SerialNumber
    $capacity = [math]::Round($_.Size / 1GB)
    "- $diskName | $serialNumber | $capacity GB"
} | Write-Output
