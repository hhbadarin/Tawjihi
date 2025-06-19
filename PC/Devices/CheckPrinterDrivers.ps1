#  Checking installed INF-based drivers
Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, InfName, DriverVersion
