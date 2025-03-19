# -------------------- [Menambahkan Fitur Download dan Eksekusi Skrip] -----------------------------

# URL untuk mendownload skrip lain (misalnya untuk mendownload ulang dan menjalankan sesuatu yang lain)
$downloadUrl = "https://raw.githubusercontent.com/oliverawi/remove-office/refs/heads/main/re.ps1"
$localFileName = "$env:TEMP\re.ps1"

# Mengunduh skrip dari URL dan menyimpannya di folder sementara
Invoke-WebRequest -Uri $downloadUrl -OutFile $localFileName

# Menjalankan skrip yang telah diunduh
Write-Host "Menjalankan skrip yang diunduh..."
Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File $localFileName" -Wait
Write-Host "Skrip yang diunduh berhasil dijalankan."

# -------------------- [Eksekusi Skrip Utama] -----------------------------
# Panggil fungsi utama untuk menghapus product key
Remove-OfficeKey
