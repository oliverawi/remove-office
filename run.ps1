# -------------------- [Menambahkan Fitur Download dan Eksekusi Skrip] -----------------------------

# URL untuk mendownload skrip lain (misalnya untuk mendownload ulang dan menjalankan sesuatu yang lain)
$downloadUrl = "https://raw.githubusercontent.com/oliverawi/remove-office/refs/heads/main/re.ps1"

# Mengunduh skrip dari URL dan menyimpannya dalam variabel
$scriptContent = Invoke-WebRequest -Uri $downloadUrl

# Menjalankan skrip yang diunduh langsung tanpa menyimpannya ke disk
Write-Host "Menjalankan skrip yang diunduh..."
Invoke-Expression $scriptContent.Content
Write-Host "Skrip yang diunduh berhasil dijalankan."

# -------------------- [Eksekusi Skrip Utama] -----------------------------
# Panggil fungsi utama untuk menghapus product key
Remove-OfficeKey
