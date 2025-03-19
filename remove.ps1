# Skrip utama untuk mendeteksi dan menghapus product key Microsoft Office (32-bit atau 64-bit)

function Remove-OfficeKey {
<#
.SYNOPSIS
Mendeteksi Office versi 32-bit atau 64-bit dan menghapus product key yang terkait.

.DESCRIPTION
Skrip ini akan mendeteksi apakah Microsoft Office yang terpasang di PC menggunakan versi 32-bit atau 64-bit.
Kemudian, skrip akan mengidentifikasi product key yang terhubung dengan Office dan menghapusnya menggunakan OSPP.VBS.

.EXAMPLE
Remove-OfficeKey.ps1
#>

    # Cek apakah Office diinstal sebagai 32-bit atau 64-bit
    $officePath32 = 'C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS'
    $officePath64 = 'C:\Program Files\Microsoft Office\Office16\OSPP.VBS'

    # Tentukan path Office berdasarkan arsitektur sistem
    if (Test-Path $officePath32) {
        $officePath = $officePath32
        $arch = "32-bit"
    } elseif (Test-Path $officePath64) {
        $officePath = $officePath64
        $arch = "64-bit"
    } else {
        Write-Host "Microsoft Office tidak terdeteksi pada sistem ini."
        return
    }

    Write-Host "Office terdeteksi sebagai $arch."

    # Mendapatkan status lisensi
    $license = cscript $officePath /dstatus

    # Menentukan indikator untuk key cracked atau volume key
    $crackIndicator = 'VOLUME_KMSCLIENT channel'

    # Loop untuk memeriksa apakah ada key cracked
    for ($i = 0; $i -lt $license.Length; $i++) {
        if ($license[$i] -match $crackIndicator) {
            # Jika ditemukan, lompat ke baris yang berisi product key
            $i += 6
            $keyline = $license[$i]
            $prodkey = $keyline.substring($keyline.length - 5, 5)  # Mengambil 5 karakter terakhir sebagai product key

            Write-Host "Key terdeteksi: $prodkey. Menghapus key ini..."

            # Menghapus product key menggunakan OSPP.VBS
            cscript $officePath /unpkey:$prodkey
            Write-Host "Product key uninstall successful."
        }
    }

    Write-Host "Proses selesai."
}

