# Aplikasi manajemen database mahasiswa

Asslamualaikum Wr. Wb. Perkenalkan saya Andy Rahman disini saya akan menjelaskan cara kerja aplikasi saya. Aplikasi ini dirancang untuk membantu administrator dalam mengelola dan menganalisis data mahasiswa secara efisien. Melalui dokumentasi ini, Anda akan mempelajari tentang fitur-fitur dan alur kerja aplikasi ini, serta petunjuk bagaimana menggunakannya untuk mencapai tujuan Anda. Di aplikasi ini saya menggunakan NET Framework 4.7.2 dengan dependencies sebagai berikut:

•	Untuk database saya menggunakan Microsoft Access dan menghubungkannya dengan namespace dalam .NET Framework yaitu System.Data.OleDb.
•	Saya juga menggunakan ClosedXML sebuah library .NET untuk memproses file Excel.

### CATATAN
Untuk menghindari terjadinya ERROR sebelum Run debug kode ini, saya sarankan anda harus menginstall dependencies seperti ClosedXML, SixLabors.Fonts, AccessDatabaseEngine dan System.Data.OleDb terlebih dahulu.
Untuk menginstall ClosedXML dan System.Data.OleDb
```
> Pergi ke Project 
> Manage NuGet Packages... 
> cari di pencarian ClosedXML dan System.Data.OleDb 
> Install
```

### Untuk menginstall SixLabors.Fonts
```
> Pergi ke Tools 
> NuGet Package Manager 
> Package Manager Console 
> copy dan paste kode ini ke terminal 
> NuGet\Install-Package SixLabors.Fonts -Version 1.0.0-beta19
```
### Untuk menginstall AccessDatabaseEngine
```
> Pergi ke link ini https://www.microsoft.com/en-in/download/details.aspx?id=13255
> Klik Download
> Pilih "AccessDatabaseEngine.exe" jangan pilih "AccessDatabaseEngine_X64.exe"
> Next dan Install
```
Note jika muncul Error di kode anda seperti ini 
System.InvalidOperationException: 'The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.' 
Coba install AccessDatabaseEngine_X64.exe atau AccessDatabaseEngine.exe

### Screenshot Aplikasi
![alt text](https://i.imgur.com/RfKrZgP.png "Gambar Aplikasi")

### Link Video Penjelasan
[Klik Aku](https://youtu.be/KzOeuxc7F4I)
