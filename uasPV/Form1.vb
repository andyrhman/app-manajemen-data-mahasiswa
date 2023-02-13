'Mengimport dependecies yang dibutuhkan di project ini
Imports System.Data.OleDb
Imports System.Drawing.Printing
Imports System.Runtime.CompilerServices
Imports System.Runtime.Remoting
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Bibliography

Public Class Form1
    'deklarasi objek baru bernama "conn" dari class OleDbConnection. 
    Dim conn As New OleDbConnection
    'deklarasi objek baru bernama "cmd" dari class OleDbCommand. 
    Dim cmd As OleDbCommand
    'deklarasi objek baru bernama "dt" dari class DataTable. 
    Dim dt As New DataTable
    'deklarasi objek baru bernama "da" dari class OleDb
    Dim da As New OleDbDataAdapter(cmd)

    'deklarasi objek bernama "bitmap" dari kelas Bitmap.
    Private bitmap As Bitmap

    'deklarasi Subroutine bernama "viewer". 
    Private Sub viewer()
        'mengatur sumber data untuk objek DataGridView1 sebagai "Nothing".
        DataGridView1.DataSource = Nothing
        'memanggil method Refresh pada objek DataGridView1, yang akan memuat ulang tampilan
        'DataGridView1 dengan data yang baru saja diterapkan.
        DataGridView1.Refresh()

        'buka koneksi ke database
        conn.Open()
        'kode yang memanggil method CreateCommand pada objek "conn" dari kelas OleDbConnection
        cmd = conn.CreateCommand()
        'mengatur CommandText pada objek "cmd" dari kelas OleDbCommand sebagai CommandType.Text. 
        cmd.CommandText = CommandType.Text
        'membuat objek "da" dari kelas OleDbDataAdapter dan menentukan
        'perintah SQL untuk mengambil semua data dari tabel "Database_uasPV". 
        da = New OleDbDataAdapter("select * from Database_uasPV", conn)
        'memanggil method Fill pada objek "da" dari kelas OleDbDataAdapter, dengan objek "dt" dari kelas DataTable sebagai parameter.
        da.Fill(dt)

        'digunakan untuk menyortir data pada DataTable (dt) menurut kolom NIM secara ascending (ASC).
        dt.DefaultView.Sort = "NIM ASC"
        'menetapkan sumber data dari DataTable (dt) untuk menampilkan data dalam bentuk tabel pada kontrol DataGridView1.
        DataGridView1.DataSource = dt.DefaultView
        'digunakan untuk menutup koneksi database yang telah terbuka.
        conn.Close()

        'mengatur lebar kolom pertama dalam tabel DataGridView1 dengan lebar 137 piksel.
        DataGridView1.Columns(0).Width = 137
        DataGridView1.Columns(1).Width = 137
        DataGridView1.Columns(2).Width = 137
        DataGridView1.Columns(3).Width = 137
        DataGridView1.Columns(4).Width = 137
        DataGridView1.Columns(5).Width = 137
    End Sub

    'event handler yang akan dijalankan saat Form1 dimuat
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'digunakan untuk menetapkan string koneksi untuk membuat koneksi dengan database.
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Kuliah\MAPEL\Pemograman Visual\Tugas\UAS\uasPV\uasPV\Database_uasPV.accdb;"
        'digunakan untuk mengatur teks default dalam kontrol txtSearch.
        txtSearch.Text = "Cari..."
        'memanggil Subroutine viewer() yang akan menampilkan data dari database dalam tabel DataGridView1.
        viewer()
        'menambahkan event handler untuk melakukan tindakan saat form ditutup.
        AddHandler Me.FormClosing, AddressOf Form1_FormClosing

    End Sub

    'memunculkan pesan konfirmasi saat pengguna ingin menutup form.
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs)
        'menampung hasil dari MessageBox yang menanyakan apakah pengguna ingin menutup form.
        Dim result As DialogResult = MessageBox.Show("Apakah Anda ingin menutup form ini?", "Konfirmasi", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        'jika hasilnya No, maka akan membatalkan penutupan form dengan mengubah nilai e.Cancel menjadi True.
        If result = DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

    'digunakan untuk menambahkan data baru ke database.
    Private Sub btnTambahData_Click(sender As Object, e As EventArgs) Handles btnTambahData.Click
        'Validasi dilakukan untuk memastikan bahwa semua field tidak kosong.
        If String.IsNullOrEmpty(txtNim.Text) Then
            MessageBox.Show("NIM tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtNamaLengkap.Text) Then
            MessageBox.Show("Nama Lengkap tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtAlamat.Text) Then
            MessageBox.Show("Alamat tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtKelas.Text) Then
            MessageBox.Show("Kelas tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtJurusan.Text) Then
            MessageBox.Show("Jurusan tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtNoHp.Text) Then
            MessageBox.Show("No. Telepon tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Jika tidak ada field yang kosong, maka koneksi dibuka, query SQL dibuat, dan data ditambahkan ke database.
        Else
            Try
                'membuka koneksi ke database.
                conn.Open()
                'membuat perintah baru untuk koneksi yang terbuka.
                cmd = conn.CreateCommand()
                'menentukan teks perintah SQL untuk memasukkan data ke database.
                cmd.CommandText = "INSERT INTO Database_uasPV (NIM, Nama, Alamat, Kelas, Jurusan, Telepon) VALUES (@NIM, @Nama, @Alamat, @Kelas, @Jurusan, @Telepon)"

                'memasukkan nilai dari kotak teks (textbox)
                cmd.Parameters.AddWithValue("@NIM", txtNim.Text)
                cmd.Parameters.AddWithValue("@Nama", txtNamaLengkap.Text)
                cmd.Parameters.AddWithValue("@Alamat", txtAlamat.Text)
                cmd.Parameters.AddWithValue("@Kelas", txtKelas.Text)
                cmd.Parameters.AddWithValue("@Jurusan", txtJurusan.Text)
                cmd.Parameters.AddWithValue("@Telepon", txtNoHp.Text)

                'mengeksekusi perintah SQL tanpa mengembalikan data.
                cmd.ExecuteNonQuery()
                'menutup koneksi ke database.
                conn.Close()

                'membersihkan objek "dt" (data table).
                dt.Clear()
                'memanggil metode "viewer".
                viewer()
                'menampilkan pesan "Data Berhasil Tersimpan" pada kotak dialog
                MessageBox.Show("Data Berhasil Tersimpan", "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'memulai blok penangkapan (catch block) untuk menangkap kesalahan yang mungkin terjadi selama eksekusi kode dalam blok "Try".
            Catch ex As Exception
                'menampilkan pesan kesalahan.
                MessageBox.Show(ex.Message, "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'menutup koneksi ke database.
                conn.Close()
            End Try
        End If
    End Sub

    'event click btnTampilkan yang akan menjalankan perintah dt.Clear() dan viewer().
    Private Sub btnTampilkan_Click(sender As Object, e As EventArgs) Handles btnTampilkan.Click
        'Perintah ini digunakan untuk menghapus isi dari data yang ditampilkan sebelumnya dan menampilkan data pada viewer.
        dt.Clear()
        viewer()
    End Sub

    ' event click btnUpdate yang akan menjalankan perintah update pada database.
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        'Validasi dilakukan untuk memastikan bahwa semua field tidak kosong.
        If String.IsNullOrEmpty(txtNim.Text) Then
            MessageBox.Show("NIM tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtNamaLengkap.Text) Then
            MessageBox.Show("Nama Lengkap tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtAlamat.Text) Then
            MessageBox.Show("Alamat tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtKelas.Text) Then
            MessageBox.Show("Kelas tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtJurusan.Text) Then
            MessageBox.Show("Jurusan tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtNoHp.Text) Then
            MessageBox.Show("No. Telepon tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Jika tidak ada field yang kosong, maka koneksi dibuka, query SQL dibuat, dan data ditambahkan ke database.
        Else
            Try
                'membuka koneksi dengan database
                conn.Open()
                'membuat perintah baru untuk koneksi yang terbuka.
                cmd = conn.CreateCommand()
                'perintah SQL update menggunakan CommandText
                cmd.CommandText = CommandType.Text
                'memperbarui data pada kolom Nama, NIM, Alamat, Kelas, Jurusan, dan Telepon dengan nilai yang diambil dari masing-masing textbox
                'dengan replace untuk menghindari kesalahan syntax SQL.
                cmd.CommandText = "UPDATE Database_uasPV SET Nama = '" & txtNamaLengkap.Text.Replace("'", "''") & "', NIM = '" & txtNim.Text.Replace("'", "''") & "', Alamat = '" & txtAlamat.Text.Replace("'", "''") & "', Kelas = '" & txtKelas.Text.Replace("'", "''") & "', Jurusan = '" & txtJurusan.Text.Replace("'", "''") & "', Telepon = '" & txtNoHp.Text.Replace("'", "''") & "' WHERE NIM = '" & txtNim.Text.Replace("'", "''") & "'"

                'mengeksekusi perintah SQL tanpa mengembalikan data.
                cmd.ExecuteNonQuery()
                'menutup koneksi ke database.
                conn.Close()
                'menampilkan pesan "Data Berhasil Tersimpan" pada kotak dialog
                MessageBox.Show("Data Berhasil Diupdate", "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'membersihkan objek "dt" (data table).
                dt.Clear()
                'memanggil metode "viewer".
                viewer()
                'memulai blok penangkapan (catch block) untuk menangkap kesalahan yang mungkin terjadi selama eksekusi kode dalam blok "Try".
            Catch ex As Exception
                'menampilkan pesan kesalahan.
                MessageBox.Show(ex.Message, "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'menutup koneksi ke database.
                conn.Close()
            End Try
        End If
    End Sub

    'deklarasi event handler untuk event CellClick pada objek DataGridView1.
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'try-catch untuk menangani kemungkinan terjadi error.
        Try
            'mengecek apakah ada baris yang dipilih pada DataGridView1.
            If DataGridView1.SelectedRows.Count > 0 Then
                'mengambil nilai dari masing-masing kolom pada baris yang dipilih dan memasukkannya ke dalam textbox yang sesuai.
                txtNim.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString()
                txtNamaLengkap.Text = DataGridView1.SelectedRows(0).Cells(1).Value.ToString()
                txtAlamat.Text = DataGridView1.SelectedRows(0).Cells(2).Value.ToString()
                txtKelas.Text = DataGridView1.SelectedRows(0).Cells(3).Value.ToString()
                txtJurusan.Text = DataGridView1.SelectedRows(0).Cells(4).Value.ToString()
                txtNoHp.Text = DataGridView1.SelectedRows(0).Cells(5).Value.ToString()
            End If
            'blok catch yang akan menampilkan pesan error jika terjadi kesalahan.
        Catch ex As Exception
            'menampilkan pesan kesalahan.
            MessageBox.Show(ex.Message, "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'menutup koneksi ke database.
            conn.Close()
        End Try
    End Sub

    'deklarasi Subroutine yang dapat dipanggil ketika tombol "btnCari" diklik.
    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click
        'deklarasi variabel "checker" bertipe integer.
        Dim checker As Integer
        Try
            'membuka koneksi dengan database
            conn.Open()
            'membuat perintah baru untuk koneksi yang terbuka.
            cmd = conn.CreateCommand()
            'perintah SQL update menggunakan CommandText
            cmd.CommandText = CommandType.Text
            'menentukan perintah SQL yang akan dikirimkan. Perintah ini akan mengambil data dari tabel "Database_uasPV"
            'yang memiliki kolom "Nama" atau "NIM"
            cmd.CommandText = "SELECT * FROM Database_uasPV WHERE Nama LIKE '%" & txtSearch.Text & "%' OR NIM LIKE '%" & txtSearch.Text & "%'"

            'mengeksekusi perintah SQL tanpa mengembalikan data.
            cmd.ExecuteNonQuery()
            'objek "dt" dideklarasikan sebagai sebuah tabel baru.
            dt = New DataTable()
            'objek "da" dideklarasikan sebagai adapter untuk mengisi tabel "dt".
            da = New OleDbDataAdapter(cmd)
            'isi dari tabel "dt" diisi dengan data yang didapatkan dari eksekusi perintah SQL.
            da.Fill(dt)
            'nilai variabel "checker" diisi dengan jumlah baris pada tabel "dt".
            checker = Convert.ToInt32(dt.Rows.Count.ToString)
            'data pada tabel "dt" ditampilkan pada DataGridView1.
            DataGridView1.DataSource = dt

            'menutup koneksi ke database.
            conn.Close()
            'terdapat percabangan apabila jumlah baris pada tabel "dt" sama dengan nol,
            'maka isi dari textbox "txtSearch" akan diubah menjadi "Cari".
            If (checker = 0) Then
                txtSearch.Text = "Cari"
            End If
            'blok catch yang akan menampilkan pesan error jika terjadi kesalahan.
        Catch ex As Exception
            'menampilkan pesan kesalahan.
            MessageBox.Show(ex.Message, "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'menutup koneksi ke database.
            conn.Close()
        End Try
    End Sub

    'deklarasi Subroutine btnHapus_Click yang akan dijalankan ketika tombol "Hapus" diklik.
    Private Sub btnHapus_Click(sender As Object, e As EventArgs) Handles btnHapus.Click
        'Validasi dilakukan untuk memastikan bahwa semua field tidak kosong.
        If String.IsNullOrEmpty(txtNim.Text) Then
            MessageBox.Show("NIM tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtNamaLengkap.Text) Then
            MessageBox.Show("Nama Lengkap tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtAlamat.Text) Then
            MessageBox.Show("Alamat tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtKelas.Text) Then
            MessageBox.Show("Kelas tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtJurusan.Text) Then
            MessageBox.Show("Jurusan tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(txtNoHp.Text) Then
            MessageBox.Show("No. Telepon tidak boleh kosong", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Jika tidak ada field yang kosong, maka koneksi dibuka, query SQL dibuat, dan data ditambahkan ke database.
        Else
            Try
                'membuka koneksi dengan database
                conn.Open()
                'membuat perintah baru untuk koneksi yang terbuka.
                cmd = conn.CreateCommand()
                'perintah SQL update menggunakan CommandText
                cmd.CommandText = CommandType.Text
                'mendefinisikan teks perintah SQL yang akan dieksekusi sebagai perintah delete dari tabel "Database_uasPV"
                'sesuai NIM yang ada pada textbox txtNim.
                cmd.CommandText = "delete * from Database_uasPV where NIM = '" + txtNim.Text + "' "

                'mengeksekusi perintah SQL tanpa mengembalikan data.
                cmd.ExecuteNonQuery()
                'menutup koneksi ke database.
                conn.Close()
                'menampilkan pesan "Data Berhasil Terhapus" menggunakan MessageBox.
                MessageBox.Show("Data Berhasil Terhapus", "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'memanggil method btnCari_Click dan mempassing object baru dan object EventArgs baru sebagai argumennya.
                btnCari_Click(New Object, New EventArgs())
                'memanggil subroutine viewer untuk mengupdate tampilan data pada DataGridView.
                viewer()

                'membersihkan isi dari textbox txtNim, txtNamaLengkap, txtAlamat, txtJurusan, txtNoHp, txtKelas,
                'dan mengatur isi dari textbox txtSearch menjadi "Cari...".
                txtNim.Text = ""
                txtNamaLengkap.Text = ""
                txtAlamat.Text = ""
                txtJurusan.Text = ""
                txtNoHp.Text = ""
                txtKelas.Text = ""
                txtSearch.Text = "Cari..."
                'blok catch yang akan menampilkan pesan error jika terjadi kesalahan.
            Catch ex As Exception
                'menampilkan pesan kesalahan.
                MessageBox.Show(ex.Message, "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'menutup koneksi ke database.
                conn.Close()
            End Try
        End If
    End Sub

    'deklarasi Subroutine btnRest_Click yang akan dijalankan ketika tombol "Reset" diklik.
    Private Sub btnRest_Click(sender As Object, e As EventArgs) Handles btnRest.Click
        'membersihkan isi dari textbox txtNim, txtNamaLengkap, txtAlamat, txtJurusan, txtNoHp, txtKelas,
        'dan mengatur isi dari textbox txtSearch menjadi "Cari...".
        txtNim.Text = ""
        txtNamaLengkap.Text = ""
        txtAlamat.Text = ""
        txtJurusan.Text = ""
        txtNoHp.Text = ""
        txtKelas.Text = ""
        txtSearch.Text = "Cari..."
    End Sub

    'deklarasi Subroutine btnCetak_Click yang akan dijalankan ketika tombol "Cetak" diklik.
    Private Sub btnCetak_Click(sender As Object, e As EventArgs) Handles btnCetak.Click
        'menentukan variabel height dengan tipe data integer dan mengisi dengan tinggi dari kontrol DataGridView1.
        Dim height As Integer = DataGridView1.Height
        Try
            'mengubah tinggi dari kontrol DataGridView1 dengan jumlah baris dalam DataGridView1
            'dikalikan dengan tinggi dari RowTemplate.
            DataGridView1.Height = DataGridView1.RowCount * DataGridView1.RowTemplate.Height
            'membuat objek baru dari tipe data Bitmap dengan lebar dan tinggi
            'sama dengan lebar dan tinggi dari kontrol DataGridView1.
            bitmap = New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
            'menggambar kontrol DataGridView1 pada objek bitmap.
            DataGridView1.DrawToBitmap(bitmap, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))
            'menentukan document yang akan dicetak.
            PrintPreviewDialog1.Document = PrintDocument1
            'menentukan zoom dari tampilan preview cetak.
            PrintPreviewDialog1.PrintPreviewControl.Zoom = 1
            'menampilkan dialog preview cetak.
            PrintPreviewDialog1.ShowDialog()
            'mengubah tinggi dari kontrol DataGridView1 kembali ke tinggi semula.
            DataGridView1.Height = height
            'blok catch yang akan menampilkan pesan error jika terjadi kesalahan.
        Catch ex As Exception
            'menampilkan pesan kesalahan.
            MessageBox.Show(ex.Message, "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'menutup koneksi ke database.
            conn.Close()
        End Try

    End Sub

    'deklarasi Subroutine yang akan dipanggil pada saat print document dipanggil.
    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Try
            'Menggambar bitmap pada posisi x = 0 dan y = 0.
            e.Graphics.DrawImage(bitmap, 0, 0)
            'Mendeklarasikan variabel bernama recP sebagai tipe RectangleF yang mengambil
            'nilai dari PrintableArea pada PrintPageEventArgs
            Dim recP As RectangleF = e.PageSettings.PrintableArea
            'jika tinggi dari DataGridView1 sama dengan tinggi dari recP dan recP lebih besar dari 0,
            'maka halaman akan terus berlanjut.
            If Me.DataGridView1.Height = recP.Height > 0 Then e.HasMorePages = True
            'blok catch yang akan menampilkan pesan error jika terjadi kesalahan.
        Catch ex As Exception
            'menampilkan pesan kesalahan.
            MessageBox.Show(ex.Message, "Database Mahasiswa", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'menutup koneksi ke database.
            conn.Close()
        End Try
    End Sub

    'event procedure saat cursor mouse masuk ke dalam area teks search.
    Private Sub txtSearch_MouseEnter(sender As Object, e As EventArgs) Handles txtSearch.MouseEnter
        'mengosongkan teks dalam textbox search.
        txtSearch.Text = ""
        'memberikan fokus ke textbox search.
        txtSearch.Focus()
        'engubah warna font teks dalam textbox search menjadi hitam.
        txtSearch.ForeColor = Color.Black
    End Sub

    'event procedure saat cursor mouse keluar dari area teks search.
    Private Sub txtSearch_MouseLeave(sender As Object, e As EventArgs) Handles txtSearch.MouseLeave
        'memeriksa apakah teks dalam textbox search kosong.
        If (txtSearch.Text = "") Then
            'jika teks dalam textbox search kosong maka teks dalam textbox akan diisi dengan string kosong.
            txtSearch.Text = ""
            'mengubah warna font teks dalam textbox search menjadi Silver.
            txtSearch.ForeColor = Color.Silver

        End If

    End Sub

    'Event handler untuk button "btnExport"
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        'Mendeklarasikan objek baru 
        Dim saveFileDialog1 As New SaveFileDialog()

        'Menentukan filter yang akan diterapkan pada jendela dialog "Save File".
        'Filter ini memungkinkan pengguna hanya menyimpan file dengan ekstensi .xlsx atau semua jenis file.
        saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"

        'Menentukan filter pertama yang diterapkan sebagai filter default saat jendela dialog "Save File" ditampilkan.
        'Dalam hal ini, filter pertama adalah "Excel Files (*.xlsx)".
        saveFileDialog1.FilterIndex = 1

        'Menentukan apakah direktori sebelumnya harus dikembalikan setelah jendela dialog "Save File" ditutup.
        saveFileDialog1.RestoreDirectory = True

        'Memeriksa apakah pengguna telah memilih file yang akan disimpan dan menekan tombol "OK".
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            'Mendeklarasikan objek baru bernama wb sebagai objek XLWorkbook.
            Dim wb As New XLWorkbook()
            'Mendeklarasikan objek baru bernama ws sebagai objek baru yang ditambahkan ke wb. Nama sheet yang ditambahkan adalah "Data Mahasiswa".
            Dim ws = wb.Worksheets.Add("Data Mahasiswa")

            'Loop untuk menambahkan header kolom pada sheet "Data Mahasiswa".
            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                'Menambahkan header kolom pada sheet "Data Mahasiswa".
                ws.Cell(1, i + 1).Value = DataGridView1.Columns(i).HeaderText
            Next

            'Loop untuk menambahkan data baris pada sheet "Data Mahasiswa".
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                'Loop untuk menambahkan data pada setiap kolom dalam baris.
                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    'Memeriksa apakah nilai pada sel tertentu adalah Nothing.
                    If DataGridView1(j, i).Value IsNot Nothing Then
                        'Menambahkan nilai sel pada sheet "Data Mahasiswa".
                        ws.Cell(i + 2, j + 1).Value = DataGridView1(j, i).Value.ToString()
                    Else
                        'Memasukkan string kosong ("") ke dalam sel (i + 2, j + 1) dari lembar kerja ws.
                        'Jika nilai dari DataGridView1 (j, i) tidak ada, maka akan ditetapkan sebagai string kosong.
                        ws.Cell(i + 2, j + 1).Value = ""
                        'Akhir dari blok If ... Then ... Else yang memeriksa apakah nilai dari DataGridView1 (j, i) ada atau tidak.
                    End If
                    'Menandakan akhir dari perulangan For yang melakukan iterasi melalui kolom dalam DataGridView1.
                Next
                'Menandakan akhir dari perulangan For yang melakukan iterasi melalui kolom dalam DataGridView1.
            Next
            'Menyimpan workbook wb dengan nama file yang ditentukan oleh saveFileDialog1.FileName.
            wb.SaveAs(saveFileDialog1.FileName)
            'Menampilkan pesan box dengan pesan "Data berhasil diekspor ke file Excel", judul "Informasi", tombol OK, dan ikon informasi.
            MessageBox.Show("Data berhasil diekspor ke file Excel", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Akhir dari blok If ... Then ... End If yang memeriksa apakah pengguna memilih file untuk disimpan melalui saveFileDialog1.
        End If
    End Sub

    'Event handler untuk button "btnBuka"
    Private Sub btnBuka_Click(sender As Object, e As EventArgs) Handles btnBuka.Click
        'Mendeklarasikan objek openFileDialog dari kelas OpenFileDialog.
        Dim openFileDialog As New OpenFileDialog()
        'Menentukan filter untuk menampilkan hanya file dengan format .xlsx pada jendela openFileDialog.
        openFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        'Statement If yang akan mengeksekusi blok kode di dalamnya jika user memilih file dan menekan tombol OK.
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            'Deklarasi string filePath yang menyimpan lokasi file yang dipilih oleh user.
            Dim filePath As String = openFileDialog.FileName
            'Deklarasi objek workbook dari kelas XLWorkbook dengan parameter filePath sebagai lokasi file yang akan dibuka.
            Dim workbook As New XLWorkbook(filePath)
            'Deklarasi objek worksheet dari interface IXLWorksheet yang mengacu pada worksheet pertama dalam workbook.
            Dim worksheet As IXLWorksheet = workbook.Worksheet(1)
            'Deklarasi objek dataTable dari kelas DataTable yang akan digunakan sebagai sumber data untuk DataGridView.
            Dim dataTable As New DataTable()
            'For loop yang akan melakukan iterasi untuk setiap kolom dalam worksheet.
            For i As Integer = 1 To worksheet.LastRowUsed().LastCellUsed().Address.ColumnNumber
                'Statement yang menambahkan kolom baru ke dataTable dengan judul sesuai dengan isi dari sel pada
                'baris pertama dan kolom i dalam worksheet.
                dataTable.Columns.Add(worksheet.Cell(1, i).Value.ToString(), GetType(String))
            Next
            'For Each Loop yang akan melakukan iterasi untuk setiap baris dalam worksheet, dengan memulai iterasi dari baris kedua (.Skip(1)).
            For Each row As IXLRow In worksheet.RowsUsed().Skip(1)
                'Deklarasi objek dataRow dari kelas DataRow yang akan digunakan untuk menambahkan baris baru ke dataTable.
                Dim dataRow As DataRow = dataTable.NewRow()
                'Perulangan untuk setiap kolom pada baris yang sedang diproses. row.LastCellUsed().Address.ColumnNumber
                'mengambil jumlah kolom terakhir yang terisi pada baris tersebut.
                For j As Integer = 1 To row.LastCellUsed().Address.ColumnNumber
                    'Proses pengisian data pada objek dataRow. worksheet.Cell(1, j).Value.ToString() mengambil nama kolom
                    'dari baris pertama (kolom header) dan row.Cell(j).Value.ToString() mengambil nilai dari sel pada baris saat ini dan kolom j.
                    dataRow(worksheet.Cell(1, j).Value.ToString()) = row.Cell(j).Value.ToString()
                Next
                'Menambahkan objek dataRow yang sudah diisi ke objek dataTable.
                dataTable.Rows.Add(dataRow)
            Next
            'Menetapkan objek dataTable sebagai sumber data untuk objek DataGridView1.
            DataGridView1.DataSource = dataTable
        End If
    End Sub
End Class
