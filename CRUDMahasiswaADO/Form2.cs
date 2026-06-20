using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace CRUDMahasiswaADO
{
    public partial class Form2 : Form
    {
        // FIX 1: Connection string dipindah ke satu tempat (tidak duplikat dengan Form2.Designer)
        static string connectionString =
            "Data Source=LAPTOP-4VAVDOFH\\WAWANLOMBOK;Initial Catalog=DBAkademikADO;User ID=sa;Password=123";

        string Prodi;
        DateTime TglMasuk;

        // FIX 2: Konstruktor menerima parameter prodi dan tglMasuk (sesuai modul langkah 12-13)
        public Form2(string prodi, DateTime tglMasuk)
        {
            InitializeComponent();
            Prodi = prodi;
            TglMasuk = tglMasuk;

            // FIX 3: Load data dan set report langsung di konstruktor (sesuai modul langkah 13),
            //        bukan di Form_Load, agar tidak terjadi double load
            LoadReport();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            // Dikosongkan — proses load sudah dilakukan di konstruktor
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
            // Dikosongkan — hanya untuk mencegah error Designer
        }

        private void LoadReport()
        {
            try
            {
                DataTable dtMahasiswa = new DataTable();

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    // FIX 4: Buka koneksi sebelum menggunakan SqlDataAdapter
                    conn.Open();

                    SqlCommand cmd = new SqlCommand("sp_Report", conn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@inProdi", SqlDbType.VarChar, 50).Value = Prodi;
                    cmd.Parameters.Add("@inTglMsuk", SqlDbType.VarChar, 4).Value = TglMasuk.Year.ToString();

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dtMahasiswa);
                }

                if (dtMahasiswa.Rows.Count == 0)
                {
                    MessageBox.Show("Data laporan tidak ditemukan.");
                    this.Close();
                    return;
                }

                // FIX 5: Gunakan instance ListMahasiswa1 yang sudah ada di Designer
                //        (bukan ReportDocument baru + Path.Combine) — sesuai modul langkah 13
                ListMahasiswa1.SetDataSource(dtMahasiswa);
                crystalReportViewer1.ReportSource = ListMahasiswa1;
                crystalReportViewer1.Refresh();
            }
            catch (Exception ex)
            {
                string pesan = ex.Message;
                if (ex.InnerException != null)
                    pesan += "\n\nInner: " + ex.InnerException.Message;
                if (ex.InnerException?.InnerException != null)
                    pesan += "\n\nInner2: " + ex.InnerException.InnerException.Message;
                MessageBox.Show(pesan);
            }
        }
    }
}