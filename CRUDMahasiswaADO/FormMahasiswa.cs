using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using ExcelDataReader;

namespace CRUDMahasiswaADO
{
    public partial class FormMahasiswa : Form
    {
        DAL dbLogic = new DAL();

        private readonly SqlConnection conn;
        private readonly string connectionString =
            "Data Source=LAPTOP-4VAVDOFH\\WAWANLOMBOK;Initial Catalog=DBAkademikADO;User ID=sa;Password=123;";

        private BindingSource bindingSource = new BindingSource();
        private DataTable dtMahasiswa = new DataTable();
        public FormMahasiswa()
        {
            InitializeComponent();
        }

        private void HitungTotal()
        {
            try
            {
                int total = (dbLogic.CountMhs().Equals(DBNull.Value)) ? 0 : dbLogic.CountMhs();

                lblTotal.Text = "Total Mahasiswa : " + total;
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("Gagal load data : " + ex.Message);
            }
        }

        private void LoadData()
        {
            try
            {
                mahasiswaBindingSource.DataSource = dbLogic.GetMhs();
                dgvMahasiswa.DataSource = mahasiswaBindingSource;
                DataGridViewImageColumn fotoColumn = (DataGridViewImageColumn)dgvMahasiswa.Columns["Foto"];
                fotoColumn.ImageLayout = DataGridViewImageCellLayout.Stretch;

                HitungTotal();
                dgvMahasiswa.Enabled = true;
                ImpDb.Enabled = false;
                btnInsert.Enabled = true;
                btnUpdate.Enabled = true;
                btnDelete.Enabled = true;
                btnCari.Enabled = true;
                btnLoad.Enabled = true;
                btnReset.Enabled = true;
                btnTest.Enabled = true;
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("Gagal load data: " + ex.Message);
            }
        }

        private void FormMahasiswa_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dBAkademikADODataSet.Mahasiswa' table. You can move, or remove it, as needed.
            cmbJK.DataSource = new string[] { "L", "P" };
            dgvMahasiswa.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvMahasiswa.MultiSelect = false;
            dgvMahasiswa.ReadOnly = true;
            dgvMahasiswa.AllowUserToAddRows = false;
            dgvMahasiswa.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            LoadData();

        }

        
        private void label1_Click(object sender, EventArgs e)
        {

        }



        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    MessageBox.Show("Koneksi Berhasil");
                }
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("Error : " + ex.Message);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
           LoadData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataRowView row = (DataRowView)bindingSource[e.RowIndex];

                txtNIM.Text = row[0].ToString();
                txtNama.Text = row[1].ToString();
                cmbJK.Text = row[2].ToString();
                dtpTanggalLahir.Value = Convert.ToDateTime(row[3]);
                txtAlamat.Text = row[4].ToString();
                txtKodeProdi.Text = row[6].ToString();

                if (row[5] != DBNull.Value)
                {
                    byte[] imgBytes = (byte[])row[5];

                    using (MemoryStream ms = new MemoryStream(imgBytes))
                    {
                        fotoMhs.Image = Image.FromStream(ms);
                        fotoMhs.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                }
                else
                {
                    fotoMhs.Image = null;
                }

                txtNIM.Enabled = false;
            }
        }

        private byte[] ConvertImageToBytes(PictureBox pb)
        {
            if (pb.Image == null)
                return null;

            using (MemoryStream ms = new MemoryStream())
            {
                pb.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                return ms.ToArray();
            }
        }
        private void btnInsert_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtNIM.Text == "" || txtNama.Text == "" || cmbJK.Text == "" ||
                    txtAlamat.Text == "" || txtKodeProdi.Text == "")
                {
                    MessageBox.Show("Semua data wajib diisi!");
                    return;
                }

                byte[] imgBytes = ConvertImageToBytes(fotoMhs);

                dbLogic.InsertMhs(
         txtNIM.Text,
         txtNama.Text,
         txtAlamat.Text,
         cmbJK.Text,
         dtpTanggalLahir.Value.Date,
         txtKodeProdi.Text,
         imgBytes
     );

                MessageBox.Show("Data mahasiswa berhasil ditambahkan");

                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error : " + ex.Message);
            }
        }
        


        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtNIM.Text == "" || txtNama.Text == "" || cmbJK.Text == "" ||
                    txtAlamat.Text == "" || txtKodeProdi.Text == "")
                {
                    MessageBox.Show("Pilih data yang ingin diubah terlebih dahulu!");
                    return;
                }

                byte[] imgBytes = ConvertImageToBytes(fotoMhs);

                dbLogic.UpdateMhs(
                   txtNIM.Text,
                   txtNama.Text,
                   txtAlamat.Text,
                   cmbJK.Text,
                   dtpTanggalLahir.Value.Date,
                   txtKodeProdi.Text,
                   imgBytes
               );

                MessageBox.Show("Data mahasiswa berhasil diubah");

                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error : " + ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dg = MessageBox.Show(
                    "Yakin ingin menghapus data?",
                    "Konfirmasi",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (dg == DialogResult.Yes)
                {
                    dbLogic.DeleteMhs(txtNIM.Text);

                    MessageBox.Show("Data mahasiswa berhasil dihapus");

                    ClearForm();
                    LoadData();
                }
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error : " + ex.Message);
            }
        }

        private void ClearForm()
        {
            txtNIM.Enabled = true;
            txtNIM.Clear();
            txtNama.Clear();
            cmbJK.SelectedIndex = -1;
            txtAlamat.Clear();
            txtKodeProdi.Clear();
            dtpTanggalLahir.Value = DateTime.Now;

            // Kalau PictureBox kamu namanya beda, sesuaikan
            fotoMhs.Image = null;

            txtNIM.Focus();
        }

        private void BindControls()
        {
            txtNIM.DataBindings.Clear();
            txtNama.DataBindings.Clear();
            cmbJK.DataBindings.Clear();
            dtpTanggalLahir.DataBindings.Clear();
            txtAlamat.DataBindings.Clear();
            txtKodeProdi.DataBindings.Clear();

            txtNIM.DataBindings.Add("Text", bindingSource, "NIM");
            txtNama.DataBindings.Add("Text", bindingSource, "Nama");
            cmbJK.DataBindings.Add("Text", bindingSource, "JenisKelamin");
            dtpTanggalLahir.DataBindings.Add("Value", bindingSource, "TanggalLahir");
            txtAlamat.DataBindings.Add("Text", bindingSource, "Alamat");
            txtKodeProdi.DataBindings.Add("Text", bindingSource, "KodeProdi");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dbLogic.resetData();

                MessageBox.Show("Data berhasil direset");

                LoadData();
            }
            catch (SqlException ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("SQL Error : " + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error : " + ex.Message);
            }
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            try
            {
                dbLogic.testInject(txtNIM.Text);

                LoadData();
            }
            catch (SqlException ex)
            {
                if (ex.Message.Contains("safe"))
                {
                    simpanLog(ex.Message);
                    MessageBox.Show("SQL Error : Unsafe UPDATE operation not allowed");
                }
                else
                {
                    simpanLog(ex.Message);
                    MessageBox.Show("SQL Error : " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("General Error : " + ex.Message);
            }
        }

        private void FormMahasiswa_Load_1(object sender, EventArgs e)
        {

        }

        private void btnRekapData_Click(object sender, EventArgs e)

        {
            Form1 fm3 = new Form1();
            fm3.Show();
            this.Hide();
        }




        public void simpanLog(string message)
        {
            dbLogic.InsertLog(message);
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fotoMhs.Image = Image.FromFile(ofd.FileName);
                fotoMhs.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }

        private void btnImpExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });

                            DataTable dt = result.Tables[0];

                            dgvMahasiswa.DataSource = dt;
                            dgvMahasiswa.Enabled = false;

                            btnLoad.Enabled = true;
                            btnInsert.Enabled = false;
                            btnUpdate.Enabled = false;
                            btnDelete.Enabled = false;
                            btnCari.Enabled = false;
                            btnLoad.Enabled = false;
                            btnReset.Enabled = false;
                            btnTest.Enabled = false;
                        }
                    }
                }
            }
        }


        private void ImpDb_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = (DataTable)dgvMahasiswa.DataSource;

                if (dt == null || dt.Rows.Count == 0)
                {
                    MessageBox.Show("Tidak ada data untuk diimport.");
                    return;
                }

                int sukses = 0;

                foreach (DataRow row in dt.Rows)
                {
                    string nim = row["NIM"].ToString().Trim();
                    string nama = row["Nama"].ToString().Trim();
                    string jk = row["JenisKelamin"].ToString().Trim();
                    string alamat = row["Alamat"].ToString().Trim();
                    string kodeProdi = row["NamaProdi"].ToString().Trim();
                    string fotoPath = row.Table.Columns.Contains("FotoPath")
                                            ? row["FotoPath"].ToString().Trim()
                                            : string.Empty;

                    if (string.IsNullOrEmpty(nim) || string.IsNullOrEmpty(nama))
                        continue;

                    DateTime tglLahir;
                    if (!DateTime.TryParse(row["TanggalLahir"].ToString(), out tglLahir))
                        continue;

                    byte[] fotoBytes = null;
                    if (!string.IsNullOrWhiteSpace(fotoPath) && File.Exists(fotoPath))
                        fotoBytes = File.ReadAllBytes(fotoPath);

                    dbLogic.InsertMhs(nim, nama, alamat, jk, tglLahir, kodeProdi, fotoBytes);
                    sukses++;
                }

                MessageBox.Show("Data mahasiswa berhasil ditambahkan");
                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                simpanLog("Rollback Insert :" + ex.Message);
                MessageBox.Show("SQL Error :" + ex.Message);
            }
            catch (Exception ex)
            {
                simpanLog("General Error :" + ex.Message);
                MessageBox.Show("General Error :" + ex.Message);
            }
        }       }
}