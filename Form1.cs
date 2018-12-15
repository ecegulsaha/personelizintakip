using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;

namespace PersonelIzinTakipProgramı
{
    public partial class Form1 : Form
    {
        string connectionString = "Data Source=ASUSPC;Initial Catalog=PersonelIzin ";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;

            timer2.Enabled = true;

            GridDoldur();
            IzinGridiDoldur();

            YILLIK_İZİN_EKLE_GUNCELLE();//30-20 İZİN EKLENMESİ
        }

        //personel bilgileriyle ilgili//
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //personel datagridview tıklanma durumu
        {
            textBox1.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[0].Value));
            textBox2.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[1].Value));
            textBox3.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[2].Value));
            textBox4.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[3].Value));
            textBox5.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[4].Value));
            dateTimePicker1.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[5].Value));
            comboBox4.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[6].Value));
            maskedTextBox1.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[7].Value));
            textBox8.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[8].Value));
            comboBox1.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[9].Value));
            comboBox5.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[10].Value));
            comboBox6.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[11].Value));
            dateTimePicker2.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[12].Value));
            dateTimePicker3.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[13].Value));
            textBox16.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[14].Value));
            textBox21.Text = Convert.ToString((dataGridView1.CurrentRow.Cells[15].Value));
        }

        private void GridDoldur()//datagridview doldurma
        {
            SqlConnection sqlConnection = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();

            cmd.CommandText = "SELECT * from Personel";
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection;

            sqlConnection.Open();

            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet);

            DataTable dt = new DataTable();
            DataRow dr;
            dt.Columns.Add(new DataColumn("Personel_Id", typeof(string)));
            dt.Columns.Add(new DataColumn("Sicil_No", typeof(string)));
            dt.Columns.Add(new DataColumn("TC", typeof(string)));
            dt.Columns.Add(new DataColumn("Ad", typeof(string)));
            dt.Columns.Add(new DataColumn("Soyad", typeof(string)));
            dt.Columns.Add(new DataColumn("Dogum_Tarihi", typeof(string)));
            dt.Columns.Add(new DataColumn("Dogum_Yeri", typeof(string)));
            dt.Columns.Add(new DataColumn("Telefon", typeof(string)));
            dt.Columns.Add(new DataColumn("Adres", typeof(string)));
            dt.Columns.Add(new DataColumn("Durumu", typeof(string)));
            dt.Columns.Add(new DataColumn("Servisi", typeof(string)));
            dt.Columns.Add(new DataColumn("Gorev_Unvanı", typeof(string)));
            dt.Columns.Add(new DataColumn("Memuriyet_Baslangic_Tarihi", typeof(string)));
            dt.Columns.Add(new DataColumn("Karayollarına_Baslangic_Tarihi", typeof(string)));
            dt.Columns.Add(new DataColumn("Görev_Süresi", typeof(string)));
            dt.Columns.Add(new DataColumn("İZİN_GECERLİLİK_TARİHİ_YIL", typeof(string)));

            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            {
                dr = dt.NewRow();
                dr[0] = dataSet.Tables[0].Rows[i]["Personel_Id"].ToString();
                dr[1] = dataSet.Tables[0].Rows[i]["Sicil_No"].ToString();
                dr[2] = dataSet.Tables[0].Rows[i]["TC"].ToString();
                dr[3] = dataSet.Tables[0].Rows[i]["Ad"].ToString();
                dr[4] = dataSet.Tables[0].Rows[i]["Soyad"].ToString();
                dr[5] = dataSet.Tables[0].Rows[i]["Dogum_Tarihi"].ToString();
                dr[6] = dataSet.Tables[0].Rows[i]["Dogum_Yeri"].ToString();
                dr[7] = dataSet.Tables[0].Rows[i]["Telefon"].ToString();
                dr[8] = dataSet.Tables[0].Rows[i]["Adres"].ToString();
                dr[9] = dataSet.Tables[0].Rows[i]["Durumu"].ToString();
                dr[10] = dataSet.Tables[0].Rows[i]["Servisi"].ToString();
                dr[11] = dataSet.Tables[0].Rows[i]["Gorev_Unvanı"].ToString();
                dr[12] = dataSet.Tables[0].Rows[i]["Memuriyet_Baslangic_Tarihi"].ToString();
                dr[13] = dataSet.Tables[0].Rows[i]["Karayollarına_Baslangic_Tarihi"].ToString();
                dr[14] = dataSet.Tables[0].Rows[i]["Gorev_Suresi"].ToString();
                dr[15] = dataSet.Tables[0].Rows[i]["İzin_Gecerlilik_Tarihi"].ToString();
                dt.Rows.Add(dr);

                //DATAGRİD RENKLENDİRME
                if (i % 2 == 0)
                { dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Silver; }
                else { }
                ///////////////////////


            }

            DataView dv = new DataView(dt);

            dataGridView1.DataSource = dv;
            sqlConnection.Close();
        }

        private void button2_Click(object sender, EventArgs e)//personel sil butonu--sicil no ile siliniyor...
        {

            if (!(textBox2.Text == ""))
            {
                if (MessageBox.Show("Personel bilgisi silmek istediğinize emin misiniz?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string Sicil_No = textBox2.Text;
                    try
                    {
                        using (var sc = new SqlConnection(connectionString))
                        using (var cmd = sc.CreateCommand())
                        {
                            sc.Open();
                            cmd.CommandText = "DELETE FROM Personel WHERE Sicil_No = @Sicil_No";
                            cmd.Parameters.AddWithValue("@Sicil_No", Sicil_No);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show(Sicil_No + " sicil nolu personel kaydı silindi!!!");
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(Sicil_No + " sicil nolu personel kaydı silinemedi!!!");
                    }
                    GridDoldur();
                }
                else MessageBox.Show("Personel bilgisi silmeyi iptal ettiniz!!!");
            }
            else
            {
                MessageBox.Show("Silinecek personel için Sicil_No girmelisiniz...");
            }
        }

        private void button1_Click(object sender, EventArgs e)//personel ekle butonu
        {
            if (!(textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == ""
                || maskedTextBox1.Text == "" || comboBox4.Text == "" || textBox8.Text == "" || comboBox5.Text == "" || comboBox6.Text == ""
                || dateTimePicker1.Text == "" || dateTimePicker2.Text == "" || dateTimePicker3.Text == "" || comboBox1.Text == ""))
            {

                if (MessageBox.Show("Personel eklemek istediğinize emin misiniz?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    //gorev suresi için
                    TimeSpan GunFarki = DateTime.Now.Date.Subtract(Convert.ToDateTime(dateTimePicker2.Text));
                    textBox16.Text = GunFarki.Days.ToString();
                    //------------------//

                    //izin geçerlilik tarihi için
                    if (comboBox1.Text == "İŞÇİ") textBox21.Text = "1";
                    if (comboBox1.Text == "MEMUR") textBox21.Text = "2";
                    //------------------//

                    SqlConnection connection = new SqlConnection(connectionString);
                    connection.Open();
                    string varmiCommandText = "select count(*) as count from Personel where Sicil_No='" + textBox2.Text + "'";

                    SqlCommand command2 = new SqlCommand(varmiCommandText);
                    command2.Connection = connection;
                    SqlDataReader reader = command2.ExecuteReader();
                    int sayii = 0;
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            try
                            {
                                sayii = Convert.ToInt32(reader["count"]);
                            }
                            catch (Exception)
                            {


                            }
                        }

                    }

                    connection.Close();

                    if (sayii == 0)
                    {

                        System.Data.DataTable dt2 = new System.Data.DataTable();
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        SqlCommand command = new SqlCommand();

                        command.Connection = connection;
                        connection.Open();

                        string sql = "INSERT INTO [dbo].[Personel]([Sicil_No],[TC],[Ad],[Soyad],[Dogum_Tarihi],[Dogum_Yeri],[Telefon],[Adres],[Durumu],[Servisi],[Gorev_Unvanı],[Memuriyet_Baslangic_Tarihi],[Karayollarına_Baslangic_Tarihi],[Gorev_Suresi],[İzin_Gecerlilik_Tarihi])  VALUES(@sicil,@tc,@ad,@soy,@dt,@dy,@tel,@adres,@d,@s,@gu,@mbt,@kbt,@gsuresi,@igt)";

                        SqlCommand cmd = new SqlCommand(sql, connection);

                        cmd.Parameters.Add("@sicil", SqlDbType.NVarChar, 8).Value = textBox2.Text;
                        cmd.Parameters.Add("@tc", SqlDbType.NVarChar, 11).Value = textBox3.Text;
                        cmd.Parameters.Add("@ad", SqlDbType.NVarChar, 50).Value = textBox4.Text;
                        cmd.Parameters.Add("@soy", SqlDbType.NVarChar, 50).Value = textBox5.Text;
                        cmd.Parameters.Add("@dt", SqlDbType.Date).Value = Convert.ToDateTime(dateTimePicker1.Text);
                        cmd.Parameters.Add("@dy", SqlDbType.NVarChar, 50).Value = comboBox4.Text;
                        cmd.Parameters.Add("@tel", SqlDbType.NVarChar, 50).Value = maskedTextBox1.Text;
                        cmd.Parameters.Add("@adres", SqlDbType.NVarChar, 50).Value = textBox8.Text;
                        cmd.Parameters.Add("@d", SqlDbType.NVarChar, 50).Value = Convert.ToString(comboBox1.Text);
                        cmd.Parameters.Add("@s", SqlDbType.NVarChar, 50).Value = comboBox5.Text;
                        cmd.Parameters.Add("@gu", SqlDbType.NVarChar, 50).Value = comboBox6.Text;
                        cmd.Parameters.Add("@mbt", SqlDbType.Date).Value = Convert.ToDateTime(dateTimePicker2.Text);
                        cmd.Parameters.Add("@kbt", SqlDbType.Date).Value = Convert.ToDateTime(dateTimePicker3.Text);
                        cmd.Parameters.Add("@gsuresi", SqlDbType.Int).Value = Convert.ToInt32(textBox16.Text);
                        cmd.Parameters.Add("@igt", SqlDbType.Int).Value = Convert.ToInt32(textBox21.Text);

                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();

                        GridDoldur();
                        MessageBox.Show("Kayit eklendi!!!");
                    }
                    else
                    {
                        MessageBox.Show("Kayit mevcut...Bilgilerinizi kontrol edin!!!");
                    }
                }
                else MessageBox.Show("Personel eklemeyi iptal ettiniz!!!");

            }
            else
            {
                MessageBox.Show("Boş alanlar mevcut önce onları doldurun !!!");
            }
        }

        private void button3_Click(object sender, EventArgs e)//personel güncelle butonu
        {
            if ((textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == ""
                 || maskedTextBox1.Text == "" || comboBox4.Text == "" || textBox8.Text == "" || comboBox5.Text == "" || comboBox6.Text == ""
                || dateTimePicker1.Text == "" || dateTimePicker2.Text == "" || dateTimePicker3.Text == "" || comboBox1.Text == ""))
            {
                MessageBox.Show("Boş alanlar mevcut önce onları doldurun !!!");
            }
            else
            {
                if (MessageBox.Show("Personel bilgisi güncellemek istediğinize emin misiniz?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    //gorev suresi için
                    TimeSpan GunFarki = DateTime.Now.Date.Subtract(Convert.ToDateTime(dateTimePicker2.Text));
                    textBox16.Text = GunFarki.Days.ToString();
                    //------------------//

                    //izin geçerlilik tarihi için
                    if (comboBox1.Text == "İŞÇİ") textBox21.Text = "1";
                    if (comboBox1.Text == "MEMUR") textBox21.Text = "2";
                    //------------------//
                    string sql = "UPDATE [dbo].[Personel] SET [Sicil_No] =@sicil,[TC] =@tc,[Ad]=@ad ,Soyad=@soy ,Dogum_Tarihi=@dt ,Dogum_Yeri=@dy ,Telefon=@tel ,Adres=@adres, Durumu=@d, Servisi=@s, Gorev_Unvanı=@gu ,Memuriyet_Baslangic_Tarihi=@mbt, Karayollarına_Baslangic_Tarihi=@kbt, Gorev_Suresi=@gsuresi,İzin_Gecerlilik_Tarihi=@igt WHERE Sicil_No=@sicil";


                    SqlConnection connection = new SqlConnection(connectionString);
                    SqlCommand cmd = new SqlCommand(sql, connection);
                    cmd.Connection = connection;
                    connection.Open();

                    cmd.Parameters.Add("@sicil", SqlDbType.NVarChar, 8).Value = textBox2.Text;
                    cmd.Parameters.Add("@tc", SqlDbType.NVarChar, 11).Value = textBox3.Text;
                    cmd.Parameters.Add("@ad", SqlDbType.NVarChar, 50).Value = textBox4.Text;
                    cmd.Parameters.Add("@soy", SqlDbType.NVarChar, 50).Value = textBox5.Text;
                    cmd.Parameters.Add("@dt", SqlDbType.Date).Value = dateTimePicker1.Value.Date;
                    cmd.Parameters.Add("@dy", SqlDbType.NVarChar, 50).Value = comboBox4.Text;
                    cmd.Parameters.Add("@tel", SqlDbType.NVarChar, 50).Value = maskedTextBox1.Text;
                    cmd.Parameters.Add("@adres", SqlDbType.NVarChar, 50).Value = textBox8.Text;
                    cmd.Parameters.Add("@d", SqlDbType.NVarChar, 50).Value = Convert.ToString(comboBox1.Text);
                    cmd.Parameters.Add("@s", SqlDbType.NVarChar, 50).Value = comboBox5.Text;
                    cmd.Parameters.Add("@gu", SqlDbType.NVarChar, 50).Value = comboBox6.Text;
                    cmd.Parameters.Add("@mbt", SqlDbType.Date).Value = dateTimePicker2.Value.Date;
                    cmd.Parameters.Add("@kbt", SqlDbType.Date).Value = dateTimePicker3.Value.Date;
                    cmd.Parameters.Add("@gsuresi", SqlDbType.Int).Value = Convert.ToInt32(textBox16.Text);
                    cmd.Parameters.Add("@igt", SqlDbType.Int).Value = Convert.ToInt32(textBox21.Text);

                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                    GridDoldur();
                    MessageBox.Show("Kayıt güncellendi !!!");
                }
                else MessageBox.Show("Personel bilgi güncellemeyi iptal ettiniz!!!");
            }
        }

        private void button4_Click(object sender, EventArgs e)//PERSONEL EXCEL E AKTAR
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);

            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            int StartCol = 1;

            int StartRow = 1;

            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {

                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];

                myRange.Value2 = dataGridView1.Columns[j].HeaderText;

            }

            StartRow++;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {

                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    try
                    {

                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];

                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;

                    }

                    catch
                    {

                        ;

                    }

                }

            }

        }

        private void button9_Click(object sender, EventArgs e)//personel bilgi arama butonu
        {
            if (textBox23.Text == "" || comboBox3.Text == "")
            {
                MessageBox.Show("Boş alanlar mevcut !!!\nÖnce onları doldurun !!!");
            }
            else
            {

                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();

                string varmiCommandText;
                int sayii = 0;

                if (comboBox3.Text == "Sicil_No")
                {
                    varmiCommandText = "select count(*) as count from Personel where Sicil_No='" + textBox23.Text + "'";
                    SqlCommand command2 = new SqlCommand(varmiCommandText);
                    command2.Connection = connection;
                    SqlDataReader reader = command2.ExecuteReader();


                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            try
                            { sayii = Convert.ToInt32(reader["count"]); }
                            catch (Exception) { }
                        }
                    }

                }

                if (comboBox3.Text == "TC")
                {
                    varmiCommandText = "select count(*) as count from Personel where TC='" + textBox23.Text + "'";
                    SqlCommand command2 = new SqlCommand(varmiCommandText);
                    command2.Connection = connection;
                    SqlDataReader reader = command2.ExecuteReader();

                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            try
                            { sayii = Convert.ToInt32(reader["count"]); }
                            catch (Exception) { }
                        }
                    }

                }

                if (comboBox3.Text == "Ad")
                {
                    varmiCommandText = "select count(*) as count from Personel where Ad='" + textBox23.Text + "'";
                    SqlCommand command2 = new SqlCommand(varmiCommandText);
                    command2.Connection = connection;
                    SqlDataReader reader = command2.ExecuteReader();

                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            try
                            { sayii = Convert.ToInt32(reader["count"]); }
                            catch (Exception) { }
                        }
                    }
                }


                connection.Close();

                if (!(sayii == 0))
                {
                    MessageBox.Show("KAYIT BULUNDU LİSTELENİYOR...");
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    SqlCommand cmd = new SqlCommand();

                    if (comboBox3.Text == "Sicil_No")
                    {
                        cmd.CommandText = "SELECT * from Personel where Sicil_No='" + textBox23.Text + "'";
                        cmd.CommandType = CommandType.Text;
                        cmd.Connection = sqlConnection;
                    }

                    if (comboBox3.Text == "TC")
                    {
                        cmd.CommandText = "SELECT * from Personel where TC='" + textBox23.Text + "'";
                        cmd.CommandType = CommandType.Text;
                        cmd.Connection = sqlConnection;
                    }

                    if (comboBox3.Text == "Ad")
                    {
                        cmd.CommandText = "SELECT * from Personel where Ad='" + textBox23.Text + "'";
                        cmd.CommandType = CommandType.Text;
                        cmd.Connection = sqlConnection;
                    }

                    sqlConnection.Open();

                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet);

                    DataTable dt = new DataTable();
                    DataRow dr;
                    dt.Columns.Add(new DataColumn("Personel_Id", typeof(string)));
                    dt.Columns.Add(new DataColumn("Sicil_No", typeof(string)));
                    dt.Columns.Add(new DataColumn("TC", typeof(string)));
                    dt.Columns.Add(new DataColumn("Ad", typeof(string)));
                    dt.Columns.Add(new DataColumn("Soyad", typeof(string)));
                    dt.Columns.Add(new DataColumn("Dogum_Tarihi", typeof(string)));
                    dt.Columns.Add(new DataColumn("Dogum_Yeri", typeof(string)));
                    dt.Columns.Add(new DataColumn("Telefon", typeof(string)));
                    dt.Columns.Add(new DataColumn("Adres", typeof(string)));
                    dt.Columns.Add(new DataColumn("Durumu", typeof(string)));
                    dt.Columns.Add(new DataColumn("Servisi", typeof(string)));
                    dt.Columns.Add(new DataColumn("Gorev_Unvanı", typeof(string)));
                    dt.Columns.Add(new DataColumn("Memuriyet_Baslangic_Tarihi", typeof(string)));
                    dt.Columns.Add(new DataColumn("Karayollarına_Baslangic_Tarihi", typeof(string)));
                    dt.Columns.Add(new DataColumn("Görev_Süresi", typeof(string)));
                    dt.Columns.Add(new DataColumn("İZİN_GECERLİLİK_TARİHİ_YIL", typeof(string)));

                    for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                    {
                        dr = dt.NewRow();
                        dr[0] = dataSet.Tables[0].Rows[i]["Personel_Id"].ToString();
                        dr[1] = dataSet.Tables[0].Rows[i]["Sicil_No"].ToString();
                        dr[2] = dataSet.Tables[0].Rows[i]["TC"].ToString();
                        dr[3] = dataSet.Tables[0].Rows[i]["Ad"].ToString();
                        dr[4] = dataSet.Tables[0].Rows[i]["Soyad"].ToString();
                        dr[5] = dataSet.Tables[0].Rows[i]["Dogum_Tarihi"].ToString();
                        dr[6] = dataSet.Tables[0].Rows[i]["Dogum_Yeri"].ToString();
                        dr[7] = dataSet.Tables[0].Rows[i]["Telefon"].ToString();
                        dr[8] = dataSet.Tables[0].Rows[i]["Adres"].ToString();
                        dr[9] = dataSet.Tables[0].Rows[i]["Durumu"].ToString();
                        dr[10] = dataSet.Tables[0].Rows[i]["Servisi"].ToString();
                        dr[11] = dataSet.Tables[0].Rows[i]["Gorev_Unvanı"].ToString();
                        dr[12] = dataSet.Tables[0].Rows[i]["Memuriyet_Baslangic_Tarihi"].ToString();
                        dr[13] = dataSet.Tables[0].Rows[i]["Karayollarına_Baslangic_Tarihi"].ToString();
                        dr[14] = dataSet.Tables[0].Rows[i]["Gorev_Suresi"].ToString();
                        dr[15] = dataSet.Tables[0].Rows[i]["İzin_Gecerlilik_Tarihi"].ToString();
                        dt.Rows.Add(dr);

                    }

                    DataView dv = new DataView(dt);
                    dataGridView1.DataSource = dv;

                    sqlConnection.Close();
                }
                else
                {
                    MessageBox.Show("KAYIT BULUNAMADI !!!\nBİLGİLERİNİZİ KONTROL EDİN...");
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)//personel kayan yazı
        {
            label19.Text = label19.Text.Substring(1) + label19.Text.Substring(0, 1);

        }

        //----------------------------------//


        //izin bilgileriyle ilgili//

        private void IzinGridiDoldur()//izin datagridview doldurma
        {
            SqlConnection sqlConnection = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();

            cmd.CommandText = "SELECT x.Personel_Id,x.Sicil_No,x.TC,x.Ad,x.Soyad, z.Gun_Sayisi,z.Devreden_Baslangic_Tarihi,y.B_Gun_Sayisi,t.Toplam_İzin,t.Kullanacagi_İzin,t.İzin_Baslangic_Tarihi, 'DD.MM.YYYY',t.İzin_Bitis_Tarihi,'DD.MM.YYYY' from Personel as x,Bu_Yila_Ait_Izin as y,Devreden_Izin as z, İZİNBİLGİSİ AS t  where x.Personel_Id=y.Personel_Id and y.Personel_Id=z.Personel_Id and z.Personel_Id=t.Personel_ID";

            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection;

            sqlConnection.Open();

            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet);

            DataTable dt = new DataTable();
            DataRow dr;
            dt.Columns.Add(new DataColumn("Personel_Id", typeof(string)));
            dt.Columns.Add(new DataColumn("Sicil_No", typeof(string)));
            dt.Columns.Add(new DataColumn("TC", typeof(string)));
            dt.Columns.Add(new DataColumn("Ad", typeof(string)));
            dt.Columns.Add(new DataColumn("Soyad", typeof(string)));
            dt.Columns.Add(new DataColumn("Devreden_Izin", typeof(string)));
            dt.Columns.Add(new DataColumn("Devreden_Baslangic_Tarihi", typeof(string)));
            dt.Columns.Add(new DataColumn("Bu_Yıla_Ait_Izin", typeof(string)));
            dt.Columns.Add(new DataColumn("Toplam_İZİN", typeof(string)));
            dt.Columns.Add(new DataColumn("Kullanacağı_izin", typeof(string)));
            dt.Columns.Add(new DataColumn("İzin_Baslangic_Tarihi", typeof(string)));
            dt.Columns.Add(new DataColumn("İzin_Bitis_Tarihi", typeof(string)));

            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            {

                dr = dt.NewRow();
                //int t=0, k=0;
                dr[0] = dataSet.Tables[0].Rows[i]["Personel_Id"].ToString();
                dr[1] = dataSet.Tables[0].Rows[i]["Sicil_No"].ToString();
                dr[2] = dataSet.Tables[0].Rows[i]["TC"].ToString();
                dr[3] = dataSet.Tables[0].Rows[i]["Ad"].ToString();
                dr[4] = dataSet.Tables[0].Rows[i]["Soyad"].ToString();
                dr[5] = dataSet.Tables[0].Rows[i]["Gun_Sayisi"].ToString();
                dr[6] = dataSet.Tables[0].Rows[i]["Devreden_Baslangic_Tarihi"].ToString();
                dr[7] = dataSet.Tables[0].Rows[i]["B_Gun_Sayisi"].ToString();
                dr[8] = dataSet.Tables[0].Rows[i]["Toplam_İzin"].ToString();
                dr[9] = dataSet.Tables[0].Rows[i]["Kullanacagi_İzin"].ToString();
                dr[10] = dataSet.Tables[0].Rows[i]["İzin_Baslangic_Tarihi"].ToString();
                dr[11] = dataSet.Tables[0].Rows[i]["İzin_Bitis_Tarihi"].ToString();
                dt.Rows.Add(dr);

                //DATAGRİD RENKLENDİRME
                if (i % 2 == 0)
                { dtgrdIzin.AlternatingRowsDefaultCellStyle.BackColor = Color.Silver; }
                else { }
                ///////////////////////

            }

            DataView dv = new DataView(dt);
            dtgrdIzin.DataSource = dv;

            sqlConnection.Close();
        }

        private void dtgrdIzin_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            textBox15.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[0].Value));
            textBox11.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[1].Value));
            textBox12.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[2].Value));
            textBox13.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[3].Value));
            textBox14.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[4].Value));
            textBox17.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[5].Value));
            dateTimePicker6.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[6].Value));
            textBox18.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[7].Value));
            textBox19.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[8].Value));
            textBox20.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[9].Value));

            dateTimePicker4.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[10].Value));
            dateTimePicker5.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[11].Value));


            //dateTimePicker4.Text= Convert.ToString((dtgrdIzin.CurrentRow.Cells[9].Value));
            //dateTimePicker5.Text = Convert.ToString((dtgrdIzin.CurrentRow.Cells[10].Value));

        }//izin datagridden textbox a aktarma

        private void button6_Click(object sender, EventArgs e)//btn izin kaydet
        {
            int kontrol = 1;
            if (!(textBox11.Text == "" || textBox12.Text == "" || textBox13.Text == "" || textBox14.Text == ""
                || textBox15.Text == "" || textBox18.Text == "" || textBox17.Text == "" || textBox19.Text == ""
                || textBox20.Text == "" || dateTimePicker4.Text == "" || dateTimePicker5.Text == ""))
            {

                if (MessageBox.Show("Personel izin bilgisi kaydetmek/güncellemek istediğinize emin misiniz?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    int d = 0, b = 0, t = 0, ki = 0;
                    int gd = 0, gb = 0, gt = 0, gki = 0;

                    gd = Convert.ToInt32(textBox17.Text);
                    gb = Convert.ToInt32(textBox18.Text);
                    //gt = Convert.ToInt32(textBox19.Text);
                    gki = Convert.ToInt32(textBox20.Text);
                    gt = gd + gb;

                    if (gki <= gd)
                    {
                        d = gd - gki;
                        b = gb;
                        t = gt - gki;
                        ki = gki;

                        textBox17.Text = Convert.ToString(d);
                        textBox18.Text = Convert.ToString(b);
                        textBox19.Text = Convert.ToString(t);
                        textBox20.Text = Convert.ToString(ki);
                    }

                    else if (gki > gd && gki <= gt)
                    {

                        d = 0;
                        int gecici = gki;
                        gki -= gd;
                        b = gb - gki;
                        t = gt - gecici;
                        ki = gecici;

                        textBox17.Text = Convert.ToString(d);
                        textBox18.Text = Convert.ToString(b);
                        textBox19.Text = Convert.ToString(t);
                        textBox20.Text = Convert.ToString(ki);

                    }

                    else if (gki > gt)
                    {
                        MessageBox.Show("İzin gününüzü aştınız ...\nTekrar Giriniz !!!");
                        textBox20.Text = "0"; kontrol = 0;

                    }

                    if (kontrol == 1)
                    {
                        SqlConnection connection = new SqlConnection(connectionString);
                        connection.Open();
                        string varmi = "select count(*) as count from İZİNBİLGİSİ where Personel_ID='" + textBox15.Text + "'";

                        SqlCommand command2 = new SqlCommand(varmi);
                        command2.Connection = connection;
                        SqlDataReader reader = command2.ExecuteReader();
                        int sayii = 0;
                        while (reader.Read())
                        {
                            if (reader.HasRows)
                            {
                                try
                                {
                                    sayii = Convert.ToInt32(reader["count"]);
                                }
                                catch (Exception)
                                {


                                }
                            }

                        }

                        connection.Close();

                        if (sayii == 0)
                        {

                            SqlCommand command = new SqlCommand();

                            command.Connection = connection;
                            connection.Open();

                            string sql = "INSERT INTO [dbo].[İZİNBİLGİSİ] ([Personel_ID] ,[Toplam_İzin],[Kullanacagi_İzin] ,[İzin_Baslangic_Tarihi] ,[İzin_Bitis_Tarihi]) VALUES(@per,@to,@ki,@ibat,@ibit)";

                            //////////////////
                            string sql1 = "INSERT INTO [dbo].[Bu_Yila_Ait_Izin] ([Personel_Id],[B_Gun_Sayisi]) VALUES (@per ,@bgun)";
                            SqlCommand cmd1 = new SqlCommand(sql1, connection);
                            cmd1.Parameters.Add("@per", SqlDbType.Int).Value = Convert.ToInt32(textBox15.Text);
                            cmd1.Parameters.Add("@bgun", SqlDbType.Int).Value = Convert.ToInt32(textBox18.Text);

                            cmd1.CommandType = CommandType.Text;
                            cmd1.ExecuteNonQuery();

                            //---------

                            string sql2 = "INSERT INTO [dbo].[Devreden_Izin] ([Personel_Id],[Gun_Sayisi],Devreden_Baslangic_Tarihi) VALUES (@per ,@gun,@dt)";
                            SqlCommand cmd2 = new SqlCommand(sql2, connection);
                            cmd2.Parameters.Add("@per", SqlDbType.Int).Value = Convert.ToInt32(textBox15.Text);
                            cmd2.Parameters.Add("@gun", SqlDbType.Int).Value = Convert.ToInt32(textBox17.Text);
                            cmd2.Parameters.Add("@dt", SqlDbType.Date).Value = Convert.ToDateTime(dateTimePicker6.Text);

                            cmd2.CommandType = CommandType.Text;
                            cmd2.ExecuteNonQuery();
                            ////////////////////////

                            SqlCommand cmd = new SqlCommand(sql, connection);

                            cmd.Parameters.Add("@per", SqlDbType.Int).Value = Convert.ToInt32(textBox15.Text);
                            cmd.Parameters.Add("@to", SqlDbType.Int).Value = Convert.ToInt32(textBox19.Text);
                            cmd.Parameters.Add("@ki", SqlDbType.Int).Value = Convert.ToInt32(textBox20.Text);
                            cmd.Parameters.Add("@ibat", SqlDbType.Date).Value = dateTimePicker4.Value.Date;
                            cmd.Parameters.Add("@ibit", SqlDbType.Date).Value = dateTimePicker5.Value.Date;

                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                            IzinGridiDoldur();
                            connection.Close();

                            MessageBox.Show("İzin Kaydı Eklendi !!!");

                        }
                        else
                        {
                            //string sql = "UPDATE [dbo].[İZİNBİLGİSİ]  SET [Personel_ID] = '" + Convert.ToInt32(textBox15.Text) + "',[Toplam_İzin] = '" + Convert.ToInt32(textBox19.Text) + "', [Kullanacagi_İzin] = '" + Convert.ToInt32(textBox20.Text) + "', WHERE Personel_ID='" + Convert.ToInt32(textBox15.Text) + "' ";
                            // --------------------------------------------
                            // string sql = "UPDATE [dbo].[İZİNBİLGİSİ]  SET [Personel_ID] = '" + Convert.ToInt32(textBox15.Text) + "',[Toplam_İzin] = '" + Convert.ToInt32(textBox19.Text) + "', [Kullanacagi_İzin] = '" + Convert.ToInt32(textBox20.Text) + "' WHERE Personel_ID='" + Convert.ToInt32(textBox15.Text) + "' ";

                            //string sql = "UPDATE [dbo].[İZİNBİLGİSİ] SET [Personel_ID] = '" + Convert.ToInt32(textBox15.Text) + "',[Toplam_İzin] = '" + Convert.ToInt32(textBox19.Text) + "', [Kullanacagi_İzin] = '" + Convert.ToInt32(textBox20.Text) + "' ,[İzin_Baslangic_Tarihi] = '" + dateTimePicker4.Value.Date + "' , [İzin_Bitis_Tarihi] = '" + dateTimePicker5.Value.Date + "' WHERE Personel_ID='" + Convert.ToInt32(textBox15.Text) + "' ";

                            string sql = "UPDATE [dbo].[İZİNBİLGİSİ] SET [Personel_ID] = @per,[Toplam_İzin] = @to, [Kullanacagi_İzin] = @ki ,[İzin_Baslangic_Tarihi] = @ibat , [İzin_Bitis_Tarihi] = @ibit WHERE Personel_ID=@per ";


                            SqlConnection connection1 = new SqlConnection(connectionString);

                            SqlCommand cmd1 = new SqlCommand(sql, connection1);
                            cmd1.Connection = connection1;
                            cmd1.CommandType = CommandType.Text;

                            cmd1.Parameters.Add("@per", SqlDbType.Int).Value = Convert.ToInt32(textBox15.Text);
                            cmd1.Parameters.Add("@to", SqlDbType.Int).Value = Convert.ToInt32(textBox19.Text);
                            cmd1.Parameters.Add("@ki", SqlDbType.Int).Value = Convert.ToInt32(textBox20.Text);
                            cmd1.Parameters.Add("@ibat", SqlDbType.Date).Value = dateTimePicker4.Value.Date;
                            cmd1.Parameters.Add("@ibit", SqlDbType.Date).Value = dateTimePicker5.Value.Date;

                            ///////////////

                            string sql1 = "UPDATE [dbo].[Devreden_Izin]  SET [Personel_Id] = '" + Convert.ToInt32(textBox15.Text) + "',[Gun_Sayisi] = '" + Convert.ToInt32(textBox17.Text) + "',Devreden_Baslangic_Tarihi=@dt  WHERE Personel_Id='" + Convert.ToInt32(textBox15.Text) + "' ";

                            SqlCommand cmd2 = new SqlCommand(sql1, connection1);
                            cmd2.Connection = connection1;
                            cmd2.CommandType = CommandType.Text;
                            cmd2.Parameters.Add("@dt", SqlDbType.Date).Value = Convert.ToDateTime(dateTimePicker6.Text);

                            ////////////////
                            string sql2 = "UPDATE [dbo].[Bu_Yila_Ait_Izin]  SET [Personel_Id] = '" + Convert.ToInt32(textBox15.Text) + "',[B_Gun_Sayisi] = '" + Convert.ToInt32(textBox18.Text) + "'   WHERE Personel_Id='" + Convert.ToInt32(textBox15.Text) + "' ";

                            SqlCommand cmd3 = new SqlCommand(sql2, connection1);
                            cmd3.Connection = connection1;
                            cmd3.CommandType = CommandType.Text;

                            connection1.Open();
                            cmd1.ExecuteNonQuery();
                            cmd2.ExecuteNonQuery();
                            cmd3.ExecuteNonQuery();
                            IzinGridiDoldur();
                            connection1.Close();
                            MessageBox.Show("İzin Kaydı güncellendi !!!");
                        }
                    }
                }//messagebox yes butonu
                else MessageBox.Show("Personel izin bilgisi kaydetmeyi/güncellemeyi iptal ettiniz !!!");
            }
            else
            {
                MessageBox.Show("Boş alanlar mevcut önce onları doldurun !!!");
            }
        }

        private void button5_Click(object sender, EventArgs e)//btn izin sil 
        {
            if (!(textBox15.Text == ""))
            {
                int Personel_ID = Convert.ToInt32(textBox15.Text);

                if (MessageBox.Show(Personel_ID + " nolu personele ait izin bilgisini \nsilmek istediğinize emin misiniz?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {

                    try
                    {
                        using (var sc = new SqlConnection(connectionString))
                        using (var cmd = sc.CreateCommand())
                        {
                            sc.Open();
                            cmd.CommandText = "DELETE FROM İZİNBİLGİSİ WHERE Personel_ID = @Personel_ID";
                            cmd.Parameters.AddWithValue("@Personel_ID", Personel_ID);
                            cmd.ExecuteNonQuery();

                            string sql2 = "DELETE FROM Devreden_Izin WHERE Personel_Id = @Personel_Id";
                            SqlCommand cmd1 = new SqlCommand(sql2, sc);
                            cmd1.Parameters.AddWithValue("@Personel_Id", Personel_ID);
                            cmd1.ExecuteNonQuery();

                            string sql3 = "DELETE FROM Bu_Yila_Ait_Izin WHERE Personel_Id = @Personel_Id";
                            SqlCommand cmd2 = new SqlCommand(sql3, sc);
                            cmd2.Parameters.AddWithValue("@Personel_Id", Personel_ID);
                            cmd2.ExecuteNonQuery();

                            MessageBox.Show(Personel_ID + " nolu personel izin kaydı silindi!!!");
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(Personel_ID + " nolu personel izin kaydı silinemedi!!!");
                    }
                    IzinGridiDoldur();
                }
                else MessageBox.Show("Personel izin bilgisi silmeyi iptal ettiniz !!!");

            }
            else
            {
                MessageBox.Show("Silinecek personel için Personel_ID girmelisiniz...");
            }
        }

        private void button7_Click(object sender, EventArgs e)//izin bilgilerini excel e aktar butonu.
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);

            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            int StartCol = 1;

            int StartRow = 1;

            for (int j = 0; j < dtgrdIzin.Columns.Count; j++)
            {

                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];

                myRange.Value2 = dtgrdIzin.Columns[j].HeaderText;

            }

            StartRow++;

            for (int i = 0; i < dtgrdIzin.Rows.Count; i++)
            {

                for (int j = 0; j < dtgrdIzin.Columns.Count; j++)
                {

                    try
                    {

                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];

                        myRange.Value2 = dtgrdIzin[j, i].Value == null ? "" : dtgrdIzin[j, i].Value;

                    }

                    catch
                    {

                        ;

                    }

                }

            }

        }

        private void timer2_Tick(object sender, EventArgs e)//izin kayan yazı
        {
            label28.Text = label28.Text.Substring(1) + label28.Text.Substring(0, 1);

        }

        private void button8_Click(object sender, EventArgs e)//personel izin bilgisi arama butonu
        {

            if (textBox22.Text == "" || comboBox2.Text == "")
            {
                MessageBox.Show("Boş alanlar mevcut !!!\nÖnce onları doldurun !!!");
            }
            else
            {

                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();

                string varmiCommandText;
                int sayii = 0;

                if (comboBox2.Text == "Sicil_No")
                {
                    varmiCommandText = "select count(*) as count from iZİNBİLGİSİ where Personel_ID=(SELECT Personel_Id from Personel where Sicil_No='" + textBox22.Text + "')";
                    SqlCommand command2 = new SqlCommand(varmiCommandText);
                    command2.Connection = connection;
                    SqlDataReader reader = command2.ExecuteReader();


                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            try
                            { sayii = Convert.ToInt32(reader["count"]); }
                            catch (Exception) { }
                        }
                    }

                }

                if (comboBox2.Text == "TC")
                {
                    varmiCommandText = "select count(*) as count from iZİNBİLGİSİ where Personel_ID=(SELECT Personel_Id from Personel where TC='" + textBox22.Text + "')";

                    SqlCommand command2 = new SqlCommand(varmiCommandText);
                    command2.Connection = connection;
                    SqlDataReader reader = command2.ExecuteReader();

                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            try
                            { sayii = Convert.ToInt32(reader["count"]); }
                            catch (Exception) { }
                        }
                    }

                }

                if (comboBox2.Text == "Ad")
                {
                    varmiCommandText = "select count(*) as count from iZİNBİLGİSİ where Personel_ID=(SELECT Personel_Id from Personel where Ad='" + textBox22.Text + "')";

                    SqlCommand command2 = new SqlCommand(varmiCommandText);
                    command2.Connection = connection;
                    SqlDataReader reader = command2.ExecuteReader();

                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            try
                            { sayii = Convert.ToInt32(reader["count"]); }
                            catch (Exception) { }
                        }
                    }
                }


                connection.Close();

                if (!(sayii == 0))
                {
                    MessageBox.Show("KAYIT BULUNDU LİSTELENİYOR...");
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    SqlCommand cmd = new SqlCommand();

                    if (comboBox2.Text == "Sicil_No")
                    {
                        cmd.CommandText = "SELECT x.Personel_Id,x.Sicil_No,x.TC,x.Ad,x.Soyad, z.Gun_Sayisi,z.Devreden_Baslangic_Tarihi,y.B_Gun_Sayisi,t.Toplam_İzin,t.Kullanacagi_İzin,t.İzin_Baslangic_Tarihi, 'DD.MM.YYYY',t.İzin_Bitis_Tarihi,'DD.MM.YYYY' from Personel as x,Bu_Yila_Ait_Izin as y,Devreden_Izin as z, İZİNBİLGİSİ AS t  where x.Personel_Id=y.Personel_Id and y.Personel_Id=z.Personel_Id and z.Personel_Id=t.Personel_ID and Sicil_No='" + textBox22.Text + "'";
                        cmd.CommandType = CommandType.Text;
                        cmd.Connection = sqlConnection;
                    }

                    if (comboBox2.Text == "TC")
                    {
                        cmd.CommandText = "SELECT x.Personel_Id,x.Sicil_No,x.TC,x.Ad,x.Soyad, z.Gun_Sayisi,z.Devreden_Baslangic_Tarihi,y.B_Gun_Sayisi,t.Toplam_İzin,t.Kullanacagi_İzin,t.İzin_Baslangic_Tarihi, 'DD.MM.YYYY',t.İzin_Bitis_Tarihi,'DD.MM.YYYY' from Personel as x,Bu_Yila_Ait_Izin as y,Devreden_Izin as z, İZİNBİLGİSİ AS t  where x.Personel_Id=y.Personel_Id and y.Personel_Id=z.Personel_Id and z.Personel_Id=t.Personel_ID and TC='" + textBox22.Text + "'";

                        cmd.CommandType = CommandType.Text;
                        cmd.Connection = sqlConnection;
                    }

                    if (comboBox2.Text == "Ad")
                    {
                        cmd.CommandText = "SELECT x.Personel_Id,x.Sicil_No,x.TC,x.Ad,x.Soyad, z.Gun_Sayisi,z.Devreden_Baslangic_Tarihi,y.B_Gun_Sayisi,t.Toplam_İzin,t.Kullanacagi_İzin,t.İzin_Baslangic_Tarihi, 'DD.MM.YYYY',t.İzin_Bitis_Tarihi,'DD.MM.YYYY' from Personel as x,Bu_Yila_Ait_Izin as y,Devreden_Izin as z, İZİNBİLGİSİ AS t  where x.Personel_Id=y.Personel_Id and y.Personel_Id=z.Personel_Id and z.Personel_Id=t.Personel_ID and Ad='" + textBox22.Text + "'";

                        cmd.CommandType = CommandType.Text;
                        cmd.Connection = sqlConnection;
                    }

                    sqlConnection.Open();

                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet);

                    DataTable dt = new DataTable();
                    DataRow dr;

                    dt.Columns.Add(new DataColumn("Personel_Id", typeof(string)));
                    dt.Columns.Add(new DataColumn("Sicil_No", typeof(string)));
                    dt.Columns.Add(new DataColumn("TC", typeof(string)));
                    dt.Columns.Add(new DataColumn("Ad", typeof(string)));
                    dt.Columns.Add(new DataColumn("Soyad", typeof(string)));
                    dt.Columns.Add(new DataColumn("Devreden_Izin", typeof(string)));
                    dt.Columns.Add(new DataColumn("Devreden_Baslangic_Tarihi", typeof(string)));
                    dt.Columns.Add(new DataColumn("Bu_Yıla_Ait_Izin", typeof(string)));
                    dt.Columns.Add(new DataColumn("Toplam_İZİN", typeof(string)));
                    dt.Columns.Add(new DataColumn("Kullanacağı_izin", typeof(string)));
                    dt.Columns.Add(new DataColumn("İzin_Baslangic_Tarihi", typeof(string)));
                    dt.Columns.Add(new DataColumn("İzin_Bitis_Tarihi", typeof(string)));

                    for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                    {

                        dr = dt.NewRow();
                        //int t=0, k=0;
                        dr[0] = dataSet.Tables[0].Rows[i]["Personel_Id"].ToString();
                        dr[1] = dataSet.Tables[0].Rows[i]["Sicil_No"].ToString();
                        dr[2] = dataSet.Tables[0].Rows[i]["TC"].ToString();
                        dr[3] = dataSet.Tables[0].Rows[i]["Ad"].ToString();
                        dr[4] = dataSet.Tables[0].Rows[i]["Soyad"].ToString();
                        dr[5] = dataSet.Tables[0].Rows[i]["Gun_Sayisi"].ToString();
                        dr[6] = dataSet.Tables[0].Rows[i]["Devreden_Baslangic_Tarihi"].ToString();
                        dr[7] = dataSet.Tables[0].Rows[i]["B_Gun_Sayisi"].ToString();
                        dr[8] = dataSet.Tables[0].Rows[i]["Toplam_İzin"].ToString();
                        dr[9] = dataSet.Tables[0].Rows[i]["Kullanacagi_İzin"].ToString();
                        dr[10] = dataSet.Tables[0].Rows[i]["İzin_Baslangic_Tarihi"].ToString();
                        dr[11] = dataSet.Tables[0].Rows[i]["İzin_Bitis_Tarihi"].ToString();
                        dt.Rows.Add(dr);

                    }

                    DataView dv = new DataView(dt);
                    dtgrdIzin.DataSource = dv;

                    sqlConnection.Close();
                }
                else
                {
                    MessageBox.Show("KAYIT BULUNAMADI !!!\nBİLGİLERİNİZİ KONTROL EDİN...");
                }
            }
        }

        //***************************//

        //yıllık izin ekleme 30--20 gün meselesi//
        private void YILLIK_İZİN_EKLE_GUNCELLE()
        {
            SqlConnection connection1 = new SqlConnection(connectionString);

            int devredenIzin, yeniDevredenIzin,
                buYilIzin, yeniBuYilIzin,
                buYil;

            DateTime  Memur_Son_Tarih, sonTrh, bugün;

            buYil = DateTime.Now.Year;
            bugün = Convert.ToDateTime(DateTime.Now.Date);
            sonTrh = Convert.ToDateTime(buYil + "-12-31");

            //MessageBox.Show("buyil: "+buYil);
            //MessageBox.Show("bugün: " + bugün);
            //MessageBox.Show("Sontarih: "+sonTrh);

            int yeni;

            //*** 1 yıl geçme durumu ***//

            if (bugün == sonTrh)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        for (int k = 0; k < dtgrdIzin.Rows.Count; k++)
                        {
                            for (int l = 0; l < dtgrdIzin.Columns.Count; l++)
                            {
                                if (Convert.ToInt32(dataGridView1[0, i].Value) == Convert.ToInt32(dtgrdIzin[0, k].Value))
                                {
                                    devredenIzin = Convert.ToInt32(dtgrdIzin[5, k].Value);
                                    buYilIzin = Convert.ToInt32(dtgrdIzin[7, k].Value);
                                    yeniDevredenIzin = devredenIzin + buYilIzin;

                                    string sql4 = "UPDATE [dbo].[Devreden_Izin]  SET [Gun_Sayisi]=@yenidevir, Devreden_Baslangic_Tarihi=@dt  WHERE Personel_Id=(select Personel_ID from Personel where Personel_ID='" + Convert.ToInt32(dataGridView1[0, i].Value) + "' and İzin_Gecerlilik_Tarihi=2 )";

                                    SqlCommand cmd4 = new SqlCommand(sql4, connection1);
                                    cmd4.Connection = connection1;
                                    cmd4.CommandType = CommandType.Text;
                                    cmd4.Parameters.Add("@dt", SqlDbType.Date).Value = sonTrh;
                                    cmd4.Parameters.Add("@yenidevir", SqlDbType.Int).Value = yeniDevredenIzin;

                                    connection1.Open();
                                    cmd4.ExecuteNonQuery();
                                    IzinGridiDoldur();
                                    connection1.Close();

                                    break; break;
                                }
                            }
                        }

                        //MessageBox.Show(dataGridView1[15, i].Value.ToString());
                        if (Convert.ToInt32(dataGridView1[14, i].Value) >= 3650) yeniBuYilIzin = 30;
                        else yeniBuYilIzin = 20;

                        string sql2 = "UPDATE [dbo].[Bu_Yila_Ait_Izin]  SET [B_Gun_Sayisi] = @yeni   WHERE Personel_Id='" + Convert.ToInt32(dataGridView1[0, i].Value) + "' ";

                        SqlCommand cmd3 = new SqlCommand(sql2, connection1);
                        cmd3.Connection = connection1;
                        cmd3.CommandType = CommandType.Text;

                        cmd3.Parameters.Add("@yeni", SqlDbType.Int).Value = yeniBuYilIzin;

                        connection1.Open();
                        cmd3.ExecuteNonQuery();
                        IzinGridiDoldur();
                        connection1.Close();

                        //
                        string sql1 = "UPDATE [dbo].[Devreden_Izin]  SET [Gun_Sayisi]=0, Devreden_Baslangic_Tarihi=@dt  WHERE Personel_Id=(select Personel_ID from Personel where Personel_ID='" + Convert.ToInt32(dataGridView1[0, i].Value) + "' and İzin_Gecerlilik_Tarihi=1 )";

                        SqlCommand cmd2 = new SqlCommand(sql1, connection1);
                        cmd2.Connection = connection1;
                        cmd2.CommandType = CommandType.Text;
                        cmd2.Parameters.Add("@dt", SqlDbType.Date).Value = Convert.ToDateTime("2000-01-01 00:00:00.0");

                        connection1.Open();
                        cmd2.ExecuteNonQuery();
                        IzinGridiDoldur();
                        connection1.Close();
                        //

                        break;
                    }
                }
            }//->1 yıl geçme durumu

            buYil += 1;
            bugün = Convert.ToDateTime(DateTime.Now.Date);
            Memur_Son_Tarih = Convert.ToDateTime(buYil + "-12-31");

            if (bugün == Memur_Son_Tarih)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        for (int k = 0; k < dtgrdIzin.Rows.Count; k++)
                        {
                            for (int l = 0; l < dtgrdIzin.Columns.Count; l++)
                            {
                                if (Convert.ToInt32(dataGridView1[0, i].Value) == Convert.ToInt32(dtgrdIzin[0, k].Value))
                                {
                                    buYilIzin = Convert.ToInt32(dtgrdIzin[7, k].Value);
                                    devredenIzin = buYilIzin;

                                    yeniDevredenIzin = buYilIzin;

                                    string sql4 = "UPDATE [dbo].[Devreden_Izin]  SET [Gun_Sayisi]=@yenidevir, Devreden_Baslangic_Tarihi=@dt  WHERE Personel_Id=(select Personel_ID from Personel where Personel_ID='" + Convert.ToInt32(dataGridView1[0, i].Value) + "' and İzin_Gecerlilik_Tarihi=2 )";

                                    SqlCommand cmd5 = new SqlCommand(sql4, connection1);
                                    cmd5.Connection = connection1;
                                    cmd5.CommandType = CommandType.Text;
                                    cmd5.Parameters.Add("@dt", SqlDbType.Date).Value = Memur_Son_Tarih;
                                    cmd5.Parameters.Add("@yenidevir", SqlDbType.Int).Value = yeniDevredenIzin;

                                    connection1.Open();
                                    cmd5.ExecuteNonQuery();
                                    IzinGridiDoldur();
                                    connection1.Close();

                                    break; break;
                                }
                            }
                        }
                    }

                }
            }
        }
    }
}
        //////////////////////////////////////////

        
        
         //yıllık izin ekleme 30--20 gün meselesi//
        /*
        private void YILLIK_İZİN_EKLE_GUNCELLE()
        {/*
            SqlConnection connection1 = new SqlConnection(connectionString);

            int devredenIzin, yeniDevredenIzin,
                buYilIzin, yeniBuYilIzin,
                buYil;

            DateTime eskiDevredenBasTrh, yeniDevredenBasTrh, sonTrh, bugün;

            buYil = DateTime.Now.Year;
            bugün = Convert.ToDateTime(DateTime.Now.Date);
            sonTrh = Convert.ToDateTime(buYil + "-12-20");

            //MessageBox.Show("buyil: "+buYil);
            //MessageBox.Show("bugün: " + bugün);
            //MessageBox.Show("Sontarih: "+sonTrh);

            int yeni;

        
            if (bugün == sonTrh)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        for (int k = 0; k < dtgrdIzin.Rows.Count; k++)
                        {
                            for (int l = 0; l < dtgrdIzin.Columns.Count; l++)
                            {
                                if (Convert.ToInt32(dataGridView1[0, i].Value)==Convert.ToInt32(dtgrdIzin[0, k].Value))
                                {
                                    devredenIzin = Convert.ToInt32(dtgrdIzin[5, k].Value);
                                    buYilIzin = Convert.ToInt32(dtgrdIzin[7, k].Value);
                                    yeniDevredenIzin = devredenIzin + buYilIzin;

                                    string sql4 = "UPDATE [dbo].[Devreden_Izin]  SET [Gun_Sayisi]=@yenidevir, Devreden_Baslangic_Tarihi=@dt  WHERE Personel_Id=(select Personel_ID from Personel where Personel_ID='" + Convert.ToInt32(dataGridView1[0, i].Value) + "' and İzin_Gecerlilik_Tarihi=2 )";

                                    SqlCommand cmd4 = new SqlCommand(sql4, connection1);
                                    cmd4.Connection = connection1;
                                    cmd4.CommandType = CommandType.Text;
                                    cmd4.Parameters.Add("@dt", SqlDbType.Date).Value = sonTrh;
                                    cmd4.Parameters.Add("@yenidevir", SqlDbType.Int).Value = yeniDevredenIzin;

                                    connection1.Open();
                                    cmd4.ExecuteNonQuery();
                                    IzinGridiDoldur();
                                    connection1.Close();

                                    break; break;
                                }
                            }
                        }

                        //MessageBox.Show(dataGridView1[15, i].Value.ToString());
                        if (Convert.ToInt32(dataGridView1[14, i].Value) >= 3650) yeniBuYilIzin = 30;
                        else yeniBuYilIzin = 20;

                        string sql2 = "UPDATE [dbo].[Bu_Yila_Ait_Izin]  SET [B_Gun_Sayisi] = @yeni   WHERE Personel_Id='" + Convert.ToInt32(dataGridView1[0, i].Value) + "' ";
                        
                        SqlCommand cmd3 = new SqlCommand(sql2, connection1);
                        cmd3.Connection = connection1;
                        cmd3.CommandType = CommandType.Text;

                        cmd3.Parameters.Add("@yeni", SqlDbType.Int).Value = yeniBuYilIzin;

                        connection1.Open();
                        cmd3.ExecuteNonQuery();
                        IzinGridiDoldur();
                        connection1.Close();

                        //
                        string sql1 = "UPDATE [dbo].[Devreden_Izin]  SET [Gun_Sayisi]=0, Devreden_Baslangic_Tarihi=@dt  WHERE Personel_Id=(select Personel_ID from Personel where Personel_ID='" + Convert.ToInt32(dataGridView1[0, i].Value) + "' and İzin_Gecerlilik_Tarihi=1 )";

                        SqlCommand cmd2 = new SqlCommand(sql1, connection1);
                        cmd2.Connection = connection1;
                        cmd2.CommandType = CommandType.Text;
                        cmd2.Parameters.Add("@dt", SqlDbType.Date).Value = Convert.ToDateTime("2000-01-01 00:00:00.0");

                        connection1.Open();
                        cmd2.ExecuteNonQuery();
                        IzinGridiDoldur();
                        connection1.Close();
                        //

                        break;
                    }
                }
            }//->1 yıl geçme durumu

        }
        //////////////////////////////////////////
         */
   