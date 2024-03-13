using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb; //Veritabanı kütüphanesini ekliyoruz
using System.Xml; //XML kütüphanesi

namespace veritabanibaglanti
{
    public partial class Form1 : Form
    {
        OleDbConnection baglanti = new OleDbConnection(); //Kullanacağımız ögeleri global olarak tanımlamamız gerekiyor. Aşağıda kullanmak için
        OleDbCommand komut = new OleDbCommand();
        OleDbDataAdapter da1 = new OleDbDataAdapter();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            vericek(); //Okunabilirlik için bu yüzden bu şekilde kullanıldı.
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Veri tabanına veri kaydetmek
            komut = new OleDbCommand("INSERT INTO ogrenci (Ad, Soyad) VALUES (@ad, @soyad)", baglanti);
            komut.Parameters.AddWithValue("@ad", textBox1.Text); //Kullanılacak olan değişkenlerin ne olduğunu tanımladık.
            komut.Parameters.AddWithValue("@soyad", textBox2.Text);

            try
            {
                baglanti.Open(); // veri tabanını açıyoruz
                komut.ExecuteNonQuery(); // verileri kaydetme
                MessageBox.Show("Veri başarıyla eklendi.");
            }
            catch (Exception ex) //Hatayı yakalamak için
            {
                MessageBox.Show("Veri eklenirken bir hata oluştu: " + ex.Message);
            }
            finally
            {
                baglanti.Close(); // açık kalması iyi değil, bu yüzden kapatıyoruz
            }
        }

        private void vericek()
        {
            //Asıl bağlantı komutumuz
            string baglan, sorgu;
            baglan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\Database2.mdb";//Provider:Sağlayıcı
            sorgu = "select * from ogrenci";
            baglanti = new OleDbConnection(baglan); //Bağlantıyı yukarıda tanımladık fakat boş oldupu için burda yukarıdaki kodu execute ettik.
            da1 = new OleDbDataAdapter(sorgu, baglanti); //Geçici olarak da e verileri kaydettik.
            DataSet al = new DataSet(); //dataset ise kalıcı olarak veriyi alır, datagridview veriyi burdan çeker.
            da1.Fill(al, "abc"); //Dataadapter dan veriyi alıp dataset e (da1-----al)
            dataGridView1.DataSource = al.Tables[0];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            vericek(); //Veritabanını tekrar açar bir nevi yenilemiş olur.
        }

        private void button3_Click(object sender, EventArgs e)
        {
            komut = new OleDbCommand("DELETE FROM ogrenci WHERE Kimlik=@Kimlik", baglanti);
            komut.Parameters.AddWithValue("@Kimlik", dataGridView1.CurrentRow.Cells[0].Value); //Dgv deki seçili olan hücrenin ilk hücresinindeki veri kimlik değişkenine eşit oldu.
            baglanti.Open();
            komut.ExecuteNonQuery();
            baglanti.Close();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            XmlDocument dosya = new XmlDocument();
            dosya.Load(Application.StartupPath + "\\deneme.xml");

            XmlElement ogrenci = dosya.CreateElement("ogrenci");
            ogrenci.SetAttribute("Kimlik", textBox3.Text);

            XmlNode ad = dosya.CreateNode(XmlNodeType.Element, "Ad", "");
            ad.InnerText = textBox1.Text;
            ogrenci.AppendChild(ad);

            XmlNode soyad = dosya.CreateNode(XmlNodeType.Element, "Soyad", "");
            soyad.InnerText = textBox2.Text;
            ogrenci.AppendChild(soyad);

            dosya.DocumentElement.AppendChild(ogrenci);
            dosya.Save(Application.StartupPath + "\\deneme.xml");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            XmlDocument dosyaa = new XmlDocument();
            dosyaa.Load(Application.StartupPath + "\\deneme.xml");

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++) //Kaç adet veri varsa o kadar döndürecek
            {
                XmlElement ogrenci = dosyaa.CreateElement("ogrenci");
                ogrenci.SetAttribute("Kimlik", dataGridView1.Rows[i].Cells[0].Value.ToString());

                XmlNode ad = dosyaa.CreateNode(XmlNodeType.Element, "Ad", "");
                ad.InnerText = dataGridView1.Rows[i].Cells[1].Value.ToString();
                ogrenci.AppendChild(ad);

                XmlNode soyad = dosyaa.CreateNode(XmlNodeType.Element, "Soyad", "");
                soyad.InnerText = dataGridView1.Rows[i].Cells[2].Value.ToString();
                ogrenci.AppendChild(soyad);

                dosyaa.DocumentElement.AppendChild(ogrenci);
                dosyaa.Save(Application.StartupPath + "\\deneme.xml");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //ÇOK ÖNEMLİ

            list();

            XmlDocument dosya = new XmlDocument();
            dosya.Load(Application.StartupPath + "\\deneme.xml");

            XmlNodeList b = dosya.SelectNodes("ogrenciler/ogrenci"); //b yi xml deki yeri olarak düşün, a ya tanımladığımız verileri b'ye yani xml deki yerine kaydediyoruz.
            foreach (XmlNode a in b)
            {
                string kimlik = a.Attributes["Kimlik"].Value; //PK olduğu için attributes kullandık, alttakiler pk değil.

                string ad = a["Ad"].InnerText;

                string soyad = a["Soyad"].InnerText;

                ListViewItem item = new ListViewItem(kimlik);

                item.SubItems.Add(ad);

                item.SubItems.Add(soyad);

                listView1.Items.Add(item);


            }
        }

        private void list()
        {
            //Görünüş İçin Önemli
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;

            listView1.Columns.Add("Kimlik");
            listView1.Columns.Add("Ad");
            listView1.Columns.Add("Soyad");
        }
    }
}
