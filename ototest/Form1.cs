using System;



using System.Drawing;

using System.Windows.Forms;
using System.Data.OleDb;
using wordeaktar = Microsoft.Office.Interop.Word;
using wordeaktar1 = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;


namespace ototest
{
    public partial class Form1 : Form
    {
        string[] sorular1 = new string[100];
        string[] aCevabi1 = new string[100];
        string[] bCevabi1 = new string[100];
        string[] cCevabi1 = new string[100];
        string[] dCevabi1 = new string[100];
        string[] eCevabi1 = new string[100];
        string[] dogrucevap1 = new string[100];

        int[] ac1 = new int[100];
        int[] bc1 = new int[100];
        int[] cc1 = new int[100];
        int[] dc1 = new int[100];
        int[] ec1 = new int[100];
        int count = 1;
        string idS1;
        string idS2;
        int index = -1;
        string[] dersid = new string[100];
        string[] sorularr = new string[100];
        string[] sid = new string[100];
        string[] dersler = new string[100];
        string[] sorusec = new string[1000];
        string imgyeri = "";
        OleDbConnection baglantim = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb");
        public Form1()
        {
            InitializeComponent();


        }
        int Say(string abc)
        {
            int count = 0;
            for (int ij = 0; ij < abc.Length; ij++)
            {
                count++;
            }
            return count;
        }
        private void anaSayfa()
        {
            dataGridView1.Visible = false;
            button1.Visible = true;
            button2.Visible = true;
            button3.Visible = true;
            button4.Visible = true;
            button5.Visible = true;
            button6.Visible = true;
            button7.Visible = true;
            button8.Visible = true;
            button9.Visible = true;
            tabControl1.TabPages.Add(tabPage1);
            tabControl1.TabPages.Add(tabPage2);
            tabControl1.TabPages.Add(tabPage3);
            tabControl1.TabPages.Add(tabPage4);
            tabControl1.TabPages.Add(tabPage5);
            tabControl1.TabPages.Add(tabPage6);
            tabControl1.TabPages.Add(tabPage7);

            tabControl1.Visible = false;

        }

        private void verig()
        {

            OleDbCommand komut2 = new OleDbCommand();
            komut2.Connection = baglantim;
            komut2.CommandText = ("select SORUADI from SORULAR");
            //Sorgumuzu ve baglantimizi parametre olarak alan bir SqlCommand nesnesi oluşturuyoruz.
            OleDbDataAdapter da1 = new OleDbDataAdapter(komut2);
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView3.DataSource = dt1;
            this.dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            dataGridView3.RowHeadersVisible = false;
            dataGridView3.BackgroundColor = Color.White;
        }
        private void verig1()
        {

            OleDbCommand komut2 = new OleDbCommand();
            komut2.Connection = baglantim;
            komut2.CommandText = ("select ID,SORUADI from SORULAR where DERSID="+idS2.ToString()+"");
            //Sorgumuzu ve baglantimizi parametre olarak alan bir SqlCommand nesnesi oluşturuyoruz.
            OleDbDataAdapter da1 = new OleDbDataAdapter(komut2);
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt1 = new System.Data.DataTable();
            da1.Fill(dt1);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView2.DataSource = dt1;
            this.dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            dataGridView2.RowHeadersVisible = false;
            dataGridView2.BackgroundColor = Color.White;
            dataGridView2.Columns[0].Width = 40;
        }


        private void veriGoster()
        {
            dataGridView1.Visible = true;

            OleDbCommand komut = new OleDbCommand();
            komut.Connection = baglantim;
            komut.CommandText = ("select * from DERSLER");



            //Sorgumuzu ve baglantimizi parametre olarak alan bir SqlCommand nesnesi oluşturuyoruz.
            OleDbDataAdapter da = new OleDbDataAdapter(komut);
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.BackgroundColor = Color.White;



        }

        private void Form1_Load(object sender, EventArgs e)
        {
            baglantim.Open();

            OleDbCommand komut3 = new OleDbCommand();
            komut3.Connection = baglantim;
            komut3.CommandText = ("select SORUADI,ID from SORULAR");
            OleDbDataReader oku1 = komut3.ExecuteReader();
            int j = 1;
            if (oku1.HasRows)
            {
                while (oku1.Read())
                {
                    sorularr[j] = oku1[0].ToString();
                    sid[j] = oku1[1].ToString();

                    j++;

                }
            }
            oku1.Close();
            baglantim.Close();
            tabControl1.Visible = false;
            dataGridView1.Visible = false;


        }



        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.Visible = true;
            tabControl1.TabPages.Remove(tabPage2);
            dataGridView1.Visible = true;
            veriGoster();
            tabControl1.TabPages.Remove(tabPage3);
            tabControl1.TabPages.Remove(tabPage4);
            tabControl1.TabPages.Remove(tabPage5);
            tabControl1.TabPages.Remove(tabPage6);
            tabControl1.TabPages.Remove(tabPage7);
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button9.Visible = false;
        }

       
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = true;
            veriGoster();
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button9.Visible = false;
            button13.Visible = false;

            verig();
            tabControl1.Visible = true;
            tabControl1.TabPages.Remove(tabPage1);
            tabControl1.TabPages.Remove(tabPage3);
            tabControl1.TabPages.Remove(tabPage4);
            tabControl1.TabPages.Remove(tabPage5);
            tabControl1.TabPages.Remove(tabPage6);
            tabControl1.TabPages.Remove(tabPage7);



            textBox9.Visible = false;
            label19.Visible = false;
            textBox11.Visible = false;
            button16.Visible = false;
            button18.Visible = false;
            button23.Visible = false;
        }

       
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
           


        }

        

        private void button3_Click(object sender, EventArgs e)
        {
          
            button41.Visible = false;
            label44.Visible = false;
            textBox19.Visible =false;
            button42.Visible = false;
            textBox10.Visible = false;
            groupBox12.Visible = false;
            button26.Visible = false;
            button27.Visible = false;
            button40.Visible =false;
            label31.Visible = false;
            button24.Visible = false;
            button25.Visible = false;
            dataGridView1.Visible = true;
            veriGoster();
            groupBox10.Visible = false;
            label29.Visible = false;
            textBox16.Visible = false;
            button37.Visible = false;
            tabControl1.Visible = true;
            tabControl1.TabPages.Remove(tabPage1);
            tabControl1.TabPages.Remove(tabPage2);
            tabControl1.TabPages.Remove(tabPage4);
            tabControl1.TabPages.Remove(tabPage5);
            tabControl1.TabPages.Remove(tabPage6);
            tabControl1.TabPages.Remove(tabPage7);
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            button9.Visible = false;
        }

       
        private void button30_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Lütfen sol taraftaki tablodan ders seçiniz");


        }

        

        private void button24_Click(object sender, EventArgs e)
        {

        }

       

        private void button12_Click_1(object sender, EventArgs e)
        {

            veriGoster();
            try
            {
                baglantim.Open();
                OleDbCommand komut2 = new OleDbCommand();

                komut2.Connection = baglantim;
                komut2.CommandText = ("select ID,DERSADI from DERSLER");
                OleDbDataReader oku = komut2.ExecuteReader();

                int i = 1;
                if (oku.HasRows)
                {
                    while (oku.Read())
                    {
                        dersid[i] = oku[0].ToString();
                        dersler[i] = oku[1].ToString();
                        i++;

                    }
                }

                oku.Close();
                baglantim.Close();




                i--;
                while (i > 0)
                {
                    if (textBox1.Text == dersler[i].ToString())
                    {
                        MessageBox.Show("Böyle bir ders veritabanında mevcut.");
                        textBox1.Clear();
                        textBox25.Clear();
                        return;

                    }
                    if (textBox25.Text == dersid[i].ToString())
                    {
                        MessageBox.Show("Bu ID'ye sahip bir ders veritabanında mevcut.");
                        textBox1.Clear();
                        textBox25.Clear();
                        return;

                    }
                    i--;

                }
                baglantim.Open();
                if (MessageBox.Show("" + textBox1.Text.ToString() + " adlı dersi eklemek istiyor musunuz?", "Dikkat", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    OleDbCommand komut = new OleDbCommand("Insert into DERSLER(ID,DERSADI) values('" + textBox25.Text.ToString() + "','" + textBox1.Text.ToString() + "')", baglantim);
                    komut.ExecuteNonQuery();
                    MessageBox.Show("Kayıt işlemi başarıyla gerçekleşti.");
                    textBox1.Clear();
                    textBox25.Clear();
                    baglantim.Close();
                }
            }

            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
            veriGoster();
        }

        private void button11_Click_1(object sender, EventArgs e)
        {

            textBox1.Clear();
            textBox25.Clear();
        }

        private void button19_Click_1(object sender, EventArgs e)
        {

            baglantim.Open();
            OleDbCommand komut2 = new OleDbCommand();
            komut2.Connection = baglantim;
            komut2.CommandText = ("select ID,DERSADI from DERSLER");
            OleDbDataReader oku = komut2.ExecuteReader();

            int i = 1;
            if (oku.HasRows)
            {
                while (oku.Read())
                {
                    dersid[i] = oku[0].ToString();
                    dersler[i] = oku[1].ToString();
                    i++;

                }
            }

            oku.Close();
            baglantim.Close();

            i--;
            while (i > 0)
            {
                if (textBox13.Text == dersler[i].ToString())
                {
                    MessageBox.Show("Böyle bir ders veritabanında mevcut.");
                    textBox13.Clear();
                    return;

                }
                i--;

            }
            MessageBox.Show("Böyle bir ders veritabanında mevcut değil.");
        }

        private void button10_Click_1(object sender, EventArgs e)
        {


            OleDbCommand kmt = new OleDbCommand();
            OleDbCommand kmt2 = new OleDbCommand();



            if (MessageBox.Show("Seçili Ögeyi Silmek İstiyor Musunuz ?", "Dikkat", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                baglantim.Open();
                kmt.Connection = baglantim;
                kmt.CommandText = "DELETE FROM DERSLER WHERE ID=@ogrnumarasi";

                kmt.Parameters.AddWithValue("@ogrnumarasi", dataGridView1.CurrentCell.Value.ToString());




                kmt.ExecuteNonQuery();

                kmt2.Connection = baglantim;
                kmt2.CommandText = "DELETE FROM DERSLER WHERE DERSADI=@dersadi";
                kmt2.Parameters.AddWithValue("@dersadi", dataGridView1.CurrentCell.Value.ToString());

                kmt2.ExecuteNonQuery();
                baglantim.Close();
                MessageBox.Show("Silme İşlemi Başarılı", "Silindi", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            else
            {

            }
            veriGoster();

        }

        private void button21_Click_1(object sender, EventArgs e)
        {
            baglantim.Open();

            OleDbCommand komut2 = new OleDbCommand();
            komut2.Connection = baglantim;


            komut2.CommandText = "update  DERSLER set ID=@ID,DERSADI=@DERSADI where ID=@ogrnumarsi";
            komut2.Parameters.AddWithValue("@ID", textBox14.Text);
            komut2.Parameters.AddWithValue("@DERSADI", textBox15.Text);
            komut2.Parameters.AddWithValue("@ogrnumarasi", dataGridView1.CurrentCell.Value.ToString());
            komut2.ExecuteNonQuery();


            OleDbCommand kmt = new OleDbCommand();
            kmt.Connection = baglantim;

            kmt.CommandText = "update  DERSLER set ID=@ID,DERSADI=@DERSADI where DERSADI=@dr";
            kmt.Parameters.AddWithValue("@ID", textBox14.Text);
            kmt.Parameters.AddWithValue("@DERSADI", textBox15.Text);
            kmt.Parameters.AddWithValue("@dr", dataGridView1.CurrentCell.Value.ToString());
            kmt.ExecuteNonQuery();

            veriGoster();
            textBox14.Clear();
            textBox15.Clear();

            baglantim.Close();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            textBox14.Clear();
            textBox15.Clear();

        }

        private void button16_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "png files(*.png)|*.png|jpg files(*.jpg)|*.jpg|All files(*.*)|*.*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                imgyeri = dialog.FileName.ToString();
                textBox9.Text = imgyeri;
            }

        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            radioButton1.Checked = false;
            textBox9.Visible = false;
            button16.Visible = false;
            button13.Visible = false;
            textBox9.Clear();
        }

        private void button15_Click_1(object sender, EventArgs e)
        {

            OleDbCommand kmt2 = new OleDbCommand();
            OleDbCommand kmt = new OleDbCommand();
            baglantim.Open();
            int id = (int)dataGridView1.CurrentRow.Cells[0].Value;



            if (radioButton1.Checked == true)
            {
                kmt2.Connection = baglantim;


                string str = id.ToString();

                kmt2.CommandText = "insert into SORULAR(SORUADI,A,B,C,D,E,DOGRUCEVAP,DERSID,SEKIL) values(@s,@a,@b,@c,@d,@e,@do,@de,@se)";

                kmt2.Parameters.AddWithValue("@s", textBox3.Text);
                kmt2.Parameters.AddWithValue("@a", textBox4.Text);
                kmt2.Parameters.AddWithValue("@b", textBox5.Text);
                kmt2.Parameters.AddWithValue("@c", textBox6.Text);
                kmt2.Parameters.AddWithValue("@d", textBox7.Text);
                kmt2.Parameters.AddWithValue("@e", textBox8.Text);
                kmt2.Parameters.AddWithValue("@do", comboBox1.Text);
                kmt2.Parameters.AddWithValue("@de", str);
                kmt2.Parameters.AddWithValue("@se", textBox9.Text);
                kmt2.ExecuteNonQuery();

            }
            else
            {
                kmt.Connection = baglantim;


                string str1 = id.ToString();

                kmt.CommandText = "insert into SORULAR(SORUADI,A,B,C,D,E,DOGRUCEVAP,DERSID) values(@s,@a,@b,@c,@d,@e,@do,@de)";

                kmt.Parameters.AddWithValue("@s", textBox3.Text);
                kmt.Parameters.AddWithValue("@a", textBox4.Text);
                kmt.Parameters.AddWithValue("@b", textBox5.Text);
                kmt.Parameters.AddWithValue("@c", textBox6.Text);
                kmt.Parameters.AddWithValue("@d", textBox7.Text);
                kmt.Parameters.AddWithValue("@e", textBox8.Text);
                kmt.Parameters.AddWithValue("@do", comboBox1.Text);
                kmt.Parameters.AddWithValue("@de", str1);

                kmt.ExecuteNonQuery();


            }
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();



            baglantim.Close();


        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            radioButton1.Checked = false;
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();
            textBox9.Visible = false;
            button16.Visible = false;
            button13.Visible = false;

        }

        private void button17_Click_1(object sender, EventArgs e)
        {
            button23.Visible = true;
            label19.Visible = true;
            textBox11.Visible = true;
            //button16.Visible = true;
            button18.Visible = true;


            string idS = dataGridView3.CurrentRow.Cells[0].Value.ToString();
            textBox11.Text = idS.ToString();

        }

        private void button23_Click_1(object sender, EventArgs e)
        {
            OleDbCommand kmt = new OleDbCommand();
            OleDbCommand kmt2 = new OleDbCommand();



            if (MessageBox.Show("Seçili Ögeyi Silmek İstiyor Musunuz ?", "Dikkat", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                baglantim.Open();


                kmt2.Connection = baglantim;
                kmt2.CommandText = "DELETE FROM SORULAR WHERE SORUADI=@dersadi";
                kmt2.Parameters.AddWithValue("@dersadi", dataGridView3.CurrentCell.Value.ToString());

                kmt2.ExecuteNonQuery();
                baglantim.Close();
                MessageBox.Show("Silme İşlemi Başarılı", "Silindi", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            else
            {

            }
            button23.Visible = false;
            label19.Visible = false;
            textBox11.Visible = false;
            button16.Visible = false;
            button18.Visible = false;
            button23.Visible = false;
            verig();

        }

        private void button18_Click_1(object sender, EventArgs e)
        {
            
            button18.Visible = false;
            label19.Visible = false;
            textBox11.Visible = false;
            button23.Visible = false;

        }

        private void button22_Click_1(object sender, EventArgs e)
        {
            if (textBox12.Text == "")
            {
                MessageBox.Show("Kelime girmeyi unuttunuz.");
            }

            for (int i = 1; i < sorularr.Length - 1; i++)
            {
                if (sorularr[i] == null)
                {
                    MessageBox.Show("Soru bulunamadı.");
                    break;
                }

                if (sorularr[i].Contains(textBox12.Text) == true)
                {
                    index = i;
                    textBox12.Text = sorularr[index].ToString();
                    break;
                }
                else
                {

                }


            }
        }

        private void button28_Click_1(object sender, EventArgs e)
        {
            if (textBox12.Text == "")
            {
                MessageBox.Show("Kelime girmeyi unuttunuz.");
            }
            if (index != -1)
            {
                OleDbCommand kmt2 = new OleDbCommand();
                baglantim.Open();
                kmt2.Connection = baglantim;

                kmt2.CommandText = "DELETE FROM SORULAR WHERE ID=" + sid[index].ToString() + "";
                kmt2.ExecuteNonQuery();
                MessageBox.Show("Silme işlemi başarı ile gerçekleşti.");
                verig();
                baglantim.Close();
                textBox12.Clear();

            }

        }

        private void button29_Click_1(object sender, EventArgs e)
        {

            button24.Visible = true;
            button25.Visible = true;
            groupBox10.Visible = true;
            button37.Visible = true;
            label29.Visible = true;
            textBox16.Visible = true;
            idS1 = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            idS2 = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            textBox16.Text = idS1.ToString();

        }

        private void button30_Click_2(object sender, EventArgs e)
        {
            MessageBox.Show("Lütfen sol taraftaki tablodan ders seçiniz");
            label29.Visible = false;
            button24.Visible = false;
            button25.Visible = false;
            groupBox10.Visible = false;
            textBox16.Visible = false;
            button37.Visible = false;
        }

        private void button37_Click_1(object sender, EventArgs e)
        {
            label29.Visible = false;
            textBox16.Visible = false;
            button37.Visible = false;
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if (textBox16.Text == "")
            {
                MessageBox.Show("Ders seçtiğinizden emin olunuz. Böyle bir ders olmayabilir. Ders ekleme işlemini yapabilirsiniz.");

                button24.Visible = false;
                button25.Visible = false;
                groupBox10.Visible = false;
                button37.Visible = false;
                label29.Visible = false;
                textBox16.Visible = false;
                return;
            }
            if (textBox20.Text == "")
            {
                MessageBox.Show("Soru sayısını boş bırakamazsınız.");
                return;
            }
            int aB = Convert.ToInt32(textBox20.Text);



            string[] sorular = new string[100];
            string[] aCevabi = new string[100];
            string[] bCevabi = new string[100];
            string[] cCevabi = new string[100];
            string[] dCevabi = new string[100];
            string[] eCevabi = new string[100];
            string[] dogrucevap = new string[100];

            int[] ac = new int[100];
            int[] bc = new int[100];
            int[] cc = new int[100];
            int[] dc = new int[100];
            int[] ec = new int[100];


            baglantim.Open();

            OleDbCommand komut3 = new OleDbCommand();
            komut3.Connection = baglantim;
            komut3.CommandText = ("select SORUADI,A,B,C,D,E,DOGRUCEVAP,SEKIL from SORULAR where DERSID=" + idS2.ToString() + "");
            OleDbDataReader oku = komut3.ExecuteReader();

            int i = 1, soruSayisi = 1;
            if (oku.HasRows)
            {
                while (oku.Read())
                {
                    sorular[soruSayisi] = oku[0].ToString();
                    aCevabi[i] = oku[1].ToString();
                    bCevabi[i] = oku[2].ToString();
                    cCevabi[i] = oku[3].ToString();
                    dCevabi[i] = oku[4].ToString();
                    eCevabi[i] = oku[5].ToString();
                    dogrucevap[i] = oku[6].ToString();

                    i++;
                    soruSayisi++;
                }
            }

            i--;
            int say = 1;
            if (i < aB)
            {
                MessageBox.Show("Veritabanınızda o kadar soru mevcut değil. Lütfen başka sayı giriniz ya da soru ekleme kısmından yeni sorular ekleyiniz.");
                return;
            }
            while (i != 0)
            {
                ac[say] = Say(aCevabi[say]);
                bc[say] = Say(bCevabi[say]);
                cc[say] = Say(cCevabi[say]);
                dc[say] = Say(dCevabi[say]);
                ec[say] = Say(eCevabi[say]);

                i--;
                say++;
            }

            oku.Close();
            baglantim.Close();
            baglantim.Close();




            wordeaktar.Application wordapp = new wordeaktar.Application();
            wordapp.Visible = true;
            wordeaktar.Document worddoc;
            object wordobj = System.Reflection.Missing.Value;
            worddoc = wordapp.Documents.Add(ref wordobj);
            Microsoft.Office.Interop.Word.Range drange = worddoc.Range();
            worddoc.PageSetup.HeaderDistance = 1;
            worddoc.PageSetup.LeftMargin = 8;
            worddoc.PageSetup.RightMargin = 8;
            worddoc.PageSetup.FooterDistance = 1;

            worddoc.PageSetup.TextColumns.SetCount(2);

            object dokson = "\\endofdoc";

            String grupAdi1 = "";
            string grupAdi2 = "";
            if (checkBox2.Checked == true)
            {
                grupAdi1 = "A GRUBU";
                grupAdi2 = "B GRUBU";
            }

            foreach (wordeaktar.Section wordSection in worddoc.Sections)
            {
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    .Range.Font.Size = 11;
                int abc = 9608;
                char x;
                x = (char)abc;
                //    label9.Text = "                                                                                                                                                                                                                                  ";
                label9.Text = "                                                                                                                                                                                                                      ";


                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + "\n" + textBox18.Text + "\n" + idS1.ToString() + "\n" + "SINAV SORULARI" + "\n " + grupAdi1.ToString() + "\n" + dateTimePicker1.Text + "        " + "Süre:" + textBox2.Text + "\n" + "Öğrenci No:                               " + "Adı Soyadı:                                              " + "\n";

                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.ColorIndex = WdColorIndex.wdBlack;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                  WdParagraphAlignment.wdAlignParagraphLeft;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[2].Alignment =
                 WdParagraphAlignment.wdAlignParagraphCenter;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[3].Alignment =
                    WdParagraphAlignment.wdAlignParagraphCenter;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[4].Alignment =
                         wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[5].Alignment =
           WdParagraphAlignment.wdAlignParagraphCenter;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[6].Alignment =
          WdParagraphAlignment.wdAlignParagraphCenter;



                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[7].Alignment =
                  WdParagraphAlignment.wdAlignParagraphLeft;


            }


            if (aB < 10)
            {
                wordapp.Selection.Font.Size = 12;

            }
            if (aB > 10 && aB <= 15)
            {
                wordapp.Selection.Font.Size = 10;
            }
            if (aB > 15 && aB <= 20)
            {
                wordapp.Selection.Font.Size = 7;
            }

            Random rand = new Random();
            int[] dizi1 = new int[soruSayisi];
            int[] dizi2 = new int[soruSayisi];

            for (int ca = 1; ca <= soruSayisi; ca++)
                dizi1[ca - 1] = ca;





            for (int j = 1; j <= aB;)
            {




                int asdee, bcd;

                bcd = rand.Next(1, soruSayisi);
                asdee = dizi1[bcd - 1];

                if (asdee != 0)
                {



                    dizi2[j] = asdee;


                    int abc = 40005;
                    char x;
                    x = (char)abc;




                    if (ac[asdee] < 10 && bc[asdee] < 10 && cc[asdee] > 30 && dc[asdee] < 30 && ec[asdee] < 30)

                    {



                        wordapp.Selection.TypeText(j + "." + sorular[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("c) " + cCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("d) " + dCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e) " + eCevabi[asdee].ToString() + "\n");

                    }
                    else if (ac[asdee] < 10 && bc[asdee] < 10 && cc[asdee] < 10 && dc[asdee] < 10 && ec[asdee] < 10)
                    {

                        wordapp.Selection.TypeText(j + "." + sorular[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("c) " + cCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("d) " + dCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("e) " + eCevabi[asdee].ToString() + "\n");
                    }
                    else if (ac[asdee] < 20 && bc[asdee] < 20 && cc[asdee] < 20 && dc[asdee] > 8 && ec[asdee] > 8)
                    {
                        wordapp.Selection.TypeText(j + "." + sorular[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("c) " + cCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("d) " + dCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("e) " + eCevabi[asdee].ToString() + "\n");
                    }
                    else if (ac[asdee] < 15 && bc[asdee] < 15 && cc[asdee] < 15 && dc[asdee] < 15)
                    {
                        wordapp.Selection.TypeText(j + "." + sorular[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("c) " + cCevabi[asdee].ToString() + " ");
                        wordapp.Selection.TypeText("d) " + dCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e) " + eCevabi[asdee].ToString() + "\n");
                    }
                    else if (ac[asdee] < 27 && bc[asdee] < 27 && cc[asdee] < 27 && dc[asdee] < 27)
                    {
                        wordapp.Selection.TypeText(j + "." + sorular[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("c) " + cCevabi[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("d) " + dCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e) " + eCevabi[asdee].ToString() + "\n");
                    }





                    else
                    {

                        wordapp.Selection.TypeText(j + "." + sorular[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a)" + aCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("b)" + bCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("c)" + cCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("d)" + dCevabi[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e)" + eCevabi[asdee].ToString() + "\n");

                    }


                    j++;

                }





                dizi1[bcd - 1] = 0;

            }

            if (checkBox1.Checked == true)
            {
                int r1 = 0, cab1 = 0;
                string strText1;
                if (aB < 10)

                {
                    wordeaktar.Table tabloA;
                    wordeaktar.Range wrdrng = worddoc.Bookmarks.get_Item(ref dokson).Range;
                    tabloA = worddoc.Tables.Add(wrdrng, aB, 6, ref wordobj, ref wordobj);
                    tabloA.Borders.OutsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleDouble;
                    tabloA.Borders.InsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;


                    for (r1 = 1; r1 <= aB; r1++)
                    {
                        for (cab1 = 1; cab1 <= 6; cab1++)
                        {


                            if (cab1 == 1)
                            {
                                strText1 = r1.ToString();
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 2)
                            {
                                strText1 = "A";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 3)
                            {
                                strText1 = "B";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 4)
                            {
                                strText1 = "C";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 5)
                            {
                                strText1 = "D";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 6)
                            {
                                strText1 = "E";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }


                            tabloA.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                            tabloA.Cell(r1, cab1).Range.Cells.Height = (float)11.50;





                        }







                    }


                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //worddoc.Sections[worddoc.Content.Sections.Count].PageSetup.TextColumns.SetCount(1);




                }
                else
                {
                    wordeaktar.Table tabloB;
                    wordeaktar.Range wrdrng1 = worddoc.Bookmarks.get_Item(ref dokson).Range;
                    tabloB = worddoc.Tables.Add(wrdrng1, (aB / 2) + 1, 12, ref wordobj, ref wordobj);


                    tabloB.Borders.OutsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleDouble;
                    tabloB.Borders.InsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;



                    for (r1 = 1; r1 <= (aB) / 2 + 1; r1++)
                        for (cab1 = 1; cab1 < 13; cab1++)
                        {



                            if (cab1 == 1)
                            {
                                strText1 = r1.ToString();
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }


                            else if (cab1 == 2)
                            {
                                strText1 = "A";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 3)
                            {
                                strText1 = "B";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 4)
                            {
                                strText1 = "C";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 5)
                            {
                                strText1 = "D";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 6)
                            {
                                strText1 = "E";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }


                            if ((r1 + (aB) / 2) + 1 <= aB)
                            {

                                if (cab1 == 7)
                                {

                                    strText1 = ((int)(r1 + (aB) / 2 + 1)).ToString();
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;

                                }


                                else if (cab1 == 8)
                                {
                                    strText1 = "A";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 9)
                                {
                                    strText1 = "B";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 10)
                                {
                                    strText1 = "C";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 11)
                                {
                                    strText1 = "D";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 12)
                                {
                                    strText1 = "E";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }
                            }

                            tabloB.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                            tabloB.Cell(r1, cab1).Range.Cells.Height = (float)11.50;


                        }

                }
            }





            foreach (wordeaktar.Section wordSection in worddoc.Sections)
            {


                int abc = 9608;
                char x;
                x = (char)abc;
                //   label9.Text = "                                                                                                                                                                                                                                  ";
                label9.Text = "                                                                                                                                                                                                                      ";

                wordSection.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + " " + x.ToString();
                wordSection.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                 WdParagraphAlignment.wdAlignParagraphLeft;

                wordSection.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    .Range.Font.Size = 11;





            }


            if (checkBox2.Checked == true)
            {
                wordeaktar1.Application wordapp1 = new wordeaktar1.Application();
                wordapp1.Visible = true;
                wordeaktar1.Document worddoc1;
                object wordobj1 = System.Reflection.Missing.Value;
                worddoc1 = wordapp1.Documents.Add(ref wordobj1);
                Microsoft.Office.Interop.Word.Range drange1 = worddoc1.Range();
                worddoc1.PageSetup.HeaderDistance = 1;
                worddoc1.PageSetup.LeftMargin = 8;
                worddoc1.PageSetup.RightMargin = 8;
                worddoc1.PageSetup.FooterDistance = 1;

                if (aB < 10)
                {
                    wordapp1.Selection.Font.Size = 12;

                }
                if (aB > 10 && aB <= 15)
                {
                    wordapp1.Selection.Font.Size = 10;
                }
                if (aB > 15 && aB <= 20)
                {
                    wordapp1.Selection.Font.Size = 7;
                }

                worddoc1.PageSetup.TextColumns.SetCount(2);
                foreach (wordeaktar1.Section wordSection1 in worddoc1.Sections)
                {
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                        .Range.Font.Size = 11;
                    int abc = 9608;
                    char x;
                    x = (char)abc;
                    //    label9.Text = "                                                                                                                                                                                                                                     ";
                    label9.Text = "                                                                                                                                                                                                                      ";


                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + "\n" + textBox18.Text + "\n" + idS1.ToString() + "\n" + "SINAV SORULARI" + "\n " + grupAdi2.ToString() + "\n" + dateTimePicker1.Text + "        " + "Süre:" + textBox2.Text + "\n" + "Öğrenci No:                               " + "Adı Soyadı:                                              " + "\n";

                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.ColorIndex = WdColorIndex.wdBlack;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                      WdParagraphAlignment.wdAlignParagraphLeft;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[2].Alignment =
                     WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[3].Alignment =
                        WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[4].Alignment =
                       WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[5].Alignment =
                     WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[6].Alignment =
                     WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[7].Alignment =
                      WdParagraphAlignment.wdAlignParagraphLeft;


                }
                for (int ca = 1; ca <= soruSayisi; ca++)
                    dizi1[ca - 1] = ca;



                for (int j = 1; j <= aB;)
                {

                    int abcd, kl;

                    kl = rand.Next(1, aB + 1);

                    abcd = dizi2[kl];

                    if (abcd != 0)
                    {



                        if (ac[abcd] < 10 && bc[abcd] < 10 && cc[abcd] > 30 && dc[abcd] < 30 && ec[abcd] < 30)

                        {

                            wordapp1.Selection.TypeText(j + "." + sorular[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("c) " + cCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("d) " + dCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e) " + eCevabi[abcd].ToString() + "\n");
                        }
                        else if (ac[abcd] < 10 && bc[abcd] < 10 && cc[abcd] < 10 && dc[abcd] < 10 && ec[abcd] < 10)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("c) " + cCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("d) " + dCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("e) " + eCevabi[abcd].ToString() + "\n");
                        }
                        else if (ac[abcd] < 20 && bc[abcd] < 20 && cc[abcd] < 20 && dc[abcd] > 8 && ec[abcd] > 8)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("c) " + cCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("d) " + dCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("e) " + eCevabi[abcd].ToString() + "\n");
                        }
                        else if (ac[abcd] < 15 && bc[abcd] < 15 && cc[abcd] < 15 && dc[abcd] < 15)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("c) " + cCevabi[abcd].ToString() + " ");
                            wordapp1.Selection.TypeText("d) " + dCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e) " + eCevabi[abcd].ToString() + "\n");
                        }
                        else if (ac[abcd] < 27 && bc[abcd] < 27 && cc[abcd] < 27 && dc[abcd] < 27)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("c) " + cCevabi[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("d) " + dCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e) " + eCevabi[abcd].ToString() + "\n");
                        }





                        else
                        {

                            wordapp1.Selection.TypeText(j + "." + sorular[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a)" + aCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("b)" + bCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("c)" + cCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("d)" + dCevabi[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e)" + eCevabi[abcd].ToString() + "\n");

                        }


                        j++;

                    }
                    dizi2[kl] = 0;

                }
                if (checkBox1.Checked == true)
                {
                    int r1 = 0, cab1 = 0;
                    string strText1;
                    if (aB < 10)

                    {
                        wordeaktar1.Table tabloA;
                        wordeaktar1.Range wrdrng = worddoc1.Bookmarks.get_Item(ref dokson).Range;


                        tabloA = worddoc1.Tables.Add(wrdrng, aB, 6, ref wordobj, ref wordobj);
                        tabloA.Borders.OutsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleDouble;
                        tabloA.Borders.InsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleSingle;


                        for (r1 = 1; r1 <= aB; r1++)
                        {
                            for (cab1 = 1; cab1 <= 6; cab1++)
                            {


                                if (cab1 == 1)
                                {
                                    strText1 = r1.ToString();
                                    tabloA.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 2)
                                {
                                    strText1 = "A";
                                    tabloA.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 3)
                                {
                                    strText1 = "B";
                                    tabloA.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 4)
                                {
                                    strText1 = "C";
                                    tabloA.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 5)
                                {
                                    strText1 = "D";
                                    tabloA.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 6)
                                {
                                    strText1 = "E";
                                    tabloA.Cell(r1, cab1).Range.Text = strText1;
                                }


                                tabloA.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                                tabloA.Cell(r1, cab1).Range.Cells.Height = (float)11.50;





                            }







                        }


                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //worddoc.Sections[worddoc.Content.Sections.Count].PageSetup.TextColumns.SetCount(1);




                    }
                    else
                    {
                        wordeaktar1.Table tabloB;
                        wordeaktar1.Range wrdrng1 = worddoc1.Bookmarks.get_Item(ref dokson).Range;
                        tabloB = worddoc1.Tables.Add(wrdrng1, (aB / 2) + 1, 12, ref wordobj, ref wordobj);


                        tabloB.Borders.OutsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleDouble;
                        tabloB.Borders.InsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleSingle;



                        for (r1 = 1; r1 <= (aB) / 2 + 1; r1++)
                            for (cab1 = 1; cab1 < 13; cab1++)
                            {



                                if (cab1 == 1)
                                {
                                    strText1 = r1.ToString();
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }


                                else if (cab1 == 2)
                                {
                                    strText1 = "A";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 3)
                                {
                                    strText1 = "B";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 4)
                                {
                                    strText1 = "C";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 5)
                                {
                                    strText1 = "D";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 6)
                                {
                                    strText1 = "E";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }


                                if ((r1 + (aB) / 2) + 1 <= aB)
                                {

                                    if (cab1 == 7)
                                    {

                                        strText1 = ((int)(r1 + (aB) / 2 + 1)).ToString();
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;

                                    }


                                    else if (cab1 == 8)
                                    {
                                        strText1 = "A";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 9)
                                    {
                                        strText1 = "B";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 10)
                                    {
                                        strText1 = "C";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 11)
                                    {
                                        strText1 = "D";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 12)
                                    {
                                        strText1 = "E";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }
                                }

                                tabloB.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                                tabloB.Cell(r1, cab1).Range.Cells.Height = (float)11.50;


                            }

                    }
                }
                foreach (wordeaktar1.Section wordSection1 in worddoc1.Sections)
                {


                    int abc = 9608;
                    char x;
                    x = (char)abc;
                    //   label9.Text = "                                                                                                                                                                                                                                  ";
                    label9.Text = "                                                                                                                                                                                                                      ";

                    wordSection1.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + " " + x.ToString();
                    wordSection1.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                     WdParagraphAlignment.wdAlignParagraphLeft;

                    wordSection1.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                        .Range.Font.Size = 11;





                }




            }


            

            wordapp = null;
            baglantim.Dispose();
            baglantim.Close();
            baglantim.Close();
            baglantim.Close();
          

        }

        private void button39_Click_1(object sender, EventArgs e)
        {

            label54.Text = count.ToString() + ".";
            label31.Visible = true;
            textBox10.Visible = true;
            button40.Visible = true;
            groupBox12.Visible = true;
            button26.Visible = true;
            button27.Visible = true;
            idS1 = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            idS2 = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox10.Text = idS1.ToString();

            verig1();

        }

        private void button38_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("Lütfen sol taraftaki tablodan ders seçiniz");
            label31.Visible = false;
            textBox10.Visible = false;
            textBox10.Clear();
            textBox21.Clear();
            button40.Visible = false;
            groupBox12.Visible = false;
            button26.Visible = false;
            button27.Visible = false;
        }

        private void button40_Click(object sender, EventArgs e)
        {
            label31.Visible = false;
            button40.Visible = false;
            textBox10.Visible=false;
        }

        private void button44_Click_1(object sender, EventArgs e)
        {
            button41.Visible = true;
            label44.Visible = true;
            textBox19.Visible = true;
            button42.Visible = true;
            idS1 = dataGridView2.CurrentRow.Cells[1].Value.ToString();

            textBox19.Text = idS1.ToString();

        }

        private void button42_Click_1(object sender, EventArgs e)
        {
            button41.Visible = false;
            sorusec[count] = dataGridView2.CurrentRow.Cells[0].Value.ToString();
            label44.Visible = false;
            textBox19.Visible = false;
            button42.Visible = false;
            textBox19.Clear();



            count++;
            label54.Text = count.ToString() + ".";
        }

        private void button41_Click_1(object sender, EventArgs e)
        {
            button41.Visible = false;
            textBox19.Clear();
            label44.Visible = false;
            textBox19.Visible = false;
            button42.Visible = false;
        }

        private void button27_Click_1(object sender, EventArgs e)
        {

            if (textBox10.Text == "")
            {
                MessageBox.Show("Ders seçtiğinizden emin olunuz. Böyle bir ders olmayabilir. Ders ekleme işlemini yapabilirsiniz.");

                button40.Visible = false;
                button26.Visible = false;
                groupBox12.Visible = false;
                button27.Visible = false;
                label31.Visible = false;
                textBox10.Visible = false;
                return;
            }

            int aB = count;

            int ijk = 1, soruSayisi = 1;






            for (ijk = 1; ijk < aB; ijk++)
            {
                baglantim.Open();

                OleDbCommand komut3 = new OleDbCommand();
                komut3.Connection = baglantim;
                komut3.CommandText = ("select SORUADI,A,B,C,D,E,DOGRUCEVAP,SEKIL from SORULAR where ID=" + sorusec[ijk].ToString() + "");
                OleDbDataReader oku = komut3.ExecuteReader();
                while (oku.Read())
                {
                    sorular1[ijk] = oku[0].ToString();
                    aCevabi1[ijk] = oku[1].ToString();
                    bCevabi1[ijk] = oku[2].ToString();
                    cCevabi1[ijk] = oku[3].ToString();
                    dCevabi1[ijk] = oku[4].ToString();
                    eCevabi1[ijk] = oku[5].ToString();
                    dogrucevap1[ijk] = oku[6].ToString();


                    soruSayisi++;
                }
                oku.Close();
                baglantim.Close();

            }




            ijk--;
            int say = 1;

            while (ijk != 0)
            {
                ac1[say] = Say(aCevabi1[say]);
                bc1[say] = Say(bCevabi1[say]);
                cc1[say] = Say(cCevabi1[say]);
                dc1[say] = Say(dCevabi1[say]);
                ec1[say] = Say(eCevabi1[say]);

                ijk--;
                say++;
            }


            baglantim.Close();
            baglantim.Close();




            wordeaktar.Application wordapp = new wordeaktar.Application();
            wordapp.Visible = true;
            wordeaktar.Document worddoc;
            object wordobj = System.Reflection.Missing.Value;
            worddoc = wordapp.Documents.Add(ref wordobj);
            Microsoft.Office.Interop.Word.Range drange = worddoc.Range();
            worddoc.PageSetup.HeaderDistance = 1;
            worddoc.PageSetup.LeftMargin = 8;
            worddoc.PageSetup.RightMargin = 8;
            worddoc.PageSetup.FooterDistance = 1;

            worddoc.PageSetup.TextColumns.SetCount(2);

            object dokson = "\\endofdoc";

            String grupAdi1 = "";
            string grupAdi2 = "";
            if (checkBox6.Checked == true)
            {
                grupAdi1 = "A GRUBU";
                grupAdi2 = "B GRUBU";
            }

            foreach (wordeaktar.Section wordSection in worddoc.Sections)
            {
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    .Range.Font.Size = 11;
                int abc = 9608;
                char x;
                x = (char)abc;
                //    label9.Text = "                                                                                                                                                                                                                                  ";
                label9.Text = "                                                                                                                                                                                                                      ";


                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + "\n" + textBox21.Text + "\n" + textBox10.Text + "\n" + "SINAV SORULARI" + "\n " + grupAdi1.ToString() + "\n" + dateTimePicker3.Text + "        " + "Süre:" + textBox17.Text + "\n" + "Öğrenci No:                               " + "Adı Soyadı:                                              " + "\n";

                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.ColorIndex = WdColorIndex.wdBlack;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                  WdParagraphAlignment.wdAlignParagraphLeft;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[2].Alignment =
                 WdParagraphAlignment.wdAlignParagraphCenter;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[3].Alignment =
                    WdParagraphAlignment.wdAlignParagraphCenter;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[4].Alignment =
                         wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[5].Alignment =
           WdParagraphAlignment.wdAlignParagraphCenter;
                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[6].Alignment =
          WdParagraphAlignment.wdAlignParagraphCenter;



                wordSection.Headers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[7].Alignment =
                  WdParagraphAlignment.wdAlignParagraphLeft;


            }


            if (aB < 10)
            {
                wordapp.Selection.Font.Size = 12;

            }
            if (aB > 10 && aB <= 15)
            {
                wordapp.Selection.Font.Size = 10;
            }
            if (aB > 15 && aB <= 20)
            {
                wordapp.Selection.Font.Size = 7;
            }






            Random rand = new Random();
            int[] dizi1 = new int[aB + 1];
            int[] dizi2 = new int[aB + 1];
            for (int ca = 1; ca <= aB; ca++)
                dizi1[ca - 1] = ca;




            for (int j = 1; j < count;)
            {
                int asdee, bcd;

                bcd = rand.Next(1, count);
                asdee = dizi1[bcd - 1];



                asdee = j;


                if (asdee != 0)
                {










                    if (ac1[asdee] < 10 && bc1[asdee] < 10 && cc1[asdee] > 30 && dc1[asdee] < 30 && ec1[asdee] < 30)

                    {



                        wordapp.Selection.TypeText(j + "." + sorular1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("c) " + cCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("d) " + dCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e) " + eCevabi1[asdee].ToString() + "\n");

                    }
                    else if (ac1[asdee] < 10 && bc1[asdee] < 10 && cc1[asdee] < 10 && dc1[asdee] < 10 && ec1[asdee] < 10)
                    {

                        wordapp.Selection.TypeText(j + "." + sorular1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("c) " + cCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("d) " + dCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("e) " + eCevabi1[asdee].ToString() + "\n");
                    }
                    else if (ac1[asdee] < 20 && bc1[asdee] < 20 && cc1[asdee] < 20 && dc1[asdee] > 8 && ec1[asdee] > 8)
                    {
                        wordapp.Selection.TypeText(j + "." + sorular1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("c) " + cCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("d) " + dCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("e) " + eCevabi1[asdee].ToString() + "\n");
                    }
                    else if (ac1[asdee] < 15 && bc1[asdee] < 15 && cc1[asdee] < 15 && dc1[asdee] < 15)
                    {
                        wordapp.Selection.TypeText(j + "." + sorular1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("c) " + cCevabi1[asdee].ToString() + " ");
                        wordapp.Selection.TypeText("d) " + dCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e) " + eCevabi1[asdee].ToString() + "\n");
                    }
                    else if (ac1[asdee] < 27 && bc1[asdee] < 27 && cc1[asdee] < 27 && dc1[asdee] < 27)
                    {
                        wordapp.Selection.TypeText(j + "." + sorular1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a) " + aCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("b) " + bCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("c) " + cCevabi1[asdee].ToString() + "  ");
                        wordapp.Selection.TypeText("d) " + dCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e) " + eCevabi1[asdee].ToString() + "\n");
                    }





                    else
                    {

                        wordapp.Selection.TypeText(j + "." + sorular1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("a)" + aCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("b)" + bCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("c)" + cCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("d)" + dCevabi1[asdee].ToString() + "\n");
                        wordapp.Selection.TypeText("e)" + eCevabi1[asdee].ToString() + "\n");

                    }


                    j++;

                }
                dizi1[bcd - 1] = 0;
            }









            aB = count;
            if (checkBox5.Checked == true)
            {
                int r1 = 0, cab1 = 0;
                string strText1;
                if (aB < 10)

                {
                    wordeaktar.Table tabloA;
                    wordeaktar.Range wrdrng = worddoc.Bookmarks.get_Item(ref dokson).Range;


                    tabloA = worddoc.Tables.Add(wrdrng, aB - 1, 6, ref wordobj, ref wordobj);
                    tabloA.Borders.OutsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleDouble;
                    tabloA.Borders.InsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;


                    for (r1 = 1; r1 < aB; r1++)
                    {
                        for (cab1 = 1; cab1 <= 6; cab1++)
                        {


                            if (cab1 == 1)
                            {
                                strText1 = r1.ToString();
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 2)
                            {
                                strText1 = "A";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 3)
                            {
                                strText1 = "B";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 4)
                            {
                                strText1 = "C";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 5)
                            {
                                strText1 = "D";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }

                            if (cab1 == 6)
                            {
                                strText1 = "E";
                                tabloA.Cell(r1, cab1).Range.Text = strText1;
                            }


                            tabloA.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                            tabloA.Cell(r1, cab1).Range.Cells.Height = (float)11.50;





                        }







                    }


                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    //worddoc.Sections[worddoc.Content.Sections.Count].PageSetup.TextColumns.SetCount(1);




                }
                else
                {
                    wordeaktar.Table tabloB;
                    wordeaktar.Range wrdrng1 = worddoc.Bookmarks.get_Item(ref dokson).Range;
                    tabloB = worddoc.Tables.Add(wrdrng1, (aB / 2)+1, 12, ref wordobj, ref wordobj);


                    tabloB.Borders.OutsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleDouble;
                    tabloB.Borders.InsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;



                    for (r1 = 1; r1 <=(aB) / 2 + 1; r1++)
                        for (cab1 = 1; cab1 < 13; cab1++)
                        {



                            if (cab1 == 1)
                            {
                                strText1 = r1.ToString();
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }


                            else if (cab1 == 2)
                            {
                                strText1 = "A";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 3)
                            {
                                strText1 = "B";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 4)
                            {
                                strText1 = "C";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 5)
                            {
                                strText1 = "D";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }

                            else if (cab1 == 6)
                            {
                                strText1 = "E";
                                tabloB.Cell(r1, cab1).Range.Text = strText1;
                            }


                            if ((r1 + (aB) / 2) + 1 < aB)
                            {

                                if (cab1 == 7)
                                {

                                    strText1 = ((int)(r1 + (aB) / 2 +1)).ToString();
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;

                                }


                                else if (cab1 == 8)
                                {
                                    strText1 = "A";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 9)
                                {
                                    strText1 = "B";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 10)
                                {
                                    strText1 = "C";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 11)
                                {
                                    strText1 = "D";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 12)
                                {
                                    strText1 = "E";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }
                            }

                            tabloB.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                            tabloB.Cell(r1, cab1).Range.Cells.Height = (float)11.50;


                        }

                }
            }





            foreach (wordeaktar.Section wordSection in worddoc.Sections)
            {


                int abc = 9608;
                char x;
                x = (char)abc;
                //   label9.Text = "                                                                                                                                                                                                                                  ";
                label9.Text = "                                                                                                                                                                                                                      ";

                wordSection.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + " " + x.ToString();
                wordSection.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                 WdParagraphAlignment.wdAlignParagraphLeft;

                wordSection.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    .Range.Font.Size = 11;





            }


            if (checkBox6.Checked == true)
            {
                wordeaktar1.Application wordapp1 = new wordeaktar1.Application();
                wordapp1.Visible = true;
                wordeaktar1.Document worddoc1;
                object wordobj1 = System.Reflection.Missing.Value;
                worddoc1 = wordapp1.Documents.Add(ref wordobj1);
                Microsoft.Office.Interop.Word.Range drange1 = worddoc1.Range();
                worddoc1.PageSetup.HeaderDistance = 1;
                worddoc1.PageSetup.LeftMargin = 8;
                worddoc1.PageSetup.RightMargin = 8;
                worddoc1.PageSetup.FooterDistance = 1;

                if (aB < 10)
                {
                    wordapp1.Selection.Font.Size = 12;

                }
                if (aB > 10 && aB <= 15)
                {
                    wordapp1.Selection.Font.Size = 10;
                }
                if (aB > 15 && aB <= 20)
                {
                    wordapp1.Selection.Font.Size = 7;
                }

                worddoc1.PageSetup.TextColumns.SetCount(2);
                foreach (wordeaktar1.Section wordSection1 in worddoc1.Sections)
                {
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                        .Range.Font.Size = 11;
                    int abc = 9608;
                    char x;
                    x = (char)abc;
                    //    label9.Text = "                                                                                                                                                                                                                                     ";
                    label9.Text = "                                                                                                                                                                                                                      ";


                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + "\n" + textBox21.Text + "\n" + textBox10.Text + "\n" + "SINAV SORULARI" + "\n " + grupAdi2.ToString() + "\n" + dateTimePicker3.Text + "        " + "Süre:" + textBox17.Text + "\n" + "Öğrenci No:                               " + "Adı Soyadı:                                              " + "\n";

                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.ColorIndex = WdColorIndex.wdBlack;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                      WdParagraphAlignment.wdAlignParagraphLeft;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[2].Alignment =
                     WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[3].Alignment =
                        WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[4].Alignment =
                       WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[5].Alignment =
                     WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[6].Alignment =
                     WdParagraphAlignment.wdAlignParagraphCenter;
                    wordSection1.Headers[wordeaktar1.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[7].Alignment =
                      WdParagraphAlignment.wdAlignParagraphLeft;


                }

                for (int ca = 1; ca < aB; ca++)
                    dizi1[ca] = ca;


                for (int j = 1; j < aB;)
                {
                    int abcd, kl;

                    kl = rand.Next(1, aB + 1);

                    abcd = dizi1[kl];

                    if (abcd != 0)
                    {


                        if (ac1[abcd] < 10 && bc1[abcd] < 10 && cc1[abcd] > 30 && dc1[abcd] < 30 && ec1[abcd] < 30)

                        {

                            wordapp1.Selection.TypeText(j + "." + sorular1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("c) " + cCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("d) " + dCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e) " + eCevabi1[abcd].ToString() + "\n");
                        }
                        else if (ac1[abcd] < 10 && bc1[abcd] < 10 && cc1[abcd] < 10 && dc1[abcd] < 10 && ec1[abcd] < 10)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("c) " + cCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("d) " + dCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("e) " + eCevabi1[abcd].ToString() + "\n");
                        }
                        else if (ac1[abcd] < 20 && bc1[abcd] < 20 && cc1[abcd] < 20 && dc1[abcd] > 8 && ec1[abcd] > 8)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("c) " + cCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("d) " + dCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("e) " + eCevabi1[abcd].ToString() + "\n");
                        }
                        else if (ac1[abcd] < 15 && bc1[abcd] < 15 && cc1[abcd] < 15 && dc1[abcd] < 15)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("c) " + cCevabi1[abcd].ToString() + " ");
                            wordapp1.Selection.TypeText("d) " + dCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e) " + eCevabi1[abcd].ToString() + "\n");
                        }
                        else if (ac1[abcd] < 27 && bc1[abcd] < 27 && cc1[abcd] < 27 && dc1[abcd] < 27)
                        {
                            wordapp1.Selection.TypeText(j + "." + sorular1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a) " + aCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("b) " + bCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("c) " + cCevabi1[abcd].ToString() + "  ");
                            wordapp1.Selection.TypeText("d) " + dCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e) " + eCevabi1[abcd].ToString() + "\n");
                        }





                        else
                        {

                            wordapp1.Selection.TypeText(j + "." + sorular1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("a)" + aCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("b)" + bCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("c)" + cCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("d)" + dCevabi1[abcd].ToString() + "\n");
                            wordapp1.Selection.TypeText("e)" + eCevabi1[abcd].ToString() + "\n");

                        }


                        j++;

                    }

                    dizi1[kl] = 0;
                }





                if (checkBox5.Checked == true)
                {
                    int r1 = 0, cab1 = 0;
                    string strText1;
                    if (aB < 10)

                    {
                        wordeaktar1.Table tabloA1;
                        wordeaktar1.Range wrdrng2 = worddoc1.Bookmarks.get_Item(ref dokson).Range;


                        tabloA1 = worddoc1.Tables.Add(wrdrng2, aB - 1, 6, ref wordobj, ref wordobj);
                        tabloA1.Borders.OutsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleDouble;
                        tabloA1.Borders.InsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleSingle;


                        for (r1 = 1; r1 < aB; r1++)
                        {
                            for (cab1 = 1; cab1 <= 6; cab1++)
                            {


                                if (cab1 == 1)
                                {
                                    strText1 = r1.ToString();
                                    tabloA1.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 2)
                                {
                                    strText1 = "A";
                                    tabloA1.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 3)
                                {
                                    strText1 = "B";
                                    tabloA1.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 4)
                                {
                                    strText1 = "C";
                                    tabloA1.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 5)
                                {
                                    strText1 = "D";
                                    tabloA1.Cell(r1, cab1).Range.Text = strText1;
                                }

                                if (cab1 == 6)
                                {
                                    strText1 = "E";
                                    tabloA1.Cell(r1, cab1).Range.Text = strText1;
                                }


                                tabloA1.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                                tabloA1.Cell(r1, cab1).Range.Cells.Height = (float)11.50;





                            }







                        }


                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //tabloA.Cell(r, cab).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //worddoc.Sections[worddoc.Content.Sections.Count].PageSetup.TextColumns.SetCount(1);




                    }
                    else
                    {
                        wordeaktar1.Table tabloB;
                        wordeaktar1.Range wrdrng1 = worddoc1.Bookmarks.get_Item(ref dokson).Range;
                        tabloB = worddoc1.Tables.Add(wrdrng1, (aB / 2)+1, 12, ref wordobj, ref wordobj);


                        tabloB.Borders.OutsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleDouble;
                        tabloB.Borders.InsideLineStyle = wordeaktar1.WdLineStyle.wdLineStyleSingle;



                        for (r1 = 1; r1 <= (aB) / 2 + 1; r1++)
                            for (cab1 = 1; cab1 < 13; cab1++)
                            {



                                if (cab1 == 1)
                                {
                                    strText1 = r1.ToString();
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }


                                else if (cab1 == 2)
                                {
                                    strText1 = "A";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 3)
                                {
                                    strText1 = "B";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 4)
                                {
                                    strText1 = "C";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 5)
                                {
                                    strText1 = "D";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }

                                else if (cab1 == 6)
                                {
                                    strText1 = "E";
                                    tabloB.Cell(r1, cab1).Range.Text = strText1;
                                }


                                if ((r1 + (aB) / 2) + 1 <= aB)
                                {

                                    if (cab1 == 7)
                                    {

                                        strText1 = ((int)(r1 + (aB) / 2 + 1)).ToString();
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;

                                    }


                                    else if (cab1 == 8)
                                    {
                                        strText1 = "A";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 9)
                                    {
                                        strText1 = "B";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 10)
                                    {
                                        strText1 = "C";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 11)
                                    {
                                        strText1 = "D";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }

                                    else if (cab1 == 12)
                                    {
                                        strText1 = "E";
                                        tabloB.Cell(r1, cab1).Range.Text = strText1;
                                    }
                                }

                                tabloB.Cell(r1, cab1).Range.Cells.Width = (float)20.50;
                                tabloB.Cell(r1, cab1).Range.Cells.Height = (float)11.50;


                            }

                    }
                }
                foreach (wordeaktar1.Section wordSection1 in worddoc1.Sections)
                {


                    int abc = 9608;
                    char x;
                    x = (char)abc;
                    //   label9.Text = "                                                                                                                                                                                                                                  ";
                    label9.Text = "                                                                                                                                                                                                                      ";

                    wordSection1.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = x.ToString() + " " + x.ToString() + " " + x.ToString() + label9.Text + x.ToString() + " " + x.ToString();
                    wordSection1.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Alignment =
                     WdParagraphAlignment.wdAlignParagraphLeft;

                    wordSection1.Footers[wordeaktar.WdHeaderFooterIndex.wdHeaderFooterPrimary]
                        .Range.Font.Size = 11;





                }





            }




            wordapp = null;
            baglantim.Dispose();
            baglantim.Close();
            baglantim.Close();
            baglantim.Close();


            Form.ActiveForm.Close();

        }

        private void button31_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabPage1);
            anaSayfa();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabPage2);
            anaSayfa();
        }

        private void button36_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabPage3);
            anaSayfa();
        }

        private void button35_Click(object sender, EventArgs e)
        {
            Form.ActiveForm.Close();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            Form.ActiveForm.Close();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            Form.ActiveForm.Close();
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
            textBox9.Visible = true;
            button16.Visible = true;
            button13.Visible = true;

        }

        private void button24_Click_1(object sender, EventArgs e)
        {
            textBox16.Clear();
            textBox20.Clear();
            textBox18.Clear();
            checkBox2.Checked = false;
            label29.Visible = false;
            button24.Visible = false;
            button25.Visible = false;
            groupBox10.Visible = false;
            textBox16.Visible = false;
            button37.Visible = false;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            groupBox12.Visible = false;

            label31.Visible = false;
            button40.Visible = false;
            textBox10.Visible = false;
            button26.Visible = false;
            button27.Visible = false;
        }
    }
       
}





        
    

