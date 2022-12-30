using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MySql.Data.MySqlClient;

namespace Personel_Takip
{
    public partial class Form1 : Form
    {
        public string str = "datasource = localhost ;port=3306; database = personel ; username = root ; password=;Convert Zero Datetime=True";
        public int  sayac = 20;
        public string islem = "";
        public string kartno  = "";

        Form2 Frm2;

        public Form1()
        {
            InitializeComponent();
            timer1.Start();
            
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            sayac = 40;
            kartno = "";

            if (e.KeyCode == Keys.Enter)
            {
                if (textBox1.Text != "0011130780")
                {
                    /*MessageBox.Show(textBox1.Text, "");*/

                    MySqlConnection Con = new MySqlConnection(str);
                    try
                    {
                        Con.Open();
                        MySqlCommand veri_cek = new MySqlCommand("SELECT * FROM personelbilgileri where kartid='" + textBox1.Text + "'", Con);
                        MySqlDataReader alindi = veri_cek.ExecuteReader();
                        if (alindi.HasRows)
                        {
                            while (alindi.Read())
                            {
                                if (alindi.GetString(4) != "")
                                {
                                    MemoryStream ms = new MemoryStream((byte[])alindi[5]);
                                    pictureBox1.Image = Image.FromStream(ms);
                                }

                                kartno= alindi.GetString(1);

                                label2.Text = alindi.GetString(2);
                            }
                        }
                        Con.Close();
                    }
                    catch (Exception eb)
                    {
                        MessageBox.Show(eb.Message);
                    }

                    if (kartno == textBox1.Text)
                    {
                        if (radioButton1.Checked == true)
                        {
                            islem = "Giris";
                        }

                        if (radioButton2.Checked == true)
                        {
                            islem = "Cıkıs";
                        }

                        try
                        {
                            Con.Open();
                            MySqlCommand veri_yaz = new MySqlCommand("Insert into puantaj  (islem,kartid,adsoyad) values ('" + islem + "','" + textBox1.Text + "','" + label2.Text + "')", Con);
                            MySqlDataReader yazildi = veri_yaz.ExecuteReader();

                            Con.Close();

                        }
                        catch (Exception eb)
                        {
                            MessageBox.Show(eb.Message);
                        }
                        int i = 0;
                        MySqlCommand veri_cek2 = new MySqlCommand("SELECT *  FROM puantaj WHERE kartid='" + textBox1.Text + "'  order by tarihsaat desc limit 15", Con);
                        Con.Open();
                        MySqlDataReader alindi2 = veri_cek2.ExecuteReader();
                        listView1.Items.Clear();
                        while (alindi2.Read())
                        {
                            ListViewItem lv = new ListViewItem(alindi2.GetDateTime(1).ToString());
                            lv.SubItems.Add(alindi2.GetString(2).ToString());
                            lv.SubItems.Add(alindi2.GetString(4).ToString());
                            listView1.Items.Add(lv);

                            switch (alindi2.GetString(2).ToString())
                            {
                                case "Giris":
                                    listView1.Items[i].BackColor = Color.Green;
                                    listView1.Items[i].ForeColor = Color.White;
                                    break;

                                case "Cikis":
                                    listView1.Items[i].BackColor = Color.Red;
                                    listView1.Items[i].ForeColor = Color.Black;
                                    break;
                            }
                            i++;

                        }
                        Con.Close();


                        textBox1.Clear();
                    }
                }
                else
                {
                    int u = 0;
                    if (radioButton1.Checked && u == 0)
                    {
                        radioButton2.PerformClick();
                        u = 1;
                    }
                    if (radioButton2.Checked && u == 0)
                    {
                        radioButton1.PerformClick();
                        u = 1;
                    }
                    textBox1.Clear();
                }
                textBox1.Clear();

            }


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(sayac==0)
            {
                textBox1.Clear();
                pictureBox1.Image = null;
                listView1.Items.Clear();
            }
            sayac--;

            textBox1.Select();
        }

        private void radioButton1_Click(object sender, EventArgs e)
        {
            radioButton1.BackColor = Color.Red;
            radioButton2.BackColor = Color.White;
        }

        private void radioButton2_Click(object sender, EventArgs e)
        {
            radioButton2.BackColor = Color.Red;
            radioButton1.BackColor = Color.White;
        }

    

        private void button1_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Frm2 = new Form2();
            Frm2.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
    }

