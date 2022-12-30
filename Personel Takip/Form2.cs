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
using Microsoft.Office.Interop.Excel;

namespace Personel_Takip
{
    public partial class Form2 : Form
    {
        public string str = "datasource = localhost ;port=3306; database = personel ; username = root ; password=;Convert Zero Datetime=True";

        public Form2()
        {
            InitializeComponent();

            listele();
        }

        private void listele()
        {
            MySqlConnection Con = new MySqlConnection(str);
            MySqlCommand veri_cek = new MySqlCommand("SELECT * FROM personelbilgileri where aktif=1", Con);
            Con.Open();
            MySqlDataReader alindi = veri_cek.ExecuteReader();
            listView1.Items.Clear();
            while (alindi.Read())
            {
                ListViewItem lv = new ListViewItem(alindi.GetString(1).ToString());
                lv.SubItems.Add(alindi.GetString(2).ToString());
                lv.SubItems.Add(alindi.GetString(3).ToString());
                lv.SubItems.Add(alindi.GetString(4).ToString());
                listView1.Items.Add(lv);
            }
            Con.Close();
        }


        private void puantajsorgula()
        {
            String[] puantajtarihbas = dateTimePicker1.Text.Split('-');
            String[] puantajtarihson = dateTimePicker2.Text.Split('-');
            int tarihbasgun = Convert.ToInt32(puantajtarihbas[2]);
            int tarihsongun = Convert.ToInt32(puantajtarihson[2]);


            if (puantajtarihbas[1] == puantajtarihson[1])
            {
                if (listView1.SelectedItems.Count > 0)
                {
                    DateTime puantajgunu = dateTimePicker1.Value;
                    listView2.Items.Clear();                                      

                    while (tarihbasgun <= tarihsongun)
                    {
                        string puantajbas = puantajtarihbas[0] + "-" + puantajtarihbas[1] + "-" + puantajtarihbas[2] + " 00:00:00.000000";
                        string puantajson = puantajtarihbas[0] + "-" + puantajtarihbas[1] + "-" + puantajtarihbas[2] + " 23:59:59.999999";

                        MySqlConnection Con = new MySqlConnection(str);
                        MySqlCommand veri_cek = new MySqlCommand("SELECT *  FROM puantaj WHERE kartid='" + listView1.SelectedItems[0].SubItems[0].Text + "'  and tarihsaat between '"+puantajbas+"' and '"+puantajson+"'  order by tarihsaat asc ", Con);
                                                                                                                                           
                        Con.Open();
                        MySqlDataReader alindi = veri_cek.ExecuteReader();

                        int i = 0;

                        while (alindi.Read())
                        {
                            ListViewItem lv = new ListViewItem(alindi.GetDateTime(1).ToString());
                            lv.SubItems.Add(alindi.GetString(2).ToString());
                            listView2.Items.Add(lv);
                            i = 1;
                        }

                        //MessageBox.Show(tarihbasgun.ToString());

                        if(i==0)
                        {
                            ListViewItem lv = new ListViewItem(puantajgunu.Date.ToString());
                            lv.SubItems.Add("---");
                            listView2.Items.Add(lv);
                        }

                        Con.Close();
                        tarihbasgun++;
                        puantajtarihbas[2] = (tarihbasgun).ToString();
                        puantajgunu = puantajgunu.AddDays(1);
                    }
                }

            }

            else
                MessageBox.Show("Başlangıç ve Bitiş Tarihlerinde Aynı Ay Seçilmelidir");
        }

        
        private void button1_Click(object sender, EventArgs e)
        {           

            if(listView1.SelectedItems.Count>0)
            {
                puantajsorgula();
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            puantajsorgula();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
                    Worksheet ws = (Worksheet)app.ActiveSheet;
                    app.Visible = false;

                    
                    ws.Range[ws.Cells[1, 1], ws.Cells[2, 1]].Interior.Color = Color.DodgerBlue;
                    ws.Range[ws.Cells[1, 1], ws.Cells[2, 1]].Font.Bold = true;
                    ws.Cells[1, 1] = "AD - SOYAD :" ;
                    ws.Cells[2, 1] = "KART ID :";

                    ws.Cells[1, 2] = listView1.SelectedItems[0].SubItems[1].Text;
                    ws.Cells[2, 2] = listView1.SelectedItems[0].SubItems[0].Text;


                    ws.Range[ws.Cells[4, 1], ws.Cells[4, 2]].Interior.Color = Color.DodgerBlue;
                    ws.Range[ws.Cells[4, 1], ws.Cells[4, 2]].Font.Bold = true;
                    ws.Cells[4, 1] = "TARİH&SAAT";
                    ws.Cells[4, 2] = "İŞLEM";

                    int li = 5;

                    foreach (ListViewItem item in listView2.Items)
                    {
                        ws.Cells[li, 1] = item.SubItems[0].Text;
                        ws.Cells[li, 2] = item.SubItems[1].Text;
                        li++;
                    }
                    ws.Columns.AutoFit();
                    ws.Range[ws.Cells[1,1],ws.Cells[10,2]].Borders[XlBordersIndex.xlEdgeTop].Weight = 1;
                    ws.SaveAs(sfd.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing);
                    app.Quit();
                }
        }
    }
}
