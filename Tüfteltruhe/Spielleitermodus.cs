﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;

namespace Tüfteltruhe
{
    public partial class Spielleitermodus : Form
    {
        Random zufall = new Random();
        public DataTable ZutatenUmgebungTB = new DataTable();
        public DataTable ZutatenRegionTB = new DataTable();
        public DataTable ZutatenTB = new DataTable();
        public DataTable HaendlerTB = new DataTable();
        public DataTable SondergewerbeTB = new DataTable();
        public Bereich Region = new Bereich(0, "", null, null, null, null);
        public Bereich Umgebung = new Bereich(0, "", null, null, null, null);
        public Bereich Region2 = new Bereich(0, "", null, null, null, null);
        public Bereich Umgebung2 = new Bereich(0, "", null, null, null, null);
        public int zwgewoehnlich = 20;
        public int zwungewoehnlich = 25;
        public int zwselten = 25;
        public int wissennatur = -5;
        public int sammeln = -5;
        public int suchstunden = 1;
        public int testergebnis = 0;
        public string ergebniszutat = "";
        public string gewaehlte_region = "";
        public string gewaehlte_umgebung = "";
        public int rowcount = 0;
        public int rowcount2 = 0;
        public int rowcount3 = 0;
        public int portionenerhöht = 0;
        public int bisherigeportionenzahl = 0;
        public int gesamtstunden = 0;
        public int kenntniszw = 100;
        public int wiederholungen = 0;
        public string regionergebnis = "";
        public string umgebungergebnis = "xyz";
        public int erforderlichesuchstunden = 1000;
        public string nachricht = "";
        public string fokusergänzung = "";
        public int nebenfundteiler = 1000;
        public int testergebnis2 = 0;
        public int heutigesuchstunden = 0;
        public bool richtigeumgebung = false;
        public bool richtigeregion = false;
        public bool istgewöhnlich = false;
        public bool istungewöhnlich = false;
        public bool istselten = false;
        public bool istsehrselten = false;
        public int händleranzahl = 0;
        public int sondergewerbe = 0;
        public string haendlerergebnis = "";
        public string voraussetzung = "";
        public string haendlerbeschreibung = "";
        public string sonderergebnis = "";
        public string sondervoraussetzung = "";
        public string sonderbeschreibung = "";
        public string groesse = "";
        public bool eintragen = true;

        public Spielleitermodus()
        {
            InitializeComponent();
        }


        //#################### ################################################
        //### ZUTATENSUCHE ### ################################################
        //#################### ################################################

        private void button2_Click(object sender, EventArgs e) //Allgemeine Suche
        {
            if (gewaehlte_region == "" || comboBox1.GetItemText(comboBox1.SelectedItem) == "")
            {
                MessageBox.Show("Keine Region gewählt!");
                return;
            }
            //if (gewaehlte_umgebung == "" || comboBox5.GetItemText(comboBox5.SelectedItem) == "")
            //{
            //    MessageBox.Show("Keine Umgebung gewählt!");
            //    return;
            //}

            wissennatur = (int)numericUpDown1.Value;
            sammeln = (int)numericUpDown2.Value;
            suchstunden = (int)numericUpDown3.Value;
            zwgewoehnlich = 0;
            zwungewoehnlich = 0;
            zwselten = 0;

            switch (wissennatur)
            {
                case -10:
                case -9:
                case -8:
                case -7:
                case -6:
                case -5:
                case -4:
                case -3:
                case -2:
                case -1:
                    kenntniszw = 0;
                    break;
                case 0:
                case 1:
                    kenntniszw = 1;
                    break;
                case 2:
                case 3:
                    kenntniszw = 2;
                    break;
                case 4:
                case 5:
                    kenntniszw = 3;
                    break;
                case 6:
                case 7:
                    kenntniszw = 4;
                    break;
                case 8:
                case 9:
                    kenntniszw = 5;
                    break;
                case 10:
                case 11:
                    kenntniszw = 6;
                    break;
                case 12:
                case 13:
                case 14:
                case 15:
                case 16:
                case 17:
                case 18:
                case 19:
                case 20:
                case 21:
                    kenntniszw = 7;
                    break;
            }


            for (int i = 0; i < suchstunden; i++)
            {
                gesamtstunden++;
                label8.Text = "Ges. Stunden: " + gesamtstunden;
                testergebnis = zufall.Next(1, 7) + zufall.Next(1, 7) + zufall.Next(1, 7) + sammeln;
                label13.Text = "Letzter Wurf: " + (testergebnis - sammeln) + " + " + sammeln + " = " + testergebnis;

                switch (testergebnis)
                {
                    case -8:
                    case -7:
                    case -6:
                    case -5:
                    case -4:
                    case -3:
                    case -2:
                    case -1:
                    case 0:
                    case 1:
                    case 2:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 9:
                        break;
                    case 10:
                    case 11:
                    case 12:
                        zwgewoehnlich++;
                        break;
                    case 13:
                    case 14:
                    case 15:
                        zwgewoehnlich++;
                        zwgewoehnlich++;
                        break;
                    case 16:
                    case 17:
                    case 18:
                        zwungewoehnlich++;
                        zwgewoehnlich++;
                        break;
                    case 19:
                    case 20:
                    case 21:
                        zwselten++;
                        zwgewoehnlich++;
                        break;
                    case 22:
                    case 23:
                    case 24:
                        zwselten++;
                        zwungewoehnlich++;
                        break;
                    case 25:
                    case 26:
                    case 27:
                    case 28:
                    case 29:
                    case 30:
                    case 31:
                    case 32:
                    case 33:
                    case 34:
                    case 35:
                    case 36:
                    case 37:
                    case 38:
                        zwselten++;
                        zwungewoehnlich++;
                        zwgewoehnlich++;
                        zwgewoehnlich++;
                        break;
                    default:
                        break;
                }
            }

            for (int i = 0; i < zwselten; i++)
            {
                portionenerhöht = 0;

                //MODUS: SCHNITTMENGE
                //wiederholungen = 0;
                //while (wiederholungen < 10000 && regionergebnis != umgebungergebnis)
                //{
                //    regionergebnis = Region.selten[zufall.Next(0, Region.selten.Length - 1)];
                //    umgebungergebnis = Umgebung.selten[zufall.Next(0, Umgebung.selten.Length - 1)];
                //    wiederholungen++;
                //}
                //ergebniszutat = regionergebnis;

                //MODUS: SUMME
                ergebniszutat = Region.selten[zufall.Next(0, Region.selten.Length - 1)];
                if (comboBox5.GetItemText(comboBox5.SelectedItem) != "" && zufall.Next(0, 1) == 1) { ergebniszutat = Umgebung.selten[zufall.Next(0, Umgebung.selten.Length - 1)]; }

                //Erhöhende Ergebnisse!
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(ergebniszutat))
                    {
                        dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value = (int)dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value + zufall.Next(1, 3);
                        portionenerhöht = 1;
                        break;
                    }
                }

                //Ohne Erhöhung -> Neues Tabellenfeld
                if (portionenerhöht == 0)
                {
                    dataGridView1.Rows.Add();
                    rowcount++;
                    dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.YellowGreen;
                    dataGridView1[0, rowcount - 1].Value = ergebniszutat;
                    dataGridView1[2, rowcount - 1].Value = "Selten";
                    dataGridView1[1, rowcount - 1].Value = zufall.Next(1, 3);
                    if (zufall.Next(1, 10) > kenntniszw)
                    {
                        dataGridView1[3, rowcount - 1].Value = "Unbekannt!";
                        dataGridView1[3, rowcount - 1].Style.BackColor = Color.Orange;
                    }
                }
            }

            for (int i = 0; i < zwungewoehnlich; i++)
            {
                portionenerhöht = 0;

                //MODUS: SUMME
                ergebniszutat = Region.ungewöhnlich[zufall.Next(0, Region.ungewöhnlich.Length - 1)];
                if (comboBox5.GetItemText(comboBox5.SelectedItem) != "" && zufall.Next(0, 1) == 1) { ergebniszutat = Umgebung.ungewöhnlich[zufall.Next(0, Umgebung.ungewöhnlich.Length - 1)]; }

                //Erhöhende Ergebnisse!
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(ergebniszutat))
                    {
                        dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value = (int)dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value + zufall.Next(1, 5);
                        portionenerhöht = 1;
                        break;
                    }
                }

                //Ohne Erhöhung -> Neues Tabellenfeld
                if (portionenerhöht == 0)
                {
                    dataGridView1.Rows.Add();
                    rowcount++;
                    dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightGreen;
                    dataGridView1[0, rowcount - 1].Value = ergebniszutat;
                    dataGridView1[2, rowcount - 1].Value = "Ungewöhnlich";
                    dataGridView1[1, rowcount - 1].Value = zufall.Next(1, 5);
                    if (zufall.Next(1, 10) > kenntniszw + 1)
                    {
                        dataGridView1[3, rowcount - 1].Value = "Unbekannt!";
                        dataGridView1[3, rowcount - 1].Style.BackColor = Color.Orange;
                    }
                }
            }

            for (int i = 0; i < zwgewoehnlich; i++)
            {
                portionenerhöht = 0;

                //MODUS: SUMME
                ergebniszutat = Region.gewöhnlich[zufall.Next(0, Region.gewöhnlich.Length - 1)];
                if (comboBox5.GetItemText(comboBox5.SelectedItem) != "" && zufall.Next(0, 1) == 1) { ergebniszutat = Umgebung.gewöhnlich[zufall.Next(0, Umgebung.gewöhnlich.Length - 1)]; }

                //Erhöhende Ergebnisse!
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(ergebniszutat))
                    {
                        dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value = (int)dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value + zufall.Next(1, 7);
                        portionenerhöht = 1;
                        break;
                    }
                }

                //Ohne Erhöhung -> Neues Tabellenfeld
                if (portionenerhöht == 0)
                {
                    dataGridView1.Rows.Add();
                    rowcount++;
                    dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightYellow;
                    dataGridView1[0, rowcount - 1].Value = ergebniszutat;
                    dataGridView1[2, rowcount - 1].Value = "Gewöhnlich";
                    dataGridView1[1, rowcount - 1].Value = zufall.Next(1, 7);
                    if (zufall.Next(1, 10) > kenntniszw + 1)
                    {
                        dataGridView1[3, rowcount - 1].Value = "Unbekannt!";
                        dataGridView1[3, rowcount - 1].Style.BackColor = Color.Orange;
                    }
                }
            }
        } 

        public void DatenbankLaden(object sender, EventArgs e)
        {
            string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string path = (System.IO.Path.GetDirectoryName(executable));
            AppDomain.CurrentDomain.SetData("DataDirectory", path);
            //Provider=Microsoft.ACE.OLEDB.12.0;
            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\TTruheAccess.accdb");
            //OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Tüfteltruhe\Tüfteltruhe\TTruheAccess.accdb");
            connection.Open();
            OleDbDataReader reader = null;
            OleDbDataReader reader2 = null;
            OleDbDataReader reader3 = null;
            OleDbDataReader reader4 = null;
            OleDbDataReader reader5 = null;
            OleDbDataReader reader6 = null;
            OleDbDataReader reader7 = null;
            OleDbDataReader reader8 = null;
            OleDbCommand command = new OleDbCommand("SELECT * FROM ZutatenRegion", connection);
            OleDbCommand command2 = new OleDbCommand("SELECT * FROM ZutatenRegion", connection);
            OleDbCommand command3 = new OleDbCommand("SELECT * FROM ZutatenUmgebung", connection);
            OleDbCommand command4 = new OleDbCommand("SELECT * FROM ZutatenUmgebung", connection);
            OleDbCommand command5 = new OleDbCommand("SELECT * FROM Zutaten", connection);
            OleDbCommand command6 = new OleDbCommand("SELECT * FROM Zutaten", connection);
            OleDbCommand command7 = new OleDbCommand("SELECT * FROM Haendler", connection);
            OleDbCommand command8 = new OleDbCommand("SELECT * FROM Sondergewerbe", connection);
            reader = command.ExecuteReader();
            reader2 = command2.ExecuteReader();
            reader3 = command3.ExecuteReader();
            reader4 = command4.ExecuteReader();
            reader5 = command5.ExecuteReader();
            reader6 = command6.ExecuteReader();
            reader7 = command7.ExecuteReader();
            reader8 = command8.ExecuteReader();
            comboBox5.Items.Clear();
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            while (reader.Read())
            {
                comboBox1.Items.Add(reader[1].ToString());
                comboBox3.Items.Add(reader[1].ToString());
            }
            while (reader3.Read())
            {
                comboBox5.Items.Add(reader3[1].ToString());
                comboBox4.Items.Add(reader3[1].ToString());
            }
            while (reader5.Read())
            {
                comboBox2.Items.Add(reader5[1].ToString());
            }

            DataTable ZutatenRegionTabelle = new DataTable();
            DataTable ZutatenUmgebungTabelle = new DataTable();
            DataTable ZutatenTabelle = new DataTable();
            DataTable HaendlerTabelle = new DataTable();
            DataTable SondergewerbeTabelle = new DataTable();

            ZutatenRegionTabelle.Load(reader2);
            ZutatenUmgebungTabelle.Load(reader4);
            ZutatenTabelle.Load(reader);
            HaendlerTabelle.Load(reader7);
            SondergewerbeTabelle.Load(reader8);

            ZutatenRegionTB = ZutatenRegionTabelle;
            ZutatenUmgebungTB = ZutatenUmgebungTabelle;
            ZutatenTB = ZutatenTabelle;
            HaendlerTB = HaendlerTabelle;
            SondergewerbeTB = SondergewerbeTabelle;

            connection.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            gewaehlte_region = comboBox1.GetItemText(comboBox1.SelectedItem);

            foreach (DataRow row in ZutatenRegionTB.Rows)
            {
                if (row["Region"].ToString() == gewaehlte_region)
                {
                    string[] tmp = row["Gewöhnlich"].ToString().Split(',') ;
                    if (tmp != null) Region.gewöhnlich = tmp;
                    else Region.gewöhnlich = null;

                    tmp = row["Ungewöhnlich"].ToString().Split(',');
                    if (tmp != null) Region.ungewöhnlich = tmp;
                    else Region.ungewöhnlich = null;

                    tmp = row["Selten"].ToString().Split(',');
                    if (tmp != null) Region.selten = tmp;
                    else Region.selten = null;

                    tmp = row["Sehr selten"].ToString().Split(',');
                    if (tmp != null) Region.sehrselten = tmp;
                    else Region.sehrselten = null;
                }
            }
            MessageBox.Show("Zutaten der Region:" + string.Join(", ", Region.gewöhnlich) + ", " + string.Join(", ", Region.ungewöhnlich) + ", " + string.Join(", ", Region.selten) + ", " + string.Join(", ", Region.sehrselten), gewaehlte_region);
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            gewaehlte_umgebung = comboBox5.GetItemText(comboBox5.SelectedItem);

            foreach (DataRow row in ZutatenUmgebungTB.Rows)
            {
                if (row["Umgebung"].ToString() == gewaehlte_umgebung)
                {
                    string[] tmp = row["Gewöhnlich"].ToString().Split(',');
                    if (tmp != null) Umgebung.gewöhnlich = tmp;
                    else Umgebung.gewöhnlich = null;

                    tmp = row["Ungewöhnlich"].ToString().Split(',');
                    if (tmp != null) Umgebung.ungewöhnlich = tmp;
                    else Umgebung.ungewöhnlich = null;

                    tmp = row["Selten"].ToString().Split(',');
                    if (tmp != null) Umgebung.selten = tmp;
                    else Umgebung.selten = null;

                    tmp = row["Sehr selten"].ToString().Split(',');
                    if (tmp != null) Umgebung.sehrselten = tmp;
                    else Umgebung.sehrselten = null;
                }
            }
            MessageBox.Show("Zutaten der Umgebung:" + string.Join(", ", Umgebung.gewöhnlich) + ", " + string.Join(", ", Umgebung.ungewöhnlich) + ", " + string.Join(", ", Umgebung.selten) + ", " + string.Join(", ", Umgebung.sehrselten), gewaehlte_umgebung);
        }

        private void button3_Click(object sender, EventArgs e) //Ergebnisse zurücksetzen
        {
            dataGridView1.Rows.Clear();
            rowcount = 0;
            gesamtstunden = 0;
            label8.Text = "Ges. Stunden: " + gesamtstunden;
        }

        private void button1_Click(object sender, EventArgs e) // 1 Stunde suchen
        {
            erforderlichesuchstunden--;
            heutigesuchstunden++;
            label14.Text = "Ges. Stunden: " + heutigesuchstunden;

            if (erforderlichesuchstunden < 1)
            {
                dataGridView3.Rows.Add();
                rowcount2++;
                dataGridView3.Rows[rowcount2 - 1].DefaultCellStyle.BackColor = Color.YellowGreen;
                int portionen = 1;
                if (istselten) portionen = zufall.Next(1, 3);
                if (istungewöhnlich) portionen = zufall.Next(1, 5);
                if (istgewöhnlich) portionen = zufall.Next(1, 7);
                dataGridView3[0, rowcount2 - 1].Value = portionen + " Portionen der gesuchten Zutat " + comboBox2.GetItemText(comboBox2.SelectedItem) + " gefunden!";
                neuesuche();
            }
            else if (erforderlichesuchstunden % nebenfundteiler == 0)
            {
                dataGridView3.Rows.Add();
                rowcount2++;
                dataGridView3.Rows[rowcount2 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
                dataGridView3[0, rowcount2 - 1].Value = "Nebenbei " + zufall.Next(1,7) + " Portionen der Zutat" + Region2.gewöhnlich[zufall.Next(0, Region2.gewöhnlich.Length - 1)] + " gefunden.";
            }
            else
            {
                dataGridView3.Rows.Add();
                rowcount2++;
                dataGridView3.Rows[rowcount2 - 1].DefaultCellStyle.BackColor = Color.LightCoral;
                dataGridView3[0, rowcount2 - 1].Value = "Nichts gefunden.";
            }

            if (heutigesuchstunden > 3)
            { 
                dataGridView3[1, rowcount2 - 1].Value = " (-1 Fokus)";
                dataGridView3[1, rowcount2 - 1].Style.BackColor = Color.Orange;
            }
            else
            {
                dataGridView3[1, rowcount2 - 1].Style.BackColor = Color.LightYellow;
            }

            if (!richtigeregion)
            {
                dataGridView3[2, rowcount2 - 1].Value = "Falsche Region!";
                dataGridView3[2, rowcount2 - 1].Style.BackColor = Color.LightCoral;
            }
            else
            {
                dataGridView3[2, rowcount2 - 1].Value = "Geeignete Region.";
                dataGridView3[2, rowcount2 - 1].Style.BackColor = Color.LightYellow;
            }

            if (!richtigeumgebung)
            {
                dataGridView3[3, rowcount2 - 1].Value = "Falsche Umgebung!";
                dataGridView3[3, rowcount2 - 1].Style.BackColor = Color.Orange;
            }
            else
            {
                dataGridView3[3, rowcount2 - 1].Value = "Geeignete Umgebung.";
                dataGridView3[3, rowcount2 - 1].Style.BackColor = Color.LightYellow;
            }
        }

        private void button4_Click(object sender, EventArgs e) //Neuer Tag
        {
            heutigesuchstunden = 0;
            label14.Text = "Ges. Stunden: " + heutigesuchstunden;
        }

        public void neuesuche()
        {
            erforderlichesuchstunden = 10000;

            gewaehlte_region = comboBox3.GetItemText(comboBox3.SelectedItem);
            foreach (DataRow row in ZutatenRegionTB.Rows)
            {
                if (row["Region"].ToString() == gewaehlte_region)
                {
                    string[] tmp = row["Gewöhnlich"].ToString().Split(',');
                    if (tmp != null) Region2.gewöhnlich = tmp;    /*Region2.gewöhnlich.Concat(tmp).ToArray();*/

                    tmp = row["Ungewöhnlich"].ToString().Split(',');
                    if (tmp != null) Region2.ungewöhnlich = tmp;

                    tmp = row["Selten"].ToString().Split(',');
                    if (tmp != null) Region2.selten = tmp;

                    tmp = row["Sehr selten"].ToString().Split(',');
                    if (tmp != null) Region2.sehrselten = tmp;
                }
            }

            gewaehlte_umgebung = comboBox4.GetItemText(comboBox4.SelectedItem);
            foreach (DataRow row in ZutatenUmgebungTB.Rows)
            {
                if (row["Umgebung"].ToString() == gewaehlte_umgebung)
                {
                    string[] tmp = row["Gewöhnlich"].ToString().Split(',');
                    if (tmp != null) Umgebung2.gewöhnlich = tmp;

                    tmp = row["Ungewöhnlich"].ToString().Split(',');
                    if (tmp != null) Umgebung2.ungewöhnlich = tmp;

                    tmp = row["Selten"].ToString().Split(',');
                    if (tmp != null) Umgebung2.selten = tmp;

                    tmp = row["Sehr selten"].ToString().Split(',');
                    if (tmp != null) Umgebung2.sehrselten = tmp;
                }
            }

            istgewöhnlich = false;
            istungewöhnlich = false;
            istselten = false;
            istsehrselten = false;

            richtigeregion = false;
            richtigeumgebung = false;
            label16.Text = "";

            if (Region2.gewöhnlich != null)
            {
                for (int i = 0; i < Region2.gewöhnlich.Length; i++)
                {
                    if (Region2.gewöhnlich[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeregion = true;
                        istgewöhnlich = true;
                        label16.Text = "Seltenheit: Gewöhnlich";
                        label16.ForeColor = Color.ForestGreen;
                    }
                }
            }
            if (Region2.ungewöhnlich != null)
            {
                for (int i = 0; i < Region2.ungewöhnlich.Length; i++)
                {
                    if (Region2.ungewöhnlich[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeregion = true;
                        istungewöhnlich = true;
                        label16.Text = "Seltenheit: Ungewöhnlich";
                        label16.ForeColor = Color.DarkBlue;
                    }
                }
            }
            if (Region2.selten != null)
            {
                for (int i = 0; i < Region2.selten.Length; i++)
                {
                    if (Region2.selten[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeregion = true;
                        istselten = true;
                        label16.Text = "Seltenheit: Selten";
                        label16.ForeColor = Color.Magenta;
                    }
                }
            }
            if (Region2.sehrselten != null)
            {
                for (int i = 0; i < Region2.sehrselten.Length; i++)
                {
                    if (Region2.sehrselten[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeregion = true;
                        istsehrselten = true;
                        label16.Text = "Seltenheit: Sehr selten!";
                        label16.ForeColor = Color.Red;
                    }
                }
            }
            if (Umgebung2.gewöhnlich != null)
            {
                for (int i = 0; i < Umgebung2.gewöhnlich.Length; i++)
                {
                    if (Umgebung2.gewöhnlich[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeumgebung = true;
                        istgewöhnlich = true;
                        label16.Text = "Seltenheit: Gewöhnlich";
                        label16.ForeColor = Color.ForestGreen;
                    }
                }
            }
            if (Umgebung2.ungewöhnlich != null)
            {
                for (int i = 0; i < Umgebung2.ungewöhnlich.Length; i++)
                {
                    if (Umgebung2.ungewöhnlich[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeumgebung = true;
                        istungewöhnlich = true;
                        label16.Text = "Seltenheit: Ungewöhnlich";
                        label16.ForeColor = Color.DarkBlue;
                    }
                }
            }
            if (Umgebung2.selten != null)
            {
                for (int i = 0; i < Umgebung2.selten.Length; i++)
                {
                    if (Umgebung2.selten[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeumgebung = true;
                        istselten = true;
                        label16.Text = "Seltenheit: Selten";
                        label16.ForeColor = Color.Magenta;
                    }
                }
            }
            if (Umgebung2.sehrselten != null)
            {
                    for (int i = 0; i < Umgebung2.sehrselten.Length; i++)
                {
                    if (Umgebung2.sehrselten[i] == " " + comboBox2.GetItemText(comboBox2.SelectedItem))
                    {
                        richtigeumgebung = true;
                        istsehrselten = true;
                        label16.Text = "Seltenheit: Sehr selten!";
                        label16.ForeColor = Color.Red;
                    }
                }
            }


            if (richtigeumgebung && richtigeregion)
            {
                if (istgewöhnlich) erforderlichesuchstunden = 1;
                if (istungewöhnlich) erforderlichesuchstunden = 6;
                if (istselten) erforderlichesuchstunden = 24;
                if (istsehrselten) erforderlichesuchstunden = 80;
            }
            else if (richtigeregion)
            {
                if (istgewöhnlich) erforderlichesuchstunden = 4;
                if (istungewöhnlich) erforderlichesuchstunden = 24;
                if (istselten) erforderlichesuchstunden = 64;
                if (istsehrselten) erforderlichesuchstunden = 10000;
            }
            else 
            {
                if (istgewöhnlich) erforderlichesuchstunden = 32;
                if (istungewöhnlich) erforderlichesuchstunden = 10000;
                if (istselten) erforderlichesuchstunden = 10000;
                if (istsehrselten) erforderlichesuchstunden = 10000;
            }

            testergebnis2 = (int)numericUpDown5.Value + zufall.Next(1, 6) + zufall.Next(1, 6) + zufall.Next(1, 6);
            label15.Text = "Letzter Wurf: " + (testergebnis2 - (int)numericUpDown5.Value) + " + " + (int)numericUpDown5.Value + " = " + testergebnis2;
            if (testergebnis2 < 11)
            {
                erforderlichesuchstunden *= 2;
                nebenfundteiler = 6;
            }
            else if (testergebnis2 < 21)
            {
                nebenfundteiler = 4;
            }
            else
            {
                erforderlichesuchstunden /= 2;
                nebenfundteiler = 2;
            }
        }

        private void button5_Click(object sender, EventArgs e) //Ergebnisse zurücksetzen
        {
            heutigesuchstunden = 0;
            label14.Text = "Ges. Stunden: " + heutigesuchstunden;

            comboBox2.SelectedIndex = -1;
            label15.Text = "";
            testergebnis2 = 0;
            nebenfundteiler = 6;
            erforderlichesuchstunden = 10000;
            dataGridView3.Rows.Clear();
            rowcount2 = 0;
        }

        public void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            neuesuche();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.GetItemText(comboBox2.SelectedItem) != null) { neuesuche(); }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.GetItemText(comboBox2.SelectedItem) != null) { neuesuche(); }
        }
        private void neuesFensterSpielermodusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spielermodus spielermodus = new Spielermodus();
            spielermodus.Show();
        }

        private void neuesFensterSpielleitermodusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spielleitermodus spielleitermodus = new Spielleitermodus();
            spielleitermodus.Show();
        }


        //#################### ################################################
        //### MARKTANGEBOT ### ################################################
        //#################### ################################################


        private void button7_Click(object sender, EventArgs e) //Alles zurücksetzen
        {
            dataGridView2.Rows.Clear();
            rowcount3 = 0;
            händleranzahl = 0;
        }

        private void button8_Click(object sender, EventArgs e) //Markt generieren
        {
            if (radioButton1.Checked) händleranzahl = 1 + zufall.Next(1,4);
            else if (radioButton2.Checked) händleranzahl = 5 + zufall.Next(1, 10);
            else if (radioButton3.Checked) händleranzahl = 15 + zufall.Next(1, 20);
            else if (radioButton4.Checked) händleranzahl = 35 + zufall.Next(1, 20);

            if (radioButton1.Checked) sondergewerbe = zufall.Next(1, 4) - 1;
            else if (radioButton2.Checked) sondergewerbe = zufall.Next(1, 6);
            else if (radioButton3.Checked) sondergewerbe = zufall.Next(1, 8) + zufall.Next(1, 8);
            else if (radioButton4.Checked) sondergewerbe = zufall.Next(1, 12) + zufall.Next(1, 12);


            for (int i = 0; i < händleranzahl; i++)
            {
                int ergebnis = zufall.Next(1, 100);
                foreach (DataRow row in HaendlerTB.Rows)
                {
                    if (row["Ergebnis"].ToString() == ergebnis.ToString())
                    {
                        string tmp = row["Haendler"].ToString();
                        if (tmp != null) haendlerergebnis = tmp;
                        else haendlerergebnis = "";

                        tmp = row["Voraussetzung"].ToString();
                        if (tmp != null) voraussetzung = tmp;
                        else voraussetzung = "";

                        tmp = row["Beschreibung"].ToString();
                        if (tmp != null) haendlerbeschreibung = tmp;
                        else haendlerbeschreibung = "";
                    }
                }

                eintragen = true;
                if (voraussetzung != "" && checkBox1.Checked == false)
                {
                    DialogResult dialogResult = MessageBox.Show(voraussetzung, "Voraussetzung!", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes) { }
                    else if (dialogResult == DialogResult.No)
                    {
                        händleranzahl++; //Damit dieser Durchgang nicht zählt.
                        eintragen = false;
                    }
                }
                if (eintragen)
                {
                    dataGridView2.Rows.Add();
                    rowcount3++;
                    dataGridView2.Rows[rowcount3 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
                    dataGridView2[0, rowcount3 - 1].Value = haendlerergebnis;
                    dataGridView2[1, rowcount3 - 1].Value = haendlerbeschreibung;
                }
            }

            for (int i = 0; i < sondergewerbe; i++)
            {
                int ergebnis2 = zufall.Next(1, 100);
                foreach (DataRow row in SondergewerbeTB.Rows)
                {
                    if (row["Ergebnis"].ToString() == ergebnis2.ToString())
                    {
                        string tmp = row["Haendler"].ToString();
                        if (tmp != null) sonderergebnis = tmp;
                        else sonderergebnis = "";

                        tmp = row["Voraussetzung"].ToString();
                        if (tmp != null) sondervoraussetzung = tmp;
                        else sondervoraussetzung = "";

                        tmp = row["Beschreibung"].ToString();
                        if (tmp != null) sonderbeschreibung = tmp;
                        else sonderbeschreibung = "";

                        tmp = row["Groesse"].ToString();
                        if (tmp != null) groesse = tmp;
                        else groesse = "";
                    }
                }

                eintragen = true;
                if (groesse == "MITTEL" && radioButton1.Checked)
                {
                    eintragen = false;
                    sondergewerbe++; //Damit dieser Durchgang nicht zählt.
                }
                if (groesse == "GROSS" && radioButton1.Checked || groesse == "GROSS" && radioButton2.Checked)
                {
                    eintragen = false;
                    sondergewerbe++; //Damit dieser Durchgang nicht zählt.
                }
                if (eintragen && sondervoraussetzung != "" && checkBox1.Checked == false)
                {
                    DialogResult dialogResult = MessageBox.Show(sondervoraussetzung, "Voraussetzung!", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes) { }
                    else if (dialogResult == DialogResult.No)
                    {
                        sondergewerbe++; //Damit dieser Durchgang nicht zählt.
                        eintragen = false;
                    }
                }

                if (eintragen)
                {
                    dataGridView2.Rows.Add();
                    rowcount3++;
                    dataGridView2.Rows[rowcount3 - 1].DefaultCellStyle.BackColor = Color.YellowGreen;
                    dataGridView2[0, rowcount3 - 1].Value = sonderergebnis;
                    dataGridView2[1, rowcount3 - 1].Value = sonderbeschreibung;
                }
            }
        }


        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Book Antiqua";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 12;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                //oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "Tüfteltruhe-Export";
                    headerRange.Font.Size = 14;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                oDoc.SaveAs2(filename);
            }
        }


        private void AllgZutatensucheExp(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "Tüfteltruhe-Export Allgemeine Zutatensuche.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }

        private void SpzZutatensucheExp(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "Tüfteltruhe-Export Spezielle Zutatensuche.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView3, sfd.FileName);
            }
        }

        private void MarktzusammensetzungExp(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "Tüfteltruhe-Export Marktzusammensetzung.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView2, sfd.FileName);
            }
        }

        private void WarenangebotExp(object sender, EventArgs e)
        {
        //    SaveFileDialog sfd = new SaveFileDialog();

        //    sfd.Filter = "Word Documents (*.docx)|*.docx";

        //    sfd.FileName = "Tüfteltruhe-Export Warenangebot.docx";

        //    if (sfd.ShowDialog() == DialogResult.OK)
        //    {
        //        Export_Data_To_Word(dataGridView1, sfd.FileName);
        //    }
        }
    }
}
