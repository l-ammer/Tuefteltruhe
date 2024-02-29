using System;
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
        public DataTable WarenTB = new DataTable();
        public DataTable HaendlerWarenTB = new DataTable();
        public DataTable KulturkontexteTB = new DataTable();
        public DataTable SchmuckTB = new DataTable();
        public DataTable MetallTB = new DataTable();
        public DataTable ZierTB = new DataTable();
        public DataTable ZauberTB = new DataTable();
        public DataTable ZaubersteinTB = new DataTable();
        public DataTable ZauberkomplexTB = new DataTable();
        public DataTable NamenTB = new DataTable();
        public DataTable SchatzTB = new DataTable();
        public DataTable SchatzGegenstandTB = new DataTable();
        public DataTable SchatzAlltagMusikTB = new DataTable();
        public DataTable SchatzWaffeTB = new DataTable();
        public DataTable SchatzRuestungTB = new DataTable();
        public DataTable WaffentypenTB = new DataTable();
        public DataTable RuestungstypenTB = new DataTable();
        public DataTable SpezialgegenstandTB = new DataTable();
        public DataTable AlltagsgegenstandTB = new DataTable();
        public DataTable KomplexringTB = new DataTable();
        public DataTable BannwaffeTB = new DataTable();
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
        public int rowcount4 = 0;
        public int rowcount5 = 0;
        public int rowcount7 = 0;
        public int rowcount8 = 0;
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
        public Color hintergrundfarbe = Color.White;
        bool keinlebewesen = false;
        bool keinezauberei = false;

        public Spielleitermodus()
        {
            InitializeComponent();
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
            OleDbDataReader reader9 = null;
            OleDbDataReader reader10 = null;
            OleDbDataReader reader11 = null;
            OleDbDataReader reader12 = null;
            OleDbDataReader reader13 = null;
            OleDbDataReader reader14 = null;
            OleDbDataReader reader15 = null;
            OleDbDataReader reader16 = null;
            OleDbDataReader reader17 = null;
            OleDbDataReader reader18 = null;
            OleDbDataReader reader19 = null;
            OleDbDataReader reader20 = null;
            OleDbDataReader reader21 = null;
            OleDbDataReader reader22 = null;
            OleDbDataReader reader23 = null;
            OleDbDataReader reader24 = null;
            OleDbDataReader reader25 = null;
            OleDbDataReader reader26 = null;
            OleDbDataReader reader27 = null;
            OleDbDataReader reader28 = null;
            OleDbDataReader reader29 = null;
            OleDbDataReader reader30 = null;
            OleDbCommand command = new OleDbCommand("SELECT * FROM ZutatenRegion", connection);
            OleDbCommand command2 = new OleDbCommand("SELECT * FROM ZutatenRegion", connection);
            OleDbCommand command3 = new OleDbCommand("SELECT * FROM ZutatenUmgebung", connection);
            OleDbCommand command4 = new OleDbCommand("SELECT * FROM ZutatenUmgebung", connection);
            OleDbCommand command5 = new OleDbCommand("SELECT * FROM Zutaten", connection);
            OleDbCommand command6 = new OleDbCommand("SELECT * FROM Zutaten", connection);
            OleDbCommand command7 = new OleDbCommand("SELECT * FROM Haendler", connection);
            OleDbCommand command8 = new OleDbCommand("SELECT * FROM Sondergewerbe", connection);
            OleDbCommand command9 = new OleDbCommand("SELECT * FROM Waren", connection);
            OleDbCommand command10 = new OleDbCommand("SELECT * FROM HaendlerWaren", connection);
            OleDbCommand command11 = new OleDbCommand("SELECT * FROM HaendlerWaren", connection);
            OleDbCommand command12 = new OleDbCommand("SELECT * FROM Kulturkontexte", connection);
            OleDbCommand command13 = new OleDbCommand("SELECT * FROM Kulturkontexte", connection);
            OleDbCommand command14 = new OleDbCommand("SELECT * FROM Schmuck", connection);
            OleDbCommand command15 = new OleDbCommand("SELECT * FROM Metall", connection);
            OleDbCommand command16 = new OleDbCommand("SELECT * FROM Zier", connection);
            OleDbCommand command17 = new OleDbCommand("SELECT * FROM Zauber", connection);
            OleDbCommand command18 = new OleDbCommand("SELECT * FROM Zauberstein", connection);
            OleDbCommand command19 = new OleDbCommand("SELECT * FROM Zauberkomplex", connection);
            OleDbCommand command20 = new OleDbCommand("SELECT * FROM Namen", connection);
            OleDbCommand command21 = new OleDbCommand("SELECT * FROM Namen", connection);
            OleDbCommand command22 = new OleDbCommand("SELECT * FROM Schatz", connection);
            OleDbCommand command23 = new OleDbCommand("SELECT * FROM SchatzGegenstand", connection);
            OleDbCommand command24 = new OleDbCommand("SELECT * FROM SchatzAlltagMusik", connection);
            OleDbCommand command25 = new OleDbCommand("SELECT * FROM Waffentypen", connection);
            OleDbCommand command26 = new OleDbCommand("SELECT * FROM Ruestungstypen", connection);
            OleDbCommand command27 = new OleDbCommand("SELECT * FROM Spezialgegenstand", connection);
            OleDbCommand command28 = new OleDbCommand("SELECT * FROM Alltagsgegenstand", connection);
            OleDbCommand command29 = new OleDbCommand("SELECT * FROM Komplexring", connection);
            OleDbCommand command30 = new OleDbCommand("SELECT * FROM Bannwaffe", connection);
            reader = command.ExecuteReader();
            reader2 = command2.ExecuteReader();
            reader3 = command3.ExecuteReader();
            reader4 = command4.ExecuteReader();
            reader5 = command5.ExecuteReader();
            reader6 = command6.ExecuteReader();
            reader7 = command7.ExecuteReader();
            reader8 = command8.ExecuteReader();
            reader9 = command9.ExecuteReader();
            reader10 = command10.ExecuteReader();
            reader11 = command11.ExecuteReader();
            reader12 = command12.ExecuteReader();
            reader13 = command13.ExecuteReader();
            reader14 = command14.ExecuteReader();
            reader15 = command15.ExecuteReader();
            reader16 = command16.ExecuteReader();
            reader17 = command17.ExecuteReader();
            reader18 = command18.ExecuteReader();
            reader19 = command19.ExecuteReader();
            reader20 = command20.ExecuteReader();
            reader21 = command21.ExecuteReader();
            reader22 = command22.ExecuteReader();
            reader23 = command23.ExecuteReader();
            reader24 = command24.ExecuteReader();
            reader25 = command25.ExecuteReader();
            reader26 = command26.ExecuteReader();
            reader27 = command27.ExecuteReader();
            reader28 = command28.ExecuteReader();
            reader29 = command29.ExecuteReader();
            reader30 = command30.ExecuteReader();
            comboBox5.Items.Clear();
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox6.Items.Clear();
            comboBox7.Items.Clear();
            comboBox9.Items.Clear();
            comboBox12.Items.Clear();

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
            while (reader10.Read())
            {
                comboBox6.Items.Add(reader10[1].ToString());
            }
            while (reader12.Read())
            {
                comboBox7.Items.Add(reader12[1].ToString());
            }
            while (reader19.Read())
            {
                comboBox9.Items.Add(reader19[1].ToString());
            }
            while (reader20.Read())
            {
                comboBox12.Items.Add(reader20[1].ToString());
            }

            DataTable ZutatenRegionTabelle = new DataTable();
            DataTable ZutatenUmgebungTabelle = new DataTable();
            DataTable ZutatenTabelle = new DataTable();
            DataTable HaendlerTabelle = new DataTable();
            DataTable SondergewerbeTabelle = new DataTable();
            DataTable WarenTabelle = new DataTable();
            DataTable HaendlerWarenTabelle = new DataTable();
            DataTable KulturkontexteTabelle = new DataTable();
            DataTable SchmuckTabelle = new DataTable();
            DataTable MetallTabelle = new DataTable();
            DataTable ZierTabelle = new DataTable();
            DataTable ZauberTabelle = new DataTable();
            DataTable ZaubersteinTabelle = new DataTable();
            DataTable NamenTabelle = new DataTable();
            DataTable SchatzTabelle = new DataTable();
            DataTable SchatzGegenstandTabelle = new DataTable();
            DataTable SchatzAlltagMusikTabelle = new DataTable();
            DataTable WaffentypenTabelle = new DataTable();
            DataTable RuestungstypenTabelle = new DataTable();
            DataTable SpezialgegenstandTabelle = new DataTable();
            DataTable AlltagsgegenstandTabelle = new DataTable();
            DataTable KomplexringTabelle = new DataTable();
            DataTable BannwaffeTabelle = new DataTable();

            ZutatenRegionTabelle.Load(reader2);
            ZutatenUmgebungTabelle.Load(reader4);
            ZutatenTabelle.Load(reader6);
            HaendlerTabelle.Load(reader7);
            SondergewerbeTabelle.Load(reader8);
            WarenTabelle.Load(reader9);
            HaendlerWarenTabelle.Load(reader11);
            KulturkontexteTabelle.Load(reader13);
            SchmuckTabelle.Load(reader14);
            MetallTabelle.Load(reader15);
            ZierTabelle.Load(reader16);
            ZauberTabelle.Load(reader17);
            ZaubersteinTabelle.Load(reader18);
            NamenTabelle.Load(reader21);
            SchatzTabelle.Load(reader22);
            SchatzGegenstandTabelle.Load(reader23);
            SchatzAlltagMusikTabelle.Load(reader24);
            WaffentypenTabelle.Load(reader25);
            RuestungstypenTabelle.Load(reader26);
            SpezialgegenstandTabelle.Load(reader27);
            AlltagsgegenstandTabelle.Load(reader28);
            KomplexringTabelle.Load(reader29);
            BannwaffeTabelle.Load(reader30);

            ZutatenRegionTB = ZutatenRegionTabelle;
            ZutatenUmgebungTB = ZutatenUmgebungTabelle;
            ZutatenTB = ZutatenTabelle;
            HaendlerTB = HaendlerTabelle;
            SondergewerbeTB = SondergewerbeTabelle;
            WarenTB = WarenTabelle;
            HaendlerWarenTB = HaendlerWarenTabelle;
            KulturkontexteTB = KulturkontexteTabelle;
            SchmuckTB = SchmuckTabelle;
            MetallTB = MetallTabelle;
            ZierTB = ZierTabelle;
            ZauberTB = ZauberTabelle;
            ZaubersteinTB = ZaubersteinTabelle;
            NamenTB = NamenTabelle;
            SchatzTB = SchatzTabelle;
            SchatzGegenstandTB = SchatzGegenstandTabelle;
            SchatzAlltagMusikTB = SchatzAlltagMusikTabelle;
            WaffentypenTB = WaffentypenTabelle;
            RuestungstypenTB = RuestungstypenTabelle;
            SpezialgegenstandTB = SpezialgegenstandTabelle;
            AlltagsgegenstandTB = AlltagsgegenstandTabelle;
            KomplexringTB = KomplexringTabelle;
            BannwaffeTB = BannwaffeTabelle;

            //comboBox2.SelectedIndex = 0;
            //comboBox3.SelectedIndex = 0;
            //comboBox4.SelectedIndex = 0;
            //comboBox5.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;
            comboBox8.SelectedIndex = 0;
            comboBox9.SelectedIndex = 0;
            comboBox12.SelectedIndex = 0;

            connection.Close();
        }

        //#################### ################################################
        //##### ZUTATEN ###### ################################################
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
                label8.Text = "Bereits gesuchte Stunden: " + gesamtstunden;
                testergebnis = zufall.Next(1, 7) + zufall.Next(1, 7) + zufall.Next(1, 7) + sammeln;
                if (numericUpDown10.Value != 0) testergebnis = (int)numericUpDown10.Value;
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
                //    regionergebnis = Region.selten[zufall.Next(0, Region.selten.Length)];
                //    umgebungergebnis = Umgebung.selten[zufall.Next(0, Umgebung.selten.Length)];
                //    wiederholungen++;
                //}
                //ergebniszutat = regionergebnis;

                //MODUS: SUMME
                ergebniszutat = Region.selten[zufall.Next(0, Region.selten.Length )];
                if (comboBox5.GetItemText(comboBox5.SelectedItem) != "" && zufall.Next(0, 2) == 1) { ergebniszutat = Umgebung.selten[zufall.Next(0, Umgebung.selten.Length)]; }

                //Erhöhende Ergebnisse!
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(ergebniszutat))
                    {
                        dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value = (int)dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value + zufall.Next(1, 4);
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
                    dataGridView1[3, rowcount - 1].Value = "Selten";
                    dataGridView1[1, rowcount - 1].Value = zufall.Next(1, 4);
                    if (zufall.Next(1, 11) > kenntniszw)
                    {
                        dataGridView1[4, rowcount - 1].Value = "Unbekannt!";
                        dataGridView1[4, rowcount - 1].Style.BackColor = Color.Orange;
                    }
                    foreach (DataRow row in ZutatenTB.Rows)
                    {
                        if (" " + row["Zutat"].ToString() == ergebniszutat)
                        {
                            dataGridView1[2, rowcount - 1].Value = row["Portionsgewicht"].ToString() + " Pfund";
                            dataGridView1[5, rowcount - 1].Value = row["Aussehen"].ToString(); 
                            dataGridView1[6, rowcount - 1].Value = row["Fundort"].ToString();
                        }
                    }
                }
            }

            for (int i = 0; i < zwungewoehnlich; i++)
            {
                portionenerhöht = 0;

                //MODUS: SUMME
                ergebniszutat = Region.ungewöhnlich[zufall.Next(0, Region.ungewöhnlich.Length)];
                if (comboBox5.GetItemText(comboBox5.SelectedItem) != "" && zufall.Next(0, 2) == 1) { ergebniszutat = Umgebung.ungewöhnlich[zufall.Next(0, Umgebung.ungewöhnlich.Length)]; }

                //Erhöhende Ergebnisse!
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(ergebniszutat))
                    {
                        dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value = (int)dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value + zufall.Next(1, 6);
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
                    dataGridView1[3, rowcount - 1].Value = "Ungewöhnlich";
                    dataGridView1[1, rowcount - 1].Value = zufall.Next(1, 6);
                    if (zufall.Next(1, 11) > kenntniszw + 1)
                    {
                        dataGridView1[4, rowcount - 1].Value = "Unbekannt!";
                        dataGridView1[4, rowcount - 1].Style.BackColor = Color.Orange;
                    }
                    foreach (DataRow row in ZutatenTB.Rows)
                    {
                        if (" " + row["Zutat"].ToString() == ergebniszutat)
                        {
                            dataGridView1[2, rowcount - 1].Value = row["Portionsgewicht"].ToString() + " Pfund";
                            dataGridView1[5, rowcount - 1].Value = row["Aussehen"].ToString();
                            dataGridView1[6, rowcount - 1].Value = row["Fundort"].ToString();
                        }
                    }
                }
            }

            for (int i = 0; i < zwgewoehnlich; i++)
            {
                portionenerhöht = 0;

                //MODUS: SUMME
                ergebniszutat = Region.gewöhnlich[zufall.Next(0, Region.gewöhnlich.Length)];
                if (comboBox5.GetItemText(comboBox5.SelectedItem) != "" && zufall.Next(0, 2) == 1) { ergebniszutat = Umgebung.gewöhnlich[zufall.Next(0, Umgebung.gewöhnlich.Length)]; }

                //Erhöhende Ergebnisse!
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(ergebniszutat))
                    {
                        dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value = (int)dataGridView1[1, dataGridView1.Rows.IndexOf(row)].Value + zufall.Next(1, 8);
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
                    dataGridView1[3, rowcount - 1].Value = "Gewöhnlich";
                    dataGridView1[1, rowcount - 1].Value = zufall.Next(1, 8);
                    if (zufall.Next(1, 11) > kenntniszw + 4)
                    {
                        dataGridView1[4, rowcount - 1].Value = "Unbekannt!";
                        dataGridView1[4, rowcount - 1].Style.BackColor = Color.Orange;
                    }
                    foreach (DataRow reihe in ZutatenTB.Rows)
                    {
                        if (" " + reihe["Zutat"].ToString() == ergebniszutat)
                        {
                            dataGridView1[2, rowcount - 1].Value = reihe["Portionsgewicht"].ToString() +  " Pfund";
                            dataGridView1[5, rowcount - 1].Value = reihe["Aussehen"].ToString();
                            dataGridView1[6, rowcount - 1].Value = reihe["Fundort"].ToString();
                        }
                    }
                }
            }
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

            if (gewaehlte_umgebung != null && Umgebung != null)
            {

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
        }

        private void button3_Click(object sender, EventArgs e) //Ergebnisse zurücksetzen
        {
            dataGridView1.Rows.Clear();
            rowcount = 0;
            gesamtstunden = 0;
            label8.Text = "Bereits gesuchte Stunden: " + gesamtstunden;
        }

        private void button1_Click(object sender, EventArgs e) // 1 Stunde suchen
        {
            erforderlichesuchstunden--;
            heutigesuchstunden++;
            label14.Text = "Gesuchte Stunden: " + heutigesuchstunden;

            if (erforderlichesuchstunden < 1)
            {
                dataGridView3.Rows.Add();
                rowcount2++;
                dataGridView3.Rows[rowcount2 - 1].DefaultCellStyle.BackColor = Color.YellowGreen;
                int portionen = 1;
                if (istselten) portionen = zufall.Next(1, 4);
                if (istungewöhnlich) portionen = zufall.Next(1, 6);
                if (istgewöhnlich) portionen = zufall.Next(1, 8);
                dataGridView3[0, rowcount2 - 1].Value = portionen + " Portionen der gesuchten Zutat " + comboBox2.GetItemText(comboBox2.SelectedItem) + " gefunden!";
                neuesuche();
            }
            else if (erforderlichesuchstunden % nebenfundteiler == 0)
            {
                dataGridView3.Rows.Add();
                rowcount2++;
                dataGridView3.Rows[rowcount2 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
                dataGridView3[0, rowcount2 - 1].Value = "Nebenbei " + zufall.Next(1,8) + " Portionen der Zutat" + Region2.gewöhnlich[zufall.Next(0, Region2.gewöhnlich.Length)] + " gefunden.";
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
            label14.Text = "Gesuchte Stunden: " + heutigesuchstunden;
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

            testergebnis2 = (int)numericUpDown5.Value + zufall.Next(1, 7) + zufall.Next(1, 7) + zufall.Next(1, 7);
            label15.Text = "Letzter Wurf: " + (testergebnis2 - (int)numericUpDown5.Value) + " + " + (int)numericUpDown5.Value + " = " + testergebnis2;

            if (numericUpDown9.Value != 0) testergebnis2 = (int)numericUpDown9.Value;
            label15.Text = "Letzter Wurf: " + (int)numericUpDown9.Value + " + " + (int)numericUpDown5.Value + " = " + testergebnis2;

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
            label14.Text = "Gesuchte Stunden: " + heutigesuchstunden;

            comboBox2.SelectedIndex = -1;
            label15.Text = "";
            testergebnis2 = 0;
            nebenfundteiler = 6;
            erforderlichesuchstunden = 10000;
            dataGridView3.Rows.Clear();
            rowcount2 = 0;
        }

        public void comboBox2_SelectedValueChanged(object sender, EventArgs e) //Zutat geädert
        {
            neuesuche();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) //Region geändert
        {
            if (comboBox2.GetItemText(comboBox2.SelectedItem) != null) { neuesuche(); }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e) //Umgebung geändert
        {
            if (comboBox2.GetItemText(comboBox2.SelectedItem) != null) { neuesuche(); }
        }


        //#################### ################################################
        //###### MARKT ####### ################################################
        //#################### ################################################

        private void button7_Click(object sender, EventArgs e) //Alles zurücksetzen
        {
            dataGridView2.Rows.Clear();
            rowcount3 = 0;
            händleranzahl = 0;
        }

        private void button8_Click(object sender, EventArgs e) //Markt generieren
        {
            if (radioButton1.Checked) händleranzahl = 1 + zufall.Next(1,5);
            else if (radioButton2.Checked) händleranzahl = 5 + zufall.Next(1, 11);
            else if (radioButton3.Checked) händleranzahl = 15 + zufall.Next(1, 21);
            else if (radioButton4.Checked) händleranzahl = 35 + zufall.Next(1, 21);

            if (radioButton1.Checked) sondergewerbe = zufall.Next(1, 5) - 1;
            else if (radioButton2.Checked) sondergewerbe = zufall.Next(1, 7);
            else if (radioButton3.Checked) sondergewerbe = zufall.Next(1, 9) + zufall.Next(1, 9);
            else if (radioButton4.Checked) sondergewerbe = zufall.Next(1, 13) + zufall.Next(1, 13);

            for (int i = 0; i < händleranzahl; i++)
            {
                int ergebnis = zufall.Next(1, 101);
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
                int ergebnis2 = zufall.Next(1, 101);
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

        //#################### ################################################
        //###### WAREN ####### ################################################
        //#################### ################################################

        public int XwY(int wuerfelzahl, int wuerfelart)
        {
            int ergebnis = 0;
            for (int i = 0; i < wuerfelzahl; i++)
            {
                ergebnis += zufall.Next(1, wuerfelart + 1);
            }
            return ergebnis;
        }

        public string ZufaelligerZauberNachArt(string art)
        {
            //alle einmal durchgehen und korrekte zeilenindizes in ein array werfen, dann array per zufall finden
            int[] auswahl = new int[100];
            int stelle = 0;
            DataRow ergebniszeile = ZauberTB.Rows[0];
            for (int i = 0; i < 285; i++)
            {
                ergebniszeile = ZauberTB.Rows[i];
                if (ergebniszeile["Art"].ToString() == art)
                {
                    auswahl[stelle] = Convert.ToInt16(ergebniszeile["ID"]);
                    stelle++;
                }
            }
            ergebniszeile = ZauberTB.Rows[auswahl[zufall.Next(0, auswahl.Length)]];

            return ergebniszeile["Zauber"].ToString();
        }


        public int ZufaelligerZauberNachKomplex(string komplex)
        {
            int[] auswahlzauber = new int[100];
            int stelle = 0;
            DataRow ergebniszeile = ZauberTB.Rows[0];
            for (int i = 0; i < ZauberTB.Rows.Count - 1; i++)
            {
                ergebniszeile = ZauberTB.Rows[i];
                if (ergebniszeile["Komplex"].ToString() == komplex)
                {
                    auswahlzauber[stelle] = Convert.ToInt16(ergebniszeile["ID"]);
                    stelle++;
                }
            }
            ergebniszeile = ZauberTB.Rows[auswahlzauber[zufall.Next(0, stelle)]];

            return Convert.ToInt16(ergebniszeile["ID"]) -2; //warum auch immer...
        }

        public List<int> AlleZauberEinesKomplexes(string komplex)
        {
            List<int> auswahlzauber = new List<int>();
            DataRow ergebniszeile = ZauberTB.Rows[0];
            for (int i = 0; i < ZauberTB.Rows.Count - 1; i++)
            {
                ergebniszeile = ZauberTB.Rows[i];
                if (ergebniszeile["Komplex"].ToString() == komplex)
                {
                    auswahlzauber.Add(Convert.ToInt16(ergebniszeile["ID"]) - 1);
                }
            }
            return auswahlzauber;
        }

        public List<int> AlleZauberEinerArt(string art)
        {
            List<int> auswahlzauber = new List<int>();
            DataRow ergebniszeile = ZauberTB.Rows[0];
            for (int i = 0; i < ZauberTB.Rows.Count - 1; i++)
            {
                ergebniszeile = ZauberTB.Rows[i];
                if (ergebniszeile["Art"].ToString() == art)
                {
                    auswahlzauber.Add(Convert.ToInt16(ergebniszeile["ID"]) - 1);
                }
            }
            return auswahlzauber;
        }

        public List<string> AlleKomplexeEinerArt(string art)
        {
            List<string> auswahlkomplexe = new List<string>();
            DataRow ergebniszeile = ZauberTB.Rows[0];
            for (int i = 0; i < ZauberTB.Rows.Count - 1; i++)
            {
                ergebniszeile = ZauberTB.Rows[i];
                if (ergebniszeile["Art"].ToString() == art && !auswahlkomplexe.Contains(ergebniszeile["Komplex"].ToString()))
                {
                    auswahlkomplexe.Add(ergebniszeile["Komplex"].ToString());
                }
            }
            return auswahlkomplexe;
        }

        public int ZauberFinden(string komplex, int komplexstufe)
        {
            DataRow ergebniszeile = ZauberTB.Rows[0];
            DataRow zeile = ZauberTB.Rows[0];
            for (int i = 0; i < 284; i++)
            {
                zeile = ZauberTB.Rows[i];
                if (zeile["Komplex"].ToString() == komplex && Convert.ToInt16(zeile["Komplexstufe"]) == komplexstufe)
                {
                    ergebniszeile = ZauberTB.Rows[i];
                }
            }
            return Convert.ToInt16(ergebniszeile["ID"]) -1;
        }

        public int ZufaelligerZaubersteinNachArt(string art)
        {
            int[] auswahl = new int[100];
            int stelle = 0;
            DataRow ergebniszeile = ZaubersteinTB.Rows[0];
            for (int i = 0; i < 92; i++)
            {
                ergebniszeile = ZaubersteinTB.Rows[i];
                if (ergebniszeile["Art"].ToString() == art)
                {
                    auswahl[stelle] = Convert.ToInt16(ergebniszeile["ID"]);
                    stelle++;
                }
            }
            ergebniszeile = ZauberTB.Rows[auswahl[zufall.Next(0, stelle)]];

            return Convert.ToInt16(ergebniszeile["ID"]) - 1;
        }

        public string ZufaelligerRegulaererKomplex()
        {
            DataRow ergebniszeile = ZauberTB.Rows[zufall.Next(1, 159)];

            return ergebniszeile["Komplex"].ToString();
        }

        public string ZufaelligerRegulaererKomplex(string zauberweg)
        {
            bool legit = false;
            string komplex = "";

            while (!legit)
            {
                DataRow ergebniszeile = ZauberTB.Rows[zufall.Next(1, 159)];
                komplex = ergebniszeile["Komplex"].ToString();

                switch (zauberweg)
                {
                    case "Feuer":
                        legit = true;
                        switch (komplex)
                        {
                            case "Blitz":
                            case "Wind":
                            case "Wasser":
                            case "Frost":
                            case "Pflanzen":
                            case "Erde":
                                legit = false;
                                break;
                        }
                        break;
                    case "Wasser":
                        legit = true;
                        switch (komplex)
                        {
                            case "Blitz":
                            case "Wind":
                            case "Feuer":
                            case "Sonne":
                            case "Pflanzen":
                            case "Erde":
                                legit = false;
                                break;
                        }
                        break;
                    case "Erde":
                        legit = true;
                        switch (komplex)
                        {
                            case "Blitz":
                            case "Wind":
                            case "Wasser":
                            case "Frost":
                            case "Feuer":
                            case "Sonne":
                                legit = false;
                                break;
                        }
                        break;
                    case "Luft":
                        legit = true;
                        switch (komplex)
                        {
                            case "Feuer":
                            case "Sonne":
                            case "Wasser":
                            case "Frost":
                            case "Pflanzen":
                            case "Erde":
                                legit = false;
                                break;
                        }
                        break;
                    case "Naturzauberei":
                        switch (komplex)
                        {
                            case "Leben":
                            case "Reinigung":
                            case "Erkenntnis":
                            case "Heilung":
                            case "Pflanzen":
                            case "Erde":
                            case "Wasser":
                            case "Sonne":
                            case "Wind":
                            case "Telepathie":
                                legit = true;
                                break;
                        }
                        break;
                    case "Ahnenzauberei1":
                        switch (komplex)
                        {
                            case "Telekinese":
                            case "Telepathie":
                            case "Erkenntnis":
                            case "Verhüllung":
                            case "Erleuchtung":
                                legit = true;
                                break;
                        }
                        break;
                    case "Ahnenzauberei2":
                        switch (komplex)
                        {
                            case "Widerstand":
                            case "Schutz":
                            case "Reinigung":
                            case "Ruhe":
                            case "Leben":
                                legit = true;
                                break;
                        }
                        break;
                    case "Baldan":
                        switch (komplex)
                        {
                            case "Widerstand":
                            case "Sonne":
                            case "Erleuchtung":
                            case "Reinigung":
                            case "Heilung":
                            case "Leben":
                                legit = true;
                                break;
                        }
                        break;
                    case "Diwan":
                        switch (komplex)
                        {
                            case "Erleuchtung":
                            case "Schutz":
                            case "Blitz":
                            case "Feuer":
                            case "Kraft":
                            case "Schaden":
                                legit = true;
                                break;
                        }
                        break;
                    case "Erda":
                        switch (komplex)
                        {
                            case "Pflanzen":
                            case "Erde":
                            case "Heilung":
                            case "Materie":
                            case "Verwandlung":
                            case "Leben":
                                legit = true;
                                break;
                        }
                        break;
                    case "Fria":
                        switch (komplex)
                        {
                            case "Geist":
                            case "Sonne":
                            case "Heilung":
                            case "Reinigung":
                            case "Spielerei":
                            case "Illusion":
                                legit = true;
                                break;
                        }
                        break;
                    case "Halla":
                        switch (komplex)
                        {
                            case "Frost":
                            case "Erde":
                            case "Materie":
                            case "Illusion":
                            case "Verhüllung":
                            case "Schaden":
                                legit = true;
                                break;
                        }
                        break;
                    case "Heimdan":
                        switch (komplex)
                        {
                            case "Schutz":
                            case "Ruhe":
                            case "Erkenntnis":
                            case "Frost":
                            case "Zeit":
                            case "Illusion":
                                legit = true;
                                break;
                        }
                        break;
                    case "Ingan":
                        switch (komplex)
                        {
                            case "Widerstand":
                            case "Feuer":
                            case "Sonne":
                            case "Pflanzen":
                            case "Schabernack":
                            case "Spielerei":
                                legit = true;
                                break;
                        }
                        break;
                    case "Lukan":
                        switch (komplex)
                        {
                            case "Blitz":
                            case "Feuer":
                            case "Erde":
                            case "Schabernack":
                            case "Materie":
                            case "Verhüllung":
                                legit = true;
                                break;
                        }
                        break;
                    case "Nertan":
                        switch (komplex)
                        {
                            case "Wind":
                            case "Wasser":
                            case "Ruhe":
                            case "Reinigung":
                            case "Schabernack":
                            case "Leben":
                                legit = true;
                                break;
                        }
                        break;
                    case "Saga":
                        switch (komplex)
                        {
                            case "Schutz":
                            case "Geist":
                            case "Ruhe":
                            case "Erleuchtung":
                            case "Erkenntnis":
                            case "Zeit":
                                legit = true;
                                break;
                        }
                        break;
                    case "Skanda":
                        switch (komplex)
                        {
                            case "Wasser":
                            case "Frost":
                            case "Pflanzen":
                            case "Kraft":
                            case "Jagd":
                            case "Spielerei":
                                legit = true;
                                break;
                        }
                        break;
                    case "Tunan":
                        switch (komplex)
                        {
                            case "Widerstand":
                            case "Wind":
                            case "Blitz":
                            case "Jagd":
                            case "Kraft":
                            case "Schaden":
                                legit = true;
                                break;
                        }
                        break;
                    case "Wodan":
                        switch (komplex)
                        {
                            case "Wind":
                            case "Geist":
                            case "Erkenntnis":
                            case "Jagd":
                            case "Verwandlung":
                            case "Verhüllung":
                                legit = true;
                                break;
                        }
                        break;
                }
                switch (zauberweg)
                { 
                    case "Baldan":
                    case "Diwan":
                    case "Erda":
                    case "Fria":
                    case "Halla":
                    case "Heimdan":
                    case "Ingan":
                    case "Lukan":
                    case "Nertan":
                    case "Saga":
                    case "Skanda":
                    case "Tunan":
                    case "Wodan":
                        switch (komplex)
                        {
                            case "Teleportation":
                            case "Telekinese":
                            case "Telepathie":
                            case "Kontrolle":
                            case "Störung":
                                legit = true;
                                break;
                        }
                        break;
                }
            }

            return komplex;
        }

        public string ZufälligeGottheit()
        {
            int zuf = zufall.Next(1, 14);
            string gh = "";

            switch (zuf)
            {
                case 1:
                    gh = "Baldan";
                    break;
                case 2:
                    gh = "Diwan";
                    break;
                case 3:
                    gh = "Erda";
                    break;
                case 4:
                    gh = "Fria";
                    break;
                case 5:
                    gh = "Halla";
                    break;
                case 6:
                    gh = "Heimdan";
                    break;
                case 7:
                    gh = "Ingan";
                    break;
                case 8:
                    gh = "Lukan";
                    break;
                case 9:
                    gh = "Nertan";
                    break;
                case 10:
                    gh = "Saga";
                    break;
                case 11:
                    gh = "Skanda";
                    break;
                case 12:
                    gh = "Tunan";
                    break;
                case 13:
                    gh = "Wodan";
                    break;
            }

            return gh;
        }

        public DataRow ZauberTrankRolleGenerator(string modus, bool Haendlertabelle)
        {
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            if (Haendlertabelle) dataGridView4.Rows.Add();
            if (Haendlertabelle) rowcount4++;
            if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
            //Zauber ermitteln
            string komplex = ZufaelligerRegulaererKomplex();
            int wurf = zufall.Next(1, 101);
            if (wurf == 98)
            {
                komplex = "Lebenszauber";
                if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Black;
                if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.ForeColor = Color.GhostWhite;
            }
            else if (wurf == 99)
            {
                komplex = "Totenzauber";
                if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Black;
                if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.ForeColor = Color.GhostWhite;
            }
            else if (wurf == 100)
            {
                komplex = "Seelenzauber";
                if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Black;
                if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.ForeColor = Color.GhostWhite;
            }
                
            //Bei Rollen: Auch Sternbilder möglich
            if (modus == "rolle")
            {
                if (wurf == 94 || wurf == 95 || wurf == 96 || wurf == 97)
                {
                    komplex = "Sternbild";
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.MidnightBlue;
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.ForeColor = Color.Yellow;
                }
            }

            int stufezufall = zufall.Next(1, 11);
            int komplexstufe = 0;
            DataRow ergebniszeile = ZauberTB.Rows[0];
            if (komplex != "Seelenzauber" && komplex != "Totenzauber" && komplex != "Lebenszauber" && komplex != "Sternbild")
            {
                switch (stufezufall)
                {
                    case 1:
                    case 2:
                    case 3:
                        komplexstufe = 1;
                        break;
                    case 4:
                    case 5:
                    case 6:
                        komplexstufe = 2;
                        break;
                    case 7:
                    case 8:
                        komplexstufe = 3;
                        break;
                    case 9:
                        komplexstufe = 4;
                        if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Goldenrod;
                        break;
                    case 10:
                        komplexstufe = 5;
                        if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Goldenrod;
                        break;
                }
                ergebniszeile = ZauberTB.Rows[ZauberFinden(komplex, komplexstufe)];
            }
            else
            {
                ergebniszeile = ZauberTB.Rows[ZufaelligerZauberNachKomplex(komplex)];
                while (ergebniszeile["Zauber"].ToString() == "Zauberhaut") //Zauberhaut darf nicht als Trank existieren, nur als Artefakt
                { 
                    ergebniszeile = ZauberTB.Rows[ZufaelligerZauberNachKomplex(komplex)]; 
                }
                    
            }     

            //Zauber anzeigen
            double bonusstufen = 0; 
            double bonusstufenmöglichkeiten = Convert.ToDouble(ergebniszeile["Bonusstufen"]);
            for (int a = 0; a < bonusstufenmöglichkeiten; a++)
            {
                bonusstufen += zufall.Next(0, 3);
            }
            bonusstufenmöglichkeiten *= 2;
            double bonusquote = (bonusstufen / bonusstufenmöglichkeiten) * 100;
            if (bonusstufenmöglichkeiten == 0) { bonusquote = 0; }
            else if (bonusquote == 100)
            {
                if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.YellowGreen;
                if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
            }
            else if (bonusquote == 0)
            {
                if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Orange;
                if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
            }
            string objekt = "";
            int preismod = 1;
            if (modus == "trank")
            {
                objekt = "Zaubertrank: ";
                preismod = 20;
            }
            else if (modus == "rolle")
            {
                objekt = "Zauberrolle: ";
                preismod = 10;
            }

            if (Haendlertabelle) dataGridView4[0, rowcount4 - 1].Value = objekt + ergebniszeile["Zauber"].ToString() + " (" + ergebniszeile["Stufe"].ToString() + ")";
            int gesamtstufe = (int)bonusstufen + Convert.ToInt16(ergebniszeile["Stufe"]);
            if (komplex == "Seelenzauber" || komplex == "Totenzauber" || komplex == "Lebenszauber") { gesamtstufe *= 3; } //dreifacher Preis für verbotene Zauberei
            if (Haendlertabelle) dataGridView4[1, rowcount4 - 1].Value = (double)(preismod * gesamtstufe);
            if (Haendlertabelle) dataGridView4[2, rowcount4 - 1].Value = "1";
            if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Value = "Bonusstufen: " + Convert.ToString(bonusstufen) + " von " + Convert.ToString(bonusstufenmöglichkeiten) + " (" + Math.Round(bonusquote) + "%)";
            if (Haendlertabelle) dataGridView4[4, rowcount4 - 1].Value = "Ja";
            if (Haendlertabelle) dataGridView4[5, rowcount4 - 1].Value = "Komplex: " + komplex + " (" + komplexstufe + ")";
            if (komplexstufe == 0 && Haendlertabelle) { dataGridView4[5, rowcount4 - 1].Value = "Komplex: " + komplex; }

            dummyergebnis["Beschreibung"] = "Komplex: " + komplex + " (" + komplexstufe + ")" + " Bonusstufen: " + Convert.ToString(bonusstufen) + " von " + Convert.ToString(bonusstufenmöglichkeiten) + " (" + Math.Round(bonusquote) + "%)";
            dummyergebnis["Wert"] = (double)(preismod * gesamtstufe);
            dummyergebnis["Name"] = objekt + ergebniszeile["Zauber"].ToString() + " (" + ergebniszeile["Stufe"].ToString() + ")";

            return dummyergebnis;

        }
        

        public void ArtefaktGenerator()
        {
            //+++
            //+++
            //+++
            dataGridView4.Rows.Add();
            rowcount4++;
            dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Red;
            dataGridView4[0, rowcount4 - 1].Value = "Artefakt";
            dataGridView4[1, rowcount4 - 1].Value = (double)999;
            dataGridView4[2, rowcount4 - 1].Value = 1;
            dataGridView4[3, rowcount4 - 1].Value = "Mittel";
            dataGridView4[4, rowcount4 - 1].Value = "Siehe Kapitel 11.3 im Regelwerk.";
            dataGridView4[5, rowcount4 - 1].Value = "";
        }

        public DataRow ZaubersteinGenerator(bool Haendlertabelle)
        {
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.

            if (Haendlertabelle) dataGridView4.Rows.Add();
            if (Haendlertabelle) rowcount4++;
            if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.LightYellow;

            //Art ermitteln
            int zufallart = zufall.Next(1, 101);
            string art = "";
            if (zufallart <= 45) { art = "Regulär"; }
            if (zufallart > 45) { art = "Ahnenzauber"; }
            if (zufallart > 60) { art = "Runenzauber"; }
            if (zufallart > 85) { art = "Bannwort"; }
            if (zufallart > 96) { art = "Totenzauber"; }
            if (zufallart > 98) { art = "Seelenzauber"; }
            //Zaubersteintabelle
            DataRow ergebniszeile = ZaubersteinTB.Rows[0];
            ergebniszeile = ZaubersteinTB.Rows[ZufaelligerZaubersteinNachArt(art)];
            switch (ergebniszeile["Art"].ToString())
            {
                case "Regulär":
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Thistle;
                    break;
                case "Bannwort":
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.PaleVioletRed;
                    break;
                case "Runenzauber":
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.BurlyWood;
                    break;
                case "Ahnenzauber":
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.PowderBlue;
                    break;
                case "Totenzauber":
                case "Lebenszauber":
                case "Seelenzauber":
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.Black;
                    if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.ForeColor = Color.GhostWhite;
                    break;
            }

            //Zauberstein und Karatzahl ermitteln
            DataRow stein = ZierTB.Rows[0];
            stein = ZierTB.Rows[zufall.Next(0, 100)];
            string steinname = "";
            bool steinlegitim = false;
            int mindestkaratzahl = 5;
            while(!steinlegitim)
            {
                stein = ZierTB.Rows[zufall.Next(0, 100)];
                steinname = stein["Schmuckstein"].ToString();
                switch (steinname)
                {
                    case "Amethyst":
                        mindestkaratzahl = 8;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Orchid;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        steinlegitim = true;
                        break;
                    case "Diamant":
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.PaleTurquoise;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        steinlegitim = true;
                        break;
                    case "Rubin":
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Red;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        steinlegitim = true;
                        break;
                    case "Saphir":
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.SlateBlue;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        steinlegitim = true;
                        break;
                    case "Smaragd":
                        steinlegitim = true;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.LimeGreen;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        break;
                    case "Granat":
                        mindestkaratzahl = 10;
                        steinlegitim = true;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Crimson;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        break;
                    case "Jade":
                        mindestkaratzahl = 10;
                        steinlegitim = true;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Lime;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        break;
                    case "Spinell":
                        mindestkaratzahl = 10;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Plum;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        steinlegitim = true;
                        break;
                    case "Mondstein":
                        mindestkaratzahl = 12;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Wheat;
                        if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Style.ForeColor = Color.Black;
                        steinlegitim = true;
                        break;
                    case "Opal":
                    case "Topas":
                        mindestkaratzahl = 6;
                        steinlegitim = true;
                        break;
                    default:
                        steinlegitim = false;
                        break;
                }
            }
            int karatzahl = XwY(Convert.ToInt16(stein["RohWurfZahl"]), Convert.ToInt16(stein["RohWuerfel"]));
            while (karatzahl < mindestkaratzahl)
            {
                karatzahl = XwY(Convert.ToInt16(stein["RohWurfZahl"]), Convert.ToInt16(stein["RohWuerfel"]));
            }

            int edelsteinwert = karatzahl * Convert.ToInt16(stein["Preis"]);

            //Zauberstein anzeigen
            if (Haendlertabelle) dataGridView4[0, rowcount4 - 1].Value = "Zauberstein: " + ergebniszeile["Name"].ToString();
            if (Haendlertabelle) dataGridView4[1, rowcount4 - 1].Value = (double)(50 + zufall.Next(1, 51) + edelsteinwert);
            if (Haendlertabelle) dataGridView4[2, rowcount4 - 1].Value = 0.1;
            if (Haendlertabelle) dataGridView4[3, rowcount4 - 1].Value = steinname + " (" + karatzahl + " kt)";
            if (Haendlertabelle) dataGridView4[4, rowcount4 - 1].Value = "Ja";
            if (Haendlertabelle) dataGridView4[5, rowcount4 - 1].Value = ergebniszeile["Wirkung"].ToString() + " (" + ergebniszeile["Komplex"].ToString() + ")";

            dummyergebnis["Beschreibung"] = steinname + " (" + karatzahl + " kt) " + ergebniszeile["Wirkung"].ToString() + " (" + ergebniszeile["Komplex"].ToString() + ")";
            dummyergebnis["Wert"] = (edelsteinwert + 50).ToString();
            dummyergebnis["Name"] = "Zauberstein: " + ergebniszeile["Name"].ToString();

            return dummyergebnis;
        }

        public DataRow ZufaelligerRohstein(bool Haendlertabelle)
        {
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            if (Haendlertabelle) dataGridView4.Rows.Add();
            if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
            DataRow schmuckstein = ZierTB.Rows[zufall.Next(1, 100)];
            int karatzahl = XwY(Convert.ToInt16(schmuckstein["RohWurfZahl"]), Convert.ToInt16(schmuckstein["RohWuerfel"]));
            if (Haendlertabelle) dataGridView4[0, rowcount4 - 1].Value = schmuckstein["Schmuckstein"].ToString() + " (" + karatzahl + " kt)";
            double realwert = karatzahl * (double)schmuckstein["Preis"];
            double preis = realwert;
            if (!checkBox2.Checked && Haendlertabelle) //Preisschwankungen (sind bei Schmucksteinen weniger extrem als sonst)
            {
                int preisschwank = zufall.Next(1, 21);
                switch (preisschwank)
                {
                    case 1:
                        preis *= 0.7;
                        break;
                    case 2:
                    case 3:
                        preis *= 0.8;
                        break;
                    case 4:
                    case 5:
                        preis *= 0.9;
                        break;
                    case 13:
                    case 14:
                    case 15:
                    case 16:
                    case 17:
                        preis *= 1.2;
                        break;
                    case 18:
                    case 19:
                        preis *= 1.5;
                        break;
                    case 20:
                        preis *= 2;
                        break;
                }
            }
            if (comboBox6.GetItemText(comboBox6.SelectedItem) == "Hehler" && Haendlertabelle
                || comboBox6.GetItemText(comboBox6.SelectedItem) == "Zwielichtiger Händler" && Haendlertabelle)
            {
                preis *= 0.7;
            }
            if (preis > realwert)
            {
                if (Haendlertabelle) dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.Orange;
            }
            else if (preis < realwert)
            {
                if (Haendlertabelle) dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.YellowGreen;
            }
            if (preis > 20) preis = (double)Math.Round(preis); //Hohe Preise sollen keine Kommabeträge mehr haben
            else preis = (double)Math.Round(preis, 2);
            if (Haendlertabelle) dataGridView4[1, rowcount4 - 1].Value = preis;
            if (Haendlertabelle) dataGridView4[2, rowcount4 - 1].Value = Math.Round(karatzahl * 0.0004, 1);
            if (!checkBox4.Checked && Haendlertabelle) //Preisschwank verbergen
            {
                dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.LightYellow;
            }
            if (Haendlertabelle)
            {
                switch (schmuckstein["Schmuckstein"].ToString())
                {
                    case "Diamant":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.PaleTurquoise;
                        break;
                    case "Rubin":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Red;
                        break;
                    case "Saphir":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.SlateBlue;
                        break;
                    case "Smaragd":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.LimeGreen;
                        break;
                    case "Bernstein":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.FromArgb(255, 235, 163, 40);
                        break;
                    case "Jade":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Lime;
                        break;
                    case "Amethyst":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Orchid;
                        break;
                    case "Lapislazuli":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.MediumBlue;
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.White;
                        break;
                    case "Granat":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Crimson;
                        break;
                }
                if (comboBox6.GetItemText(comboBox6.SelectedItem) == "Hehler" && Haendlertabelle
                    || comboBox6.GetItemText(comboBox6.SelectedItem) == "Zwielichtiger Händler" && Haendlertabelle)
                {
                    dataGridView4[4, rowcount4 - 1].Value = "Diebesgut!";
                    dataGridView4[4, rowcount4 - 1].Style.BackColor = Color.Black;
                    dataGridView4[4, rowcount4 - 1].Style.ForeColor = Color.GhostWhite;
                }
            }
            dummyergebnis["Beschreibung"] = karatzahl + " kt";
            dummyergebnis["Wert"] = preis.ToString();
            dummyergebnis["Name"] = "Schmuckstein: " + schmuckstein["Schmuckstein"].ToString();

            return dummyergebnis;
        }

        public DataRow ZufaelligesSchmuckstueck(bool Haendlertabelle, string schmucktyp)
        {
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.

            string schmuckbezeichnung = "";
            double realwert = 0;
            double gesamtgewicht = 0;
            DataRow ergebniszeilemetall = MetallTB.Rows[0];
            DataRow ergebniszeilezier = ZierTB.Rows[0];
            //Gegenstand
            DataRow ergebniszeile = SchmuckTB.Rows[zufall.Next(1, 130)];
            if (schmucktyp != "" || schmucktyp == null) 
            {
                int k = 0;
                while (ergebniszeile["Schmuck"].ToString() != schmucktyp)
                {
                    ergebniszeile = SchmuckTB.Rows[k];
                    k++;
                }
            }
            schmuckbezeichnung += ergebniszeile["Schmuck"].ToString();
            //Metall
            int metallgewicht = XwY(Convert.ToInt16(ergebniszeile["MetallWurfZahl"]), Convert.ToInt16(ergebniszeile["MetallWuerfel"])) + Convert.ToInt16(ergebniszeile["MetallMod"]);
            if (metallgewicht > 0)
            {
                schmuckbezeichnung += " aus " + (double)metallgewicht * 0.01 + " Pfund " /*+ "(=" + (double)metallgewicht * 25 + " Karat) "*/;
                ergebniszeilemetall = MetallTB.Rows[zufall.Next(1, 100)];
                schmuckbezeichnung += ergebniszeilemetall["Metall"].ToString();
                realwert += metallgewicht * (double)ergebniszeilemetall["Preis"];
                gesamtgewicht += metallgewicht * 0.01;
            }
            //Zierelemente
            int anzahlzierelemente = XwY(Convert.ToInt16(ergebniszeile["ZierWurfZahl"]), Convert.ToInt16(ergebniszeile["ZierWuerfel"])) + Convert.ToInt16(ergebniszeile["ZierMod"]);
            if (anzahlzierelemente > 0)
            {
                int ergebnisverschiedenheitsgrad = zufall.Next(1, 21);
                int anzahlelementtypen = 0;
                //anzahlzierelemente verschwindet...! +++
                switch (ergebnisverschiedenheitsgrad)
                {
                    case 1:
                    case 2:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 9:
                    case 10:
                    case 11:
                    case 12:
                    case 13:
                    case 14:
                        anzahlelementtypen = 1;
                        break;
                    case 15:
                    case 16:
                    case 17:
                        anzahlelementtypen = 2;
                        break;
                    case 18:
                    case 19:
                        anzahlelementtypen = 2;
                        break;
                    case 20:
                        anzahlelementtypen = anzahlzierelemente;
                        break;
                }

                for (int a = 0; a < anzahlelementtypen; a++)
                {
                    ergebniszeilezier = ZierTB.Rows[zufall.Next(1, 100)];
                    if (a == 0) { schmuckbezeichnung += " mit "; }
                    else { schmuckbezeichnung += " und "; }
                    schmuckbezeichnung += ergebniszeilezier["Schmuckstein"].ToString() + " (";
                    int karatzahl = XwY(Convert.ToInt16(ergebniszeilezier["ZierWurfZahl"]), Convert.ToInt16(ergebniszeilezier["ZierWuerfel"]));
                    schmuckbezeichnung += karatzahl + " kt)";
                    realwert += karatzahl * (double)ergebniszeilezier["Preis"];
                    gesamtgewicht += karatzahl * 0.0004; //1 Pfund sind 2500 Karat
                }
            }

            if (Haendlertabelle) dataGridView4.Rows.Add();
            if (Haendlertabelle) rowcount4++;
            if (Haendlertabelle) dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
            if (Haendlertabelle) dataGridView4[0, rowcount4 - 1].Value = schmuckbezeichnung;
            double preis = realwert;
            //Preisschwankungen (sind bei Schmuck weniger extrem als sonst)
            if (!checkBox2.Checked && Haendlertabelle)
            {
                int preisschwank = zufall.Next(1, 21);
                switch (preisschwank)
                {
                    case 1:
                        preis *= 0.7;
                        break;
                    case 2:
                    case 3:
                        preis *= 0.8;
                        break;
                    case 4:
                    case 5:
                        preis *= 0.9;
                        break;
                    case 13:
                    case 14:
                    case 15:
                    case 16:
                    case 17:
                        preis *= 1.2;
                        break;
                    case 18:
                    case 19:
                        preis *= 1.5;
                        break;
                    case 20:
                        preis *= 2;
                        break;
                }
            }
            //Hehler verkaufen 30% günstiger
            if (comboBox6.GetItemText(comboBox6.SelectedItem) == "Hehler" && Haendlertabelle
                || comboBox6.GetItemText(comboBox6.SelectedItem) == "Zwielichtiger Händler" && Haendlertabelle)
            {
                preis *= 0.7;
            }
            //Preisschwank farblich anzeigen
            if (preis > realwert)
            {
                if (Haendlertabelle) dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.Orange;
            }
            else if (preis < realwert)
            {
                if (Haendlertabelle) dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.YellowGreen;
            }
            //Preisschwank verbergen
            if (!checkBox4.Checked && Haendlertabelle)
            {
                dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.LightYellow;
            }

            if (realwert > 20) preis = (double)Math.Round(preis);
            else preis = (double)Math.Round(preis, 2); //Hohe Preise sollen keine Kommabeträge mehr haben
            if (Haendlertabelle)
            { 
                dataGridView4[1, rowcount4 - 1].Value = preis;
                dataGridView4[2, rowcount4 - 1].Value = Math.Round(gesamtgewicht, 1) + Convert.ToInt16(ergebniszeile["Zusatzgewicht"]);
                dataGridView4[3, rowcount4 - 1].Value = "Mittel";
                dataGridView4[5, rowcount4 - 1].Value = ergebniszeile["Beschreibung"].ToString();
                dataGridView4[4, rowcount4 - 1].Value = "Ja";
                dataGridView4.Columns[0].Width = 500;
                dataGridView4.Columns[3].Visible = false;
                dataGridView4.Columns[4].Visible = false;
                switch (ergebniszeilemetall["Metall"].ToString())
                {
                    case "Asterium":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.LightSteelBlue;
                        break;
                    case "Adamant":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Firebrick;
                        break;
                    case "Gold":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Gold;
                        break;
                    case "Bernstein":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.FromArgb(255, 235, 163, 40);
                        break;
                    case "Elektrum":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.LightGray;
                        break;
                    case "Silber":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Silver;
                        break;
                    case "Bronze":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.DarkOrange;
                        break;
                    case "Messing":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Goldenrod;
                        break;
                    case "Kupfer":
                        dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.Tomato;
                        break;
                }
                if (metallgewicht == 0) { dataGridView4[0, rowcount4 - 1].Style.BackColor = Color.LightYellow; }
                switch (ergebniszeilezier["Schmuckstein"].ToString())
                {
                    case "Diamant":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.PaleTurquoise;
                        break;
                    case "Rubin":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.Red;
                        break;
                    case "Saphir":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.SlateBlue;
                        break;
                    case "Smaragd":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.LimeGreen;
                        break;
                    case "Bernstein":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.FromArgb(255, 235, 163, 40);
                        break;
                    case "Jade":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.Lime;
                        break;
                    case "Amethyst":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.Orchid;
                        break;
                    case "Lapislazuli":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.MediumBlue;
                        break;
                    case "Granat":
                        dataGridView4[0, rowcount4 - 1].Style.ForeColor = Color.Crimson;
                        break;
                }
                if (comboBox6.GetItemText(comboBox6.SelectedItem) == "Hehler" || comboBox6.GetItemText(comboBox6.SelectedItem) == "Zwielichtiger Händler")
                {
                    dataGridView4[4, rowcount4 - 1].Value = "Diebesgut!";
                    dataGridView4[4, rowcount4 - 1].Style.BackColor = Color.Black;
                    dataGridView4[4, rowcount4 - 1].Style.ForeColor = Color.GhostWhite;
                }
            }

            dummyergebnis["Beschreibung"] = "";
            dummyergebnis["Wert"] = preis.ToString();
            dummyergebnis["Name"] = schmuckbezeichnung;

            return dummyergebnis;
        }

        public void SchmuckHaendlerGenerator()
        {
            int warenzahl = zufall.Next(1, 9) + zufall.Next(1, 9) + zufall.Next(1, 9);
            if (zufall.Next(1, 5) == 1) { warenzahl += zufall.Next(1, 31); } //in 25% der Fälle: Reichhaltigeres Angebot, also +1W30
            int rohsteinzahl = zufall.Next(1, 13);

            for (int i = 0; i < warenzahl; i++)
            {
                ZufaelligesSchmuckstueck(true, "");
            }

            for (int i = 0; i < rohsteinzahl; i++)
            {
                ZufaelligerRohstein(true);
            }
            //dataGridView4.Sort(dataGridView4.Columns[4], ListSortDirection.Descending);
        }

        public void WarenDurchsuchen(string Haendlertyp, string[] NurDieseWaren, double Haufigkeitsmod, bool Limit)
        {
            string kulturkontexthatkeineneinfluss = "Der Kulturkontext hat keine Auswirkungen.";
            if (Haendlertyp == "Schmuck") {
                SchmuckHaendlerGenerator();
                return;
            }
            else if (Haendlertyp == "Zaubertrank") {
                int trankzahl = zufall.Next(1, 11) + zufall.Next(1, 11) + zufall.Next(1, 11);
                for (int i = 0; i < trankzahl; i++)
                {
                    ZauberTrankRolleGenerator("trank", true);
                }
                return;
            }
            else if (Haendlertyp == "Zauberrolle")
            {
                int trankzahl = zufall.Next(1, 11) + zufall.Next(1, 11) + zufall.Next(1, 11);
                for (int i = 0; i < trankzahl; i++)
                {
                    ZauberTrankRolleGenerator("rolle", true);
                }
                return;
            }
            else if (Haendlertyp == "Artefakte")
            {
                ArtefaktGenerator();
                return;
            }
            else if (Haendlertyp == "Zaubersteine")
            {
                int steinzahl = zufall.Next(1, 11) + zufall.Next(1, 11) + zufall.Next(1, 11);
                for (int i = 0; i < steinzahl; i++)
                {
                    ZaubersteinGenerator(true);
                }
                    return;
            }

            //Beim Geldwechsler keine Preis- und Qualitätsschwankungen
            if (Haendlertyp == "Geldwechsler")
            {
                checkBox3.Checked = true;
                checkBox2.Checked = true;
            }

            foreach (DataRow row in WarenTB.Rows)
            {
                if (row["Haendler"].ToString() == Haendlertyp)
                {
                    bool warenichtabbilden = false;
                    if (NurDieseWaren != null) //Wenn es sich um einen Verweis mit nur exklusiven Waren handelt.
                    {
                        warenichtabbilden = true;
                        for (int i = 0; i < NurDieseWaren.Length; i++)
                        {
                            if (NurDieseWaren[i] == " " + row["Ware"].ToString())
                            {
                                warenichtabbilden = false;
                            }
                        }
                    }
                    
                    //Beim (durch Verweis erzeugten) Hehler das Verweislimit nicht setzen
                    if (Haendlertyp == "Hehler") { Limit = false; }
                    //Beim durch Verweis erzeugten Tränke-Händler ebenso nicht
                    if (Haendlertyp == "Tränke") { Limit = false; }

                    if (row["Verweis"].ToString() == "" && !warenichtabbilden)
                    {
                        string[] kulturkontexte = row["Kulturkontexte"].ToString().Split(',');
                        bool richtigekultur = false;
                        for (int i = 0; i < kulturkontexte.Length; i++)
                        {
                            if (kulturkontexte[i] == " " + comboBox7.GetItemText(comboBox7.SelectedItem))
                            {
                                richtigekultur = true;
                            }
                        }
                        if (richtigekultur || row["Kulturkontexte"].ToString() == "")
                        {
                            dataGridView4.Rows.Add();
                            rowcount4++;
                            dataGridView4[0, rowcount4 - 1].Value = row["Ware"].ToString();
                            dataGridView4[2, rowcount4 - 1].Value = row["Gewicht"].ToString();
                            dataGridView4[3, rowcount4 - 1].Value = "Mittel";
                            dataGridView4[5, rowcount4 - 1].Value = row["Beschreibung"].ToString(); ;
                            dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
                            double preis = (double)row["Realwert"];
                            int gewichttemp = 0;
                            if (row["Gewicht"].ToString() != "")
                            {
                                gewichttemp = Convert.ToInt32(row["Gewicht"]);
                            }
                            if (gewichttemp < -100) //Pferdegewichte (WX)
                            {
                                double zufallsgewicht = zufall.Next((int)(Math.Abs(gewichttemp) * 0.7), Math.Abs(gewichttemp) + 1);
                                dataGridView4[2, rowcount4 - 1].Value = zufallsgewicht;
                            }
                            else if (gewichttemp < 0) //Negative Werte bedeuten Zufallswerte (WX)
                            {
                                double zufallsgewicht = zufall.Next(1, (int)(Math.Abs(gewichttemp)) + 1);
                                dataGridView4[2, rowcount4 - 1].Value = zufallsgewicht;
                                preis = (double)row["Realwert"] * zufallsgewicht;
                            }

                            //Qualitätsschwankungen
                            if (!checkBox3.Checked)
                            {
                                int quali = zufall.Next(1, 21);
                                switch (quali)
                                {
                                    case 1:
                                    case 2:
                                        dataGridView4[3, rowcount4 - 1].Value = "Beschädigt, nicht verwendbar!";
                                        dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.OrangeRed;
                                        preis *= 0.5;
                                        break;
                                    case 3:
                                    case 4:
                                    case 5:
                                    case 6:
                                        dataGridView4[3, rowcount4 - 1].Value = "Niedrige Qualität: -1 auf Aktionen (bei Waffen: auf Schaden)";
                                        dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Orange;
                                        preis *= 0.5;
                                        break;
                                    case 18:
                                    case 19:
                                        dataGridView4[3, rowcount4 - 1].Value = "Hohe Qualität: +1 auf Aktionen (bei Waffen: auf Schaden)";
                                        dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.YellowGreen;
                                        preis *= 1.5;
                                        break;
                                    case 20:
                                        dataGridView4[3, rowcount4 - 1].Value = "Herausragende Qualität: +3 auf Aktionen (bei Waffen: auf Schaden)";
                                        dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.Gold;
                                        preis *= 3;
                                        break;
                                }
                            }
                            //Preisschwankungen
                            if (!checkBox2.Checked)
                            {
                                int preisschwank = zufall.Next(1, 21);
                                switch (preisschwank)
                                {
                                    case 1:
                                        preis *= 0.5;
                                        //dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.Gold;
                                        break;
                                    case 2:
                                    case 3:
                                        preis *= 0.7;
                                        //dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.YellowGreen;
                                        break;
                                    case 4:
                                    case 5:
                                        preis *= 0.9;
                                        //dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.LightGreen;
                                        break;
                                    case 13:
                                    case 14:
                                    case 15:
                                    case 16:
                                    case 17:
                                        preis *= 1.5;
                                        //dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.Orange;
                                        break;
                                    case 18:
                                    case 19:
                                        preis *= 2;
                                        //dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.OrangeRed;
                                        break;
                                    case 20:
                                        preis *= 3;
                                        //dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.LightCoral;
                                        break;
                                }
                            }

                            //Hehler verkaufen 30% günstiger //EDIT: und zwielichtige H. auch, wenn es sich um Verweise handelt.
                            if (comboBox6.GetItemText(comboBox6.SelectedItem) == "Hehler" || comboBox6.GetItemText(comboBox6.SelectedItem) == "Zwielichtiger Händler" && Haendlertyp != "Zwielichtiger Händler")
                            {
                                preis *= 0.7;
                            }

                            dataGridView4[1, rowcount4 - 1].Value = (double)Math.Round(preis, 2);
                            if (preis > 20) { dataGridView4[1, rowcount4 - 1].Value = (double)Math.Round(preis); } //Hohe Preise sollen keine Kommabeträge mehr haben
                            if (preis > (double)row["Realwert"])
                            {
                                dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.Orange;
                            }
                            else if (preis < (double)row["Realwert"])
                            {
                                dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.YellowGreen;
                            }
                            if (!checkBox4.Checked) //Preisschwank verbergen
                            {
                                dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.LightYellow;
                            }

                            double nettowahrscheinlichkeit = (double)row["Wahrscheinlichkeit"];
                            if (Haufigkeitsmod != 1000) { nettowahrscheinlichkeit += Haufigkeitsmod; }

                            if (nettowahrscheinlichkeit >= zufall.Next(1, 11) || checkBox5.Checked) { dataGridView4[4, rowcount4 - 1].Value = "Ja"; }
                            else //Nicht verfügbar
                            {
                                dataGridView4[4, rowcount4 - 1].Value = row["Dauer"].ToString();
                                if ((double)row["Wahrscheinlichkeit"] != 0)
                                {
                                    dataGridView4[1, rowcount4 - 1].Value = (double)0;
                                    dataGridView4[2, rowcount4 - 1].Value = "";
                                    dataGridView4[3, rowcount4 - 1].Value = "";

                                    dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.LightCoral;
                                    dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.LightCoral;
                                    dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.LightCoral;
                                }
                                else //Nur nach Auftrag
                                {
                                    dataGridView4.Rows[rowcount4 - 1].DefaultCellStyle.BackColor = Color.LightBlue;
                                    dataGridView4[1, rowcount4 - 1].Style.BackColor = Color.LightBlue;
                                    dataGridView4[3, rowcount4 - 1].Style.BackColor = Color.LightBlue;
                                    dataGridView4[3, rowcount4 - 1].Value = "Nur nach Auftrag!";
                                }
                                if (Limit || comboBox6.GetItemText(comboBox6.SelectedItem) == "Zwielichtiger Händler") //ergo: wenn diese Ware durch einen Verweis ins Inventar kommt... //EDIT: ...oder der Zw. H. gewählt ist...
                                {
                                    dataGridView4.Rows[rowcount4 - 1].Visible = false; //=> ...gar nicht erst als Zeile anzeigen.
                                }
                            }
                            //Alle Hehlerwaren sind Diebesgut
                            //Alle durch Verweis erzeugten Zwielichtigen Waren sind Diebesgut
                            if (comboBox6.GetItemText(comboBox6.SelectedItem) == "Hehler" || comboBox6.GetItemText(comboBox6.SelectedItem) == "Zwielichtiger Händler" && Haendlertyp != "Zwielichtiger Händler")
                            {
                                dataGridView4[4, rowcount4 - 1].Value = "Diebesgut!";
                                dataGridView4[4, rowcount4 - 1].Style.BackColor = Color.Black;
                                dataGridView4[4, rowcount4 - 1].Style.ForeColor = Color.GhostWhite;
                                //Für unter 1 S stehlen, lohnt nicht.
                                if (preis < 1)
                                {
                                    dataGridView4.Rows[rowcount4 - 1].Visible = false; //=> ...gar nicht erst als Zeile anzeigen.
                                }
                            }
                        }
                        else {
                            kulturkontexthatkeineneinfluss = "Der Kulturkontext hat Auswirkungen!";
                        }
                    }
                    else if ((double)row["Wahrscheinlichkeit"] >= zufall.Next(1, 11) && !Limit)
                    {
                        double VHMod = 0;
                        if (row["VerweisHaufigkeitsMod"].ToString() != "") { VHMod = (double)row["VerweisHaufigkeitsMod"]; }
                        string[] VWaren = null;
                        if (row["VerweisNurDieseWaren"].ToString() != "") { VWaren = row["VerweisNurDieseWaren"].ToString().Split(','); }
                        WarenDurchsuchen(row["Verweis"].ToString(), VWaren, VHMod, true);
                    }
                }
            }
            dataGridView4.Columns[0].Width = 200;
            dataGridView4.Columns[3].Visible = true;
            dataGridView4.Columns[4].Visible = true;
            //dataGridView4.Sort(dataGridView4.Columns[4], ListSortDirection.Descending);
            if (comboBox7.GetItemText(comboBox7.SelectedItem) != "") { label21.Text = kulturkontexthatkeineneinfluss; }
            else { label21.Text = ""; }
            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView4.Rows.Clear();
            rowcount4 = 0;
            string[] leer = null;
            WarenDurchsuchen(comboBox6.GetItemText(comboBox6.SelectedItem), leer, 1000, false);
        }

        //#################### ################################################
        //###### ZAUBER ###### ################################################
        //#################### ################################################

        private string ZufälligerKomplexnachKontext(string kontext)
        {
            string komplex = "";
            int wurf = zufall.Next(1, 101);
            switch (kontext)
            {
                case "":
                    break;
                case "Gesprochene Formel":
                    komplex = ZufaelligerRegulaererKomplex();
                    if (wurf > 45)
                    {
                        komplex = "Ahnenzauber";
                    }
                    if (wurf > 70)
                    {
                        komplex = "Naturruf";
                    }
                    if (wurf > 90)
                    {
                        komplex = "Runenzauber";
                    }
                    if (wurf == 93 || wurf == 94)
                    {
                        komplex = "Sternbild";
                    }
                    if (wurf == 93 || wurf == 94)
                    {
                        komplex = "Bannwort";
                    }
                    if (wurf == 97)
                    {
                        komplex = "Titanenkraft";
                    }
                    if (wurf == 98)
                    {
                        komplex = "Lebenszauber";
                    }
                    else if (wurf == 99)
                    {
                        komplex = "Totenzauber";
                    }
                    else if (wurf == 100)
                    {
                        komplex = "Seelenzauber";
                    }
                    break;
                case "Zauberrolle":
                    komplex = ZufaelligerRegulaererKomplex();
                    if (wurf > 93)
                    {
                        komplex = "Sternbild";
                    }
                    if (wurf == 98)
                    {
                        komplex = "Lebenszauber";
                    }
                    else if (wurf == 99)
                    {
                        komplex = "Totenzauber";
                    }
                    else if (wurf == 100)
                    {
                        komplex = "Seelenzauber";
                    }
                    break;
                case "Zaubertrank":
                    komplex = ZufaelligerRegulaererKomplex();
                    if (wurf == 98)
                    {
                        komplex = "Lebenszauber";
                    }
                    else if (wurf == 99)
                    {
                        komplex = "Totenzauber";
                    }
                    else if (wurf == 100)
                    {
                        komplex = "Seelenzauber";
                    }
                    break;
                case "Zauberstein":
                    komplex = ZufaelligerRegulaererKomplex();
                    if (wurf > 45)
                    {
                        komplex = "Ahnenzauber";
                    }
                    if (wurf > 65)
                    {
                        komplex = "Runenzauber";
                    }
                    if (wurf > 85)
                    {
                        komplex = "Bannwort";
                    }
                    if (wurf == 97 || wurf == 98)
                    {
                        komplex = "Totenzauber";
                    }
                    if (wurf == 99 || wurf == 100)
                    {
                        komplex = "Seelenzauber";
                    }
                    break;
                case "Repertoire eines Zauberers":
                case "Repertoire eines Druiden":
                case "Repertoire eines Priesters":
                case "Repertoire eines Schamanen":
                case "Repertoire eines Runenmeisters":
                case "Repertoire eines Sterndeuters":
                case "Repertoire eines Kultmeisters":
                    komplex = "rep";
                    break;
                default:
                    break;
            }
            return komplex;   
        }

        private List<string> AlleKomplexeNachKontext(string kontext)
        {
            List<string> komplexliste = new List<string>();
            switch (kontext)
            {
                case "":
                    break;
                case "Gesprochene Formel":
                    komplexliste.AddRange(AlleKomplexeEinerArt("Regulär"));
                    komplexliste.Add("Ahnenzauber");
                    komplexliste.Add("Naturruf");
                    komplexliste.Add("Runenzauber");
                    komplexliste.Add("Sternbild");
                    komplexliste.Add("Bannwort");
                    komplexliste.Add("Titanenkraft");
                    komplexliste.Add("Lebenszauber");
                    komplexliste.Add("Totenzauber");
                    komplexliste.Add("Seelenzauber");
                    break;
                case "Zauberrolle":
                    komplexliste.AddRange(AlleKomplexeEinerArt("Regulär"));
                    komplexliste.Add("Sternbild");
                    komplexliste.Add("Lebenszauber");
                    komplexliste.Add("Totenzauber");
                    komplexliste.Add("Seelenzauber");
                    break;
                case "Zaubertrank":
                    komplexliste.AddRange(AlleKomplexeEinerArt("Regulär"));
                    komplexliste.Add("Lebenszauber");
                    komplexliste.Add("Totenzauber");
                    komplexliste.Add("Seelenzauber");
                    break;
                case "Zauberstein":
                    komplexliste.AddRange(AlleKomplexeEinerArt("Regulär"));
                    komplexliste.Add("Ahnenzauber");
                    komplexliste.Add("Runenzauber");
                    komplexliste.Add("Bannwort");
                    komplexliste.Add("Lebenszauber");
                    komplexliste.Add("Totenzauber");
                    komplexliste.Add("Seelenzauber");
                    break;
                case "Repertoire eines Zauberers":
                    komplexliste.AddRange(AlleKomplexeEinerArt("Regulär"));
                    komplexliste.Add("Lebenszauber");
                    komplexliste.Add("Totenzauber");
                    komplexliste.Add("Seelenzauber");
                    break;
                case "Repertoire eines Druiden":
                    komplexliste.Add("Naturruf");
                    komplexliste.Add("Leben");
                    komplexliste.Add("Reinigung");
                    komplexliste.Add("Erkenntnis");
                    komplexliste.Add("Heilung");
                    komplexliste.Add("Pflanzen");
                    komplexliste.Add("Erde");
                    komplexliste.Add("Wasser");
                    komplexliste.Add("Sonne");
                    komplexliste.Add("Wind");
                    komplexliste.Add("Telepathie");
                    break;
                case "Repertoire eines Priesters":
                    komplexliste.AddRange(AlleKomplexeEinerArt("Regulär"));
                    break;
                case "Repertoire eines Schamanen":
                    komplexliste.Add("Ahnenzauber");
                    komplexliste.Add("Telekinese");
                    komplexliste.Add("Telepathie");
                    komplexliste.Add("Erkenntnis");
                    komplexliste.Add("Verhüllung");
                    komplexliste.Add("Erleuchtung");
                    komplexliste.Add("Widerstand");
                    komplexliste.Add("Schutz");
                    komplexliste.Add("Reinigung");
                    komplexliste.Add("Ruhe");
                    komplexliste.Add("Leben");
                    break;
                case "Repertoire eines Runenmeisters":
                    komplexliste.Add("Runenzauber");
                    break;
                case "Repertoire eines Sterndeuters":
                    komplexliste.Add("Sternbild");
                    break;
                case "Repertoire eines Kultmeisters":
                    komplexliste.Add("Bannwort");
                    komplexliste.Add("Titanenkraft");
                    break;
                default:
                    break;
            }
            return komplexliste;
        }

        private void ZauberAbbilden(DataRow ergebniszeile, string komplex)
        {
            dataGridView5.Rows.Add();
            rowcount5++;
            switch (komplex)
            {
                case "Ahnenzauber":
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.PowderBlue;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.Black;
                    break;
                case "Naturruf":
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.DarkGreen;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.Wheat;
                    break;
                case "Runenzauber":
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.BurlyWood;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.DarkRed;
                    break;
                case "Sternbild":
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.MidnightBlue;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.Yellow;
                    break;
                case "Bannwort":
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.PaleVioletRed;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.Black;
                    break;
                case "Titanenkraft":
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.PaleVioletRed;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.Black;
                    break;
                case "Lebenszauber":
                case "Totenzauber":
                case "Seelenzauber":
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.Black;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.GhostWhite;
                    break;
                default: //Regulärer Komplex
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.BackColor = Color.Thistle;
                    dataGridView5.Rows[rowcount5 - 1].DefaultCellStyle.ForeColor = Color.Black;
                    break;
            }

            dataGridView5[0, rowcount5 - 1].Value = ergebniszeile["Art"].ToString();
            dataGridView5[1, rowcount5 - 1].Value = ergebniszeile["Komplex"].ToString();
            dataGridView5[2, rowcount5 - 1].Value = ergebniszeile["Komplexstufe"].ToString();
            dataGridView5[3, rowcount5 - 1].Value = ergebniszeile["Zauber"].ToString();
            dataGridView5[4, rowcount5 - 1].Value = ergebniszeile["Stufe"].ToString();
            dataGridView5[5, rowcount5 - 1].Value = ergebniszeile["Bonusstufen"].ToString();
            dataGridView5[6, rowcount5 - 1].Value = ergebniszeile["Wirkung"].ToString();
        }

        private void button10_Click(object sender, EventArgs e) //1 Zauber generieren
        {
            string komplex = "";
            string gottheit = "";
            string element = "";

            if (label37.Text != "") //Wenn schon das Ergebnis einer Zaubersuche abgebildet ist. => Zurücksetzen
            {
                dataGridView5.Rows.Clear();
                rowcount5 = 0;
            }
            if (comboBox9.GetItemText(comboBox9.SelectedItem) != "(alle)") //wenn Komplex gewählt
            {
                komplex = comboBox9.GetItemText(comboBox9.SelectedItem);
            }
            else //wenn kein Komplex gewählt
            {
                komplex = ZufälligerKomplexnachKontext(comboBox8.GetItemText(comboBox8.SelectedItem)); 
            }
            if (comboBox9.GetItemText(comboBox9.SelectedItem) == "(alle)" && comboBox8.GetItemText(comboBox8.SelectedItem) == "") // wenn weder Komplex noch Kontext gewählt
            {
                komplex = ZufälligerKomplexnachKontext("Gesprochene Formel");
            }
            if (komplex == "rep")
            {
                dataGridView5.Rows.Clear();
                rowcount5 = 0;
                int zufallszahl = 0;
                int zufallszahl2 = 0;
                int zufallszahl3 = 0;
                int id = 0;
                List<int> zauberliste = new List<int>();
                DataRow ergebniszeile = ZauberTB.Rows[0];
                switch (comboBox8.GetItemText(comboBox8.SelectedItem))
                {
                    case "Repertoire eines Zauberers":
                        zufallszahl = zufall.Next(2, 8); //beherrscht 2 bis 7 Komplexe
                        zufallszahl3 = zufall.Next(1, 5);
                        switch (zufallszahl3)
                        {
                            case 1:
                                element = "Feuer";
                                break;
                            case 2:
                                element = "Wasser";
                                break;
                            case 3:
                                element = "Erde";
                                break;
                            case 4:
                                element = "Luft";
                                break;
                        }
                        for (int i = 0; i < zufallszahl; i++)
                        {
                            komplex = ZufaelligerRegulaererKomplex(element);
                            foreach (int idx in AlleZauberEinesKomplexes(komplex))
                            {
                                if (!zauberliste.Contains(idx)) zauberliste.Add(idx);
                            }
                        }
                        break;
                    case "Repertoire eines Druiden":
                        zufallszahl = zufall.Next(4, 10); //beherrscht 4-9 Naturrufe
                        while (zauberliste.Count < zufallszahl)
                        {
                            id = ZufaelligerZauberNachKomplex("Naturruf");
                            if (!zauberliste.Contains(id)) zauberliste.Add(id);
                        }
                        zufallszahl2 = zufall.Next(0, 5); //beherrscht 0-4 Komplexe (nur bestimmte!)
                        for (int i = 0; i < zufallszahl2; i++)
                        {
                            komplex = ZufaelligerRegulaererKomplex("Naturzauberei");
                            foreach (int idx in AlleZauberEinesKomplexes(komplex))
                            {
                                if (!zauberliste.Contains(idx)) zauberliste.Add(idx);
                            }
                        }
                        break;
                    case "Repertoire eines Priesters":
                        gottheit = ZufälligeGottheit();
                        while (zauberliste.Count < 55) //beherrscht alle 11 Komplexe einer bestimmten Gottheit 
                        {
                            komplex = ZufaelligerRegulaererKomplex(gottheit);
                            foreach (int idx in AlleZauberEinesKomplexes(komplex))
                            {
                                if (!zauberliste.Contains(idx)) zauberliste.Add(idx);
                            }
                        }
                        break;
                    case "Repertoire eines Schamanen":
                        zufallszahl = zufall.Next(4, 17); //beherrscht 4-16 Ahnenzauber
                        while (zauberliste.Count < zufallszahl)
                        {
                            id = ZufaelligerZauberNachKomplex("Ahnenzauber");
                            if (!zauberliste.Contains(id)) zauberliste.Add(id);
                        }
                        //2 Komplexe (nur bestimmte!)
                        foreach (int idx in AlleZauberEinesKomplexes(ZufaelligerRegulaererKomplex("Ahnenzauberei1")))
                        {
                            if (!zauberliste.Contains(idx)) zauberliste.Add(idx);
                        }
                        foreach (int idx in AlleZauberEinesKomplexes(ZufaelligerRegulaererKomplex("Ahnenzauberei2")))
                        {
                            if (!zauberliste.Contains(idx)) zauberliste.Add(idx);
                        }
                        break;
                    case "Repertoire eines Runenmeisters":
                        zufallszahl = zufall.Next(15, 25); //beherrscht 15-24 Runenzauber
                        while (zauberliste.Count < zufallszahl)
                        {
                            id = ZufaelligerZauberNachKomplex("Runenzauber");
                            if (!zauberliste.Contains(id)) zauberliste.Add(id);
                        }
                        break;
                    case "Repertoire eines Sterndeuters":
                        zufallszahl = zufall.Next(13, 27); //beherrscht 13-26 Sternbilder (nach Regeln 6W6 + INT)
                        while (zauberliste.Count < zufallszahl)
                        {
                            id = ZufaelligerZauberNachKomplex("Sternbild");
                            if (!zauberliste.Contains(id)) zauberliste.Add(id);
                        }
                        break;
                    case "Repertoire eines Kultmeisters":
                        zufallszahl = zufall.Next(6, 16); //beherrscht 6-15 Bannwörter
                        while (zauberliste.Count < zufallszahl)
                        {
                            id = ZufaelligerZauberNachKomplex("Bannwort");
                            if (!zauberliste.Contains(id)) zauberliste.Add(id);
                        }
                        zufallszahl2 = zufall.Next(0, 3) + zufall.Next(1, 3); //beherrscht 1-4 Titanenkräfte
                        while (zauberliste.Count < zufallszahl + zufallszahl2)
                        {
                            id = ZufaelligerZauberNachKomplex("Titanenkraft");
                            if (!zauberliste.Contains(id)) zauberliste.Add(id);
                        }
                        break;
                }

                for (int a = 0; a < zauberliste.Count; a++)
                {
                    ergebniszeile = ZauberTB.Rows[zauberliste[a]];
                    if (Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) >= numericUpDown6.Value && Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) <= numericUpDown4.Value)
                    {
                        ZauberAbbilden(ergebniszeile, ergebniszeile["Komplex"].ToString());
                    }
                    if (Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) == 0)
                    {
                        ZauberAbbilden(ergebniszeile, ergebniszeile["Komplex"].ToString());
                    }
                }
            }
            else
            {
                DataRow ergebniszeile = ZauberTB.Rows[0];
                ergebniszeile = ZauberTB.Rows[ZufaelligerZauberNachKomplex(komplex)];
                if (numericUpDown6.Value > numericUpDown4.Value || Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) == 0)
                {
                    ZauberAbbilden(ergebniszeile, ergebniszeile["Komplex"].ToString());
                }
                else
                {
                    while (Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) < numericUpDown6.Value || Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) > numericUpDown4.Value)
                    {
                        ergebniszeile = ZauberTB.Rows[ZufaelligerZauberNachKomplex(komplex)];
                    }
                    ZauberAbbilden(ergebniszeile, ergebniszeile["Komplex"].ToString());
                }
            }

            if (comboBox8.GetItemText(comboBox8.SelectedItem) == "Repertoire eines Priesters" && gottheit != "") 
            { label36.Text = "Der Priester dient der Gottheit " + gottheit + "."; }
            else if (comboBox8.GetItemText(comboBox8.SelectedItem) == "Repertoire eines Zauberers" && element != "")
            { label36.Text = "Der Zauberer hat das Element " + element + "."; }
            else { label36.Text = ""; }
            label37.Text = "";
        }

        private void button11_Click(object sender, EventArgs e) //Alle betreffenden Zauber auflisten
        {
            dataGridView5.Rows.Clear();
            rowcount5 = 0;
            string komplex = "";
            List<int> zauberliste = new List<int>();
            List<string> komplexliste = new List<string>();
            List<string> artenliste = new List<string>();
            DataRow ergebniszeile = ZauberTB.Rows[0];

            //Nur den Komplex abbilden
            if (comboBox9.GetItemText(comboBox9.SelectedItem) != "(alle)")
            {
                komplex = comboBox9.GetItemText(comboBox9.SelectedItem);
                zauberliste = AlleZauberEinesKomplexes(komplex);
            }
            //Nur den Kontext abbilden
            else if (comboBox8.GetItemText(comboBox8.SelectedItem) != "")
            {
                komplexliste = AlleKomplexeNachKontext(comboBox8.GetItemText(comboBox8.SelectedItem));
                for (int i = 0; i < komplexliste.Count; i++)
                {
                    zauberliste.AddRange(AlleZauberEinesKomplexes(komplexliste[i]));
                }
            }
            //Alle Zauberformeln abbilden
            else 
            {
                komplexliste = AlleKomplexeNachKontext("Gesprochene Formel");
                for (int i = 0; i < komplexliste.Count; i++)
                {
                    zauberliste.AddRange(AlleZauberEinesKomplexes(komplexliste[i]));
                }
            }

            for (int a = 0; a < zauberliste.Count; a++)
            {
                ergebniszeile = ZauberTB.Rows[zauberliste[a]];
                if (Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) >= numericUpDown6.Value && Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) <= numericUpDown4.Value
                    || Convert.ToDecimal(ergebniszeile["Stufe"].ToString()) == 0)
                {
                    ZauberAbbilden(ergebniszeile, ergebniszeile["Komplex"].ToString());
                }
            }
            label37.Text = rowcount5.ToString();
        }


        private void checkBox6_CheckedChanged(object sender, EventArgs e) //Zeilenumbrüche
        {
            if (checkBox6.Checked)
            {
                dataGridView5.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                dataGridView5.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dataGridView5.Columns[6].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            }
            else
            {
                dataGridView5.Columns[6].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            }
        }

        //#################### ################################################
        //####### NAMEN ###### ################################################
        //#################### ################################################

        private void button13_Click(object sender, EventArgs e) //Namen generieren
        {
            dataGridView7.Rows.Add();
            rowcount7++;
            string gewaehltes_volk = comboBox12.GetItemText(comboBox12.SelectedItem);
            string ergebnisname = "";

            if (radioButton6.Checked) //Weiblich
            {
                if (gewaehltes_volk == "Zwerge (Schwarzalben)")
                {
                    radioButton5.Checked = true;
                    radioButton6.Checked = false;
                }
            }
            else //Männlich
            {
                if (gewaehltes_volk == "Nymphen" || gewaehltes_volk == "Feen (Lichtalben)")
                {
                    radioButton6.Checked = true;
                    radioButton5.Checked = false;
                }
            }

            foreach (DataRow row in NamenTB.Rows)
            {
                if (row["Volk"].ToString() == gewaehltes_volk)
                {
                    string[] array = row["VornameM"].ToString().Split(',');
                    dataGridView7.Rows[rowcount7 - 1].DefaultCellStyle.BackColor = Color.PowderBlue;
                    if (radioButton6.Checked)
                    {
                        array = row["VornameW"].ToString().Split(',');
                        dataGridView7.Rows[rowcount7 - 1].DefaultCellStyle.BackColor = Color.Pink;
                    }
                    ergebnisname = array[zufall.Next(0, array.Length)];
                }
            }
            if (ergebnisname.StartsWith(" "))
            {
                ergebnisname = ergebnisname.TrimStart(' ');
            }

            dataGridView7[0, rowcount7 - 1].Value = ergebnisname;
            dataGridView7[1, rowcount7 - 1].Value = gewaehltes_volk;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            dataGridView7.Rows.Clear();
            rowcount7 = 0;
        }

        //#################### ################################################
        //###### SCHÄTZE ##### ################################################
        //#################### ################################################

        private DataRow ZufaelligeWaffe()
        {
            int row = 0;
            int rowwaffe = 10;
            double wert = 0;
            string material = "";
            DataRow waffenzeile = WaffentypenTB.Rows[rowwaffe];
            DataRow typenzeile = WaffentypenTB.Rows[row];
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            int wuerfel = zufall.Next(1, 100);
            
            while(wuerfel > Convert.ToInt16(typenzeile["Ergebnis"].ToString()) && Convert.ToInt16(typenzeile["Ergebnis"].ToString()) < 101)
            {
                typenzeile = WaffentypenTB.Rows[row];
                row++;
            }

            if (row == 0) row++;
            wuerfel = zufall.Next(1, 20);

            while (wuerfel + (1000*row) > Convert.ToInt16(waffenzeile["Ergebnis"].ToString()) && Convert.ToInt16(waffenzeile["Ergebnis"].ToString()) < (1000 * row) + 25)
            {
                waffenzeile = WaffentypenTB.Rows[rowwaffe];
                rowwaffe++;
            }

            wert = Convert.ToDouble(waffenzeile["Wert"].ToString());
            wuerfel = zufall.Next(1, 100);
            material = "";

            switch (wuerfel)
            {
                case 76:
                case 77:
                case 78:
                case 79:
                case 80:
                case 81:
                case 82:
                case 83:
                case 84:
                case 85:
                case 86:
                case 87:
                case 88:
                case 89:
                case 90:
                    material = "Bronze";
                    wert *= 2;
                    break;
                case 91:
                    material = "Adamant";
                    wert *= 250;
                    break;
                case 92:
                    material = "Asterium";
                    wert *= 425;
                    break;
                case 93:
                    material = "Kupfer";
                    break;
                case 94:
                case 95:
                case 96:
                case 97:
                    material = "Obsidian";
                    wert *= 100;
                    break;
                case 98:
                case 99:
                    material = "Gold";
                    wert *= 625;
                    break;
                case 100:
                    material = "Holz";
                    wert *= 0.1;
                    break;
                default:
                    material = "Eisen";
                    wert *= 1;
                    break;
            }

            if (row == 7 || row == 8 || row == 9)
            {
                material = "Holz";
                wert = Convert.ToDouble(waffenzeile["Wert"].ToString());
            }

            dummyergebnis["Beschreibung"] = "Material: " + material + ".";
            dummyergebnis["Wert"] = wert.ToString();
            dummyergebnis["Name"] = waffenzeile["Name"].ToString();

            return dummyergebnis;
        }

        private DataRow ZufaelligeRuestung()
        {
            int rowruestung = 0;
            double wert = 0;
            string material = "";
            DataRow ruestungszeile = RuestungstypenTB.Rows[rowruestung];
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            int wuerfel = zufall.Next(1, 100);
            string beschreibung = "";

            while (wuerfel > Convert.ToInt16(ruestungszeile["Ergebnis"].ToString()) && Convert.ToInt16(ruestungszeile["Ergebnis"].ToString()) < 101)
            {
                ruestungszeile = RuestungstypenTB.Rows[rowruestung];
                rowruestung++;
            }

            if (rowruestung == 0)
            {
                rowruestung += 13; //springe zu Rumpfrüstungen
                wuerfel = zufall.Next(1, 100);

                while (wuerfel + 1000 > Convert.ToInt16(ruestungszeile["Ergebnis"].ToString()))
                {
                    ruestungszeile = RuestungstypenTB.Rows[rowruestung];
                    rowruestung++;
                }
            }

            wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
            wuerfel = zufall.Next(1, 100);
            material = "";

            switch (wuerfel)
            {
                case 76:
                case 77:
                case 78:
                case 79:
                case 80:
                case 81:
                case 82:
                case 83:
                case 84:
                case 85:
                case 86:
                case 87:
                case 88:
                case 89:
                case 90:
                    material = "Bronze";
                    wert *= 2;
                    break;
                case 91:
                    material = "Adamant";
                    wert *= 250;
                    break;
                case 92:
                    material = "Asterium";
                    wert *= 425;
                    break;
                case 93:
                case 94:
                case 95:
                case 96:
                case 97:
                    material = "Bronze";
                    wert *= 2;
                    break;
                case 98:
                case 99:
                    material = "Drachenschuppen";
                    wert *= 200;
                    break;
                case 100:
                    material = "Chitin";
                    wert *= 20;
                    break;
                default:
                    material = "Eisen";
                    break;
            }

            switch (rowruestung)
            {
                case 26:
                case 27: //Thorax
                    if (zufall.Next(1, 2) == 1) material = "Leder";
                    else material = "Leinen";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    break;
                case 28:
                case 29:
                case 30:
                case 31: //Gambeson
                    material = "Leinen";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    break;
                case 19: 
                case 20:
                case 21:
                case 22: //Schuppenrüstung
                    if (zufall.Next(1, 2) == 1) material = "Leder";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    break;
                case 23: //Brustharnisch
                    if (zufall.Next(1, 3) == 1) material = "Leder";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    break;
                case 24: //Lamellen
                case 25: 
                    if (zufall.Next(1,2) == 1) material = "Leder";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    break;
                default:
                    break;
            }

            beschreibung = ruestungszeile["Beschreibung"].ToString();

            if (beschreibung == "Verstärkung:")
            {
                wuerfel = zufall.Next(1, 10);
                switch (wuerfel)
                {
                    case 1:
                    case 2:
                        beschreibung += " Asteriumschuppen.";
                        wert += 25000;
                        break;
                    case 3:
                        beschreibung += " Eisenschuppen.";
                        wert += 0;
                        break;
                    case 4:
                        beschreibung += " Bronzeschuppen.";
                        wert += 180;
                        break;
                    case 5:
                        beschreibung += " Adamantschuppen.";
                        wert += 30000;
                        break;
                    case 6:
                        beschreibung += " Chitin.";
                        wert += 0;
                        break;
                    case 7:
                        beschreibung += " Riesenknochenplatten.";
                        wert += 300;
                        break;
                    case 8:
                        beschreibung += " Lederschuppen.";
                        wert += -40;
                        break;
                    case 9:
                        beschreibung += " Leinenschuppen.";
                        wert += -40;
                        break;
                    case 10:
                        beschreibung += " Drachenschuppen.";
                        wert += 1000;
                        break;
                    default:
                        beschreibung += " Eisenschuppen.";
                        wert += 0;
                        break;
                }
                beschreibung += " Kernmaterial: " + material + ".";
            }
            else if (beschreibung == "Metall")
            {
                if (zufall.Next(1, 2) == 1) beschreibung = "Material: Eisen.";
                else if (zufall.Next(1, 10) == 1) 
                { 
                    beschreibung = "Material: Asterium.";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    wert *= 425;
                }
                else if (zufall.Next(1, 10) == 1) 
                { 
                    beschreibung = "Material: Adamant.";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    wert *= 250;
                }
                else
                {
                    beschreibung = "Material: Bronze.";
                    wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
                    wert *= 2;
                }
            }
            else if (beschreibung == "Material: Leinen." || beschreibung == "Material: Leder.")
            {
                wert = Convert.ToDouble(ruestungszeile["Wert"].ToString());
            }
            else
            {
                beschreibung = "Material: " + material + ".";
            }

            dummyergebnis["Beschreibung"] = beschreibung;
            dummyergebnis["Wert"] = wert.ToString();
            dummyergebnis["Name"] = ruestungszeile["Name"].ToString();

            return dummyergebnis;
        }

        private DataRow ZufaelligesMusikinstrument()
        {
            DataRow musikzeile = SchatzAlltagMusikTB.Rows[zufall.Next(100, 119)];
            return musikzeile;
        }

        private DataRow ZufaelligerAlltagsgegenstand()
        {
            DataRow alltagszeile = SchatzAlltagMusikTB.Rows[zufall.Next(0, 99)];

            if (keinezauberei)
            {
                while (alltagszeile["Zauberei"].ToString() == "JA")
                {
                    alltagszeile = SchatzAlltagMusikTB.Rows[zufall.Next(0, 99)];
                }
            }
            return alltagszeile;
        }

        private DataRow ZufaelligerSpezialGegenstand(string typ)
        {

            int row = 0;
            double wert = 0;
            string beschreibung = "";
            DataRow zeile = SpezialgegenstandTB.Rows[row];
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            int wuerfel = zufall.Next(1, 20);
            int wuerfel2 = zufall.Next(1, 20);
            if (typ == "Urkriegsgegenstand" || typ == "Schutzamulett") wuerfel = zufall.Next(1, 100);

            while (wuerfel > Convert.ToInt16(zeile["Ergebnis"].ToString()) || zeile["Typ"].ToString() != typ)
            {
                zeile = SpezialgegenstandTB.Rows[row];
                row++;
            }

            dummyergebnis["Beschreibung"] = zeile["Beschreibung"].ToString();
            dummyergebnis["Wert"] = zeile["Wert"].ToString();

            if (typ == "Edelmetall")
            {
                wuerfel = zufall.Next(1, 20);
                wert = 0;
                wert = Convert.ToDouble(zeile["Wert"].ToString());
                switch (wuerfel)
                {
                    case 1:
                        beschreibung = "Ein Talent (50-Pfund-Barren).";
                        wert *= 50;
                        break;
                    case 2:
                    case 3:
                    case 4:
                        beschreibung = "Ein 5-Pfund-Barren.";
                        wert *= 5;
                        break;
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 9:
                        wuerfel2 = zufall.Next(1, 10);
                        beschreibung = wuerfel2.ToString() + " kleine Barren (je 0,5 Pfund).";
                        wert *= 0.5 * wuerfel2;
                        break;
                    case 10:
                    case 11:
                    case 12:
                        wuerfel2 = zufall.Next(1, 20);
                        beschreibung = wuerfel2.ToString() + " kleine Nuggets (je 0,2 Pfund).";
                        wert *= 0.2 * wuerfel2;
                        break;
                    case 13:
                    case 14:
                    case 15:
                    case 16:
                        wuerfel2 = zufall.Next(1, 20);
                        beschreibung = wuerfel2.ToString() + " kleine Bruchstücke (je 0,2 Pfund).";
                        wert *= 0.2 * wuerfel2;
                        break;
                    case 17:
                    case 18:
                    case 19:
                    case 20:
                        wuerfel2 = zufall.Next(1, 6);
                        beschreibung = wuerfel2.ToString() + " Pfund schwerer, unreiner (unverhütteter) Erzbrocken.";
                        wert *= 0.5 * wuerfel2;
                        break;
                }
                dummyergebnis["Beschreibung"] = beschreibung;
                dummyergebnis["Wert"] = wert.ToString();
            }

            dummyergebnis["Name"] = zeile["Name"].ToString();

            return dummyergebnis;
        }

        private DataRow ZufaelligeBannwaffe()
        {
            int row = 0;
            double wert = 0;
            DataRow zeile = BannwaffeTB.Rows[row];
            DataRow waffe = ZufaelligeWaffe();
            DataRow effekt = BannwaffeTB.Rows[26];
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            int wuerfel = zufall.Next(1, 100);

            while (wuerfel > Convert.ToInt16(zeile["Ergebnis"].ToString()) && Convert.ToInt16(zeile["Ergebnis"].ToString()) < 101)
            {
                zeile = BannwaffeTB.Rows[row];
                row++;
            }

            row = 26;
            while (wuerfel + 1000 > Convert.ToInt16(effekt["Ergebnis"].ToString()))
            {
                effekt = BannwaffeTB.Rows[row];
                row++;
            }

            wert = Convert.ToDouble(waffe["Wert"].ToString()) + 500;

            dummyergebnis["Beschreibung"] = "Bannwaffe: " + waffe["Name"].ToString() + " gegen " + zeile["Beschreibung"].ToString() + ". Effekt: " + effekt["Name"].ToString();
            dummyergebnis["Wert"] = wert.ToString();
            dummyergebnis["Name"] = zeile["Name"].ToString();

            return dummyergebnis;
        }

        private DataRow ZufaelligerKomplexring()
        {
            int row = 0;
            double wert = 0;
            DataRow zeile = KomplexringTB.Rows[row];
            DataRow ring = ZufaelligesSchmuckstueck(false, "Ring");
            string zauberkomplex = ZufaelligerRegulaererKomplex();
            DataRow wirkweise = KomplexringTB.Rows[5];
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            int wuerfel = zufall.Next(1, 20);

            while (wuerfel > Convert.ToInt16(zeile["Ergebnis"].ToString()) && Convert.ToInt16(zeile["Ergebnis"].ToString()) < 101)
            {
                zeile = KomplexringTB.Rows[row];
                row++;
            }

            row = 5;
            wuerfel = zufall.Next(1, 20);
            while (wuerfel + 1000 > Convert.ToInt16(wirkweise["Ergebnis"].ToString()))
            {
                wirkweise = KomplexringTB.Rows[row];
                row++;
            }

            wert = Convert.ToDouble(zeile["Wert"].ToString()) * 100; 
            wert += Convert.ToDouble(ring["Wert"].ToString()); 

            dummyergebnis["Beschreibung"] = ring["Name"].ToString() + " Der Träger des Ringes kann " + zeile["Beschreibung"].ToString() + " " + zauberkomplex + " wirken. " + wirkweise["Beschreibung"].ToString();
            dummyergebnis["Wert"] = wert.ToString();
            dummyergebnis["Name"] = "Zauberring: " + zauberkomplex;

            return dummyergebnis;
        }

        private DataRow ZufaelligesGeschoss()
        {
            DataRow zeile = ZufaelligerSpezialGegenstand("Geschosse");
            DataRow dummyergebnis = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            double wert = Convert.ToDouble(zeile["Wert"].ToString());
            string beschreibung = zeile["Beschreibung"].ToString();
            string name = zeile["Name"].ToString();
            string material = "Eisen";
            int anzahl = zufall.Next(1, 20) + zufall.Next(1, 20);
            wert *= anzahl;

            if (beschreibung == "Material")
            {
                int wuerfel = zufall.Next(1, 2);
                if (wuerfel == 1)
                {
                    material = "Eisen";
                    wert *= 1;
                }
                else
                    {
                        wuerfel = zufall.Next(1, 10);
                        switch (wuerfel)
                        {
                            case 1:
                            case 2:
                                material = "Bronze";
                                wert *= 2;
                                break;
                            case 3:
                                material = "Adamant";
                                wert *= 250;
                                break;
                            case 4:
                                material = "Asterium";
                                wert *= 425;
                            break;
                            case 5:
                                material = "Kupfer";
                                wert *= 1;
                                break;
                            case 6:
                                material = "Obsidian";
                                wert *= 100;
                                break;
                            case 7:
                            case 8:
                                material = "Feuerstein";
                                wert *= 0.2;
                                break;
                            case 9:
                            case 10:
                                material = "Knochen";
                                wert *= 0.2;
                            break;
                        }
                    }
                
            }

            dummyergebnis["Beschreibung"] = anzahl.ToString() + " " + name + " aus Holz mit Spitzen aus " + material + ".";
            dummyergebnis["Wert"] = wert.ToString();
            dummyergebnis["Name"] = zeile["Name"].ToString();

            return dummyergebnis;
        }

        private DataRow ZufaelligerGegenstand()
        {
            DataRow ergebnisgegenstand = SchatzGegenstandTB.Rows[zufall.Next(0, SchatzGegenstandTB.Rows.Count - 1)];

            switch (ergebnisgegenstand["Name"].ToString())
            {
                case "Waffe":
                    ergebnisgegenstand = ZufaelligeWaffe();
                    hintergrundfarbe = Color.Gray;
                    break;
                case "Rüstung":
                    ergebnisgegenstand = ZufaelligeRuestung();
                    hintergrundfarbe = Color.DimGray;
                    break;
                case "Musikinstrument":
                    ergebnisgegenstand = ZufaelligesMusikinstrument();
                    hintergrundfarbe = Color.Yellow;
                    break;
                case "Alltagsgegenstand":
                    ergebnisgegenstand = ZufaelligerAlltagsgegenstand();
                    hintergrundfarbe = Color.BurlyWood;
                    break;
                case "Schmuckstück":
                    ergebnisgegenstand = ZufaelligesSchmuckstueck(false, "");
                    hintergrundfarbe = Color.Goldenrod;
                    break;
                default:
                    hintergrundfarbe = Color.BurlyWood;
                    break;
            }

            return ergebnisgegenstand;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DataRow kategoriezeile = SchatzTB.Rows[0];
            DataRow ergebniszeile = SchatzGegenstandTB.Rows[100]; // dummyzeile.
            string ergebnisname = "(Kein Name)";
            string ergebniswert = "(Kein Wert)";
            string ergebniswirkung = "(Keine Wirkung)";
            string kategorie = "";
            string ziertabellen = "NEIN";
            if (checkBox7.Checked) keinlebewesen = true;
            if (checkBox8.Checked) keinezauberei = true;


            for (int i = 0; i < numericUpDown8.Value; i++)
            {
                dataGridView8.Rows.Add();
                rowcount8++;
                ergebnisname = "(Kein Name)";
                ergebniswert = "(Kein Wert)";
                ergebniswirkung = "(Keine Wirkung)";

                kategoriezeile = SchatzTB.Rows[zufall.Next(0, SchatzTB.Rows.Count - 1)];
                while (kategoriezeile["Zauberei"].ToString() == "JA" && keinezauberei
                    || kategoriezeile["Lebendig"].ToString() == "JA" && keinlebewesen)
                {
                    kategoriezeile = SchatzTB.Rows[zufall.Next(0, SchatzTB.Rows.Count -1)];
                }
               
                kategorie = kategoriezeile["Name"].ToString();
                ziertabellen = kategoriezeile["Ziertabellen"].ToString();

                switch (kategorie)
                {
                    case "Gegenstand":
                        ergebniszeile = ZufaelligerGegenstand();
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        break;
                    case "Komplexring":
                        ergebniszeile = ZufaelligerKomplexring(); 
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.DodgerBlue;
                        break;
                    case "Bannwaffe":
                        ergebniszeile = ZufaelligeBannwaffe();
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.OrangeRed;
                        break;
                    case "Schutzamulett":
                        ergebniszeile = ZufaelligerSpezialGegenstand("Schutzamulett");
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.Plum;
                        break;
                    case "Edelmetall":
                        ergebniszeile = ZufaelligerSpezialGegenstand("Edelmetall");
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.LightBlue;
                        break;
                    case "Zaubernahrung":
                        ergebniszeile = ZufaelligerSpezialGegenstand("Zaubernahrung");
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.Tomato;
                        break;
                    case "Urkriegsgegenstand":
                        ergebniszeile = ZufaelligerSpezialGegenstand("Urkriegsgegenstand");
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.LightSlateGray;
                        break;
                    case "Schriftstück":
                        ergebniszeile = ZufaelligerSpezialGegenstand("Schriftstück");
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.Moccasin;
                        break;
                    case "Geschosse":
                        ergebniszeile = ZufaelligesGeschoss();
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        ziertabellen = "JA";
                        hintergrundfarbe = Color.LightGray;
                        break;
                    case "Schmuckstein":
                        ergebniszeile = ZufaelligerRohstein(false);
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.LimeGreen;
                        break;
                    case "Zauberstein":
                        ergebniszeile = ZaubersteinGenerator(false);
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.DarkViolet;
                        break;
                    case "Zauberrolle":
                        ergebniszeile = ZauberTrankRolleGenerator("rolle", false); 
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.Purple;
                        break;
                    case "Zaubertrank":
                        ergebniszeile = ZauberTrankRolleGenerator("trank", false);
                        ergebnisname = ergebniszeile["Name"].ToString();
                        ergebniswert = ergebniszeile["Wert"].ToString();
                        ergebniswirkung = ergebniszeile["Beschreibung"].ToString();
                        hintergrundfarbe = Color.MediumTurquoise;
                        break;
                    case "Münzen":
                        ergebnisname = "Silbermünzen";
                        ergebniswert = zufall.Next(1, 100).ToString();
                        int waehrungswurf = zufall.Next(1, 10);
                        switch (waehrungswurf)
                        {
                            case 1:
                            case 2:
                            case 3:
                                ergebniswirkung = "Drachmen";
                                break;
                            case 4:
                            case 5:
                            case 6:
                                ergebniswirkung = "Denare";
                                break;
                            case 7:
                            case 8:
                                ergebniswirkung = "Ungeprägt";
                                break;
                            case 9:
                                ergebniswirkung = "Loi (Hochelfisch)";
                                break;
                            case 10:
                                ergebniswirkung = "Fremdartige Prägung";
                                break;
                        }
                        hintergrundfarbe = Color.Silver;
                        break;
                    case "Ringgeld":
                        int ringwurf = zufall.Next(1, 2);
                        switch (ringwurf)
                        {
                            case 1:
                                ergebnisname = "Goldener Armring";
                                ergebniswert = "1000";
                                ergebniswirkung = "Ringgeld";
                                break;
                            case 2:
                                int ringgeldwurf = zufall.Next(5, 20);
                                ergebnisname = (ringgeldwurf).ToString() + " silberne Finggerringe";
                                ergebniswert = (ringgeldwurf * 10).ToString();
                                ergebniswirkung = "Ringgeld";
                                break;
                        }
                        hintergrundfarbe = Color.Gold;
                        break;
                    case "Alkohol":
                        int alkwurf = zufall.Next(1, 4);
                        switch (alkwurf)
                        {
                            case 1:
                                ergebnisname = "Flasche exquisiter Wein";
                                ergebniswert = "11";
                                ergebniswirkung = "1 Kornmaß (Liter)";
                                break;
                            case 2:
                                ergebnisname = "Fass exquisiter Met";
                                ergebniswert = "150";
                                ergebniswirkung = "2 Amphoren (54 Liter)";
                                break;
                            case 3:
                                ergebnisname = "Flasche teurer Schnaps";
                                ergebniswert = "7";
                                ergebniswirkung = "Halbes Kornmaß (Liter)";
                                break;
                            case 4:
                                ergebnisname = "Amphore exquisiter Wein";
                                ergebniswert = "300";
                                ergebniswirkung = "(27 Liter)";
                                break;
                        }
                        hintergrundfarbe = Color.Crimson;
                        break;
                    case "Pferd":
                        int pferdewurf = zufall.Next(1, 4);
                        switch (pferdewurf)
                        {
                            case 1:
                                ergebnisname = "Junges Packpferd";
                                ergebniswert = "600";
                                ergebniswirkung = "Gewicht: 700 Pfund. Tragkraft: 350 Pfund. Bewegung: 90. Kampfgeist: 2";
                                break;
                            case 2:
                                ergebnisname = "Junges Reitpferd";
                                ergebniswert = "750";
                                ergebniswirkung = "Gewicht: 900 Pfund. Tragkraft: 300 Pfund. Bewegung: 120. Kampfgeist: 3";
                                break;
                            case 3:
                                ergebnisname = "Ausgezeichnetes Reitpferd";
                                ergebniswert = "1800";
                                ergebniswirkung = "Gewicht: 1000 Pfund. Tragkraft: 400 Pfund. Bewegung: 140. Kampfgeist: 3";
                                break;
                            case 4:
                                ergebnisname = "Kriegspferd";
                                ergebniswert = "2500";
                                ergebniswirkung = "Gewicht: 1100 Pfund. Tragkraft: 400 Pfund. Bewegung: 120. Kampfgeist: 4";
                                break;
                        }
                        hintergrundfarbe = Color.SaddleBrown;
                        break;
                    case "Rind":
                        int rinderwurf = zufall.Next(1, 4);
                        switch (rinderwurf)
                        {
                            case 1:
                                ergebnisname = "Stattlicher Ochse";
                                ergebniswert = "250";
                                ergebniswirkung = "";
                                break;
                            case 2:
                                ergebnisname = "Stattliche Kuh";
                                ergebniswert = "200";
                                ergebniswirkung = "";
                                break;
                            case 3:
                                ergebnisname = "Rinderherde";
                                ergebniswert = "5000";
                                ergebniswirkung = "25 Rinder";
                                break;
                            case 4:
                                ergebnisname = "Hekatombe";
                                ergebniswert = "20000";
                                ergebniswirkung = "100 Rinder";
                                break;
                        }
                        hintergrundfarbe = Color.Cornsilk;
                        break;
                    case "Sklave":
                        int sklavenwurf = zufall.Next(1, 4);
                        switch (sklavenwurf)
                        {
                            case 1:
                                ergebnisname = "Junger ungelernter Sklave";
                                ergebniswert = "600";
                                ergebniswirkung = "";
                                break;
                            case 2:
                                ergebnisname = "Junge ungelernte Sklavin";
                                ergebniswert = "800";
                                ergebniswirkung = "";
                                break;
                            case 3:
                                ergebnisname = "Gelernter Sklave";
                                ergebniswert = "850";
                                ergebniswirkung = "Beherrscht 3 Fähigkeiten/Handwerke sehr gut.";
                                break;
                            case 4:
                                ergebnisname = "Gelernte Sklavin";
                                ergebniswert = "1150";
                                ergebniswirkung = "Beherrscht 3 Fähigkeiten/Handwerke sehr gut.";
                                break;
                        }
                        hintergrundfarbe = Color.MistyRose;
                        break;
                    case "Falke":
                        ergebnisname = "Abgerichteter Falke";
                        ergebniswert = "160";
                        ergebniswirkung = "Kann zur Beizjagd verwendet werden.";
                        hintergrundfarbe = Color.Tan;
                        break;
                    case "Hund":
                        int hundewurf = zufall.Next(1, 3);
                        switch (hundewurf)
                        {
                            case 1:
                                ergebnisname = "Wachhund";
                                ergebniswert = "40";
                                ergebniswirkung = "Weder für Kampf, noch für Jagd gut geeignet. Loyal.";
                                break;
                            case 2:
                                ergebnisname = "Jagdhund";
                                ergebniswert = "160";
                                ergebniswirkung = "Gut für die Jagd aber kaum für den Kampf geeignet. Loyal.";
                                break;
                            case 3:
                                ergebnisname = "Kampfhund";
                                ergebniswert = "200";
                                ergebniswirkung = "Gut für den Kampf aber kaum für die Jagd geeignet. Loyal.";
                                break;
                        }
                        hintergrundfarbe = Color.Peru;
                        break;
                    default:
                        ergebnisname = "Nichts";
                        hintergrundfarbe = Color.White;
                        break;
                }

                if (ziertabellen == "JA")
                {
                    double wert = Convert.ToDouble(ergebniswert);
                    if (ergebniswirkung != "") ergebniswirkung += " ";

                    //Zier
                    int zierwurf = zufall.Next(1, 20);
                    if (zierwurf < 5)
                    {
                        ergebniswirkung += "Verziert. ";
                        wert += 10;
                    }

                    zierwurf = zufall.Next(1, 20);
                    if (zierwurf < 9) ergebniswirkung += "Kulturelle Eigenart. ";

                    zierwurf = zufall.Next(1, 20);
                    if (zierwurf < 4)
                    {
                        ergebniswirkung += "Von einem Helden/König. ";
                        wert += 50;
                    }
                    zierwurf = zufall.Next(1, 20);
                    if (zierwurf < 4)
                    {
                        ergebniswirkung += "Mit Goldbestandteilen. ";
                        wert += 100;
                    }
                    zierwurf = zufall.Next(1, 20);
                    if (zierwurf < 4)
                    {
                        ergebniswirkung += "Enthält Bestandteile einer Jagdtrophäe. ";
                        wert += 20;
                    }

                    //Qualität
                    zierwurf = zufall.Next(1, 20);
                    switch (zierwurf)
                    {
                        case 1:
                        case 2:
                        case 3:
                        case 4:
                            ergebniswirkung += "Geringe Qualität (-1 auf Tests). ";
                            wert *= 0.5;
                            break;
                        case 13:
                        case 14:
                        case 15:
                        case 16:
                        case 17:
                            ergebniswirkung += "Hohe Qualität (+1 auf Tests). ";
                            wert *= 1.5;
                            break;
                        case 18:
                        case 19:
                        case 20:
                            ergebniswirkung += "Herausragende Qualität (+3 auf Tests). ";
                            wert *= 3;
                            break;
                    }


                    //Inschriften
                    zierwurf = zufall.Next(1, 20);
                    switch (zierwurf)
                    {
                        case 1:
                            ergebniswirkung += "Mit seltsamen Zauberzeichen. ";
                            break;
                        case 2:
                        case 3:
                            ergebniswirkung += "Mit Gravur. ";
                            break;
                        case 4:
                        case 5:
                        case 6:
                            ergebniswirkung += "Mit unlesbarer Gravur. ";
                            break;
                        case 7:
                            ergebniswirkung += "Mit Weiheinschrift. ";
                            break;
                        case 8:
                        case 9:
                        case 10:
                            ergebniswirkung += "Mit Namensinschrift. ";
                            break;
                    }

                    //Schmucksteine
                    zierwurf = zufall.Next(1, 20);
                    switch (zierwurf)
                    {
                        case 1:
                        case 2:
                        case 3:
                        case 4:
                            DataRow schmucksteinzeile = ZufaelligerRohstein(false);
                            wert += Convert.ToDouble(schmucksteinzeile["Wert"].ToString());
                            ergebniswirkung += "Mit Schmuckstein: " + schmucksteinzeile["Name"].ToString() + " " + schmucksteinzeile["Beschreibung"].ToString() + " ";
                            break;
                        case 5:
                        case 6:
                        case 7:
                            if (!keinezauberei)
                            {
                                DataRow zauberzeile = ZaubersteinGenerator(false);
                                wert += Convert.ToDouble(zauberzeile["Wert"].ToString());
                                ergebniswirkung += "Mit Zauberstein: " + zauberzeile["Name"].ToString() + " " + zauberzeile["Beschreibung"].ToString() + " ";
                            }
                            break;
                    }

                    //Artefakt
                    if (!keinezauberei)
                    {
                        zierwurf = zufall.Next(1, 20);
                        if (zierwurf > 3)
                        { 
                            ergebniswirkung += "Verzaubert! Siehe Abschnitt 11.3 im Regelwerk.";
                            //Zufälliger Zauber
                            wert += 500;
                        }
                    }

                    ergebniswert = wert.ToString();
                }

                dataGridView8[0, rowcount8 - 1].Value = ergebnisname;
                dataGridView8[1, rowcount8 - 1].Value = ergebniswert;
                dataGridView8[2, rowcount8 - 1].Value = ergebniswirkung;
                dataGridView8[0, rowcount8 - 1].Style.BackColor = hintergrundfarbe;
                hintergrundfarbe = Color.White;
            }


            
        }

        private void button16_Click(object sender, EventArgs e)
        {
            dataGridView8.Rows.Clear();
            rowcount8 = 0;
        }

        //#################### ################################################
        //##### EXPORTE ###### ################################################
        //#################### ################################################

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

                for (int d = 1; d <= RowCount - 1; d++)
                {
                    oDoc.Application.Selection.Tables[1].Rows[d].Range.Font.Size = 8;
                    oDoc.Application.Selection.Tables[1].Rows[d].Range.Font.Name = "Book Antiqua";
                    oDoc.Application.Selection.Tables[1].Rows[d].Range.Bold = 0;
                }

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 10;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                //oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Borders.Enable = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "Tüfteltruhe-Export";
                    headerRange.Font.Size = 12;
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
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "Tüfteltruhe-Export Warenangebot.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView4, sfd.FileName);
            }
        }

        private void ZauberlisteExp(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "Tüfteltruhe-Export Zauberliste.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView5, sfd.FileName);
            }
        }

        private void NamenslisteExp(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "Tüfteltruhe-Export Namensliste.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                Export_Data_To_Word(dataGridView7, sfd.FileName);
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Export von " + sfd.FileName + " erfolgreich"); 
            }
        }

        //#################### ################################################
        //### NEUE FENSTER ### ################################################
        //#################### ################################################


        private void neuerSpielermodusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spielermodus spielermodus = new Spielermodus();
            spielermodus.Show();
        }

        private void neuerSLModusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spielleitermodus spielleitermodus = new Spielleitermodus();
            spielleitermodus.Show();
        }
    }
}
