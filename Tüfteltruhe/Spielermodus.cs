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

namespace Tüfteltruhe
{ 
    //+++Tooltips
    //+++Seefahrt: Wetter vorher bestimmen, dann Fahrweise ausrechnen
    public partial class Spielermodus : Form
    {
        //Waffentestrechner
        public DataTable WaffenTB = new DataTable();
        public Waffe Waffenwahl = new Waffe("", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", 0, 0, 0, 0, 0);
        public int stichangriff_ergebnis = -100;
        public int wuchtangriff_ergebnis = -100;
        public int schnittangriff_ergebnis = -100;
        public int schuss_ergebnis = -100;
        public int parieren_ergebnis = -100;
        public int blocken_ergebnis = -100;
        public int werfen_ergebnis = -100;
        public int oeffnen_ergebnis = -100;
        public int ausschalten_ergebnis = -100;
        public string gewaehlte_waffe = "";
        public string gewaehlte_gg = "";
        public int rowcount = 0;
        public int nachteil = 0;
        public decimal htk;
        public int handzahl;
        public int fixwert1 = -5; //Waffenfähigkeit
        public int fixwert2 = -1; //Griffgewöhnungs-Abzug
        public int fixwert3 = 0; //Gewichtsabzug

        //Reisetestrechner
        public int rowcount2 = 0;
        Random zufall = new Random();
        public int wuerfelgesamt;
        public int reisetest_ergebnis;
        public int reisegeschwindigkeit;
        public string gewaehlte_reiseart;
        public string gewaehltes_terrain;
        public int fortbewegungsmod = 0;
        public int reisenachteil = 0;
        public int zielwert = 8;
        public int tag = 1;
        public int resultat;
        public int ausdauerverlust = 0;
        public int willenskraftverlust = 0;
        public int extremfehl = 0;
        public int extremerfolg = 0;
        public int orientierungszw = 0;
        public int fortschrittdurchorientierung = 0;
        public int fortschrittausgegeben = 0;
        public int orientierungswurf;

        //Seereise-Rechner
        public DataTable SchiffstypenTB = new DataTable();
        public int rowcount3 = 0;
        public decimal knoten;
        public string wetter = "";
        public string fahrweise;
        public string seereiseumgebung;
        public int schwierigkeit;
        public decimal seefahrtsfähigkeit;
        public decimal seefahrtswurf;
        public decimal effektiveruderer;
        public decimal effektivesegler;

        public Spielermodus()
        {
            InitializeComponent();
        }

        // ###### Waffentest ######
        public void button1_Click(object sender, EventArgs e)
        {

            //Ergebnistabelle initialisieren
            dataGridView1.Rows.Clear();
            dataGridView1.ReadOnly = false;
            rowcount = 0;
            nachteil = 0;

            //Werte ermitteln
            fixwert1 = (int)numericUpDown1.Value; //Waffenfähigkeit
            
            gewaehlte_gg = comboBox1.GetItemText(comboBox1.SelectedItem); //Griffgewöhnung
            switch (gewaehlte_gg) //Griffgewöhnung berechnen
            {
                case "A":
                case "B":
                    fixwert2 = -1;
                    break;
                case "C":
                case "D":
                    fixwert2 = 0;
                    break;
                case "E":
                case "F":
                    fixwert2 = 1;
                    break;
                case "G":
                case "H":
                    fixwert2 = 2;
                    break;
            }

            
            if (!checkBox1.Checked) //Warn-Hinweise bei SP SC LH und SP, wenn nicht zweihändig verwendet
            { switch (Waffenwahl.waffentyp)
                {
                    case "SP":
                    case "LH":
                        MessageBox.Show("Spieße und Lange Hiebwaffen sollten zweihändig verwendet werden, da sonst Nachteil (ab 4 Pfund Doppel-Nachteil) besteht.");
                        if (Waffenwahl.gewicht < 4) nachteil++;
                        else { nachteil += 2; }
                        break;
                    case "SW":
                    case "SD":
                        MessageBox.Show("Spannwaffen und Schleuderwaffen müssen zweihändig verwendet werden!");
                        checkBox1.Checked = true;
                        break;
                }
            }
            htk = 3 + (numericUpDown2.Value / 2); //HTK berechnen
            label11.Text = "Handtragkraft (HTK): " + htk.ToString();
            label11.ForeColor = System.Drawing.Color.Indigo;

            handzahl = 1;
            if (checkBox1.Checked) { handzahl = 2; } //Wenn Zweihändig -> HTK faktisch verdoppelt

            fixwert3 = 0;
            if ((Waffenwahl.gewicht * 10 / handzahl) > (int)(htk * 10)) //Wenn Waffengewicht größer HTK
            fixwert3 = ((int)(htk*10) - (int)(Waffenwahl.gewicht*10 / handzahl))/10; //HTK minus Waffengewicht

            //Schuss
            if (Waffenwahl.waffentyp == "SW" || Waffenwahl.waffentyp == "SD")
            {
                schuss_ergebnis = fixwert1 + fixwert2 + fixwert3;
                dataGridView1.Rows.Add();
                rowcount++;
                dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightCoral;
                dataGridView1[0, rowcount - 1].Value = "Schuss mit " + Waffenwahl.name;
                if (schuss_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + schuss_ergebnis; }
                else { dataGridView1[1, rowcount - 1].Value = schuss_ergebnis; }
                switch (nachteil)
                {
                    case 1:
                        dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                        break;
                    case 2:
                        dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                        break;
                    case 3:
                        dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                        break;
                    case 4:
                        dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                        break;
                }
                dataGridView1[2, rowcount - 1].Value = Waffenwahl.waffenschaden + " Schaden (Typ je nach Munition)";
            }

            //Stichangriff
            if (Waffenwahl.stichmod > -20)
            {
                if (Waffenwahl.waffentyp != "(Kein Waffentyp)") //Wenn Waffe -> reguläre Berechnung
                {stichangriff_ergebnis = Waffenwahl.stichmod + fixwert1 + fixwert2 + fixwert3;}
                else //Wenn keine Waffe -> auch keine Waffenfähigkeit mit einberechnen
                { stichangriff_ergebnis = Waffenwahl.stichmod + fixwert2 + fixwert3;}
                dataGridView1.Rows.Add();
                rowcount++;
                dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightCoral;
                dataGridView1[0, rowcount - 1].Value = Waffenwahl.name + " Stichangriff"; 
                if (stichangriff_ergebnis >= 0){dataGridView1[1, rowcount - 1].Value = "+" + stichangriff_ergebnis;}
                else { dataGridView1[1, rowcount - 1].Value = stichangriff_ergebnis; }
                switch (nachteil)
                {
                    case 1: dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                        break;
                    case 2: dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                        break;
                    case 3: dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                        break;
                    case 4: dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                        break;
                }
                dataGridView1[2, rowcount - 1].Value = Waffenwahl.waffenschaden + " Stichschaden";
            }

            //Wuchtangriff
            if (Waffenwahl.wuchtmod > -20)
            {
                if (Waffenwahl.waffentyp != "(Kein Waffentyp)") //Wenn Waffe -> reguläre Berechnung
                { wuchtangriff_ergebnis = Waffenwahl.wuchtmod + fixwert1 + fixwert2 + fixwert3;}
                else //Wenn keine Waffe -> auch keine Waffenfähigkeit mit einberechnen
                { wuchtangriff_ergebnis = Waffenwahl.wuchtmod + fixwert2 + fixwert3;}
                dataGridView1.Rows.Add();
                rowcount++;
                dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightCoral;
                dataGridView1[0, rowcount - 1].Value = Waffenwahl.name + " Wuchtangriff";
                if (wuchtangriff_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + wuchtangriff_ergebnis; }
                else { dataGridView1[1, rowcount - 1].Value = wuchtangriff_ergebnis; }
                switch (nachteil)
                {
                    case 1:
                        dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                        break;
                    case 2:
                        dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                        break;
                    case 3:
                        dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                        break;
                    case 4:
                        dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                        break;
                }
                dataGridView1[2, rowcount - 1].Value = Waffenwahl.waffenschaden + " Wuchtschaden";
            }

            //Schnittangriff
            if (Waffenwahl.schnittmod > -20)
            {
                if (Waffenwahl.waffentyp != "(Kein Waffentyp)") //Wenn Waffe -> reguläre Berechnung
                { schnittangriff_ergebnis = Waffenwahl.schnittmod + fixwert1 + fixwert2 + fixwert3;}
                else //Wenn keine Waffe -> auch keine Waffenfähigkeit mit einberechnen
                { schnittangriff_ergebnis = Waffenwahl.schnittmod + fixwert2 + fixwert3;}
                dataGridView1.Rows.Add();
                rowcount++;
                dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightCoral;
                dataGridView1[0, rowcount - 1].Value = Waffenwahl.name + " Schnittangriff";
                if (schnittangriff_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + schnittangriff_ergebnis; }
                else { dataGridView1[1, rowcount - 1].Value = schnittangriff_ergebnis; }
                switch (nachteil)
                {
                    case 1:
                        dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                        break;
                    case 2:
                        dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                        break;
                    case 3:
                        dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                        break;
                    case 4:
                        dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                        break;
                }
                dataGridView1[2, rowcount - 1].Value = Waffenwahl.waffenschaden + " Schnittschaden";
            }
            
            //Waffe werfen
            werfen_ergebnis = (int)numericUpDown6.Value + Waffenwahl.werfenmod + fixwert1 + fixwert2 + fixwert3;
            dataGridView1.Rows.Add();
            rowcount++;
            dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightCoral;
            dataGridView1[0, rowcount - 1].Value = Waffenwahl.name + " werfen";
            if (werfen_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + werfen_ergebnis; }
            else { dataGridView1[1, rowcount - 1].Value = werfen_ergebnis; }
            switch (nachteil)
            {
                case 1:
                    dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                    break;
                case 2:
                    dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                    break;
                case 3:
                    dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                    break;
                case 4:
                    dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                    break;
            }
            dataGridView1[2, rowcount - 1].Value = Waffenwahl.wurfschaden + " Schaden (siehe Regeln)";

            //Parieren
            parieren_ergebnis = (int)numericUpDown3.Value + Waffenwahl.parierenmod + fixwert1 + fixwert2 + fixwert3;
            dataGridView1.Rows.Add();
            rowcount++;
            dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1[0, rowcount - 1].Value = "Parieren mit " + Waffenwahl.name;
            if (parieren_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + parieren_ergebnis; }
            else { dataGridView1[1, rowcount - 1].Value = parieren_ergebnis; }
            switch (nachteil)
            {
                case 1:
                    dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                    break;
                case 2:
                    dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                    break;
                case 3:
                    dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                    break;
                case 4:
                    dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                    break;
            }
            dataGridView1[2, rowcount - 1].Value = "";

            //Blocken
            blocken_ergebnis = (int)numericUpDown4.Value + Waffenwahl.blockenmod + fixwert1 + fixwert2 + fixwert3;
            dataGridView1.Rows.Add();
            rowcount++;
            dataGridView1.Rows[rowcount -1].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1[0, rowcount - 1].Value = "Blocken mit " + Waffenwahl.name;
            if (blocken_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + blocken_ergebnis; }
            else { dataGridView1[1, rowcount - 1].Value = blocken_ergebnis; }
            switch (nachteil)
            {
                case 1:
                    dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                    break;
                case 2:
                    dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                    break;
                case 3:
                    dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                    break;
                case 4:
                    dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                    break;
            }
            dataGridView1[2, rowcount - 1].Value = "";

            //Ausschaltversuch
            ausschalten_ergebnis = (int)numericUpDown5.Value + Waffenwahl.ausschaltenmod + fixwert1 + fixwert2 + fixwert3;
            dataGridView1.Rows.Add();
            rowcount++;
            dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightYellow;
            dataGridView1[0, rowcount - 1].Value = "Ausschaltversuch mit " + Waffenwahl.name;
            if (ausschalten_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + ausschalten_ergebnis; }
            else { dataGridView1[1, rowcount - 1].Value = ausschalten_ergebnis; }
            switch (nachteil)
            {
                case 1:
                    dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                    break;
                case 2:
                    dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                    break;
                case 3:
                    dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                    break;
                case 4:
                    dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                    break;
            }
            dataGridView1[2, rowcount - 1].Value = "";

            //Öffnungsversuch
            oeffnen_ergebnis = (int)numericUpDown7.Value + Waffenwahl.oeffnenmod + fixwert1 + fixwert2 + fixwert3;
            dataGridView1.Rows.Add();
            rowcount++;
            dataGridView1.Rows[rowcount - 1].DefaultCellStyle.BackColor = Color.LightYellow;
            dataGridView1[0, rowcount - 1].Value = "Öffnungsversuch mit " + Waffenwahl.name;
            if (oeffnen_ergebnis >= 0) { dataGridView1[1, rowcount - 1].Value = "+" + oeffnen_ergebnis; }
            else { dataGridView1[1, rowcount - 1].Value = oeffnen_ergebnis; }
            switch (nachteil)
            {
                case 1:
                    dataGridView1[1, rowcount - 1].Value += " (Nachteil!)";
                    break;
                case 2:
                    dataGridView1[1, rowcount - 1].Value += " (Doppel-Nachteil!)";
                    break;
                case 3:
                    dataGridView1[1, rowcount - 1].Value += " (Dreifach-Nachteil!)";
                    break;
                case 4:
                    dataGridView1[1, rowcount - 1].Value += " (Vierfach-Nachteil!)";
                    break;
            }
            dataGridView1[2, rowcount - 1].Value = "";

        }

        // ###### Reisetest ######
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Add();
            rowcount2++;
            dataGridView2.Rows[rowcount2 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
            dataGridView2[0, rowcount2 - 1].Value = rowcount2 + " . Etappe, Tag " + tag;
            textBox1.Text += "\r\nTag " + tag + " - Etappe " + rowcount2 + "\r\n\r\n";

            ausdauerverlust = 0;
            willenskraftverlust = 0;
            extremfehl = 0;
            extremerfolg = 0;

            gewaehlte_reiseart = comboBox3.GetItemText(comboBox3.SelectedItem);
            gewaehltes_terrain = comboBox4.GetItemText(comboBox4.SelectedItem);
            switch (gewaehlte_reiseart)
            {
                case "Langsames Gehen":
                    switch (gewaehltes_terrain)
                    {
                        case "Straße/Weg":
                            reisegeschwindigkeit = 2;
                            orientierungszw = 0;
                            break;
                        case "Gelände (Wiesen, lichter Wald, verschneite Straße)":
                            reisegeschwindigkeit = 2;
                            orientierungszw = 10;
                            break;
                        case "Dickicht (Urwald, Gebirge, Sumpf, Tiefschnee":
                            reisegeschwindigkeit = zufall.Next(1, 2);
                            orientierungszw = 12;
                            break;
                    }
                    fortbewegungsmod = 1;
                    break;
                case "Gehen (Standard)":
                    switch (gewaehltes_terrain)
                    {
                        case "Straße/Weg":
                            reisegeschwindigkeit = 4;
                            orientierungszw = 0;
                            break;
                        case "Gelände (Wiesen, lichter Wald, verschneite Straße)":
                            reisegeschwindigkeit = 3;
                            orientierungszw = 10;
                            break;
                        case "Dickicht (Urwald, Gebirge, Sumpf, Tiefschnee":
                            reisegeschwindigkeit = 2;
                            orientierungszw = 12;
                            break;
                    }
                    fortbewegungsmod = 0;
                    break;
                case "Laufschritt":
                    switch (gewaehltes_terrain)
                    {
                        case "Straße/Weg":
                            reisegeschwindigkeit = 6;
                            orientierungszw = 0;
                            break;
                        case "Gelände (Wiesen, lichter Wald, verschneite Straße)":
                            reisegeschwindigkeit = 5;
                            orientierungszw = 11;
                            break;
                        case "Dickicht (Urwald, Gebirge, Sumpf, Tiefschnee":
                            reisegeschwindigkeit = 3;
                            orientierungszw = 13;
                            break;
                    }
                    fortbewegungsmod = -1;
                    break;
                case "Laufen":
                    switch (gewaehltes_terrain)
                    {
                        case "Straße/Weg":
                            reisegeschwindigkeit = 8;
                            orientierungszw = 0;
                            break;
                        case "Gelände (Wiesen, lichter Wald, verschneite Straße)":
                            reisegeschwindigkeit = 6;
                            orientierungszw = 12;
                            break;
                        case "Dickicht (Urwald, Gebirge, Sumpf, Tiefschnee":
                            reisegeschwindigkeit = 3;
                            orientierungszw = 14;
                            break;
                    }
                    fortbewegungsmod = -2;
                    break;
                case "Wagenfahrt":
                    switch (gewaehltes_terrain)
                    {
                        case "Straße/Weg":
                            reisegeschwindigkeit = zufall.Next(2,6);
                            orientierungszw = 0;
                            break;
                        case "Gelände (Wiesen, lichter Wald, verschneite Straße)":
                            reisegeschwindigkeit = zufall.Next(1,3);
                            orientierungszw = 10;
                            break;
                        case "Dickicht (Urwald, Gebirge, Sumpf, Tiefschnee":
                            reisegeschwindigkeit = 0;
                            orientierungszw = 0;
                            break;
                    }
                    break;
                case "Reiten":
                    switch (gewaehltes_terrain)
                    {
                        case "Straße/Weg":
                            reisegeschwindigkeit = 8;
                            orientierungszw = 0;
                            break;
                        case "Gelände (Wiesen, lichter Wald, verschneite Straße)":
                            reisegeschwindigkeit = 5;
                            orientierungszw = 11;
                            break;
                        case "Dickicht (Urwald, Gebirge, Sumpf, Tiefschnee":
                            reisegeschwindigkeit = 0;
                            orientierungszw = 0;
                            break;
                    }
                    break;
                case "Galoppieren":
                    switch (gewaehltes_terrain)
                    {
                        case "Straße/Weg":
                            reisegeschwindigkeit = 16;
                            orientierungszw = 10;
                            break;
                        case "Gelände (Wiesen, lichter Wald, verschneite Straße)":
                            reisegeschwindigkeit = 12;
                            orientierungszw = 12;
                            break;
                        case "Dickicht (Urwald, Gebirge, Sumpf, Tiefschnee":
                            reisegeschwindigkeit = 0;
                            orientierungszw = 0;
                            break;
                    }
                    break;
            }

            //Belastung
            if (radioButton1.Checked) reisenachteil = 0;
            else if (radioButton2.Checked) reisenachteil = 1;
            else if (radioButton3.Checked) reisenachteil = 1;
            else if (radioButton4.Checked) reisenachteil = 2;

            //Würfelwurf            
            for (int i = 0; i < (int)numericUpDown8.Value; i++)
            {
                textBox1.Text += "Reisetest: ";
                switch (reisenachteil)
                {
                    case 0:
                        wuerfelgesamt = zufall.Next(1, 7) + zufall.Next(1, 7) + zufall.Next(1, 7);
                        textBox1.Text += "Wurf: ";
                        break;
                    case 1: //Nachteil
                        wuerfelgesamt = nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7));
                        wuerfelgesamt += nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7));
                        wuerfelgesamt += nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7));
                        textBox1.Text += "Nachteilswurf: ";
                        break;
                    case 2: //Doppel-Nachteil
                        wuerfelgesamt = nimmdiekleinerezahl(nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7)), zufall.Next(1, 7));
                        wuerfelgesamt += nimmdiekleinerezahl(nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7)), zufall.Next(1, 7));
                        wuerfelgesamt += nimmdiekleinerezahl(nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7)), zufall.Next(1, 7));
                        textBox1.Text += "Doppelnachteilswurf: ";
                        break;
                    case 3: //Dreifach-Nachteil
                        wuerfelgesamt = nimmdiekleinerezahl(nimmdiekleinerezahl(nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7)), zufall.Next(1, 7)), zufall.Next(1, 7));
                        wuerfelgesamt += nimmdiekleinerezahl(nimmdiekleinerezahl(nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7)), zufall.Next(1, 7)), zufall.Next(1, 7));
                        wuerfelgesamt += nimmdiekleinerezahl(nimmdiekleinerezahl(nimmdiekleinerezahl(zufall.Next(1, 7), zufall.Next(1, 7)), zufall.Next(1, 7)), zufall.Next(1, 7));
                        textBox1.Text += "Dreifachnachteilswurf: ";
                        break;
                }
                reisetest_ergebnis = wuerfelgesamt + (int)numericUpDown9.Value + fortbewegungsmod;
                textBox1.Text += wuerfelgesamt + " (Würfel) + " + (int)numericUpDown9.Value + " (Kon/Reiten) + " + fortbewegungsmod + " (Mod) = " + reisetest_ergebnis + " gegen ZW " + zielwert;
                resultat = test3w6gegenzw(zielwert, reisetest_ergebnis, (int)numericUpDown9.Value + fortbewegungsmod);
                switch (resultat)
                {
                    case 1: textBox1.Text += " ► Extremer Fehlschlag!\r\n";
                        break;
                    case 2: textBox1.Text += " ► Fehlschlag\r\n";
                        break;
                    case 3: textBox1.Text += " ► Teilerfolg\r\n";
                        break;
                    case 4: textBox1.Text += " ► Erfolg\r\n";
                        break;
                    case 5: textBox1.Text += " ► Extremer Erfolg!\r\n";
                        break;
                }
                switch (resultat)
                {
                    case 1:
                        extremfehl++;
                        break;
                    case 2:
                        ausdauerverlust++;
                        break;
                    case 3:
                        willenskraftverlust++;
                        break;
                    case 5:
                        extremerfolg++;
                        break;
                }
                zielwert++;
            }


            //Orientierungstest
            fortschrittdurchorientierung = 0;
            fortschrittausgegeben = 0;
            if (orientierungszw > 0)
            {
                for (int i = 0; i < (int)numericUpDown8.Value; i++)
                {
                    orientierungswurf = zufall.Next(1, 7) + zufall.Next(1, 7) + zufall.Next(1, 7) + (int)numericUpDown10.Value;
                    textBox1.Text += "Orientierungstest: " + orientierungswurf + " gegen ZW " + orientierungszw;
                    resultat = test3w6gegenzw(orientierungszw, orientierungswurf, (int)numericUpDown10.Value);
                    switch (resultat)
                    {
                        case 1:
                            textBox1.Text += " ► Extremer Fehlschlag!\r\n";
                            break;
                        case 2:
                            textBox1.Text += " ► Fehlschlag\r\n";
                            break;
                        case 3:
                            textBox1.Text += " ► Teilerfolg\r\n";
                            break;
                        case 4:
                            textBox1.Text += " ► Erfolg\r\n";
                            break;
                        case 5:
                            textBox1.Text += " ► Extremer Erfolg!\r\n";
                            break;
                    }
                    switch (resultat)
                    {
                        case 1:
                            dataGridView2[1, rowcount2 - 1].Value = "Irrweg! -" + (int)numericUpDown10.Value + " Meilen"; //Eine Meile Verlust pro Stunde
                            fortschrittausgegeben = 1;
                            break;
                        case 2:
                            dataGridView2[1, rowcount2 - 1].Value = "Im Kreis gelaufen! Kein Fortschritt!";
                            fortschrittausgegeben = 1;
                            break;
                        case 3:
                            fortschrittdurchorientierung--;
                            break;
                        case 5:
                            fortschrittdurchorientierung++;
                            break;
                    }
                    
                }
            }

            //Strecke
            if (fortschrittausgegeben != 1) dataGridView2[1, rowcount2 - 1].Value = ((int)numericUpDown8.Value * reisegeschwindigkeit + fortschrittdurchorientierung) + " Meilen";
            dataGridView2[2, rowcount2 - 1].Value = ("-" + ausdauerverlust + " Ausdauer");
            dataGridView2[3, rowcount2 - 1].Value = ("-" + willenskraftverlust + " Willenskraft");
            dataGridView2[4, rowcount2 - 1].Value = (extremfehl);
            dataGridView2[5, rowcount2 - 1].Value = (extremerfolg);
            if (extremerfolg > 0) dataGridView2[5, rowcount2 - 1].Style.BackColor = Color.LightGreen;
            if (extremfehl > 0) dataGridView2[4, rowcount2 - 1].Style.BackColor = Color.LightCoral;


        }

        // ###### Seereisetest ######
        private void button5_Click(object sender, EventArgs e)
        {
            //Daten aus der geladenen Datenbank ziehen
            int gewgrundgeschwindigkeit = 0;
            int gewrudersoll = 0;
            int gewsegelsoll = 0;
            string tmp = "";

            foreach (DataRow reihe in SchiffstypenTB.Rows)
            {
                if (comboBox5.GetItemText(this.comboBox5.SelectedItem) == "Auswählen...")
                {
                    MessageBox.Show("Es muss ein Schiffstyp ausgewählt werden");
                    comboBox5.BackColor = Color.LightCoral;
                    return;
                }
                else if (reihe["Schiffstyp"].ToString() == comboBox5.GetItemText(this.comboBox5.SelectedItem))
                {
                    tmp = reihe["Grundgeschwindigkeit"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") gewgrundgeschwindigkeit = int.Parse(tmp);
                    else gewgrundgeschwindigkeit = 0;
                    tmp = reihe["Rudersoll"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") gewrudersoll = int.Parse(tmp);
                    else gewrudersoll = 0;
                    tmp = reihe["Segelsoll"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") gewsegelsoll = int.Parse(tmp);
                    else gewsegelsoll = 0;
                }
            }

            if (radioButton5.Checked) fahrweise = "Rudern";
            else if (radioButton6.Checked) fahrweise = "Segeln";

            if (radioButton7.Checked) seereiseumgebung = "Binnensee";
            else if (radioButton8.Checked) seereiseumgebung = "Küstennähe";
            else if (radioButton9.Checked) seereiseumgebung = "Offenes Meer";
            else if (radioButton10.Checked) seereiseumgebung = "Flussabwärts"; 
            else if (radioButton11.Checked) seereiseumgebung = "Flussaufwärts";

            //Wetterberechnung
            int wuerfel = zufall.Next(1, 21);
            switch (seereiseumgebung)
            {
                case "Binnensee":
                    if (wuerfel < 7) wetter = "Windstill";
                    else if (wuerfel > 6 && wuerfel < 17) wetter = "Brise";
                    else if (wuerfel > 16 && wuerfel < 20) wetter = "Wind";
                    else wetter = "Sturm";
                    break;
                case "Küstennähe":
                    if (wuerfel < 6) wetter = "Windstill";
                    else if (wuerfel > 5 && wuerfel < 15) wetter = "Brise";
                    else if (wuerfel > 14 && wuerfel < 20) wetter = "Wind";
                    else wetter = "Sturm";
                    break;
                case "Offenes Meer":
                    if (wuerfel < 4) wetter = "Windstill";
                    else if (wuerfel > 3 && wuerfel < 10) wetter = "Brise";
                    else if (wuerfel > 9 && wuerfel < 18) wetter = "Wind";
                    else if (wuerfel > 17 && wuerfel < 20) wetter = "Sturm";
                    else wetter = "Orkan";
                    break;
                case "Flussabwärts":
                    if (fahrweise == "Rudern") schwierigkeit =6;
                    else if (fahrweise == "Segeln") schwierigkeit = 10;
                    wetter = "Flussreise - Kein Wettermod.";
                    break;
                case "Flussaufwärts":
                    if (fahrweise == "Rudern") schwierigkeit = 20;
                    else if (fahrweise == "Segeln") schwierigkeit = 40;
                    wetter = "Flussreise - Kein Wettermod.";
                    break;
                default:
                    MessageBox.Show("Es muss eine Fahrweise gewählt werden!");
                    radioButton7.BackColor = Color.LightCoral;
                    radioButton8.BackColor = Color.LightCoral;
                    radioButton9.BackColor = Color.LightCoral;
                    radioButton10.BackColor = Color.LightCoral;
                    radioButton11.BackColor = Color.LightCoral;
                    return;
            }

            //Schwierigkeitsberechnung
            switch (wetter)
            {
                case "Windstill":
                    if (fahrweise == "Rudern") schwierigkeit = 10;
                    else if (fahrweise == "Segeln") schwierigkeit = 24;
                    break;
                case "Brise":
                    if (fahrweise == "Rudern") schwierigkeit = 13;
                    else if (fahrweise == "Segeln") schwierigkeit = 16;
                    break;
                case "Wind":
                    if (fahrweise == "Rudern") schwierigkeit = 16;
                    else if (fahrweise == "Segeln") schwierigkeit = 14;
                    break;
                case "Sturm":
                    if (fahrweise == "Rudern") schwierigkeit = 30;
                    else if (fahrweise == "Segeln") schwierigkeit = 24;
                    break;
                case "Orkan":
                    if (fahrweise == "Rudern") schwierigkeit = 60;
                    else if (fahrweise == "Segeln") schwierigkeit = 60;
                    break;

            }

            //Ohne eingetragene Mannschaft
            if (checkBox2.Checked == false)
            {
                //Fähigkeitsberechnung
                seefahrtsfähigkeit = numericUpDown12.Value;
                //Falls Besatzung größer als Sollwert -> Zurechtstutzen
                if (numericUpDown11.Value > gewrudersoll) effektiveruderer = gewrudersoll;
                else effektiveruderer = numericUpDown11.Value;
                if (numericUpDown11.Value > gewsegelsoll) effektivesegler = gewsegelsoll;
                else effektivesegler = numericUpDown11.Value;
            }
            //Bei manuell eingetragener Mannschaft
            else
            {
                //Fähigkeitsberechnung
                decimal fähigkeitssumme = 0;
                decimal anzahlsumme = 0;
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    if (Convert.ToInt32(row.Cells[0].Value) > 0) anzahlsumme += Convert.ToDecimal(row.Cells[0].Value);
                    if (Convert.ToInt32(row.Cells[0].Value) > 0) fähigkeitssumme += (Convert.ToDecimal(row.Cells[0].Value) * Convert.ToInt32(row.Cells[2].Value));
                }
                //Falls Besatzung größer als Sollwert -> Zurechtstutzen
                seefahrtsfähigkeit = fähigkeitssumme / anzahlsumme;
                seefahrtsfähigkeit = Math.Round(seefahrtsfähigkeit, 2);
                if (anzahlsumme > gewrudersoll) effektiveruderer = gewrudersoll;
                else effektiveruderer = anzahlsumme;
                if (anzahlsumme > gewsegelsoll) effektivesegler = gewsegelsoll;
                else effektivesegler = anzahlsumme;
            }

            //Würfel
            int naturlicherseefahrtswurf = zufall.Next(1, 7) + zufall.Next(1, 7) + zufall.Next(1, 7);
            seefahrtswurf = naturlicherseefahrtswurf + seefahrtsfähigkeit;

            //Knotenberechnung
            if (fahrweise == "Rudern")
            {
                if (gewrudersoll == 0) 
                {
                    MessageBox.Show("Der Schiffstyp " + comboBox5.GetItemText(this.comboBox5.SelectedItem) + " erlaubt kein Rudern!");
                    return;
                }
                knoten = gewgrundgeschwindigkeit * (effektiveruderer / gewrudersoll) * seefahrtswurf / schwierigkeit;
                knoten = Math.Round(knoten, 2);
            }
            else if (fahrweise == "Segeln")
            {
                if (gewsegelsoll == 0)
                {
                    MessageBox.Show("Der Schiffstyp " + comboBox5.GetItemText(this.comboBox5.SelectedItem) + "  erlaubt kein Segeln!");
                    return;
                }
                knoten = gewgrundgeschwindigkeit * (effektivesegler / gewsegelsoll) * seefahrtswurf / schwierigkeit;
                knoten = Math.Round(knoten, 2);
            }
            
            //Ausgabe
            dataGridView3.Rows.Add();
            rowcount3++;
            dataGridView3.Rows[rowcount3 - 1].DefaultCellStyle.BackColor = Color.LightYellow;
            dataGridView3[0, rowcount3 - 1].Value = rowcount3 + " .";
            dataGridView3[3, rowcount3 - 1].Value = wetter;
            dataGridView3[1, rowcount3 - 1].Value = comboBox5.GetItemText(this.comboBox5.SelectedItem);
            dataGridView3[2, rowcount3 - 1].Value = seereiseumgebung + " - " + fahrweise;
            dataGridView3[4, rowcount3 - 1].Value = schwierigkeit;
            dataGridView3[5, rowcount3 - 1].Value = naturlicherseefahrtswurf + " + " + (seefahrtsfähigkeit) + " = " + (seefahrtswurf);
            dataGridView3[6, rowcount3 - 1].Value = knoten + " Knoten";
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
            OleDbCommand command = new OleDbCommand("SELECT * FROM Schiffstypen", connection);
            OleDbCommand command2 = new OleDbCommand("SELECT * FROM Schiffstypen", connection);
            OleDbCommand command3 = new OleDbCommand("SELECT * FROM Waffen", connection);
            OleDbCommand command4 = new OleDbCommand("SELECT * FROM Waffen", connection);
            reader = command.ExecuteReader();
            reader2 = command2.ExecuteReader();
            reader3 = command3.ExecuteReader();
            reader4 = command4.ExecuteReader();
            comboBox5.Items.Clear();
            comboBox2.Items.Clear();

            while (reader.Read())
            {
                comboBox5.Items.Add(reader[0].ToString());
            }
            while (reader3.Read())
            {
                comboBox2.Items.Add(reader3[1].ToString());
            }

            DataTable SchiffstypenTabelle = new DataTable();
            DataTable WaffenTabelle = new DataTable();

            SchiffstypenTabelle.Load(reader2);
            WaffenTabelle.Load(reader4);

            SchiffstypenTB = SchiffstypenTabelle;
            WaffenTB = WaffenTabelle;

            connection.Close();




            //SQL Server Verbindung
            //using (SqlConnection con = new SqlConnection(@"Data Source=MORELAM\ALPHA;Initial Catalog=TüfteltruheDatenbank;Integrated Security=True"))
            //{
            //    using (SqlDataAdapter sda = new SqlDataAdapter("SELECT TOP (1000) [Schiffstyp ],[Grundgeschwindigkeit],[Rudersoll],[Segelsoll] FROM[TTruheAccess].[accdb].[Schiffstypen]", con))
            //    {
            //        //Datentabelle mit sda füllen
            //        DataTable SchiffstypenTabelle = new DataTable();
            //        sda.Fill(SchiffstypenTabelle);
            //        //try
            //        //{
            //        //    sda.Fill(SchiffstypenTabelle);
            //        //}
            //        //catch
            //        //{
            //        //    MessageBox.Show("Kein Zugriff auf die Schiffstypen-Datenbank.", "Datenbank nicht erreichbar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        //    DataRow r = SchiffstypenTabelle.NewRow();
            //        //    SchiffstypenTabelle.Columns.Add("Schiffstyp ");
            //        //}

            //        //Neue Default-Zeile in der Tabelle
            //        DataRow row = SchiffstypenTabelle.NewRow();
            //        row[0] = "Auswählen...";
            //        SchiffstypenTabelle.Rows.InsertAt(row, 0);

            //        comboBox5.DataSource = SchiffstypenTabelle;
            //        comboBox5.DisplayMember = SchiffstypenTabelle.Columns[0].ToString();

            //        SchiffstypenTB = SchiffstypenTabelle;
            //    }
            //}

            //SQL Server Verbindung
            //using (SqlConnection con = new SqlConnection(@"Data Source = E:\Tüfteltruhe\Tüfteltruhe\TTruheAccess.accdb; Persist Security Info=False;"))
            //{
            //    string selector = "SELECT TOP (1000) [Waffe ],[Waffenfähigkeit],[Stich],[Wucht],[Schnitt],[Parieren],[Blocken],[Ausschalten],[Werfen],[Öffnen],[Waffenenergie],[Waffenschaden],[Wurfschaden],[Schirmwert],[Reichweite (Nah)],[Reichweite (Fern)],[Gewicht]FROM[TTruheAccess].[accdb].[Waffen]";
            //    using (SqlDataAdapter sda = new SqlDataAdapter(selector, con))
            //    {
            //        //Datentabelle mit sda füllen
            //        DataTable WaffenTabelle = new DataTable();
            //        sda.Fill(WaffenTabelle);
            //        //try
            //        //{
            //        //    sda.Fill(WaffenTabelle);
            //        //}
            //        //catch
            //        //{
            //        //    MessageBox.Show("Kein Zugriff auf die Waffen-Datenbank.", "Datenbank nicht erreichbar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        //    DataRow r = WaffenTabelle.NewRow();
            //        //    WaffenTabelle.Columns.Add("Waffe ");
            //        //}

            //        //+++Nach Alphabet sortieren
            //        //WaffenTabelle.DefaultView.Sort = WaffenTabelle.Columns[0].ToString();

            //        //Neue Default-Zeile in der Tabelle
            //        DataRow row = WaffenTabelle.NewRow();
            //        row[0] = "Auswählen...";
            //        WaffenTabelle.Rows.InsertAt(row, 0);

            //        comboBox2.DataSource = WaffenTabelle;
            //        comboBox2.DisplayMember = WaffenTabelle.Columns[0].ToString();
            //        //MessageBox.Show(WaffenTabelle.Columns[0].ToString());

            //        WaffenTB = WaffenTabelle;
            //    }
            //}
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            gewaehlte_waffe = comboBox2.GetItemText(comboBox2.SelectedItem);
            //MessageBox.Show(WaffenTB.Rows[1].ToString() + " in tabelle");

            foreach (DataRow row in WaffenTB.Rows)
            {
                //MessageBox.Show(row["Waffe"].ToString() + " in tabelle");
                //string z1 = WaffenTB.Columns[1].ToString();
                //MessageBox.Show(z1);
                if (row["Waffe"].ToString() == gewaehlte_waffe)
                {
                    string tmp = row["Waffe"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.name = tmp;
                    else Waffenwahl.name = null;
                    tmp = row["Waffenfähigkeit"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.waffentyp = tmp;
                    else Waffenwahl.waffentyp = null;
                    tmp = row["Gewicht"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.gewicht = (double)row["Gewicht"];
                    else Waffenwahl.gewicht = -100;
                    tmp = row["Stich"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.stichmod = int.Parse(tmp);
                    else Waffenwahl.stichmod = -100;
                    tmp = row["Schnitt"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.schnittmod = int.Parse(tmp);
                    else Waffenwahl.schnittmod = -100;
                    tmp = row["Wucht"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.wuchtmod = int.Parse(tmp);
                    else Waffenwahl.wuchtmod = -100;
                    tmp = row["Parieren"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.parierenmod = int.Parse(tmp);
                    else Waffenwahl.parierenmod = -100;
                    tmp = row["Blocken"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.blockenmod = int.Parse(tmp);
                    else Waffenwahl.blockenmod = -100;
                    tmp = row["Ausschalten"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.ausschaltenmod = int.Parse(tmp);
                    else Waffenwahl.ausschaltenmod = -100;
                    tmp = row["Werfen"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.werfenmod = int.Parse(tmp);
                    else Waffenwahl.werfenmod = -100;
                    tmp = row["Öffnen"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.oeffnenmod = int.Parse(tmp);
                    else Waffenwahl.oeffnenmod = -100;
                    tmp = row["Waffenschaden"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.waffenschaden = tmp;
                    else Waffenwahl.waffenschaden = null;
                    tmp = row["Wurfschaden"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.wurfschaden = tmp;
                    else Waffenwahl.wurfschaden = null;
                    tmp = row["Reichweite (Nah)"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.reichweiteNah = int.Parse(tmp);
                    else Waffenwahl.reichweiteNah = -100;
                    tmp = row["Reichweite (Fern)"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.reichweiteFern = int.Parse(tmp);
                    else Waffenwahl.reichweiteFern = -100;
                    tmp = row["Waffenenergie"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.waffenenergie = int.Parse(tmp);
                    else Waffenwahl.waffenenergie = -100;
                    tmp = row["Schirmwert"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") Waffenwahl.schirmwert = int.Parse(tmp);
                    else Waffenwahl.schirmwert = -100;
                }   
            }

            //Waffentyp für Fähigkeit anzeigen
            label10.Text = "Waffenfähigkeit " + Waffenwahl.waffentyp;
            label10.ForeColor = System.Drawing.Color.Indigo;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox3.GetItemText(comboBox3.SelectedItem))
            {
                case "Wagenfahrt":
                    label14.Text = "Reiten-Fähigkeit des Fahrers";
                    break;
                case "Reiten":
                case "Galoppieren":
                    label14.Text = "Reiten-Fähigkeit des Reiters";
                    break;
                default:
                    label14.Text = "Konstitutions-Mod.";
                    break;
            }
            label14.ForeColor = System.Drawing.Color.Indigo;
        }

        public int nimmdiekleinerezahl(int vgl1, int vgl2)
        {
            if (vgl1 >= vgl2) return vgl2;
            else return vgl1;
        }

        public int nimmdiegroesserezahl(int vgl1, int vgl2)
        {
            if (vgl1 >= vgl2) return vgl1;
            else return vgl2;
        }

        public int test3w6gegenzw(int zielwert, int testwert, int abzuziehen)
        {
            int nat = testwert - abzuziehen;
            if (nat <= 5) return 1; //Natürliches Ergebnis: Extremer Fehlschlag
            else if (nat >= 16) return 5; //Natürliches Ergebnis: Extremer Erfolg
            else if (testwert == zielwert || testwert == zielwert + 1 || testwert == zielwert - 1) return 3; //Teilerfolg
            else if (testwert > zielwert) return 4; //Erfolg
            else return 2;
        }

        private void button3_Click(object sender, EventArgs e) //Alles zurücksetzen
        {
            dataGridView2.Rows.Clear();
            rowcount2 = 0;
            textBox1.Text = "Protokoll:";
            tag = 1;
            zielwert = 8;
        }

        private void button4_Click(object sender, EventArgs e) //Neuer Tag
        {
            zielwert = 8;
            tag++;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Spielermodus spielermodus = new Spielermodus();
            spielermodus.Show();
        }


        private void neuesFensterSpielermodusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spielermodus spielermodus = new Spielermodus();
            spielermodus.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView3.Rows.Clear();
            rowcount3 = 0;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            foreach (DataRow reihe in SchiffstypenTB.Rows)
            {
                if (comboBox5.GetItemText(this.comboBox5.SelectedItem) != "Auswählen...")
                {
                    comboBox5.BackColor = Color.White;
                    erforderlichebesatzungaktualisieren();
                }
            }
        }
        private void erforderlichebesatzungaktualisieren()
        {
            int gewrudersoll = 0;
            int gewsegelsoll = 0;

            foreach (DataRow reihe in SchiffstypenTB.Rows)
            {
                string tmp = "";
                if (reihe["Schiffstyp"].ToString() == comboBox5.GetItemText(this.comboBox5.SelectedItem))
                {
                    tmp = reihe["Rudersoll"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") gewrudersoll = int.Parse(tmp);
                    else gewrudersoll = 0;
                    tmp = reihe["Segelsoll"].ToString();
                    if (!string.IsNullOrEmpty(tmp) && tmp != "-") gewsegelsoll = int.Parse(tmp);
                    else gewsegelsoll = 0;
                }
            }
            if (radioButton5.Checked) label19.Text = "Erforderlich: " + gewrudersoll.ToString();
            else if (radioButton6.Checked) label19.Text = "Erforderlich: " + gewsegelsoll.ToString();

        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (comboBox5.GetItemText(this.comboBox5.SelectedItem) != "Auswählen...")
            {
                comboBox5.BackColor = Color.White;
                erforderlichebesatzungaktualisieren();
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (comboBox5.GetItemText(this.comboBox5.SelectedItem) != "Auswählen...")
            {
                comboBox5.BackColor = Color.White;
                erforderlichebesatzungaktualisieren();
            }
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            radioButton7.BackColor = Color.White;
            radioButton8.BackColor = Color.White;
            radioButton9.BackColor = Color.White;
            radioButton10.BackColor = Color.White;
            radioButton11.BackColor = Color.White;
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            radioButton7.BackColor = Color.White;
            radioButton8.BackColor = Color.White;
            radioButton9.BackColor = Color.White;
            radioButton10.BackColor = Color.White;
            radioButton11.BackColor = Color.White;
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            radioButton7.BackColor = Color.White;
            radioButton8.BackColor = Color.White;
            radioButton9.BackColor = Color.White;
            radioButton10.BackColor = Color.White;
            radioButton11.BackColor = Color.White;
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            radioButton7.BackColor = Color.White;
            radioButton8.BackColor = Color.White;
            radioButton9.BackColor = Color.White;
            radioButton10.BackColor = Color.White;
            radioButton11.BackColor = Color.White;
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            radioButton7.BackColor = Color.White;
            radioButton8.BackColor = Color.White;
            radioButton9.BackColor = Color.White;
            radioButton10.BackColor = Color.White;
            radioButton11.BackColor = Color.White;
        }

        private void neuesFensterSpielleitermodusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Spielleitermodus spielleitermodus = new Spielleitermodus();
            spielleitermodus.Show();
        }
    }    
}
