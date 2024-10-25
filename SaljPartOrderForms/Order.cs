﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaljPartOrderForms
{
    public class Order : IDisposable
    {
        private Garp.Application oGarp;
        private Garp.IComponents CompOrder;
        private bool disposed = false;
        private Garp.Dataset dsOGR, dsOGA;

        //private Garp.Dataset dsOGA;
        private Garp.ITable tblOGR;
        private Version mVersion;

        private Garp.ITable oArtReg, oKOHReg;
        private Garp.ITable oKundReg, oTabReg;
        private Garp.ITabField oTabNyckel, oAvtalNr, oKundKat, oArtNr, oArtPris, oArtNyckel, oArtLevNr, oArtKat, oArtAntDec;
        private Garp.ITabField[] aoTabTxt = new Garp.ITabField[3];
        private Garp.ITabField[] aoTabNum = new Garp.ITabField[13];
        private Garp.ITabField[] aoTabKod = new Garp.ITabField[13];

        private Garp.ITabField oKoTxtOrderNr, oKOTxtRadNr, oKOTxtSekvNr, oKOTxt, oKOTxtOBFl, oKOTxtPLFl, oKOTxtFSFl, oKOTxtFAFl;
        private Garp.ITabField oKOHResKod, oKORadResKod;
        private Garp.ITable oKOTxtReg, oKORadReg;
        private Garp.IComponent edbBruttoPris, edbPallRab, edbKvantRab, edbAvtalsRab, edbAktRab, edbKundRab, edbProvision, edbRadText, edbRabattUtr;
        private Garp.IComponent lblPallRab, lblKvantRab, lblAvtalsRab, lblAktRab, lblKundRab, lblRabattUtr, lblProvision;
        private string[,] asRabattBas = new string[5, 3];
        private string sOrderNr = "", currentRowNo;
        private int iOrderRadNr;
        private int iLevFlagga = -1;
        private bool bFelkod;
        private string sUrsprPris;
        private decimal cUrsprAntal;

        private bool blndebugg = false;



        public Order()
        {
            try
            {
                oGarp = new Garp.Application();
                CompOrder = oGarp.Components;
                //dsOGA = oGarp.Datasets.Item("ogrMcDataSet");
                dsOGR = oGarp.Datasets.Item("ogrMcDataSet");
                InitTablesAndFields();

                //dsOGR.BeforePost += on_BeforePostOrderRow;
                //dsOGA.AfterScroll += on_AfterScrollOrder;
                dsOGR.AfterScroll += on_AfterScrollOrderRow;
                //MessageBox.Show("FORMS CREATED");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in the Order constructor" + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

    private void InitTablesAndFields()
    {
        try
        {
            oKundReg = oGarp.Tables.Item("KA");
            oAvtalNr = oKundReg.Fields.Item("NM1");
            oKundKat = oKundReg.Fields.Item("KAT");

            oTabReg = oGarp.Tables.Item("TA");
            oArtReg = oGarp.Tables.Item("AGA");
            oArtNr = oArtReg.Fields.Item("ANR");
            oArtPris = oArtReg.Fields.Item("PRI");
            oArtNyckel = oArtReg.Fields.Item("SES");
            oArtLevNr = oArtReg.Fields.Item("LNR");
            oArtKat = oArtReg.Fields.Item("KAT");
            oArtAntDec = oArtReg.Fields.Item("ADE");

            oKOHReg = oGarp.Tables.Item("OGA");
            oKOHResKod = oKOHReg.Fields.Item("RES");
            oKORadReg = oGarp.Tables.Item("OGR");
            oKORadResKod = oKORadReg.Fields.Item("RES");
            oKOTxtReg = oGarp.Tables.Item("OGK");

            oKoTxtOrderNr = oKOTxtReg.Fields.Item("ONR");
            oKOTxtRadNr = oKOTxtReg.Fields.Item("RDC");
            oKOTxtSekvNr = oKOTxtReg.Fields.Item("SQC");
            oKOTxt = oKOTxtReg.Fields.Item("TX1");
            oKOTxtOBFl = oKOTxtReg.Fields.Item("OBF");
            oKOTxtPLFl = oKOTxtReg.Fields.Item("PLF");
            oKOTxtFSFl = oKOTxtReg.Fields.Item("FSF");
            oKOTxtFAFl = oKOTxtReg.Fields.Item("FAF");


        }
        catch (Exception ex)
        {
            MessageBox.Show("Error in the InitTablesAndFields" + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

    }

    private void InitForNoRabatt()
    { 
        try
        {
                CompOrder.Item("vrabattGroupBox").Visible = true;
                CompOrder.Item("prisinfoBitBtn").Visible = true;
                CompOrder.Item("mceOgrREV").Visible = true;
                CompOrder.Item("lOgrREV").Visible = true;
                CompOrder.Item("btnOgrREV").Visible = true;

                CompOrder.Item("artikelinfoEdit").Top = 130;
                CompOrder.Item("artikelinfoEdit").Left = 308;
                CompOrder.Item("artikelinfoEdit").Width = 245;

                CompOrder.Item("Label53").Top = 130; // Note: Adjusted, possibly a typo in VB
                CompOrder.Item("Label53").Top = 330;
                CompOrder.Item("ogrPriMcEdit").Visible = true;
                CompOrder.Item("ogrResMcEdit").TabStop = true;
                CompOrder.Item("ogr2SesMcEdit").Visible = true; //Nya fält i 4.03
                CompOrder.Item("ogr2SesLookupBtn").Visible = true; //Nya fält i 4.03
                CompOrder.Item("lblogr2McEdit").Visible = true; //Nya fält i 4.03

                if (CompOrder.Item("edbBruttoPris") != null)
                {
                    CompOrder.Delete("edbBruttoPris");
                    CompOrder.Delete("edbPallRab");
                    CompOrder.Delete("edbKvantRab");
                    CompOrder.Delete("edbAvtalsRab");
                    CompOrder.Delete("edbAktRab");
                    CompOrder.Delete("edbKundRab");
                    CompOrder.Delete("edbRadText");
                }

                if (CompOrder.Item("lblPallRab") != null)
                {
                    CompOrder.Delete("lblPallRab");
                    CompOrder.Delete("lblKvantRab");
                    CompOrder.Delete("lblAvtalsRab");
                    CompOrder.Delete("lblAktRab");
                    CompOrder.Delete("lblKundRab");
                }

                //oGarp.FieldEnter -= FieldEnter;
                //oGarp.FieldExit -= FieldExit;
                sOrderNr = "";
        }         
            
        catch (Exception ex)
        {
            MessageBox.Show("Error in the InitForRabatt " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }

    private void InitForRabatt()
    {
            int a = 0;
            try
            {          
                mVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

                for (int iX = 1; iX <= 2; iX++)
                {
                    aoTabTxt[iX] = oTabReg.Fields.Item("TX" + iX.ToString());
                }
                a = 1;
                for (int iX = 1; iX <= 12; iX++)
                {
                    aoTabNum[iX] = oTabReg.Fields.Item("FX" + iX.ToString());
                    aoTabKod[iX] = oTabReg.Fields.Item("KD" + iX.ToString());
                }
                a = 2;

                sOrderNr = CompOrder.Item("onrEdit").Text;
                iOrderRadNr = Convert.ToInt32(CompOrder.Item("oradEdit").Text);

                //MessageBox.Show("iOrderRadNr " + iOrderRadNr);

                iLevFlagga = Convert.ToInt32(CompOrder.Item("ogrLvfMcedit").Text.Substring(0, 1));

                // Hide certain components
                CompOrder.Item("vrabattGroupBox").Visible = false;
                CompOrder.Item("prisinfoBitBtn").Visible = false;
                CompOrder.Item("mceOgrREV").Visible = false;
                CompOrder.Item("lOgrREV").Visible = false;
                CompOrder.Item("btnOgrREV").Visible = false;

                CompOrder.Item("artikelinfoEdit").Top = 130;
                CompOrder.Item("artikelinfoEdit").Left = 308;
                CompOrder.Item("artikelinfoEdit").Width = 245;

                CompOrder.Item("Label53").Top = 130; // Note: Adjusted, possibly a typo in VB
                CompOrder.Item("Label53").Top = 330;
                CompOrder.Item("ogrPriMcEdit").Visible = false;
                CompOrder.Item("ogrResMcEdit").TabStop = false;
                CompOrder.Item("ogr2SesMcEdit").Visible = false; //Nya fält i 4.03
                CompOrder.Item("ogr2SesLookupBtn").Visible = false; //Nya fält i 4.03
                CompOrder.Item("lblogr2McEdit").Visible = false; //Nya fält i 4.03

                a = 3;
                InitFormKundOrder();
                a = 4;
                InitTxtkundOrder();
                a = 5;
                InitKontroll();
                a = 6;


                if (iLevFlagga < 5) // Undelivered order
                {
                    a = 7;
                    LäsInAPris();
                    a = 8;
                    LäsInRabatter();
                    a = 9;
                    BeräknaNetto();
                    a = 10;
                }

                oGarp.FieldEnter += FieldEnter;
                oGarp.FieldExit += FieldExit;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in the InitForRabatt " + a + "-" + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void InitKontroll()
        {
            if (iLevFlagga < 5) // Undelivered order
            {
                DatumKontroll();
                if (oArtReg.Find(CompOrder.Item("ogrAnrMcEdit").Text))
                {
                    CompOrder.Item("ogrKstMcEdit").Text = oArtKat.Value; // Cost center = ArtKat
                    if (string.IsNullOrEmpty(CompOrder.Item("ogaSesMcEdit").Text))
                    {
                        CompOrder.Item("ogaSesMcEdit").Text = oArtLevNr.Value; // Sets supplier number in the season field
                    }
                    else if (CompOrder.Item("ogaSesMcEdit").Text != oArtLevNr.Value)
                    {
                        // Additional validation based on supplier number
                        if (CompOrder.Item("ogaSesMcEdit").Text == "10" || oArtLevNr.Value == "10")
                        {
                            MessageBox.Show("Lecoraartiklar och artiklar från andra leverantörer får inte förekomma på samma order.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            CompOrder.Item("ogrAnrMcedit").SetFocus();
                        }
                        else if (CompOrder.Item("ogaSesMcEdit").Text == "60" || oArtLevNr.Value == "60")
                        {
                            MessageBox.Show("Dr Pers-artiklar och artiklar från andra leverantörer får inte förekomma på samma order.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            CompOrder.Item("ogrAnrMcedit").SetFocus();
                        }
                    }
                }

                if (CompOrder.Item("ogaSesMcEdit").Text == "10")
                {
                    CompOrder.Item("ogaBvkMcEdit").Text = "ZZ"; // Billing conditions not invoiced
                }
            }
        }


        private void DatumKontroll()
        {
            try
            { 
            string sDatum;

                while (string.IsNullOrEmpty(CompOrder.Item("ogaBltMcEdit").Text))
                {
                    if (string.IsNullOrEmpty(CompOrder.Item("ogrLdtMcEdit").Text)) // If delivery date is missing
                    {
                        inputdate levdate = new inputdate();
                        levdate.ShowDialog();

                        sDatum = levdate.TheDate;
                       
                        CompOrder.Item("ogaBltMcEdit").Text = sDatum;
                        CompOrder.Item("ogrLdtMcEdit").Text = sDatum;
                        CompOrder.Item("ogrOraMcEdit").SetFocus();
                        System.Threading.Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("InitFormKundOrder " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void InitFormKundOrder()
        {
            int a = 0;
            try
            {

                if (iLevFlagga < 5)
                {
                    a = 1;
                    CompOrder.BaseComponent = "Tabsheet3";
                    a = 2;
                    if(CompOrder.Item("edbBruttoPris") == null)
                    {
                        edbBruttoPris = CompOrder.AddEdit("edbBruttoPris");
                        edbPallRab = CompOrder.AddEdit("edbPallRab");
                        edbKvantRab = CompOrder.AddEdit("edbKvantRab");
                        edbAvtalsRab = CompOrder.AddEdit("edbAvtalsRab");
                        edbAktRab = CompOrder.AddEdit("edbAktRab");
                        edbKundRab = CompOrder.AddEdit("edbKundRab");
                        edbRadText = CompOrder.AddEdit("edbRadText");
                    }

                    a = 3;

                    // Set properties for edbBruttoPris
                    edbBruttoPris.Top = 28;
                    edbBruttoPris.Left = 365;
                    edbBruttoPris.MaxLength = 11;
                    edbBruttoPris.Width = 75;
                    edbBruttoPris.Height = 20;
                    edbBruttoPris.Text = CompOrder.Item("ogrPriMcEdit").Text; //TODO Format(CompOrder.Item("ogrPriMcEdit").Text, "@@@@@@@@@");
                    edbBruttoPris.TabOrder = 8;
                    edbBruttoPris.TabStop = true;
                    a = 4;
                    edbPallRab.Top = 125;
                    edbPallRab.Left = 3;
                    edbPallRab.MaxLength = 6;
                    edbPallRab.Width = 58;
                    edbPallRab.Height = 20;
                    edbPallRab.Text = "0.00";  //TODO Format(CompOrder.Item("ogrPriMcEdit").Text, "@@@@@@@@@");
                    edbPallRab.TabStop = false;

                    edbKvantRab.Top = 125;
                    edbKvantRab.Left = 64;
                    edbKvantRab.MaxLength = 11;
                    edbKvantRab.Width = 58;
                    edbKvantRab.Height = 20;
                    edbKvantRab.Text = "0.00"; //TODO Format(CompOrder.Item("ogrPriMcEdit").Text, "@@@@@@@@@");
                    edbKvantRab.TabStop = false;

                    edbAvtalsRab.Top = 125;
                    edbAvtalsRab.Left = 124;
                    edbAvtalsRab.MaxLength = 11;
                    edbAvtalsRab.Width = 58;
                    edbAvtalsRab.Height = 20;
                    edbAvtalsRab.Text = "0.00"; //TODO Format(CompOrder.Item("ogrPriMcEdit").Text, "@@@@@@@@@");
                    edbAvtalsRab.TabStop = false;

                    edbAktRab.Top = 125;
                    edbAktRab.Left = 184;
                    edbAktRab.MaxLength = 11;
                    edbAktRab.Width = 58;
                    edbAktRab.Height = 20;
                    edbAktRab.Text = "0.00"; //TODO Format(CompOrder.Item("ogrPriMcEdit").Text, "@@@@@@@@@");
                    edbAktRab.TabStop = false;

                    edbKundRab.Top = 125;
                    edbKundRab.Left = 244;
                    edbKundRab.MaxLength = 11;
                    edbKundRab.Width = 58;
                    edbKundRab.Height = 20;
                    edbKundRab.Text = "0.00"; //TODO Format(CompOrder.Item("ogrPriMcEdit").Text, "@@@@@@@@@");
                    edbKundRab.TabStop = false;

                    edbRadText.Top = 160;
                    edbRadText.Left = 3;
                    edbRadText.MaxLength = 60;
                    edbRadText.Width = 300;
                    edbRadText.Height = 20;
                    a = 5;
                    CompOrder.Item("artikelinfoedit").Top = 160;
                    a = 6;
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show("InitFormKundOrder " + a + " " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void InitTxtkundOrder()
        {
            if (iLevFlagga < 5)
            {
                if (CompOrder.Item("lblPallRab") == null)
                {
                    lblPallRab = CompOrder.AddLabel("lblPallRab");
                    lblKvantRab = CompOrder.AddLabel("lblKvantRab");
                    lblAvtalsRab = CompOrder.AddLabel("lblAvtalsRab");
                    lblAktRab = CompOrder.AddLabel("lblAktRab");
                    lblKundRab = CompOrder.AddLabel("lblKundRab");
                }

                lblPallRab.Top = 105;
                lblPallRab.Left = 5;
                lblPallRab.Text = "Helpall";

                lblKvantRab.Top = 93;
                lblKvantRab.Left = 66;
                lblKvantRab.Text = "Kvant/\r\nHämt";

                lblAvtalsRab.Top = 105;
                lblAvtalsRab.Left = 126;
                lblAvtalsRab.Text = "Avtal";

                lblAktRab.Top = 105;
                lblAktRab.Left = 186;
                lblAktRab.Text = "Aktivitet";

                lblKundRab.Top = 105;
                lblKundRab.Left = 246;
                lblKundRab.Text = "Kund";
            }
        }

        private void LäsInAPris()
        {
            decimal NX1Pris;
            string sArtNr;

            sArtNr = CompOrder.Item("ogrAnrMcEdit").Text; // Nyinlagt 200512
            if (oArtReg.Find(sArtNr))
            {
                sUrsprPris = oArtPris.Value;
                if (edbBruttoPris.Text !="" && decimal.Parse(edbBruttoPris.Text.Replace(".", ",")) != 0)
                {
                    edbBruttoPris.Text = sUrsprPris.ToString();
                }
                else
                {
                    CompOrder.Item("ogrPriMcEdit").Text = edbBruttoPris.Text;
                }
            }

            NX1Pris = decimal.Parse(CompOrder.Item("ogrNX1McEdit").Text.Replace(".", ","));

            if (CompOrder.Item("ogrNX1McEdit").Text == "0")
            {
                NX1Pris = decimal.Parse(CompOrder.Item("edbBruttoPris").Text.Replace(".", ",")) * 100;
                CompOrder.Item("ogrNX1McEdit").Text = NX1Pris.ToString().Replace(",", ".");
            }
            else if (CompOrder.Item("ogrNX1McEdit").Text != "0")
            {
                edbBruttoPris.Text = (NX1Pris / 100).ToString("#0.00").Replace(",", ".");
                edbBruttoPris.Text = edbBruttoPris.Text; //TODO Format(edbBruttoPris.Text, "@@@@@@@@@");
            }
        }

        private void FieldEnter()
        {
            try
            {
                if(!disposed)
                {
                     if(CompOrder.Item("edbBruttoPris") != null && iLevFlagga < 5) { 
                        switch (CompOrder.CurrentField)
                        {
                            case "ogrOraMcEdit":
                                cUrsprAntal = decimal.Parse(CompOrder.Item("ogrOraMcEdit").Text.Replace(".", ","));
                                break;
                            case "edbBruttoPris":
                                sUrsprPris = CompOrder.Item("edbBruttoPris").Text;
                                break;
                            case "edbPallRab":
                                CompOrder.Item("ogrOraMcEdit").SetFocus();
                                break;
                            case "edbKvantRab":
                                CompOrder.Item("ogrOraMcEdit").SetFocus();
                                break;
                            case "edbAvtalsRab":
                                CompOrder.Item("ogrOraMcEdit").SetFocus();
                                break;
                            case "edbAktRab":
                                CompOrder.Item("ogrOraMcEdit").SetFocus();
                                break;
                            case "edbKundRab":
                                CompOrder.Item("ogrOraMcEdit").SetFocus();
                                break;
                        }
                     }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("FieldEnter " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void FieldExit()
        {
            try
            {
                if (!disposed)
                {
                    string sFältVärde;
                    string sPrisfråga;
                    //Sholud only perform below if rabattarticle

                    if (CompOrder.Item("edbBruttoPris") != null && iLevFlagga < 5)
                    {
                        switch (CompOrder.CurrentField)
                        {
                            case "ogrAnrMcEdit": // Artikelnummer
                                InitKontroll(); // Subanrop
                                break;
                            case "ogrOraMcEdit": // Orderantal
                                LoggaAntalsÄndringar(); // Subanrop
                                LäsInRabatter(); // Subanrop
                                BeräknaNetto(); // Subanrop
                                break;
                            case "edbBruttoPris":
                                if (edbBruttoPris.Text.Trim() != sUrsprPris.Trim())
                                {
                                    sPrisfråga = MessageBox.Show("Vill du verkligen ändra priset?", "Fråga?", MessageBoxButtons.YesNo, MessageBoxIcon.Question).ToString();
                                    if (sPrisfråga == "Yes")
                                    {
                                        edbBruttoPris.Text = edbBruttoPris.Text.Trim().Replace(".", ",");
                                        edbBruttoPris.Text = Math.Round(decimal.Parse(edbBruttoPris.Text), 2).ToString("#0.00");
                                        CompOrder.Item("ogrNX1McEdit").Text = (decimal.Parse(edbBruttoPris.Text) * 100).ToString().Replace(",", ".");
                                        edbBruttoPris.Text = edbBruttoPris.Text.Trim().Replace(",", ".");
                                        edbBruttoPris.Text = edbBruttoPris.Text; //TODO Format(edbBruttoPris.Text, "@@@@@@@@@");
                                    }
                                    else if (sPrisfråga == "No")
                                    {
                                        edbBruttoPris.Text = sUrsprPris;
                                    }
                                }
                                CompOrder.Item("ogrNX1McEdit").Text = edbBruttoPris.Text.Trim().Replace(".", "");
                                LäsInAPris(); // Subanrop
                                LäsInRabatter(); // Subanrop
                                BeräknaNetto(); // Subanrop
                                break;
                            case "ogrLdtMcEdit": // Levdatum
                                LäsInAPris(); // Subanrop
                                LäsInRabatter(); // Subanrop
                                BeräknaNetto(); // Subanrop
                                break;
                        }
                    }
                    if (CompOrder.CurrentField == "edbRadText") // Lägger upp textvärdet som en extern orderradstext.
                    {
                        if (edbRadText.Text.Trim() != string.Empty)
                        {
                            oGarp.InsertOrderText(255, edbRadText.Text);
                            edbRadText.Text = string.Empty;
                            edbRadText.SetFocus();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("FieldExit " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }




        private void LoggaAntalsÄndringar() // Tillägg 170213 för Axfood. Logg av ändringar.
        {
            try { 
                if (decimal.Parse(CompOrder.Item("ogrOraMcEdit").Text) != cUrsprAntal) // Om antal ändrats
                {
                    if (oKundReg.Find(CompOrder.Item("knrEdit").Text))
                    {
                        if (oKundKat.Value == "50")
                        {
                        CompOrder.Item("ogaResMcEdit").Text = "4"; // Märkerar på OH att ändring skett
                            if (CompOrder.Item("ogrOraMcEdit").Text == "0") // Nytt antal = 0
                            CompOrder.Item("ogrResMcEdit").Text = "7"; // Ej accepterad
                            else
                            CompOrder.Item("ogrResMcEdit").Text = "3"; // Ändrat antal
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("LoggaAntalsÄndringar " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void LäsInRabatter()
        {
            string sNyckel = string.Empty;
            int iAvtal;
            int iRabTyp;
            int iPrisAlt;
            NollaRabatter(); // Subanrop
            LäsInRabattBen(); // Subanrop
            for (iAvtal = 1; iAvtal <= 2; iAvtal++)
            {
                for (iRabTyp = 0; iRabTyp <= 4; iRabTyp++) // Rabattyper
                {
                    SkapaTabNyckel(ref sNyckel, iAvtal, iRabTyp); // Subanrop
                    for (iPrisAlt = 5; iPrisAlt >= 0; iPrisAlt--) // Prisalternativ
                    {

                        sNyckel = sNyckel.Substring(0, 9) + iPrisAlt.ToString();
                        if (oTabReg.Find(sNyckel))
                        {

                            if (string.IsNullOrWhiteSpace(aoTabTxt[1].Value.Substring(13, 6)) ||
                                CompareDates(aoTabTxt[1].Value.Substring(13, 6).Trim(), CompOrder.Item("ogrLdtMcEdit").Text,"lte"))
                            {
                                if (string.IsNullOrWhiteSpace(aoTabTxt[1].Value.Substring(19, 6)) ||
                                    CompareDates(aoTabTxt[1].Value.Substring(19, 6).Trim(),CompOrder.Item("ogrLdtMcEdit").Text,"gte"))
                                {
                                    LäsInRabattBaser(iRabTyp); // Subanrop
                                      LäsInRabattTyper(iRabTyp); // Subanrop
                                }
                            }
                        }
                    }
                }
            }
        }



        private void SkapaTabNyckel(ref string sNyckel, int iAvtal, int iRabTyp)
        { 
            sNyckel = "9";

            if (iRabTyp > 0) // > Pallrabatt
            {
                if (iAvtal == 1)
                {
                    if (oKundReg.Find(CompOrder.Item("knrEdit").Text)) // Söker upp ev. avtalsnummer
                    {
                        sNyckel += oAvtalNr.Value.ToString(); // Läser in avtalsnummer (num 1)
                    }
                }
                else if (iAvtal == 2)
                {
                    sNyckel += CompOrder.Item("knrEdit").Text.Substring(0, 4); // &Kundnummer 4 tkn
                }
            }
            sNyckel = sNyckel.PadRight(5); // Fyller ut till 1+4tkn (kundnummer)

            if (iRabTyp < 4) // < Kundrabatt
            {
                if (oArtReg.Find(CompOrder.Item("ogrAnrMcEdit").Text))
                {
                    sNyckel += oArtNyckel.Value.ToString(); // Läser in artikelnyckel
                }
            }

            sNyckel = sNyckel.PadRight(8) + iRabTyp.ToString(); // Fyller ut till 8 tkn + rabattyp
        }


        private void LäsInRabattTyper(int iRabTyp)
        {
            int iX;
            decimal cPallRab;
            decimal cAntal;

            cAntal = Convert.ToDecimal(CompOrder.Item("ogrOraMcEdit").Text.Replace(".", ","));

            switch (iRabTyp)
            {
                case 0: // Pallrabatt
                    if (Convert.ToDecimal(aoTabNum[7].Value) <= Math.Abs(cAntal))
                    {
                        // Om Antal = Helpall
                        if (cAntal % Convert.ToDecimal(aoTabNum[7].Value) == 0) // Om antal = helpall
                        {
                            edbPallRab.Text = (Convert.ToDecimal(aoTabNum[1].Value) / 100).ToString("0.00").Replace(",", ".");
                        }
                        else
                        {
                            cPallRab = cAntal % Convert.ToDecimal(aoTabNum[7].Value); // Heltalsrest
                            cPallRab = cAntal - cPallRab; // Antal med rabatt = Helpall
                            cPallRab = cPallRab / cAntal; // Orderantal / antal med rabatt
                            cPallRab = cPallRab * Convert.ToDecimal(aoTabNum[1].Value); // Vägd rabatt
                            edbPallRab.Text = (cPallRab / 100).ToString("0.00").Replace(",", ".");
                        }
                        edbPallRab.Text = edbPallRab.Text.PadLeft(15); // Centrera värdet
                    }
                    else
                    {
                        edbPallRab.Text = "0.00"; // Nollrabatt
                    }
                    edbPallRab.Text = edbPallRab.Text.PadLeft(15); // Centrera värdet
                    break;
                case 1: // Kvant-/Hämtrabatt
                    for (iX = 5; iX >= 1; iX--)
                    {
                        if (Convert.ToDecimal(aoTabNum[iX].Value) != 0) // Om staffling finns
                        {
                            if (Convert.ToDecimal(aoTabNum[iX + 6].Value) <= Math.Abs(cAntal)) // Staffling <= Orderantal
                            {
                                edbKvantRab.Text = (Convert.ToDecimal(aoTabNum[iX].Value) / 100).ToString("0.00").Replace(",", ".");
                                break;
                            }
                        }
                        else
                        {
                            edbKvantRab.Text = "0.00"; // Nollrabatt
                        }
                    }
                    edbKvantRab.Text = edbKvantRab.Text.PadLeft(15); // Centrera värdet
                    break;
                case 2: // Avtalsrabatt
                    edbAvtalsRab.Text = (Convert.ToDecimal(aoTabNum[1].Value) / 100).ToString("0.00").Replace(",", ".");
                    edbAvtalsRab.Text = edbAvtalsRab.Text.PadLeft(15); // Centrera värdet
                    break;
                case 3: // Aktivitetsrabatt
                    for (iX = 5; iX >= 1; iX--)
                    {
                        if (Convert.ToDecimal(aoTabNum[iX].Value) != 0) // Om staffling finns
                        {
                            if (Convert.ToDecimal(aoTabNum[iX + 6].Value) <= Math.Abs(cAntal)) // Staffling <= Orderantal
                            {
                                edbAktRab.Text = (Convert.ToDecimal(aoTabNum[iX].Value) / 100).ToString("0.00").Replace(",", ".");
                                break;
                            }
                        }
                        else
                        {
                            edbAktRab.Text = "0.00"; // Nollrabatt
                        }
                    }
                    edbAktRab.Text = edbAktRab.Text.PadLeft(15); // Centrera värdet
                    break;
                case 4: // Kundrabatt
                    edbKundRab.Text = (Convert.ToDecimal(aoTabNum[1].Value) / 100).ToString("0.00").Replace(",", ".");
                    edbKundRab.Text = edbKundRab.Text.PadLeft(15); // Centrera värdet
                    break;
            }
        }



        private void NollaRabatter()
        {
            edbPallRab.Text = "0.00".PadLeft(15);
            edbKvantRab.Text = "0.00".PadLeft(15);
            edbAvtalsRab.Text = "0.00".PadLeft(15);
            edbAktRab.Text = "0.00".PadLeft(15);
            edbKundRab.Text = "0.00".PadLeft(15);
        }

        private void LäsInRabattBen() // Anropas från: LäsInRabatter
        {
            asRabattBas[0, 1] = "Pallrabatt";
            asRabattBas[1, 1] = "Kvant-/Hämtrabatt";
            asRabattBas[2, 1] = "Avtalsrabatt";
            asRabattBas[3, 1] = "Aktivitetsrabatt";
            asRabattBas[4, 1] = "Kundrabatt";
        }


        private void LäsInRabattBaser(int iRabTyp) // Anropas från: LäsInRabatter (arg 0 - 4)
        {
            if (Convert.ToDecimal(aoTabKod[1].Value) == 0)
            {
                asRabattBas[iRabTyp, 2] = "kr";
            }
            else if (Convert.ToDecimal(aoTabKod[1].Value) == 1)
            {
                asRabattBas[iRabTyp, 2] = "%";
            }

            switch (iRabTyp)
            {
                case 0:
                    lblPallRab.Text = "Helpall " + asRabattBas[iRabTyp, 2];
                    break;
                case 1:
                    lblKvantRab.Text = "Kvant/" + Environment.NewLine + "Hämt " + asRabattBas[iRabTyp, 2];
                    break;
                case 2:
                    lblAvtalsRab.Text = "Avtal " + asRabattBas[iRabTyp, 2];
                    break;
                case 3:
                    lblAktRab.Text = "Aktivitet " + asRabattBas[iRabTyp, 2];
                    break;
                case 4:
                    lblKundRab.Text = "Kund " + asRabattBas[iRabTyp, 2];
                    break;
            }
        }


        private void BeräknaNetto()
        {
            decimal cNettoBel;
            /*
            MessageBox.Show(asRabattBas[0, 2]);
            MessageBox.Show(asRabattBas[1, 2]);
            MessageBox.Show(asRabattBas[2, 2]);
            MessageBox.Show(asRabattBas[3, 2]);
            MessageBox.Show(asRabattBas[4, 2]);
            */

            try
            {
                // Replace '.' with ',' in the input and parse it as decimal
                cNettoBel = decimal.Parse(edbBruttoPris.Text.Replace(".", ","));

                if (asRabattBas[0, 2] == "kr") // Beloppsrabatt
                {
                    cNettoBel += decimal.Parse(edbPallRab.Text.Replace(".", ","));
                }
                else if (asRabattBas[0, 2] == "%") // Procentrabatt
                {
                    cNettoBel *= (1 + (decimal.Parse(edbPallRab.Text.Replace(".", ",")) / 100));
                }

                if (asRabattBas[1, 2] == "kr") // Beloppsrabatt
                {
                    cNettoBel += decimal.Parse(edbKvantRab.Text.Replace(".", ","));
                }
                else if (asRabattBas[1, 2] == "%") // Procentrabatt
                {
                    cNettoBel *= (1 + (decimal.Parse(edbKvantRab.Text.Replace(".", ",")) / 100));
                }

                if (asRabattBas[2, 2] == "kr") // Beloppsrabatt
                {
                    cNettoBel += decimal.Parse(edbAvtalsRab.Text.Replace(".", ","));
                }
                else if (asRabattBas[2, 2] == "%") // Procentrabatt
                {
                    cNettoBel *= (1 + (decimal.Parse(edbAvtalsRab.Text.Replace(".", ",")) / 100));
                }

                   
                if (asRabattBas[3, 2] == "kr") // Beloppsrabatt
                {
                    cNettoBel += decimal.Parse(edbAktRab.Text.Replace(".", ","));
                }
                else if (asRabattBas[3, 2] == "%") // Procentrabatt
                {
                    cNettoBel *= (1 + (decimal.Parse(edbAktRab.Text.Replace(".", ",")) / 100));
                }
                if (asRabattBas[4, 2] == "kr") // Beloppsrabatt
                {
                    cNettoBel += decimal.Parse(edbKundRab.Text.Replace(".", ","));
                }
                else if (asRabattBas[4, 2] == "%") // Procentrabatt
                {
                    cNettoBel *= (1 + (decimal.Parse(edbKundRab.Text.Replace(".", ",")) / 100));
                }

                cNettoBel = Math.Round(cNettoBel, 2);
    
                if (decimal.Parse(edbBruttoPris.Text.Replace(".", ",")) != 0)
                {
                  CompOrder.Item("ogrPriMcEdit").Text = cNettoBel.ToString("#0.00").Replace(",", ".");
                }

                CompOrder.Item("nettoprisLabel").Text = cNettoBel.ToString("#0.00").Replace(",", ".");


                // Om 0 i pris
                cNettoBel -= decimal.Parse(CompOrder.Item("ogrLVPMcEdit").Text.Replace(".", ","));
                cNettoBel = cNettoBel / decimal.Parse(CompOrder.Item("ogrPriMcEdit").Text.Replace(".", ",")) * 100; // TB / Nettopris = TG
                CompOrder.Item("nettoTbLabel").Text = cNettoBel.ToString("#0.00").Replace(",", "."); // TG
            }
            catch(Exception ex)
            {
                // Handle exception if needed
                MessageBox.Show("BeräknaNetto " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }




        private void RaderaRabattTexter() // Called from: Dispose
        {
            try
            {
                debugg("RaderaRabattTexter Find: " + sOrderNr + iOrderRadNr.ToString("D3") + "  0");
                oKOTxtReg.Find(sOrderNr + iOrderRadNr.ToString("D3") + "  0"); // Set pointer before 1st text
                do
                {
                    debugg("oKoTxtOrderNr.Value: " + oKoTxtOrderNr.Value + "-" + sOrderNr);
                    //MessageBox.Show("oKOTxtRadNr.Value: " + oKOTxtRadNr.Value + "-" + iOrderRadNr.ToString());
                    debugg("oKOTxt.Value: " + oKOTxt.Value);

                    if (oKoTxtOrderNr.Value == sOrderNr && !string.IsNullOrEmpty(oKOTxtRadNr.Value) && Convert.ToInt32(oKOTxtRadNr.Value) == iOrderRadNr) //TODO osäker på om oKOTxtRadNr.Value & iOrderRadNr har samma format här
                    {
                        debugg("oKOTxtFAFl.Value " + oKOTxtFAFl.Value);
                        if (oKOTxtFAFl.Value == "R")
                        {
                            oKOTxtReg.Delete(); // Delete all discount texts
                            oKOTxtReg.Prior();
                        }
                    }
                    else
                    {
                        debugg("No match on row " + oKOTxtRadNr.Value);
                        if (!string.IsNullOrEmpty(oKoTxtOrderNr.Value) && !string.IsNullOrEmpty(oKOTxtRadNr.Value))
                        {
                            debugg("break");
                            break;
                        }
                    }
                    debugg("FÖRE NEXT oKOTxtRadNr.Value: " + oKOTxtRadNr.Value);
                    oKOTxtReg.Next();
                    debugg("EFTER NEXT oKOTxtRadNr.Value: " + oKOTxtRadNr.Value);
                } while (!oKOTxtReg.Eof);
    
            }
            catch (Exception ex)
            {
                // Handle exception if needed
                MessageBox.Show("RaderaRabattTexter " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }



        private void LäggUppRabattTexter() // Called from: Dispose
        {
            try
            {

                //MessageBox.Show("LäggUppRabattTexter onr: " + sOrderNr + " -" + iOrderRadNr);
                for (int iX = 0; iX <= 4; iX++) // Log texts
                {
                    debugg("iX: " + iX);
                    string sStr = string.Empty;
                    switch (iX)
                    {
                        case 0:
                            if (edbPallRab.Text !="" && decimal.Parse(edbPallRab.Text.Replace(".", ",")) != 0)
                            {
                                sStr = asRabattBas[iX, 1]; // Discount type + % / kr
                                sStr += new string(' ', 16 - sStr.Length) + edbPallRab.Text + " " + asRabattBas[iX, 2];
                            }
                            break;
                        case 1:
                            if (edbKvantRab.Text != "" && decimal.Parse(edbKvantRab.Text.Replace(".", ",")) != 0)
                            {
                                sStr = asRabattBas[iX, 1].Substring(0, 14) + "."; // Discount type + % / kr
                                sStr += new string(' ', 16 - sStr.Length) + edbKvantRab.Text + " " + asRabattBas[iX, 2];
                            }
                            break;
                        case 2:
                            if (edbAvtalsRab.Text != "" && decimal.Parse(edbAvtalsRab.Text.Replace(".", ",")) != 0)
                            {
                                sStr = asRabattBas[iX, 1]; // Discount type + % / kr
                                sStr += new string(' ', 16 - sStr.Length) + edbAvtalsRab.Text + " " + asRabattBas[iX, 2];
                            }
                            break;
                        case 3:
                            if (edbAktRab.Text != "" && decimal.Parse(edbAktRab.Text.Replace(".", ",")) != 0)
                            {
                                sStr = asRabattBas[iX, 1]; // Discount type + % / kr
                                sStr += new string(' ', 16 - sStr.Length) + edbAktRab.Text + " " + asRabattBas[iX, 2];
                            }
                            break;
                        case 4:
                            if (edbKundRab.Text != "" && decimal.Parse(edbKundRab.Text.Replace(".", ",")) != 0)
                            {
                                sStr = asRabattBas[iX, 1]; // Discount type + % / kr
                                sStr += new string(' ', 16 - sStr.Length) + edbKundRab.Text + " " + asRabattBas[iX, 2];
                            }
                            break;
                    }
                    if (!string.IsNullOrEmpty(sStr))
                    {
                        debugg("LäggUppRabattTexter sStr: " + sStr + "-" + sOrderNr);

                        //oKOTxtReg.Find(sOrderNr + iOrderRadNr.ToString("D3") + "  0"); // Set pointer before 1st text

                        oKOTxtReg.Insert();

                        oKoTxtOrderNr.Value = sOrderNr;
                        oKOTxtRadNr.Value = iOrderRadNr.ToString("D3");
                        oKOTxtSekvNr.Value = "255";
                        oKOTxt.Value = sStr;
                        oKOTxtOBFl.Value = "R";
                        oKOTxtPLFl.Value = "R";
                        oKOTxtFSFl.Value = "R";
                        oKOTxtFAFl.Value = "R";

                        oKOTxtReg.Post();
                        //MessageBox.Show("Postade in ny text: " + sStr + "-" + sOrderNr +"-" + iOrderRadNr.ToString("D3"));

                        //oGarp.InsertOrderTextEx(255, sStr, "R", "R", "R", "R");
                    }

                }

                Array.Clear(asRabattBas, 0, asRabattBas.Length);
            }
            catch (Exception ex)
            {
                // Handle exception if needed
                MessageBox.Show("LäggUppRabattTexter " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void on_AfterScrollOrder()
        {
            if (!disposed)
            {
                try
                {
                    debugg("on_AfterScrollOrder");
                    sOrderNr = CompOrder.Item("onrEdit").Text;
                }
                catch (Exception ex)
                {
                    // Handle exception if needed
                    MessageBox.Show("on_AfterScrollOrderRow " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

            }
        }

        private void on_AfterScrollOrderRow()
        {
            if (!disposed)
            {
                try
                {
                    //MessageBox.Show("AA");
                    //MessageBox.Show("on_AfterScrollOrderRow sOrderNr: " + sOrderNr);
                    //MessageBox.Show("on_AfterScrollOrderRow iLevFlagga: " + iLevFlagga);
                    //MessageBox.Show(string.IsNullOrEmpty(sOrderNr).ToString());
                    //sOrderNr is not blank if we scrolled from a row which was a rabattrow so save orderrowtext
                    if (false && !string.IsNullOrEmpty(sOrderNr) && iLevFlagga < 5)
                    {
                        debugg("on_AfterScrollOrderRow sOrderNr: " + sOrderNr);

                        RaderaRabattTexter();
                        debugg("EEE");
                        //MessageBox.Show("edbBruttoPris.Text: " + edbBruttoPris.Text);
                        if (edbBruttoPris.Text != "" && decimal.Parse(edbBruttoPris.Text.Replace(".",",")) != 0)
                        {
                            LäggUppRabattTexter();
                            if (edbBruttoPris.Text.Trim() != sUrsprPris.Trim())
                            {
                                MessageBox.Show("Kolla rabatten. Kan vara dubbelräknad.", "Varning!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            }
                        }
                    }

                    if (!checkNewRow())
                    {
                        //MessageBox.Show("new row");
                        
                        if (oArtReg.Find(CompOrder.Item("ogrAnrMcEdit").Text) && Convert.ToInt32(CompOrder.Item("ogrLvfMcedit").Text.Substring(0, 1)) < 5)
                        {
                            if (oArtReg.Fields.Item("KD1").Value == "R")
                            {
                                //MessageBox.Show("AKTIVERA rabatthanteringen " + CompOrder.Item("ogrAnrMcEdit").Text);

                                InitForRabatt();

                                if (!string.IsNullOrEmpty(sOrderNr) && iLevFlagga < 5)
                                {
                                    RaderaRabattTexter();
                                    debugg("edbBruttoPris.Text: " + edbBruttoPris.Text);
                                    if (edbBruttoPris.Text != "" && decimal.Parse(edbBruttoPris.Text.Replace(".", ",")) != 0)
                                    {
                                        LäggUppRabattTexter();
                                        if (edbBruttoPris.Text.Trim() != sUrsprPris.Trim())
                                        {
                                            //MessageBox.Show("Kolla rabatten. Kan vara dubbelräknad.", "Varning!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        }
                                    }
                                }
                            }
                            else {
                                //MessageBox.Show("IN-AKTIVERA rabatthanteringenA " + CompOrder.Item("ogrAnrMcEdit").Text);
                                InitForNoRabatt();
                            }
                        }
                        else
                        {
                            //MessageBox.Show("IN-AKTIVERA rabatthanteringenB " + CompOrder.Item("ogrAnrMcEdit").Text);
                            InitForNoRabatt();
                        }
                    }
  
                }
                catch (Exception ex)
                {
                    MessageBox.Show("on_AfterScrollOrderRow " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }


        private bool checkNewRow()
        {
            try
            {
                //then, save some data!
                string savedRowNo = currentRowNo;
                currentRowNo = dsOGR.Fields.Item("RDC").Value;

                return (savedRowNo != currentRowNo);

            }
            catch (Exception e)
            {
                return false;
            }
        }



        //Informat yymmdd
        private bool CompareDates(string indate1, string indate2, string comp)
        {
            try
            {
                switch (comp)
                {
                    case "lte":
                        return DateTime.ParseExact("20" + indate1, "yyyyMMdd", CultureInfo.InvariantCulture) <= DateTime.ParseExact("20" + indate2, "yyyyMMdd", CultureInfo.InvariantCulture);
                    case "lt":
                        return DateTime.ParseExact("20" + indate1, "yyyyMMdd", CultureInfo.InvariantCulture) < DateTime.ParseExact("20" + indate2, "yyyyMMdd", CultureInfo.InvariantCulture);
                    case "gte":
                        return DateTime.ParseExact("20" + indate1, "yyyyMMdd", CultureInfo.InvariantCulture) >= DateTime.ParseExact("20" + indate2, "yyyyMMdd", CultureInfo.InvariantCulture);
                    case "gt":
                        return DateTime.ParseExact("20" + indate1, "yyyyMMdd", CultureInfo.InvariantCulture) > DateTime.ParseExact("20" + indate2, "yyyyMMdd", CultureInfo.InvariantCulture);
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CompareDates " + ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }


        private void debugg(string txt)
        {
            if (blndebugg)
            {
                MessageBox.Show(txt);
            }
            
        }


        protected virtual void Dispose(bool disposing)
        {
            try
            {
                if (!disposed)
                {
                    if (disposing)
                    {
                        if (oGarp != null)
                        {
                            oGarp.FieldExit -= FieldExit;
                            oGarp.FieldEnter -= FieldEnter;

                           //dsOGA.AfterScroll -= on_AfterScrollOrder;
                            dsOGR.AfterScroll -= on_AfterScrollOrderRow;

                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(dsOGA);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(dsOGR);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGarp);
                        }
                    }
                }
                disposed = true;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Forms", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        ~Order()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

    }

   

}