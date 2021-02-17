using System;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Geared.Winforms.SpeedTest;
using LiveCharts.Geared;
using LiveCharts.Wpf;
using System.Windows.Media;
using Net.SourceForge.Koogra.Excel;
using System.Data.OleDb;

namespace Test
{
    public partial class Form1 : Form
    {

        private SpeedTestVm _viewModel = new SpeedTestVm();
        //string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"C:\\Users\\maintenance1\\Desktop\\Automatisation des numero de programme.xlsx" + ";Extended Properties=\"Excel 8.0\";";
        //string filePath = @"C:\\Users\\maintenance1\\Desktop\\BDD_COEX.xls";
        //Workbook workbook = new Workbook(filePath);
        const string File = @"D:\Workspace\Visual Studio\ProgCOEX\src\SimpleClientDA\BDD\COEX.xls";
        Worksheet worksheet = new Workbook(File).Sheets.GetByName("BDD");
        Worksheet BDDCoexGFH = new Workbook(File).Sheets.GetByName("GFH");
        Worksheet BDDCoexDENSITE = new Workbook(File).Sheets.GetByName("DENSITE");
        Worksheet BDDCoexPLC = new Workbook(File).Sheets.GetByName("BDD_Vitesse");
        Worksheet BDDCoexZumbach = new Workbook(File).Sheets.GetByName("BDD_Zumbach");
        Worksheet BDDLongueur = new Workbook(File).Sheets.GetByName("LongueurRajout");
        Worksheet BDDTemperature = new Workbook(File).Sheets.GetByName("BDD_Temperature");
        Worksheet BDDVitesse = new Workbook(File).Sheets.GetByName("BDD_VitessePROD");
        Worksheet BDDEnregistrementOF = new Workbook(File).Sheets.GetByName("EnregistrementOF");
        Worksheet BDDPID = new Workbook(File).Sheets.GetByName("BDD_PID");

        private Resources.properties ConfProperties = Resources.properties.Default;
        private string e1_consigne_vitesse_rotation = "";
        private string e2_consigne_vitesse_rotation= "";
        private string e3_consigne_vitesse_rotation = "";
        private string e4_consigne_vitesse_rotation= "";
        private string banc_tirage_consigne = "";

        public Form1()
        {
            InitializeComponent(); 
            cartesianChart1.Series.Add(new GLineSeries
            {
                Values = _viewModel.Values,
                StrokeThickness = 2,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(28, 142, 196)),
                Fill = System.Windows.Media.Brushes.Transparent,
                LineSmoothness = 1,
                PointGeometrySize = 15,
                PointForeground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(34, 46, 49))
            });
            cartesianChart1.DisableAnimations = true;
            cartesianChart1.AxisY.Add(
            new Axis
            {
                MinValue = -0.05,
                MaxValue = 0.05,
                Sections = new SectionsCollection
                {
                    new AxisSection
                    {
                        Value = 0,
                        Stroke = new SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 213, 72))
                    },
                    new AxisSection

                    {
                        Label = "Good",
                        Value = -0.03,
                        SectionWidth = 0.06,
                        Fill = new SolidColorBrush
                        {
                            Color = System.Windows.Media.Color.FromArgb(100,10,204,10),
                            Opacity = .9
                        }
                    },
                    new AxisSection
                    {
                        Value = -0.05,
                        SectionWidth = 0.02,
                        Fill = new SolidColorBrush
                        {
                            Color = System.Windows.Media.Color.FromArgb(50,254,132,132),
                            Opacity = 10
                        }
                    },
                    new AxisSection
                    {
                        Value = 0.03,
                        SectionWidth = 0.02,
                        Fill = new SolidColorBrush
                        {
                            Color = System.Windows.Media.Color.FromArgb(100,254,132,132),
                            Opacity = 50
                        }
                    }
                }
            });
        }

        void ReadDBVitesse(string numPorg)
        {
            int numeroProg;
            bool IsNumeric = int.TryParse(numPorg, out numeroProg);

            if (IsNumeric == true)
            {
                int DBW = (numeroProg * 16) + 16; //16 = 30-14
                int W1 = DBW;
                int W2 = DBW + 2;
                int W3 = DBW + 4;
                int W4 = DBW + 6;
                int W5 = DBW + 8;
                e1_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W" + W1; //Extru 1
                e2_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W" + W2; //Extru 2
                e3_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W" + W3; //Extru 3
                e4_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W" + W4;//Extru 4
                banc_tirage_consigne = "S7:[Liaison_S7_1],DB31,W" + W5;  //Extru Banc Tirage
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.cpb4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(128)),
                     ((System.Byte)(255)), ((System.Byte)(255)));
            this.cpb4.ForeColor = System.Drawing.Color.Black;
            OFFEnCour();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _viewModel.Read();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _viewModel.Stop();
        }

        private void TestWriteExcel()
        {
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + File + ";Extended Properties=\"Excel 8.0\";";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                // Le nom des colonnes se trouve en première ligne de la feuille. 
                // Ici les colonnes concernées s'intitulent COL 1 et COL 2 : 
                String NameCol1 = "NumeroProcessus";
                String NameCol2 = "GFH";
                String NameCol3 = "Densité";
                String NameCol4 = "Longueur";
                String NameCol5 = "Tete";
                String NameCol6 = "Type";
                String NameCol7 = "Matiere";
                String NameCol8 = "DebitEau1";
                String NameCol9 = "DebitEau2";
                String NameCol10 = "DebitEau3";
                String NameCol11= "Outillage";
                String NameCol12= "EcartementBT";
                String NameCol13= "ColorVis1";
                String NameCol14= "ColorVis2";
                String NameCol15= "VitesseVis1";
                String NameCol16 = "VitesseVis2";
                String NameCol17 = "VitesseVis3";
                String NameCol18 = "VitesseVis4";
                String NameCol19 = "VitesseVisBT";
                String NameCol20 = "Videbac1";
                String NameCol21 = "Videbac2";

                String DataCol1 = "TEST 1";
                String DataCol2 = "TEST 2";
                String DataCol3 = "TEST 2";
                String DataCol4 = "TEST 2";
                String DataCol5 = "TEST 2";
                String DataCol6 = "TEST 2";
                String DataCol7 = "TEST 2";
                String DataCol8 = "TEST 2";
                String DataCol9 = "TEST 2";
                String DataCol10 = "TEST 2";
                String DataCol11= "TEST 2";
                String DataCol12 = "TEST 2";
                String DataCol13 = "TEST 2";
                String DataCol14 = "TEST 2";
                String DataCol15 = "TEST 2";
                String DataCol16 = "TEST 2";
                String DataCol17 = "TEST 2";
                String DataCol18 = "TEST 2";
                String DataCol19 = "TEST 2";
                String DataCol20 = "TEST 2";
                String DataCol21 = "TEST 2";

                string cmdText = "INSERT INTO [EnregistrementOF$] " +
                    "(["+ NameCol1+ "], [" + NameCol2 + "], [" + NameCol3 + "], [" + NameCol4 + "], [" + NameCol5 + "], [" + NameCol6 + "], [" + NameCol7 + "], [" + NameCol8 + "], [" + NameCol9 + "], [" + NameCol10 + "], [" + NameCol11 + "], [" + NameCol12 + "], [" + NameCol13 + "], [" + NameCol14 + "], [" + NameCol15 + "], [" + NameCol16 + "], [" + NameCol17 + "], [" + NameCol18 + "], [" + NameCol19 + "], [" + NameCol20 + "], [" + NameCol21 + "])" +
                    "VALUES " +
                    "('" + DataCol1 + "', '" + DataCol2 + "', '" + DataCol3 + "', '" + DataCol4 + "', '" + DataCol5 + "', '" + DataCol6 + "', '" + DataCol7 + "', '" + DataCol8 + "', '" + DataCol9 + "', '" + DataCol10 + "', '" + DataCol11 + "', '" + DataCol12 + "', '" + DataCol13 + "', '" + DataCol14 + "', '" + DataCol15 + "', '" + DataCol16 + "', '" + DataCol17 + "', '" + DataCol18 + "', '" + DataCol19 + "', '" + DataCol20 + "', '" + DataCol21 + "')";
                using (OleDbCommand command = new OleDbCommand(cmdText, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //_viewModel.Values.Clear();
            TestWriteExcel();
        }

        private int gestionTemperature(string intituler, bool green)
        {
            for (ushort i = 1; i <= BDDTemperature.Rows.LastRow; i++)
            {
                if (green == false)
                {
                    if (BDDTemperature.Rows[i].Cells[2].Value.ToString() == intituler)
                    {
                        return Convert.ToInt32(BDDTemperature.Rows[i].Cells[3].Value.ToString());
                        //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                }
                else if(green == true)
                {
                    if (BDDTemperature.Rows[i].Cells[2].Value.ToString() == intituler)
                    {
                        return Convert.ToInt32(BDDTemperature.Rows[i].Cells[4].Value.ToString());
                        //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                } 
            }
            return 0;
        }

        private void Test()
        {
            float longueurTube = (float)Convert.ToDouble("105,5") * 100;
            Console.WriteLine(longueurTube);
        }

        private void searchVitesse()
        {
            for (ushort i = 1; i <= BDDVitesse.Rows.LastRow; i++)
            {
                if (textBox12.Text.Contains(BDDVitesse.Rows[i].Cells[0].Value.ToString()))
                {
                    textBox13.Text = BDDVitesse.Rows[i].Cells[3].Value.ToString();
                    ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }

            }
        }

        private void NewVitessse()
        {
            MessageBox.Show("VITESSE PROG DEMARRAG");
            //VITESSE PROG DEMARRAGE ------------------------------------------------------------------------------------
            float VitesseParDefaultE1 = (float)Convert.ToDouble("137");
            float VitesseParDefaultE2 = (float)Convert.ToDouble(("220"));
            float VitesseParDefaultE3 = (float)Convert.ToDouble(("170"));
            float VitesseParDefaultE4 = (float)Convert.ToDouble(("267"));
            float VitesseParDefaultBT = (float)Convert.ToDouble(("1560"));
            //VITESSE PROG DEMARRAGE ------------------------------------------------------------------------------------

            MessageBox.Show("WritePLCnextProg");
            //WritePLC(nextProg, lblProgProd.Text);

            MessageBox.Show("ReadDBVitesselblProgProd");
            ReadDBVitesse("1");
            Console.WriteLine("Avant : "+e1_consigne_vitesse_rotation + " " + e2_consigne_vitesse_rotation + " " + e3_consigne_vitesse_rotation + " " + e4_consigne_vitesse_rotation + " " + banc_tirage_consigne);
            ReadDBVitesse("1");
            Console.WriteLine("Aprés : " + e1_consigne_vitesse_rotation + " " + e2_consigne_vitesse_rotation + " " + e3_consigne_vitesse_rotation + " " + e4_consigne_vitesse_rotation + " " + banc_tirage_consigne);

            MessageBox.Show("WritePLCtimer");
            //WritePLC(timer, "18000");

            MessageBox.Show("NewVitessetxtCadenceProd");
            float NewVitesse = (float)Convert.ToDouble(textBox12.Text); //New cadencce
            MessageBox.Show("Vitessecadence_coupe");
            float Vitesse = (float)Convert.ToDouble("150"); // Cadence actuelle

            MessageBox.Show("VITESSE PROG PROD");
            //VITESSE PROG PROD -----------------------------------------------------------------------------------------
            int NewVitesseParDefaultE1 = (int)Math.Round(((NewVitesse * VitesseParDefaultE1) / Vitesse), 0);
            int NewVitesseParDefaultE2 = (int)Math.Round(((NewVitesse * VitesseParDefaultE2) / Vitesse), 0);
            int NewVitesseParDefaultE3 = (int)Math.Round(((NewVitesse * VitesseParDefaultE3) / Vitesse), 0);
            int NewVitesseParDefaultE4 = (int)Math.Round(((NewVitesse * VitesseParDefaultE4) / Vitesse), 0);
            int NewVitesseParDefaultBT = (int)Math.Round(((NewVitesse * VitesseParDefaultBT) / Vitesse), 0);
            //VITESSE PROG PROD -----------------------------------------------------------------------------------------

            MessageBox.Show("WritePLC e1_consigne_vitesse_rotation");
            Console.WriteLine("NewVitesseParDefaultE1 "+NewVitesseParDefaultE1.ToString());
            MessageBox.Show("WritePLC e2_consigne_vitesse_rotation");
            Console.WriteLine("NewVitesseParDefaultE2 " + NewVitesseParDefaultE2.ToString());
            MessageBox.Show("WritePLC e3_consigne_vitesse_rotation");
            Console.WriteLine("NewVitesseParDefaultE3 " + NewVitesseParDefaultE3.ToString());
            MessageBox.Show("WritePLC e4_consigne_vitesse_rotation");
            Console.WriteLine("NewVitesseParDefaultE4 " + NewVitesseParDefaultE4.ToString());
            MessageBox.Show("WritePLC banc_tirage_consigne");
            Console.WriteLine("NewVitesseParDefaultBT " + NewVitesseParDefaultBT.ToString()); //PB

            MessageBox.Show("Clique BTN");
            //WritePLC(btnK1, "true"); //Clique bouton K1
            MessageBox.Show("Relache BTN");
            //WritePLC(btnK1, "false"); //Relache bouton K1

            MessageBox.Show("Les prog PRODUCTION vient d'être charger.\n" +
                            "Les nouvelles vitesses des vis ont étais chargés.\n\n" +
                            "Les paramètres ainsi que les recettes vont être enregistés toutes les 10min.\n\n" +
                            "Vous pouvez aussi changer la 'Vitesse de PROD' en appuyant sur l'ONGLET 'DASHBOARD'\n\n" +
                            "Merci de cliquer sur 'Nouvelle PRODUCTION'\n" +
                            "Une fois la ligne prête à changer de commande.", "Information");

        }


        private void OFFEnCour()
        {
            if(ConfProperties.Processus != "")
            {
                textBox1.Text = ConfProperties.Processus;
                textBox2.Text = ConfProperties.Diametre;
                textBox3.Text = ConfProperties.Densite;
                textBox4.Text = ConfProperties.GFH;
                textBox10.Text = ConfProperties.Longueur;
            }
        }

        private void EnregConfOFFEnCour()
        {
            ConfProperties.Processus = textBox1.Text;
            ConfProperties.Diametre = textBox2.Text;
            ConfProperties.Densite = textBox3.Text;
            ConfProperties.GFH = textBox4.Text;
            ConfProperties.Longueur = textBox10.Text;
            Console.WriteLine(ConfProperties.Processus);
            ConfProperties.Save();
        }

        private void ChoixPID()
        {

            string diametre = textBox2.Text;//textBox4.Text;
            string densite = textBox3.Text;
            Console.WriteLine(diametre + " " + densite);
            for (ushort i = 1; i <= BDDPID.Rows.LastRow; i++)
            {
                if (BDDPID.Rows[i].Cells[0].Value.ToString() == diametre && densite.Contains(BDDPID.Rows[i].Cells[1].Value.ToString()) && BDDPID.Rows[i].Cells[2].Value.ToString() == "BAC 1")
                {
                    Console.WriteLine("BAC 1");
                    Console.WriteLine(BDDPID.Rows[i].Cells[3].Value.ToString());
                    Console.WriteLine(BDDPID.Rows[i].Cells[4].Value.ToString());
                }
            }
            for (ushort i = 1; i <= BDDPID.Rows.LastRow; i++)
            {
                if (BDDPID.Rows[i].Cells[0].Value.ToString() == diametre && densite.Contains(BDDPID.Rows[i].Cells[1].Value.ToString()) && BDDPID.Rows[i].Cells[2].Value.ToString() == "BAC 2")
                {
                    Console.WriteLine("BAC 2");
                    Console.WriteLine(BDDPID.Rows[i].Cells[3].Value.ToString());
                    Console.WriteLine(BDDPID.Rows[i].Cells[4].Value.ToString());

                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //ReadExcelBDD();
            //Longueur();
            //Test();
            //bool test = false;
            //if (comboBox1.Text == "true")
            //{
            //    test = true;
            //}
            //EnregConfOFFEnCour();
            ChoixPID();

            //textBox13.Text = gestionTemperature(textBox12.Text, test).ToString();
            //searchVitesse();
            //NewVitessse();


            //ushort ligne = (ushort)(Convert.ToUInt16("7"));
            //string VitesseParDefaultE1 = BDDCoexPLC.Rows[ligne].Cells[3].Value.ToString();
            //string VitesseParDefaultE2 = BDDCoexPLC.Rows[ligne].Cells[4].Value.ToString();
            //string VitesseParDefaultE3 = BDDCoexPLC.Rows[ligne].Cells[5].Value.ToString();
            //string VitesseParDefaultE4 = BDDCoexPLC.Rows[ligne].Cells[6].Value.ToString();
            //string VitesseParDefaultBT = BDDCoexPLC.Rows[ligne].Cells[7].Value.ToString();

            //Console.WriteLine(VitesseParDefaultE1 + " " + VitesseParDefaultE2 + " " + VitesseParDefaultE3 + " " + VitesseParDefaultE4 + " " + VitesseParDefaultBT);


        }
        private void WriteExcelBDD()
        {

        }

        private string ChoixProgZumbach()
        {

            string nomProduit = GFH(textBox4.Text, 2);//textBox4.Text;
            nomProduit = nomProduit + " D" + textBox2.Text;
            Console.WriteLine(nomProduit);
            string densite = textBox3.Text;
            Console.WriteLine(densite);
            for (ushort i = 1; i <= BDDCoexZumbach.Rows.LastRow; i++)
            {
                if (BDDCoexZumbach.Rows[i].Cells[2].Value.ToString() == nomProduit && BDDCoexZumbach.Rows[i].Cells[3].Value.ToString() == densite)
                {
                    return BDDCoexZumbach.Rows[i].Cells[1].Value.ToString();
                    break;

                    //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }
            }
            return "Aucune PROG Trouvé";
        }

        private string ChoixProgPLC()
        {
            string diametre = textBox2.Text;//textBox4.Text;
            string densite = textBox3.Text;
            for (ushort i = 1; i <= BDDCoexPLC.Rows.LastRow; i++)
            {
                if (BDDCoexPLC.Rows[i].Cells[1].Value.ToString() == diametre && BDDCoexPLC.Rows[i].Cells[2].Value.ToString() == densite)
                {
                    return BDDCoexPLC.Rows[i].Cells[0].Value.ToString();
                    break;

                    //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }
            }
            return "Aucun PROG Trouvé";
        }

        private bool ReadExcelBDD()
        {
            bool trouver = false;
            for (ushort i = 1; i <= worksheet.Rows.LastRow; i++)
            {
                if (worksheet.Rows[i].Cells[2].Value.ToString() == textBox1.Text)
                {
                    Console.WriteLine("Numero processus : " + worksheet.Rows[i].Cells[2].Value.ToString());
                    Console.WriteLine("Diametre tube : " + worksheet.Rows[i].Cells[4].Value.ToString());
                    textBox2.Text = worksheet.Rows[i].Cells[4].Value.ToString();
                    Console.WriteLine("Longueur tube : " + worksheet.Rows[i].Cells[5].Value.ToString());
                    textBox10.Text = worksheet.Rows[i].Cells[5].Value.ToString();
                    Console.WriteLine("Densité tube : " + worksheet.Rows[i].Cells[9].Value.ToString());
                    Console.WriteLine("Destination tube : " + worksheet.Rows[i].Cells[11].Value.ToString().Substring(2, 3));
                    Console.WriteLine("i : "+i+1);
                    trouver = true;
                    DENSITE(worksheet.Rows[i].Cells[9].Value.ToString());
                    GFH(worksheet.Rows[i].Cells[11].Value.ToString(),1);
                    break;

                //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }
            }

            

            if (trouver == true)
            {
                Console.WriteLine("\nNumero de processus trouvé");
                return true;
            } else
            {
                Console.WriteLine("\nNumero de processus non trouvé");
                return false;
            }
        }
        private string DENSITE(string norme)
        {
            for (ushort i = 1; i <= BDDCoexDENSITE.Rows.LastRow; i++)
            {
                if (BDDCoexDENSITE.Rows[i].Cells[0].Value.ToString() == norme)
                {
                    Console.WriteLine("Numero processus : " + BDDCoexDENSITE.Rows[i].Cells[0].Value.ToString());
                    Console.WriteLine("Diametre tube : " + BDDCoexDENSITE.Rows[i].Cells[1].Value.ToString());
                    Console.WriteLine("Longueur tube : " + BDDCoexDENSITE.Rows[i].Cells[2].Value.ToString());
                    textBox3.Text = BDDCoexDENSITE.Rows[i].Cells[2].Value.ToString();
                    Console.WriteLine("i : " + i + 1);
                    return BDDCoexDENSITE.Rows[i].Cells[2].Value.ToString();
                    break;

                    //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }
            }
            return "";
        }
        private string GFH(string norme, int choix)
        {
            if (choix == 1)
            {
                for (ushort i = 1; i <= BDDCoexGFH.Rows.LastRow; i++)
                {
                    if (BDDCoexGFH.Rows[i].Cells[0].Value.ToString() == norme)
                    {
                        Console.WriteLine("Norme : " + BDDCoexGFH.Rows[i].Cells[0].Value.ToString());
                        Console.WriteLine("Nom : " + BDDCoexGFH.Rows[i].Cells[1].Value.ToString());
                        Console.WriteLine("code : " + BDDCoexGFH.Rows[i].Cells[2].Value.ToString());
                        Console.WriteLine("i : " + i + 1);
                        textBox4.Text = BDDCoexGFH.Rows[i].Cells[2].Value.ToString();// + " D" + textBox2.Text;
                        return BDDCoexGFH.Rows[i].Cells[2].Value.ToString();

                        break;

                        //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                }
            }
            else if (choix == 2)
            {
                for (ushort i = 1; i <= BDDCoexGFH.Rows.LastRow; i++)
                {
                    if (BDDCoexGFH.Rows[i].Cells[2].Value.ToString() == norme)
                    {
                        Console.WriteLine("Norme : " + BDDCoexGFH.Rows[i].Cells[0].Value.ToString());
                        Console.WriteLine("Nom : " + BDDCoexGFH.Rows[i].Cells[1].Value.ToString());
                        Console.WriteLine("code : " + BDDCoexGFH.Rows[i].Cells[2].Value.ToString());
                        Console.WriteLine("i : " + i + 1);
                        return BDDCoexGFH.Rows[i].Cells[1].Value.ToString();

                        break;

                        //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                }
            }
            return "";
        }

        private void Longueur()
        {
            string diametre = textBox2.Text;
            string densite = textBox3.Text;
            string gfh = textBox4.Text;
            string typeTete = textBox7.Text;
            Double LongueurSansTete = (int)Convert.ToInt32(textBox10.Text);
            string LongueurTete = "";
            Console.WriteLine("diametre: '" + diametre + "' densite: '" + densite+ "' gfh: '"+ gfh+ "' typeTete: '" + typeTete+"'");
            Console.WriteLine("\n--------------------------------------------------\n");

            for (ushort i = 0; i <= BDDLongueur.Rows.LastRow-18; i++)
            {
                //Console.WriteLine(BDDLongueur.Rows.LastRow);=
                if (BDDLongueur.Rows[i].Cells[6].Value.ToString() == typeTete)
                {
                    typeTete = BDDLongueur.Rows[i].Cells[7].Value.ToString();
                    Console.WriteLine(typeTete);
                }
            }

            Console.WriteLine("diametre: '" + diametre + "' densite: '" + densite+ "' typeTete: '" + typeTete + "'");

            for (ushort i = 0; i <= BDDLongueur.Rows.LastRow; i++)
            {//BDDLongueur.Rows[i].Cells[1].Value.ToString() == gfh && 
                //Console.WriteLine("diametre: '" + diametre + "' densite: '" + densite + "' typeTete: '" + typeTete + "'");
                if (BDDLongueur.Rows[i].Cells[3].Value.ToString() == densite && BDDLongueur.Rows[i].Cells[5].Value.ToString().Contains(diametre) && BDDLongueur.Rows[i].Cells[2].Value.ToString() == typeTete)
                {
                    LongueurTete = BDDLongueur.Rows[i].Cells[4].Value.ToString();
                    Console.WriteLine("OK");
                    break;
                }
            }

            Console.WriteLine(LongueurTete);
            //LongueurTete = LongueurTete.Replace(",", ".");
            //Console.WriteLine(LongueurTete);

            float test = (float)Convert.ToDouble(LongueurTete);

            textBox11.Text = (LongueurSansTete+ test).ToString();
        }


        private void TestNumero()
        {
            int numProgramme = (int)Convert.ToDouble(textBox1.Text);
            Console.WriteLine(numProgramme);

            int DBW = (numProgramme * 16) + 16; //16 = 30-14
            label1.Text = DBW.ToString(); //Extru 1
            label2.Text = (DBW + 2).ToString(); //Extru 2
            label3.Text = (DBW + 4).ToString(); //Extru 3
            label4.Text = (DBW + 6).ToString(); //Extru 4
            label5.Text = (DBW + 8).ToString(); //Extru Banc Tirage
        }

        private void cartesianChart1_ChildChanged(object sender, System.Windows.Forms.Integration.ChildChangedEventArgs e)
        {

        }

        private void cpb4_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string[] plop = textBox1.Text.Split('/');
            if (plop.Length == 2)
            {
                textBox6.Text = plop[1].Length.ToString();
                if (Int32.Parse(plop[1].Length.ToString()) == 2)
                {
                    bool TestProcessus = ReadExcelBDD();
                    if(TestProcessus == true)
                    {
                        MessageBox.Show("Numero de processus trouvé","Information");
                    }
                    else
                    {
                        MessageBox.Show("Numero de processus non trouvé.\nMerci de remplir les informations manuellements.", "Information");
                    }
                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox5.Text = ChoixProgPLC();

            int ProgPLC = (int)Convert.ToInt32(ChoixProgPLC());
            Console.WriteLine(ProgPLC % 3);
            Console.WriteLine(ProgPLC % 4);


            textBox9.Text = ChoixProgPLC();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '/'))
            {
                e.Handled = true;
            }
            if ((e.KeyChar == '/') && ((sender as TextBox).Text.IndexOf('/') > -1))
            {
                e.Handled = true;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //Longueur();
            
            HashMDP();
        }
        
        private string SelectMDP()
        {
            if (txtUser.Text != "")
            {
                Worksheet BDDCompte = new Workbook(File).Sheets.GetByName("Compte");
                int te = 0;
                //sSourceData = txtUser.Text;
                //tmpSource = ASCIIEncoding.ASCII.GetBytes(sSourceData);0
                //tmpHash = new MD5CryptoServiceProvider().ComputeHash(tmpSource);
                //Console.WriteLine("MDP : " + ByteArrayToString(tmpHash));
                for (ushort i = 0; i <= BDDCompte.Rows.LastRow; i++)
                {
                    if (BDDCompte.Rows[i].Cells[0].Value.ToString() == txtUser.Text)
                    {
                        string tmpStringHash = BDDCompte.Rows[i].Cells[1].Value.ToString();
                        //Console.WriteLine(tmpStringHash + " OK");
                        return tmpStringHash;
                    }
                }
            }
            return null;
        }

        protected void HashMDP()
        {
            string sSourceData;
            byte[] tmpSource;
            string tmpHash = SelectMDP();
            byte[] tmpNewHash;

            //sSourceData = txtUser.Text;
            //tmpSource = ASCIIEncoding.ASCII.GetBytes(sSourceData);
            //tmpHash = new MD5CryptoServiceProvider().ComputeHash(tmpSource);
            //Console.WriteLine("MDP : " + ByteArrayToString(tmpHash));

            if (tmpHash != null) {
                sSourceData = txtPasswd.Text;
                tmpSource = ASCIIEncoding.ASCII.GetBytes(sSourceData);
                tmpNewHash = new MD5CryptoServiceProvider().ComputeHash(tmpSource);
                //Console.WriteLine("New MDP : " + ByteArrayToString(tmpNewHash));

                string MDP = tmpHash;
                Console.WriteLine(MDP);
                string MDP2 = ByteArrayToString(tmpNewHash);
                Console.WriteLine(MDP2);

                if (MDP == MDP2)
                {
                    Console.WriteLine("Bon MDP");
                } else
                {
                    Console.WriteLine("Mauvais MDP");
                }
            }
        }

        static string ByteArrayToString(byte[] arrInput)
        {
            int i;
            StringBuilder sOutput = new StringBuilder(arrInput.Length);
            for (i = 0; i < arrInput.Length - 1; i++)
            {
                sOutput.Append(arrInput[i].ToString("X2"));
            }
            return sOutput.ToString();
        }

        private void EnregistrerOFF()
        {


        }
        private void button7_Click(object sender, EventArgs e)
        {
            SelectMDP();
        }

        private void selectProgProd()
        {
            textBox9.Text = "";
            if (textBox5.Text != "")
            {
                int ProgPLC = (int)Convert.ToInt32(textBox5.Text);
                Console.WriteLine(ProgPLC);
                if (ProgPLC == 1 || ProgPLC == 2)
                {
                    textBox9.Text = textBox5.Text;
                }
                else
                {
                    int choixProgProd = ProgPLC % 4;
                    if (choixProgProd == 3)
                    {
                        textBox9.Text = (ProgPLC + 2).ToString();
                    }
                    else if (choixProgProd == 0)
                    {
                        textBox9.Text = (ProgPLC + 1).ToString();
                    }
                }
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            selectProgProd();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if(textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "")
            {
                textBox5.Text = ChoixProgPLC();
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            //EnregistrerOFF();
            timer1.Interval = 3000;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //Console.WriteLine("Test");
        }
    } 
}

