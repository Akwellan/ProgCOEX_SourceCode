using System;
using System.Windows.Forms;
using Net.SourceForge.Koogra.Excel;
using Siemens.Opc.Da;
using System.Collections.Specialized;
using Dashboard;
using System.Data.OleDb;

namespace Siemens.Opc.DaClient.Controller
{
    public partial class Recette : UserControl
    {

        #region Variable

        enum ClientHandles
        {
            Item1 = 0,
            Item2,
            Item3 = 0,
            Item4,
            Item5 = 0,
            Item6,
            ItemBlockRead,
            ItemBlockWrite
        };

        private BDD.OFF_enCour OFFEnCourSVG = BDD.OFF_enCour.Default;
        
        const string LinkBDD = @"BDD\\COEX.xls";
        const string LinkBDD_Enregistrement = @"BDD\\BDD_Enregistrement_COEX.xls";//C:\\Users\\maintenance1\\Desktop\\ProgCOEX\\C#\\ProgCOEX\\src\SimpleClientDA\\

        Worksheet BDDCoexRecette = new Workbook(LinkBDD).Sheets.GetByName("BDD");
        Worksheet BDDCoexGFH = new Workbook(LinkBDD).Sheets.GetByName("GFH");
        Worksheet BDDCoexDENSITE = new Workbook(LinkBDD).Sheets.GetByName("DENSITE");
        Worksheet BDDCoexPLC = new Workbook(LinkBDD).Sheets.GetByName("BDD_Vitesse");
        Worksheet BDDCoexZumbach = new Workbook(LinkBDD).Sheets.GetByName("BDD_Zumbach");
        Worksheet BDDTemperature = new Workbook(LinkBDD).Sheets.GetByName("BDD_Temperature");
        Worksheet BDDLongueur = new Workbook(LinkBDD).Sheets.GetByName("LongueurRajout");
        Worksheet BDDVitesse = new Workbook(LinkBDD).Sheets.GetByName("BDD_VitessePROD");
        Worksheet BDDPID = new Workbook(LinkBDD).Sheets.GetByName("BDD_PID");

        private Server servZumbach = null;
        private Server servPLC = null;
        private Server servEurotherm = null;

        // ----------------------- CONNEXION ---------------------------------

        const string serverUrlZUMBACH = "opcda://localhost/OPC.ZUMBACH";
        const string serverUrlSimaticNET = "opcda://localhost/OPC.SimaticNET";
        const string serverUrlModbusServer = "opcda://localhost/Eurotherm.ModbusServer.1";

        // ----------------------- CONNEXION ---------------------------------


        // ----------------------- PLC ---------------------------------------

        // ----------------------- Extrudeuse 1-------------------------------

        // ----------------------- Régulation Température Zone 1---------------

        const string e1rtz1_instant = "S7:[Liaison_S7_1],DB106,W184"; // 1850 - 185
        const string e1rtz1_consigne = "S7:[Liaison_S7_1],DB106,W22";  // 1850 - 185

        // ----------------------- Régulation Température Zone 2---------------

        const string e1rtz2_instant = "S7:[Liaison_S7_1],DB106,W450"; // 1850 - 185
        const string e1rtz2_consigne = "S7:[Liaison_S7_1],DB106,W288";  // 1850 - 185

        // ----------------------- Régulation Température Zone 3---------------

        const string e1rtz3_instant = "S7:[Liaison_S7_1],DB106,W716"; // 1850 - 185
        const string e1rtz3_consigne = "S7:[Liaison_S7_1],DB106,W554";  // 1850 - 185

        // ----------------------- Régulation Température Zone 4---------------

        const string e1rtz4_instant = "S7:[Liaison_S7_1],DB106,W982"; // 1850 - 185
        const string e1rtz4_consigne = "S7:[Liaison_S7_1],DB106,W820";  // 1850 - 185

        // ----------------------- Régulation Température Zone 5---------------

        const string e1rtz5_instant = "S7:[Liaison_S7_1],DB106,W1245"; // 1850 - 185
        const string e1rtz5_consigne = "S7:[Liaison_S7_1],DB106,W1086";  // 1850 - 185


        // ----------------------- Complément Extrudeuse 1 ---------------------

        // ----------------------- Consigne vitesse de rotation-----------------

        string e1_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W48";  // 158 - 15.8

        // ---------------------------------vitesse de rotation-----------------

        const string e1_instant_vitesse_rotation = "S7:[Liaison_S7_1],DB30,W2";  // 158 - 15.8

        // -------------------------------Charge -------------------------------

        const string e1_charge = "S7:[Liaison_S7_1],DB30,W0";  // 158 - 15.8

        // ----------------------- Temperature matiere--------------------------

        const string e1_temperature_matiere = "S7:[Liaison_S7_1],DB106,W4174";  // 1890 - 189

        // ----------------------- Pression matiere ----------------------------

        const string e1_pression_matiere = "S7:[Liaison_S7_1],DB30,W6";  // 296 - 296

        // ----------------------- Pression matiere Mini -----------------------

        const string e1_pression_matiere_mini = "S7:[Liaison_S7_1],DB30,W166";  // 500 - 500

        // ----------------------- Pression matiere Maxi -----------------------

        const string e1_pression_matiere_maxi = "S7:[Liaison_S7_1],DB30,W164";  // 50 - 50





        // ----------------------- Tete d'extrusion ----------------------------

        // ----------------------- Régulation Température Zone 6---------------

        const string e1rtz6_instant = "S7:[Liaison_S7_1],DB106,W1780"; // 1850 - 185
        const string e1rtz6_consigne = "S7:[Liaison_S7_1],DB106,W1352";  // 1850 - 185

        // ----------------------- Régulation Température Zone 7---------------

        const string e1rtz7_instant = "S7:[Liaison_S7_1],DB106,W2046"; // 1850 - 185
        const string e1rtz7_consigne = "S7:[Liaison_S7_1],DB106,W1618";  // 1850 - 185

        // ----------------------- Régulation Température Zone 8---------------

        const string e1rtz8_instant = "S7:[Liaison_S7_1],DB106,W2046"; // 1850 - 185
        const string e1rtz8_consigne = "S7:[Liaison_S7_1],DB106,W1884";  // 1850 - 185

        // ----------------------- Régulation Température Zone 9---------------

        const string e1rtz9_instant = "S7:[Liaison_S7_1],DB106,W2312"; // 1850 - 185
        const string e1rtz9_consigne = "S7:[Liaison_S7_1],DB106,W2150";  // 1850 - 185

        // ----------------------- Régulation Température Zone 10---------------

        const string e1rtz10_instant = "S7:[Liaison_S7_1],DB106,W2578"; // 1850 - 185
        const string e1rtz10_consigne = "S7:[Liaison_S7_1],DB106,W2416";  // 1850 - 185

        // ----------------------- Régulation Température Zone 11---------------

        const string e1rtz11_instant = "S7:[Liaison_S7_1],DB106,W2844"; // 1850 - 185
        const string e1rtz11_consigne = "S7:[Liaison_S7_1],DB106,2682";  // 1850 - 185

        // ----------------------- Régulation Température Zone 12---------------

        const string e1rtz12_instant = "S7:[Liaison_S7_1],DB106,W3110"; // 1850 - 185
        const string e1rtz12_consigne = "S7:[Liaison_S7_1],DB106,W2948";  // 1850 - 185




        // ----------------------- Extrudeuse 2-------------------------------

        // ----------------------- Régulation Température Zone 1---------------

        const string e2rtz1_instant = "S7:[Liaison_S7_1],DB107,W184"; // 1850 - 185
        const string e2rtz1_consigne = "S7:[Liaison_S7_1],DB107,W22";  // 1850 - 185

        // ----------------------- Régulation Température Zone 2---------------

        const string e2rtz2_instant = "S7:[Liaison_S7_1],DB107,W450"; // 1850 - 185
        const string e2rtz2_consigne = "S7:[Liaison_S7_1],DB107,W288";  // 1850 - 185

        // ----------------------- Régulation Température Zone 3---------------

        const string e2rtz3_instant = "S7:[Liaison_S7_1],DB107,W716"; // 1850 - 185
        const string e2rtz3_consigne = "S7:[Liaison_S7_1],DB107,W554";  // 1850 - 185

        // ----------------------- Régulation Température Zone 4---------------

        const string e2rtz4_instant = "S7:[Liaison_S7_1],DB107,W982"; // 1850 - 185
        const string e2rtz4_consigne = "S7:[Liaison_S7_1],DB107,W820";  // 1850 - 185


        // ----------------------- Complément Extrudeuse 2 ---------------------

        // ----------------------- Consigne vitesse de rotation-----------------

        string e2_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W50";  // 158 - 15.8

        // ---------------------------------vitesse de rotation-----------------

        const string e2_instant_vitesse_rotation = "S7:[Liaison_S7_1],DB30,W18";  // 158 - 15.8

        // -------------------------------Charge -------------------------------

        const string e2_charge = "S7:[Liaison_S7_1],DB30,W16";  // 158 - 15.8

        // ----------------------- Temperature matiere--------------------------

        const string e2_temperature_matiere = "S7:[Liaison_S7_1],DB30,W20";  // 1890 - 189

        // ----------------------- Pression matiere ----------------------------

        const string e2_pression_matiere = "S7:[Liaison_S7_1],DB30,W22";  // 296 - 296

        // ----------------------- Pression matiere Mini -----------------------

        const string e2_pression_matiere_mini = "S7:[Liaison_S7_1],DB30,W170";  // 500 - 500

        // ----------------------- Pression matiere Maxi -----------------------

        const string e2_pression_matiere_maxi = "S7:[Liaison_S7_1],DB30,W168";  // 50 - 50






        // ----------------------- Extrudeuse 3-------------------------------

        // ----------------------- Régulation Température Zone 1---------------

        const string e3rtz1_instant = "S7:[Liaison_S7_1],DB108,W184"; // 1850 - 185
        const string e3rtz1_consigne = "S7:[Liaison_S7_1],DB108,W22";  // 1850 - 185

        // ----------------------- Régulation Température Zone 2---------------

        const string e3rtz2_instant = "S7:[Liaison_S7_1],DB108,W450"; // 1850 - 185
        const string e3rtz2_consigne = "S7:[Liaison_S7_1],DB108,W288";  // 1850 - 185

        // ----------------------- Régulation Température Zone 3---------------

        const string e3rtz3_instant = "S7:[Liaison_S7_1],DB108,W716"; // 1850 - 185
        const string e3rtz3_consigne = "S7:[Liaison_S7_1],DB108,W554";  // 1850 - 185

        // ----------------------- Régulation Température Zone 4---------------

        const string e3rtz4_instant = "S7:[Liaison_S7_1],DB108,W982"; // 1850 - 185
        const string e3rtz4_consigne = "S7:[Liaison_S7_1],DB108,W820";  // 1850 - 185


        // ----------------------- Complément Extrudeuse 3 ---------------------

        // ----------------------- Consigne vitesse de rotation-----------------

        string e3_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W52";  // 158 - 15.8

        // ---------------------------------vitesse de rotation-----------------

        const string e3_instant_vitesse_rotation = "S7:[Liaison_S7_1],DB30,W34";  // 158 - 15.8

        // -------------------------------Charge -------------------------------

        const string e3_charge = "S7:[Liaison_S7_1],DB30,W32";  // 158 - 15.8

        // ----------------------- Temperature matiere--------------------------

        const string e3_temperature_matiere = "S7:[Liaison_S7_1],DB30,W36";  // 1890 - 189

        // ----------------------- Pression matiere ----------------------------

        const string e3_pression_matiere = "S7:[Liaison_S7_1],DB30,W38";  // 296 - 296

        // ----------------------- Pression matiere Mini -----------------------

        const string e3_pression_matiere_mini = "S7:[Liaison_S7_1],DB30,W174";  // 500 - 500

        // ----------------------- Pression matiere Maxi -----------------------

        const string e3_pression_matiere_maxi = "S7:[Liaison_S7_1],DB30,W172";  // 50 - 50






        // ----------------------- Extrudeuse 4-------------------------------

        // ----------------------- Régulation Température Zone 1---------------

        const string e4rtz1_instant = "S7:[Liaison_S7_1],DB109,W184"; // 1850 - 185
        const string e4rtz1_consigne = "S7:[Liaison_S7_1],DB109,W22";  // 1850 - 185

        // ----------------------- Régulation Température Zone 2---------------

        const string e4rtz2_instant = "S7:[Liaison_S7_1],DB109,W450"; // 1850 - 185
        const string e4rtz2_consigne = "S7:[Liaison_S7_1],DB109,W288";  // 1850 - 185

        // ----------------------- Régulation Température Zone 3---------------

        const string e4rtz3_instant = "S7:[Liaison_S7_1],DB109,W716"; // 1850 - 185
        const string e4rtz3_consigne = "S7:[Liaison_S7_1],DB109,W554";  // 1850 - 185

        // ----------------------- Régulation Température Zone 4---------------

        const string e4rtz4_instant = "S7:[Liaison_S7_1],DB109,W982"; // 1850 - 185
        const string e4rtz4_consigne = "S7:[Liaison_S7_1],DB109,W820";  // 1850 - 185

        // ----------------------- Régulation Température Zone 6---------------

        const string e4rtz6_instant = "S7:[Liaison_S7_1],DB109,W1780"; // 1850 - 185 //A MDOFIER ---------------------------------------
        const string e4rtz6_consigne = "S7:[Liaison_S7_1],DB109,W1352";  // 1850 - 185 //A MDOFIER ---------------------------------------

        // ----------------------- Régulation Température Zone 7---------------

        const string e4rtz7_instant = "S7:[Liaison_S7_1],DB109,W2046"; // 1850 - 185 //A MDOFIER ---------------------------------------
        const string e4rtz7_consigne = "S7:[Liaison_S7_1],DB109,W1618";  // 1850 - 185 //A MDOFIER ---------------------------------------


        // ----------------------- Complément Extrudeuse 4 ---------------------

        // ----------------------- Consigne vitesse de rotation-----------------

        string e4_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W54";  // 158 - 15.8

        // ---------------------------------vitesse de rotation-----------------

        const string e4_instant_vitesse_rotation = "S7:[Liaison_S7_1],DB30,W50";  // 158 - 15.8

        // -------------------------------Charge -------------------------------

        const string e4_charge = "S7:[Liaison_S7_1],DB30,W48";  // 158 - 15.8

        // ----------------------- Temperature matiere--------------------------

        const string e4_temperature_matiere = "S7:[Liaison_S7_1],DB30,W52";  // 1890 - 189

        // ----------------------- Pression matiere ----------------------------

        const string e4_pression_matiere = "S7:[Liaison_S7_1],DB30,W54";  // 296 - 296

        // ----------------------- Pression matiere Mini -----------------------

        const string e4_pression_matiere_mini = "S7:[Liaison_S7_1],DB30,W178";  // 500 - 500

        // ----------------------- Pression matiere Maxi -----------------------

        const string e4_pression_matiere_maxi = "S7:[Liaison_S7_1],DB30,W176";  // 50 - 50





        // ----------------------- Banc de tirage-------------------------------

        const string banc_tirage_instant = "S7:[Liaison_S7_1],DB30,W82"; // 1850 - 18.50
        string banc_tirage_consigne = "S7:[Liaison_S7_1],DB31,W56";  // 1850 - 18.50


        // ----------------------- Banc de coupe--------------------------------

        const string longueur_coupe_instant = "S7:[Liaison_S7_1],DB70,W2"; // 6700 - 67
        const string longueur_coupe_consigne = "S7:[Liaison_S7_1],DB30,W86";  // 6700 - 67
        const string cadence_coupe = "S7:[Liaison_S7_1],DB30,W84";  // 1850 - 185.0
        const string compteur_piece = "S7:[Liaison_S7_1],DB30,W94";  // 0 - 0

        const string activeProgPLC = "S7:[Liaison_S7_1],DB31,W2";  // 0 - 0
        const string nextProg = "S7:[Liaison_S7_1],DB31,W6";  // 0 - 0
        const string timer = "S7:[Liaison_S7_1],DB31,W0";  // 0 - 0
        const string btnK1 = "S7:[Liaison_S7_1],DB21,X114.6"; //false - false

        // ----------------------- Zumbach---------------------------------------

        // ----------------------- Regulation diametre---------------------------

        const string numero_prog_zumbach = "Usys_200_Coex/Product/c1100";  // 
        const string activer_progZumbach = "Usys_200_Coex/Product/c1040";  // W
        const string nom_prog_zumbach = "Usys_200_Coex/Product/c1110";
        const string diametre_mesure = "Usys_200_Coex/Data/d1100";
        const string diametre_cible = "Usys_200_Coex/Product/c2100.1";
        const string diametre_max = "Usys_200_Coex/Product/c2120.1";
        const string diametre_min = "Usys_200_Coex/Product/c2140.1";
        const string numero_of = "Usys_200_Coex/Data/d8000";
        const string code_operateur = "Usys_200_Coex/Data/d8010";
        const string deviation = "Usys_200_Coex/Data/d2010.1";





        // ----------------------- Eurotherm--------------------------------------

        // ----------------------- Bac sous vide----------------------------------

        const string videbac1_instant = "COM4.ID001-2408.Operator.MAIN.PV"; // -219 - -219
        const string videbac1_consigne = "COM4.ID001-2408.Operator.MAIN.tSP"; // -219 - -219
        const string videbac2_instant = "COM3.ID001-2408.Operator.MAIN.PV"; // -219 - -219

        #endregion

        #region Connection et Deconnexion Server
        /// <summary>
        /// Handles connect procedure
        /// </summary>
        private void OnConnect()
        {
            if (servZumbach == null)
            {
                // Create a server object
                servZumbach = new Server();//Serveur ZUMBACH
            }
            try
            {
                // connect to the server
                servZumbach.Connect(serverUrlZUMBACH);


                // Change GUI settings
                //Console.WriteLine("Connecté au Zumbach");

            }
            catch (Exception exception)
            {
                // Cleanup
                servZumbach = null;

                System.Windows.MessageBox.Show(exception.Message, "Connect failed");
            }
        }
        private void OnConnect2()
        {
            if (servPLC == null)
            {
                // Create a server object
                servPLC = new Server();
            }
            try
            {
                // connect to the server
                servPLC.Connect(serverUrlSimaticNET);

                // Change GUI settings
                //Console.WriteLine("Connecté au PLC");
            }
            catch (Exception exception)
            {
                // Cleanup
                servPLC = null;

                MessageBox.Show(exception.Message, "Connect failed");
            }
        }
        private void OnConnect3()
        {
            if (servEurotherm == null)
            {
                // Create a server object
                servEurotherm = new Server();
            }
            try
            {
                // connect to the server
                servEurotherm.Connect(serverUrlModbusServer);

                // Change GUI settings
                //Console.WriteLine("Connecté au Eurotherm");
            }
            catch (Exception exception)
            {
                // Cleanup
                servEurotherm = null;

                MessageBox.Show(exception.Message, "Connect failed");
            }
        }
        private void OnDisconnect()
        {
            if (servZumbach == null)
            {
                return;
            }
            if (servPLC == null)
            {
                return;
            }
            if (servEurotherm == null)
            {
                return;
            }

            try
            {
                // Change GUI settings
                //Console.WriteLine("Déconnecté de l'Eurotherm");
                //Console.WriteLine("Déconnecté du PLC");
                //Console.WriteLine("Déconnecté du Zumbach");

                // Disconnect
                servZumbach.Disconnect();
                servPLC.Disconnect();
                servEurotherm.Disconnect();

                // Cleanup
                servZumbach = null;
                servPLC = null;
                servEurotherm = null;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Disconnect failed");
            }
        }
        private void ConnectOPC_Zumbach()
        {
            if (servZumbach == null)
            {
                OnConnect();
            }
            else if (servZumbach.IsConnected)
            {
                OnDisconnect();
            }
            else
            {
                OnDisconnect();
            }
        }
        private void ConnectOPC_PLC()
        {
            if (servPLC == null)
            {
                OnConnect2();
            }
            else if (servPLC.IsConnected)
            {
                OnDisconnect();
            }
            else
            {
                OnDisconnect();
            }
        }
        private void ConnectOPC_Eurotherm()
        {
            if (servEurotherm == null)
            {
                OnConnect3();
            }
            else if (servEurotherm.IsConnected)
            {
                OnDisconnect();
            }
            else
            {
                OnDisconnect();
            }
        }

        #endregion

        #region Construction
        public Recette()
        {
            InitializeComponent();
            reini();

            // ----------------------- CONNEXION ---------------------------------
            ConnectOPC_Eurotherm(); //Connect OPC EUROTHERM
            ConnectOPC_PLC(); //Connect OPC PLC
            ConnectOPC_Zumbach(); //Connect OPC Zumbach
            // ----------------------- CONNEXION ---------------------------------
            string numPorg = ReadPLC(activeProgPLC);
            ReadDBVitesse(numPorg);

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

        #endregion

        #region Fonction Read / Write

        private string ReadEurotherm(string NomItem)
        {
            //try
            //{
            bool bResult;
            StringCollection itemIds = new StringCollection();
            object[] values;
            int[] pErrors;

            // ItemIds to read
            itemIds.Add(NomItem);

            bResult = servEurotherm.Read(itemIds, out values, out pErrors);

            // everything went fine
            if (bResult)
            {
                return values[0].ToString();
            }
            // we have to check the individual result codes
            else
            {
                if (pErrors[0] == 0)
                {
                    return values[0].ToString();
                }
                else
                {
                    string errorString;
                    servEurotherm.GetErrorString(pErrors[0], 0, out errorString);
                    return errorString;
                }
            }
            //}
            //catch (Exception exception)
            //{
            //    return exception.Message;
            //}
        }
        private string ReadPLC(string NomItem)
        {
            try
            {
            bool bResult;
            StringCollection itemIds = new StringCollection();
            object[] values;
            int[] pErrors;

            // ItemIds to read
            itemIds.Add(NomItem);

            bResult = servPLC.Read(itemIds, out values, out pErrors);

            // everything went fine
            if (bResult)
            {
                return values[0].ToString();
            }
            // we have to check the individual result codes
            else
            {
                if (pErrors[0] == 0)
                {
                    return values[0].ToString();
                }
                else
                {
                    string errorString;
                    servPLC.GetErrorString(pErrors[0], 0, out errorString);
                    return errorString;
                }
            }
            }
            catch (Exception exception)
            {
                return exception.Message;
            }
        }
        private string ReadZumbach(string NomItem)
        {
            //try
            //{
            bool bResult;
            StringCollection itemIds = new StringCollection();
            object[] values;
            int[] pErrors;

            // ItemIds to read
            itemIds.Add(NomItem);

            bResult = servZumbach.Read(itemIds, out values, out pErrors);

            // everything went fine
            if (bResult)
            {
                return values[0].ToString();
            }
            // we have to check the individual result codes
            else
            {
                if (pErrors[0] == 0)
                {
                    return values[0].ToString();
                }
                else
                {
                    string errorString;
                    servZumbach.GetErrorString(pErrors[0], 0, out errorString);
                    return errorString;
                }
            }
            //}
            //catch (Exception exception)
            //{
            //    return exception.Message;
            //}
        }

        private void WriteEurotherm(string NomItemID, string Valeur)
        {
            try
            {
                this.servEurotherm.Write(NomItemID, Valeur);
            }
            catch (Exception exception)
            {
                //Console.WriteLine("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

        }
        private void WritePLC(string NomItemID, string Valeur) 
        {
            try
            {
                this.servPLC.Write(NomItemID, Valeur);
            }
            catch (Exception exception)
            {

                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

        }
        private void WriteZumbach(string NomItemID, string Valeur)
        {
            try
            {
                this.servZumbach.Write(NomItemID, Valeur);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);

            }

        }
        #endregion

        #region Fonction Evenement PROG

        private void reini()
        {
            txtNumProcess.Text = "";

            txtDiametre.Text = "";
            cbGFXDestination.Text = "";
            txtLongueur.Text = "";
            txtCadenceProd.Text = "";
            cbDensite.Text = "";

            txtDebitEau1.Text = "";
            txtDebitEau2.Text = "";
            txtDebitEau3.Text = "";

            txtColoriVis1.Text = "";
            txtColoriVis4.Text = "";

            lblProgPLC.Text = "";
            lblProgProd.Text = "";
            lblProgZumbach.Text = "";

            txtMatiere.Text = "";  
            cbOutillage.Text = "";
            txtEcartement.Text = "";
            txtDebitEau1.Text = "";

            cbTete.Text = "";
            cbType.Text = "";
        }

        private void confirmTrue()
        {
            txtNumProcess.Enabled = false;

            txtDiametre.Enabled = false;
            cbGFXDestination.Enabled = false;
            txtLongueur.Enabled = false;
            txtCadenceProd.Enabled = false;
            cbDensite.Enabled = false;

            lblProgPLC.Enabled = false;
            lblProgProd.Enabled = false;
            lblProgZumbach.Enabled = false;

            cbTete.Enabled = false;
            cbType.Enabled = false;
            txtMatiere.Enabled = false;
        }

        private void confirmFalse()
        {
            txtNumProcess.Enabled = true;

            txtDiametre.Enabled = true;
            cbGFXDestination.Enabled = true;
            txtLongueur.Enabled = true;
            cbDensite.Enabled = true;
            txtCadenceProd.Enabled = true;

            lblProgPLC.Enabled = true;
            lblProgProd.Enabled = true;
            lblProgZumbach.Enabled = true;

            cbTete.Enabled = true;
            cbType.Enabled = true;
            txtMatiere.Enabled = true;
        }

        private bool ReadExcelBDD()
        {
            bool trouver = false;
            for (ushort i = 1; i <= BDDCoexRecette.Rows.LastRow; i++)
            {
                if (BDDCoexRecette.Rows[i].Cells[2].Value.ToString() == txtNumProcess.Text)
                {
                    txtDiametre.Text = BDDCoexRecette.Rows[i].Cells[4].Value.ToString();
                    txtLongueur.Text = BDDCoexRecette.Rows[i].Cells[5].Value.ToString();
                    cbDensite.Text = DENSITE(BDDCoexRecette.Rows[i].Cells[9].Value.ToString());
                    cbGFXDestination.Text = GFH(BDDCoexRecette.Rows[i].Cells[11].Value.ToString(),1);
                    txtMatiere.Text = BDDCoexRecette.Rows[i].Cells[7].Value.ToString();
                    //typeTete = txtTest.Text;//BDDCoexRecette.Rows[i].Cells[5].Value.ToString(); //A REMPLIRE
                    string type = BDDCoexRecette.Rows[i].Cells[3].Value.ToString();
                    if(type == "S")
                    {
                        cbType.Text = "Symétrique";
                    }else if (type == "A")
                    {
                        cbType.Text = "Asymètrique";
                    }
                    else if (type == "UV")
                    {
                        cbType.Text = "Asymètrique";
                    }
                    
                    ////Console.WriteLine("Fin");
                    //Console.WriteLine("Numero processus : " + BDDCoexRecette.Rows[i].Cells[2].Value.ToString());
                    //Console.WriteLine("Diametre tube : " + BDDCoexRecette.Rows[i].Cells[4].Value.ToString());
                    //Console.WriteLine("Longueur tube : " + BDDCoexRecette.Rows[i].Cells[5].Value.ToString());
                    //Console.WriteLine("Densité tube : " + BDDCoexRecette.Rows[i].Cells[9].Value.ToString());
                    //Console.WriteLine("i : " + i + 1);
                    trouver = true;
                    break;

                    ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }

            }

            if (trouver == true)
            {
                //Console.WriteLine("\nNumero de processus trouvé");
                lblCaract.Text = "Caractéristiques du tube :";
                return true;
            }
            else
            {
                //Console.WriteLine("\nNumero de processus non trouvé");

                txtDiametre.Text = "";
                cbGFXDestination.Text = "";
                txtLongueur.Text = "";
                cbDensite.Text = "";
                return false;
            }

        }

        private string DENSITE(string norme)
        {
            for (ushort i = 1; i <= BDDCoexDENSITE.Rows.LastRow; i++)
            {
                if (BDDCoexDENSITE.Rows[i].Cells[0].Value.ToString() == norme)
                {
                    //Console.WriteLine("Numero processus : " + BDDCoexDENSITE.Rows[i].Cells[0].Value.ToString());
                    //Console.WriteLine("Diametre tube : " + BDDCoexDENSITE.Rows[i].Cells[1].Value.ToString());
                    //Console.WriteLine("Longueur tube : " + BDDCoexDENSITE.Rows[i].Cells[2].Value.ToString());
                    //Console.WriteLine("i : " + i + 1);
                    return BDDCoexDENSITE.Rows[i].Cells[2].Value.ToString();
                    break;

                    ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
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
                        //Console.WriteLine("Norme : " + BDDCoexGFH.Rows[i].Cells[0].Value.ToString());
                        //Console.WriteLine("Nom : " + BDDCoexGFH.Rows[i].Cells[1].Value.ToString());
                        //Console.WriteLine("code : " + BDDCoexGFH.Rows[i].Cells[2].Value.ToString());
                        //Console.WriteLine("i : " + i + 1);
                        return BDDCoexGFH.Rows[i].Cells[2].Value.ToString();

                        break;

                        ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                }
            }
            else if (choix == 2)
            {
                for (ushort i = 1; i <= BDDCoexGFH.Rows.LastRow; i++)
                {
                    if (BDDCoexGFH.Rows[i].Cells[2].Value.ToString() == norme)
                    {
                        //Console.WriteLine("Norme : " + BDDCoexGFH.Rows[i].Cells[0].Value.ToString());
                        //Console.WriteLine("Nom : " + BDDCoexGFH.Rows[i].Cells[1].Value.ToString());
                        //Console.WriteLine("code : " + BDDCoexGFH.Rows[i].Cells[2].Value.ToString());
                        //Console.WriteLine("i : " + i + 1);
                        return BDDCoexGFH.Rows[i].Cells[1].Value.ToString();

                        break;

                        ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                }
            }
            return "";
        }

        private string ChoixProgZumbach()
        {

            string nomProduit = GFH(cbGFXDestination.Text, 2);//textBox4.Text;
            nomProduit = nomProduit + " D" + txtDiametre.Text;
            //Console.WriteLine(nomProduit);
            string densite = cbDensite.Text;
            //Console.WriteLine(densite);
            for (ushort i = 1; i <= BDDCoexZumbach.Rows.LastRow; i++)
            {
                if (BDDCoexZumbach.Rows[i].Cells[2].Value.ToString() == nomProduit && BDDCoexZumbach.Rows[i].Cells[3].Value.ToString() == densite)
                {
                    return BDDCoexZumbach.Rows[i].Cells[1].Value.ToString();
                    break;

                    ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }
            }
            return "Aucun PROG Trouvé";
        }

        private string ChoixProgPLC()
        {

            string diametre = txtDiametre.Text;//textBox4.Text;
            string densite = cbDensite.Text;
            for (ushort i = 1; i <= BDDCoexPLC.Rows.LastRow; i++)
            {
                if (BDDCoexPLC.Rows[i].Cells[1].Value.ToString() == diametre && BDDCoexPLC.Rows[i].Cells[2].Value.ToString() == densite)
                {
                    return BDDCoexPLC.Rows[i].Cells[0].Value.ToString();
                    break;

                    ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }
            }
            return "Aucun PROG Trouvé";
        }

        private void selectProgProd()
        {
            lblProgProd.Text = "";
            if (lblProgPLC.Text != "")
            {
                int ProgPLC = (int)Convert.ToInt32(lblProgPLC.Text);
                //Console.WriteLine(ProgPLC);
                if (ProgPLC == 1 || ProgPLC == 2)
                {
                    lblProgProd.Text = lblProgPLC.Text;
                }
                else
                {
                    int choixProgProd = ProgPLC % 4;
                    if (choixProgProd == 3)
                    {
                        lblProgProd.Text = (ProgPLC + 2).ToString();
                    }
                    else if (choixProgProd == 0)
                    {
                        lblProgProd.Text = (ProgPLC + 1).ToString();
                    }
                }
            }
            if(cbDensite.Text == "Asymètrique")
            {
                lblProgProd.Text = lblProgPLC.Text;
            }
        }

        private void searchProg()
        {
            if(txtDiametre.Text != "" && cbGFXDestination.Text != "" && txtLongueur.Text != "" && cbDensite.Text != "")
            {
                lblProgZumbach.Text = ChoixProgZumbach();
                lblProgPLC.Text = ChoixProgPLC();
                
            } else
            {
                lblProgZumbach.Text = "";
                lblProgPLC.Text = "";
                lblProgProd.Text = "";
            }

            if(lblProgPLC.Text != "")
            {
                selectProgProd();
            }
        }

        private void Longueur()
        {
            string diametre = txtDiametre.Text;
            string densite = cbDensite.Text;
            string gfh = cbGFXDestination.Text;
            string typeTete = cbTete.Text;
            Double LongueurSansTete = (int)Convert.ToInt32(txtLongueur.Text);
            string LongueurTete = "";
            ////Console.WriteLine("diametre: '" + diametre + "' densite: '" + densite + "' gfh: '" + gfh + "' typeTete: '" + typeTete + "'");
            ////Console.WriteLine("\n--------------------------------------------------\n");

            for (ushort i = 0; i <= BDDLongueur.Rows.LastRow - 18; i++)
            {
                ////Console.WriteLine(BDDLongueur.Rows.LastRow);=
                if (BDDLongueur.Rows[i].Cells[6].Value.ToString() == typeTete)
                {
                    typeTete = BDDLongueur.Rows[i].Cells[7].Value.ToString();
                    ////Console.WriteLine(typeTete);
                }
            }

            ////Console.WriteLine("diametre: '" + diametre + "' densite: '" + densite + "' typeTete: '" + typeTete + "'");

            for (ushort i = 0; i <= BDDLongueur.Rows.LastRow; i++)
            {//BDDLongueur.Rows[i].Cells[1].Value.ToString() == gfh && 
                ////Console.WriteLine("diametre: '" + diametre + "' densite: '" + densite + "' typeTete: '" + typeTete + "'");
                if (BDDLongueur.Rows[i].Cells[3].Value.ToString() == densite && BDDLongueur.Rows[i].Cells[5].Value.ToString().Contains(diametre) && BDDLongueur.Rows[i].Cells[2].Value.ToString() == typeTete)
                {
                    LongueurTete = BDDLongueur.Rows[i].Cells[4].Value.ToString();
                    ////Console.WriteLine("OK");
                    break;
                }
            }

            //Console.WriteLine(LongueurTete);
            //LongueurTete = LongueurTete.Replace(",", ".");
            ////Console.WriteLine(LongueurTete);

            float test = (float)Convert.ToDouble(LongueurTete);

            txtLongueur.Text = (LongueurSansTete + test).ToString();
        }

        private void gestionTemperature()
        {
            bool green = false;
            if(txtMatiere.Text.Contains("GREEN") == true)
            {
                green = true;
            }

            int E1Z1 = recuperationTemperature("E1Z1", green);
            int E1Z2 = recuperationTemperature("E1Z2", green);
            int E1Z3 = recuperationTemperature("E1Z3", green);
            int E1Z4 = recuperationTemperature("E1Z4", green);
            int E1Z5 = recuperationTemperature("E1Z5", green);

            int E2Z1 = recuperationTemperature("E2Z1", green);
            int E2Z2 = recuperationTemperature("E2Z2", green);
            int E2Z3 = recuperationTemperature("E2Z3", green);
            int E2Z4 = recuperationTemperature("E2Z4", green);

            int E3Z1 = recuperationTemperature("E3Z1", green);
            int E3Z2 = recuperationTemperature("E3Z2", green);
            int E3Z3 = recuperationTemperature("E3Z3", green);
            int E3Z4 = recuperationTemperature("E3Z4", green);

            int E4Z1 = recuperationTemperature("E4Z1", green);
            int E4Z2 = recuperationTemperature("E4Z2", green);
            int E4Z3 = recuperationTemperature("E4Z3", green);
            int E4Z4 = recuperationTemperature("E4Z4", green);
            int E4Z6 = recuperationTemperature("E4Z6", green);
            int E4Z7 = recuperationTemperature("E4Z7", green);

            int TZ6 = recuperationTemperature("TZ6", green);
            int TZ7 = recuperationTemperature("TZ7", green);
            int TZ8 = recuperationTemperature("TZ8", green);
            int TZ9 = recuperationTemperature("TZ9", green);
            int TZ10 = recuperationTemperature("TZ10", green);
            int TZ11 = recuperationTemperature("TZ11", green);
            int TZ12 = recuperationTemperature("TZ12", green);

            Console.WriteLine(E1Z1 + " " + E1Z2 + " " + E1Z3 + " " + E1Z4 + " " + E1Z5);
            Console.WriteLine(E2Z1 + " " + E2Z2 + " " + E2Z3 + " " + E2Z4);
            Console.WriteLine(E3Z1 + " " + E3Z2 + " " + E3Z3 + " " + E3Z4);
            Console.WriteLine(E4Z1 + " " + E4Z2 + " " + E4Z3 + " " + E4Z4 + " " + E4Z6 + " " + E4Z7);
            Console.WriteLine(TZ6 + " " + TZ7 + " " + TZ8 + " " + TZ9 + " " + TZ10 + " " + TZ11 + " " + TZ12);

            WritePLC(e1rtz1_consigne, (E1Z1 * 10).ToString());
            WritePLC(e1rtz2_consigne, (E1Z2 * 10).ToString());
            WritePLC(e1rtz3_consigne, (E1Z3 * 10).ToString());  //EXTRUSION 1
            WritePLC(e1rtz4_consigne, (E1Z4 * 10).ToString());
            WritePLC(e1rtz5_consigne, (E1Z5 * 10).ToString());

            WritePLC(e1rtz6_consigne, (TZ6 * 10).ToString());
            WritePLC(e1rtz7_consigne, (TZ7 * 10).ToString());
            WritePLC(e1rtz8_consigne, (TZ8 * 10).ToString());
            WritePLC(e1rtz9_consigne, (TZ9 * 10).ToString());  //TETE EXTRUSION
            WritePLC(e1rtz10_consigne, (TZ10 * 10).ToString());
            WritePLC(e1rtz11_consigne, (TZ11 * 10).ToString());
            WritePLC(e1rtz12_consigne, (TZ12 * 10).ToString());

            WritePLC(e2rtz1_consigne, (E2Z1 * 10).ToString());
            WritePLC(e2rtz2_consigne, (E2Z2 * 10).ToString());
            WritePLC(e2rtz3_consigne, (E2Z3 * 10).ToString());  //EXTRUSION 2
            WritePLC(e2rtz4_consigne, (E2Z4 * 10).ToString());

            WritePLC(e3rtz1_consigne, (E3Z1 * 10).ToString());
            WritePLC(e3rtz2_consigne, (E3Z2 * 10).ToString());
            WritePLC(e3rtz3_consigne, (E3Z3 * 10).ToString());  //EXTRUSION 3
            WritePLC(e3rtz4_consigne, (E3Z4 * 10).ToString()); 

            WritePLC(e3rtz1_consigne, (E3Z1 * 10).ToString());
            WritePLC(e3rtz2_consigne, (E3Z2 * 10).ToString());
            WritePLC(e3rtz3_consigne, (E3Z3 * 10).ToString());  //EXTRUSION 4
            WritePLC(e3rtz4_consigne, (E3Z4 * 10).ToString());

            WritePLC(e4rtz1_consigne, (E4Z1 * 10).ToString());
            WritePLC(e4rtz2_consigne, (E4Z2 * 10).ToString());
            WritePLC(e4rtz3_consigne, (E4Z3 * 10).ToString());
            WritePLC(e4rtz4_consigne, (E4Z4 * 10).ToString());
            WritePLC(e4rtz6_consigne, (E4Z6 * 10).ToString());
            WritePLC(e4rtz7_consigne, (E4Z7 * 10).ToString());

        }

        private int recuperationTemperature(string intituler, bool green)
        {
            for (ushort i = 1; i <= BDDTemperature.Rows.LastRow; i++)
            {
                if (green == false)
                {
                    if (BDDTemperature.Rows[i].Cells[2].Value.ToString() == intituler)
                    {
                        return Convert.ToInt32(BDDTemperature.Rows[i].Cells[3].Value.ToString());
                        ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                }
                else if (green == true)
                {
                    if (BDDTemperature.Rows[i].Cells[2].Value.ToString() == intituler)
                    {
                        return Convert.ToInt32(BDDTemperature.Rows[i].Cells[4].Value.ToString());
                        ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                    }
                }
            }
            return 0;
        }

        private void searchVitesse()
        {
            bool b = false;
            for (ushort i = 1; i <= BDDVitesse.Rows.LastRow; i++)
            {
                if (BDDVitesse.Rows[i].Cells[0].Value.ToString() == txtDiametre.Text)
                {
                    txtCadenceProd.Text = BDDVitesse.Rows[i].Cells[3].Value.ToString();
                    b = true;
                    ////Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                } else if (b == false)
                {
                    txtCadenceProd.Text = "";
                }                                      
            }
        }

        private void OFFEnCour()
        {
            if (OFFEnCourSVG.Processus != "")
            {
                txtNumProcess.Text = OFFEnCourSVG.Processus;
                txtDiametre.Text = OFFEnCourSVG.Diametre;
                cbDensite.Text = OFFEnCourSVG.Densite;
                cbGFXDestination.Text = OFFEnCourSVG.GFH;
                txtLongueur.Text = OFFEnCourSVG.Longueur;
                lblProgPLC.Text = OFFEnCourSVG.ProgPLCDem;
                lblProgProd.Text = OFFEnCourSVG.ProgPLCProd;
                lblProgZumbach.Text = OFFEnCourSVG.ProgZumbach;
                txtCadenceProd.Text = OFFEnCourSVG.CadenceProd;
                txtDebitEau1.Text = OFFEnCourSVG.DebitEau1;
                txtDebitEau2.Text = OFFEnCourSVG.DebitEau2;
                txtDebitEau3.Text = OFFEnCourSVG.DebitEau3;
                txtEcartement.Text = OFFEnCourSVG.EcartementBT;
                cbOutillage.Text = OFFEnCourSVG.Outillage;
                cbTete.Text = OFFEnCourSVG.Tete;
                cbType.Text = OFFEnCourSVG.Type;
                txtMatiere.Text = OFFEnCourSVG.Matiere;

                if(OFFEnCourSVG.EtatBTN == "Confirmer")
                {
                    btnNewRecette.Hide();
                    btnDemarrageProd.Hide();
                    btnConfirmer.Hide();
                    btnCharger.Show();
                    btnReinitialiser.Show();

                    confirmTrue();
                } 
                else if (OFFEnCourSVG.EtatBTN == "Charger")
                {
                    btnNewRecette.Hide();
                    btnDemarrageProd.Show();
                    btnConfirmer.Hide();
                    btnCharger.Hide();
                    btnReinitialiser.Hide();
                    confirmTrue();

                    lblNumProcessus.Text = "Numero Processus en cour";
                    lblCaract.Text = "Caractéristique en cour";
                }
                else if (OFFEnCourSVG.EtatBTN == "Production")
                {
                    btnCharger.Hide();
                    btnDemarrageProd.Hide();
                    btnConfirmer.Hide();
                    btnReinitialiser.Hide();
                    btnNewRecette.Show();

                    confirmTrue();
                }
            } else
            {
                btnCharger.Hide();
                btnDemarrageProd.Hide();
                btnConfirmer.Show();
                btnReinitialiser.Show();
                btnNewRecette.Hide();
                reini();
                confirmFalse();
                
            }
        }

        private void EnregConfOFFEnCour()
        {
            OFFEnCourSVG.Processus = txtNumProcess.Text;
            OFFEnCourSVG.Diametre = txtDiametre.Text;
            OFFEnCourSVG.Densite = cbDensite.Text;
            OFFEnCourSVG.GFH = cbGFXDestination.Text;
            OFFEnCourSVG.Longueur = txtLongueur.Text;
            OFFEnCourSVG.ProgPLCDem = lblProgPLC.Text;
            OFFEnCourSVG.ProgPLCProd = lblProgProd.Text;
            OFFEnCourSVG.ProgZumbach = lblProgZumbach.Text;
            OFFEnCourSVG.CadenceProd = txtCadenceProd.Text;
            OFFEnCourSVG.DebitEau1 = txtDebitEau1.Text;
            OFFEnCourSVG.DebitEau2 = txtDebitEau2.Text;
            OFFEnCourSVG.DebitEau3 = txtDebitEau3.Text;
            OFFEnCourSVG.EcartementBT = txtEcartement.Text;
            OFFEnCourSVG.Outillage = cbOutillage.Text;
            OFFEnCourSVG.Tete = cbTete.Text;
            OFFEnCourSVG.Type = cbType.Text;
            OFFEnCourSVG.Matiere = txtMatiere.Text;
            OFFEnCourSVG.Save();
        }

        private string ChoixPID()
        {
            string diametre = txtDiametre.Text; //textBox4.Text;
            string densite = cbDensite.Text;
            string BAC;

            for (ushort i = 1; i <= BDDPID.Rows.LastRow; i++)
            {
                if (BDDPID.Rows[i].Cells[1].Value.ToString() == diametre && BDDPID.Rows[i].Cells[2].Value.ToString() == densite)
                {
                    return BDDPID.Rows[i].Cells[0].Value.ToString();
                }
            }
            return "Aucun PID Trouvé";
        }

        void ReadDBVitesse()
        {
            string numPorg = ReadPLC(activeProgPLC);
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


        private void WriteExcel_Enregistrement()
        {
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + LinkBDD_Enregistrement + ";Extended Properties=\"Excel 8.0\";";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                // Le nom des colonnes se trouve en première ligne de la feuille. 
                // Ici les colonnes concernées s'intitulent COL 1 et COL 2 : 
                String NameCol1 = "NumeroProcessus";
                String NameCol2 = "GFH";
                String NameCol22 = "Diametre";
                String NameCol3 = "Densité";
                String NameCol4 = "Longueur";
                String NameCol5 = "Tete";
                String NameCol6 = "Type";
                String NameCol7 = "Matiere";
                String NameCol8 = "DebitEau1";
                String NameCol9 = "DebitEau2";
                String NameCol10 = "DebitEau3";
                String NameCol11 = "Outillage";
                String NameCol12 = "EcartementBT";
                String NameCol13 = "ColorVis1";
                String NameCol14 = "ColorVis4";
                String NameCol15 = "VitesseVis1";
                String NameCol16 = "VitesseVis2";
                String NameCol17 = "VitesseVis3";
                String NameCol18 = "VitesseVis4";
                String NameCol19 = "VitesseVisBT";
                String NameCol20 = "Videbac1";
                String NameCol21 = "Videbac2";

                String DataCol1 = txtNumProcess.Text; 
                String DataCol2 = cbGFXDestination.Text;
                String DataCol22 = txtDiametre.Text;
                String DataCol3 = cbDensite.Text;
                String DataCol4 = txtLongueur.Text;
                String DataCol5 = cbTete.Text;
                String DataCol6 = cbType.Text;
                String DataCol7 = txtMatiere.Text;
                String DataCol8 = txtDebitEau1.Text;
                String DataCol9 = txtDebitEau2.Text;
                String DataCol10 = txtDebitEau3.Text;
                String DataCol11 = cbOutillage.Text;
                String DataCol12 = txtEcartement.Text;
                String DataCol13 = txtColoriVis1.Text;
                String DataCol14 = txtColoriVis4.Text;
                //Mettre un IsNumeric ERREUR ENREGISTREMENT - POSSIBILITE 1
                String DataCol15 = (Convert.ToDouble(ReadPLC(e1_consigne_vitesse_rotation))/10).ToString(); //A FAIRE
                String DataCol16 = (Convert.ToDouble(ReadPLC(e2_consigne_vitesse_rotation))/10).ToString();
                String DataCol17 = (Convert.ToDouble(ReadPLC(e3_consigne_vitesse_rotation))/10).ToString();
                String DataCol18 = (Convert.ToDouble(ReadPLC(e4_consigne_vitesse_rotation))/10).ToString();
                String DataCol19 = (Convert.ToDouble(ReadPLC(banc_tirage_consigne))/100).ToString();
                String DataCol20 = ReadEurotherm(videbac1_consigne);
                String DataCol21 = ReadEurotherm(videbac2_instant);

                string cmdText = "INSERT INTO [EnregistrementOF$] " +
                    "([" + NameCol1 + "], [" + NameCol2 + "], [" + NameCol22 + "], [" + NameCol3 + "], [" + NameCol4 + "], [" + NameCol5 + "], [" + NameCol6 + "], [" + NameCol7 + "], [" + NameCol8 + "], [" + NameCol9 + "], [" + NameCol10 + "], [" + NameCol11 + "], [" + NameCol12 + "], [" + NameCol13 + "], [" + NameCol14 + "], [" + NameCol15 + "], [" + NameCol16 + "], [" + NameCol17 + "], [" + NameCol18 + "], [" + NameCol19 + "], [" + NameCol20 + "], [" + NameCol21 + "])" +
                    "VALUES " +
                    "('" + DataCol1 + "', '" + DataCol2 + "', '" + DataCol22 + "', '" + DataCol3 + "', '" + DataCol4 + "', '" + DataCol5 + "', '" + DataCol6 + "', '" + DataCol7 + "', '" + DataCol8 + "', '" + DataCol9 + "', '" + DataCol10 + "', '" + DataCol11 + "', '" + DataCol12 + "', '" + DataCol13 + "', '" + DataCol14 + "', '" + DataCol15 + "', '" + DataCol16 + "', '" + DataCol17 + "', '" + DataCol18 + "', '" + DataCol19 + "', '" + DataCol20 + "', '" + DataCol21 + "')";
                using (OleDbCommand command = new OleDbCommand(cmdText, connection))
                {
                    command.ExecuteNonQuery();
                }

                MessageBox.Show("Recette enregistrer.", "Information");
            }
        }


        #endregion

        #region EvenementPanel

        private void btnConfirmer_Click(object sender, EventArgs e)
        {
            if (txtDiametre.Text != "" && cbGFXDestination.Text != "" && txtLongueur.Text != "" && cbDensite.Text != "" && txtCadenceProd.Text != "" && cbTete.Text != "")
            {
                Longueur();
                btnCharger.Show();
                btnConfirmer.Hide();
                confirmTrue();
            }
            else if (txtDiametre.Text == "")
            {
                MessageBox.Show("Veuillez indiquer le diametre du tube");
            }
            else if (cbGFXDestination.Text == "")
            {
                MessageBox.Show("Veuillez indiquer le GFH de destination du tube");
            }
            else if (txtLongueur.Text == "")
            {
                MessageBox.Show("Veuillez indiquer la longueur du tube");
            }
            else if (cbDensite.Text == "")
            {
                MessageBox.Show("Veuillez indiquer la densité du tube");
            }
            else if (txtCadenceProd.Text == "")
            {
                MessageBox.Show("Veuillez indiquer une cadence de production du tube");
            }
            else if (cbTete.Text == "")
            {
                MessageBox.Show("Veuillez indiquer un type de tête");
            }
            EnregConfOFFEnCour();
            OFFEnCourSVG.EtatBTN = "Confirmer";
            OFFEnCourSVG.Save();
        }

        private void btnCharger_Click(object sender, EventArgs e)
        {  
            btnNewRecette.Hide();
            btnDemarrageProd.Show();
            btnConfirmer.Hide();
            btnCharger.Hide();
            btnReinitialiser.Hide();
            lblNumProcessus.Text = "Numero Processus en cour";
            lblCaract.Text = "Caractéristique en cour";


            WriteZumbach(activer_progZumbach, lblProgZumbach.Text);

            WritePLC(activeProgPLC, lblProgPLC.Text);

            ReadDBVitesse(lblProgPLC.Text);

            ushort ligne = (ushort)(Convert.ToUInt16(lblProgPLC.Text));
            Double VitesseParDefaultE1 = (Double)Convert.ToDouble(BDDCoexPLC.Rows[ligne].Cells[3].Value) * 10;
            Double VitesseParDefaultE2 = (Double)Convert.ToDouble(BDDCoexPLC.Rows[ligne].Cells[4].Value) * 10;
            Double VitesseParDefaultE3 = (Double)Convert.ToDouble(BDDCoexPLC.Rows[ligne].Cells[5].Value) * 10;
            Double VitesseParDefaultE4 = (Double)Convert.ToDouble(BDDCoexPLC.Rows[ligne].Cells[6].Value) * 10;
            Double VitesseParDefaultBT = (Double)Convert.ToDouble(BDDCoexPLC.Rows[ligne].Cells[7].Value) * 100;

            //Console.WriteLine(VitesseParDefaultE1 + " " + VitesseParDefaultE2 + " " + VitesseParDefaultE3 + " " + VitesseParDefaultE4 + " " + VitesseParDefaultBT);

            WritePLC(e1_consigne_vitesse_rotation, VitesseParDefaultE1.ToString());
            WritePLC(e2_consigne_vitesse_rotation, VitesseParDefaultE2.ToString());
            WritePLC(e3_consigne_vitesse_rotation, VitesseParDefaultE3.ToString());
            WritePLC(e4_consigne_vitesse_rotation, VitesseParDefaultE4.ToString());
            //MessageBox.Show("banc_tirage_consigne");
            WritePLC(banc_tirage_consigne, VitesseParDefaultBT.ToString()); //PB

            float longueurTube = (float)Convert.ToDouble(txtLongueur.Text) * 100;
            WritePLC(longueur_coupe_consigne, longueurTube.ToString());

            //MessageBox.Show("gestionTemperature");
            gestionTemperature();

            MessageBox.Show("Numero de programme PLC et Zumbach chargé et pret pour le démarrage.\n" +
                            "Temperature par default chergé\n\n" +
                            "Merci de cliquer sur 'Lancer la PRODUCTION'\n" +
                            "Une fois la ligne prête à produire à la cadence voulu.", "Information");

            EnregConfOFFEnCour();
            OFFEnCourSVG.EtatBTN = "Charger";
            OFFEnCourSVG.Save();
        }

        private void btnDemarrageProd_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("VITESSE PROG DEMARRAG");
            //VITESSE PROG DEMARRAGE ------------------------------------------------------------------------------------
            double VitesseParDefaultE1 = (double)Convert.ToDouble(ReadPLC(e1_consigne_vitesse_rotation));
            double VitesseParDefaultE2 = (double)Convert.ToDouble(ReadPLC(e2_consigne_vitesse_rotation));
            double VitesseParDefaultE3 = (double)Convert.ToDouble(ReadPLC(e3_consigne_vitesse_rotation));
            double VitesseParDefaultE4 = (double)Convert.ToDouble(ReadPLC(e4_consigne_vitesse_rotation));
            double VitesseParDefaultBT = (double)Convert.ToDouble(ReadPLC(banc_tirage_consigne));
            //VITESSE PROG DEMARRAGE ------------------------------------------------------------------------------------
            //MessageBox.Show(VitesseParDefaultE1 + "\n" + VitesseParDefaultE2 + "\n" + VitesseParDefaultE3 + "\n" + VitesseParDefaultE4 + "\n" + VitesseParDefaultBT, "Vitesse maintenant");


            WritePLC(nextProg, lblProgProd.Text);

            ReadDBVitesse(lblProgProd.Text);

            WritePLC(timer, "12000");

            double NewVitesse = (double)Convert.ToDouble(txtCadenceProd.Text); //New cadencce
            //MessageBox.Show("NewVitessetxtCadenceProd " + NewVitesse.ToString());

            //A CHANGER ==================================================================================
            //double Vitesse = (double)Convert.ToDouble(ReadPLC(cadence_coupe)); // Cadence actuelle
            //A CHANGER ==================================================================================

            double Vitesse = VitesseParDefaultBT / ((double)Convert.ToDouble(txtLongueur.Text)/10);
            

            //MessageBox.Show("Vitessecadence_coupe "+ Vitesse.ToString());

            //MessageBox.Show("VITESSE PROG PROD");
            //VITESSE PROG PROD -----------------------------------------------------------------------------------------
            int NewVitesseParDefaultE1 = (int)Math.Round(((NewVitesse * VitesseParDefaultE1) / Vitesse), 0);
            int NewVitesseParDefaultE2 = (int)Math.Round(((NewVitesse * VitesseParDefaultE2) / Vitesse), 0);
            int NewVitesseParDefaultE3 = (int)Math.Round(((NewVitesse * VitesseParDefaultE3) / Vitesse), 0);
            int NewVitesseParDefaultE4 = (int)Math.Round(((NewVitesse * VitesseParDefaultE4) / Vitesse), 0);
            int NewVitesseParDefaultBT = (int)Math.Round(((NewVitesse * VitesseParDefaultBT) / Vitesse), 0);
            //VITESSE PROG PROD -----------------------------------------------------------------------------------------
            
            //MessageBox.Show(NewVitesseParDefaultE1 + "\n" + NewVitesseParDefaultE2 + "\n" + NewVitesseParDefaultE3 + "\n" + NewVitesseParDefaultE4 + "\n" + NewVitesseParDefaultBT, "Vitesse maintenant");
            //MessageBox.Show("WritePLC e1_consigne_vitesse_rotation");
            WritePLC(e1_consigne_vitesse_rotation, NewVitesseParDefaultE1.ToString());
            //MessageBox.Show("WritePLC e2_consigne_vitesse_rotation");
            WritePLC(e2_consigne_vitesse_rotation, NewVitesseParDefaultE2.ToString());
            //MessageBox.Show("WritePLC e3_consigne_vitesse_rotation");
            WritePLC(e3_consigne_vitesse_rotation, NewVitesseParDefaultE3.ToString());
            //MessageBox.Show("WritePLC e4_consigne_vitesse_rotation");
            WritePLC(e4_consigne_vitesse_rotation, NewVitesseParDefaultE4.ToString());
            //MessageBox.Show("WritePLC banc_tirage_consigne");
            WritePLC(banc_tirage_consigne, NewVitesseParDefaultBT.ToString()); //PB

            //MessageBox.Show("Clique BTN");
            WritePLC(btnK1, "true"); //Clique bouton K1
            //MessageBox.Show("Relache BTN");
            WritePLC(btnK1, "false"); //Relache bouton K1

            btnDemarrageProd.Hide();

            MessageBox.Show("Les prog PRODUCTION vient d'être charger.\n"+
                            "Les nouvelles vitesses des vis ont étais chargés.\n\n" +
                            "Les paramètres ainsi que les recettes vont être enregistés toutes les 10min.\n\n" +
                            "Vous pouvez aussi changer la 'Vitesse de PROD' en appuyant sur l'ONGLET 'DASHBOARD'\n\n" +
                            "Merci de cliquer sur 'Nouvelle PRODUCTION'\n" +
                            "Une fois la ligne prête à changer de commande.", "Information");

            btnNewRecette.Show();
            OFFEnCourSVG.EtatBTN = "Production";
            OFFEnCourSVG.Save();

        }

        private void btnReinitialiser_Click(object sender, EventArgs e)
        {
            btnCharger.Hide();
            btnDemarrageProd.Hide();
            btnConfirmer.Show();
            btnReinitialiser.Show();
            btnNewRecette.Hide();
            reini();
            confirmFalse();
            EnregConfOFFEnCour();
            OFFEnCourSVG.EtatBTN = "Vide";
        }

        private void btnNewRecette_Click(object sender, EventArgs e)
        {
            lblNumProcessus.Text = "Indiquer le numero de Processus";
            btnReinitialiser_Click(sender, e);
        }

        private void Recette_Load(object sender, EventArgs e)
        {
            //btnReinitialiser_Click(sender, e);
            OFFEnCour();
        }

        private void txtNumProcess_TextChanged(object sender, EventArgs e)
        {
            string[] plop = txtNumProcess.Text.Split('/');
            if (plop.Length == 2)
            {
                if (Int32.Parse(plop[1].Length.ToString()) == 2)
                {
                    bool TestProcessus = ReadExcelBDD();
                    if (TestProcessus == true)
                    {
                        MessageBox.Show("Numero de processus trouvé", "Information");
                       


                    }
                    else
                    {
                        MessageBox.Show("Numero de processus non trouvé.\nMerci de remplir les informations manuellements.", "Information");
                    }
                }
            }
            else
            {
                ReadExcelBDD();
            }
        }

        private void txtLongueur_TextChanged(object sender, EventArgs e)
        {
            searchProg();
            //Longueur();
            ////Console.WriteLine("txtLongueur_TextChanged");
        }

        private void txtMatiere_TextChanged(object sender, EventArgs e)
        {
            //searchMatiereGREEN();
        }

        private void txtDiametre_TextChanged(object sender, EventArgs e)
        {
            searchProg();
            searchVitesse();
            ////Console.WriteLine("txtDiametre_TextChanged");
        }

        private void cbDensite_SelectedValueChanged(object sender, EventArgs e)
        {
            searchProg();
        }

        private void txtNumProcess_KeyPress(object sender, KeyPressEventArgs e)
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

        private void cbGFXDestination_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbDensite_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchProg();
        }

        private void cbGFXDestination_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchProg();
        }

        #endregion

        private void bntEnregistrer_Click(object sender, EventArgs e)
        {
            WriteExcel_Enregistrement();
        }

        private void pnlRegManu_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}




