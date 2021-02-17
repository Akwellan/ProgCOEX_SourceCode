using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Net.SourceForge.Koogra.Excel;
using Siemens.Opc.Da;

namespace Siemens.Opc.DaClient.Controller
{
    public partial class Dash : UserControl
    {

        enum ClientHandles
        {
            Item1 = 0, //Zumbach OF
            Item2, //Zumbach Pross
            Item3, //Zumbach Matiere Tube
            Item4, //Zumbach Matiere Tube
            Item15, //Zumbach Matiere Tube
            Item5 = 0, //PLC Vitesse Extru 1
            Item6, //PLC Vitesse Extru 2
            Item7, //PLC Vitesse Extru 3
            Item8, //PLC Vitesse Extru 4
            Item9, //PLC Vitesse Extru 5
            Item10, //Longueurur
            Item11, //Cadence
            Item13 = 0,
            Item14,
            ItemBlockRead,
            ItemBlockWrite
        };

        #region Variable


        Worksheet BDDVitesse = new Workbook(@"BDD\\COEX.xls").Sheets.GetByName("BDD_VitessePROD");



        private Server servZumbach = null;
        private Subscription m_Subscription = null;
        private Server servPLC = null;
        private Subscription m_Subscription2 = null;
        private Server servEurotherm = null;
        private Subscription m_Subscription3 = null;

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

        string e1_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W160";  // 158 - 15.8

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

        string e2_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W162";  // 158 - 15.8

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

        string e3_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W164";  // 158 - 15.8

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


        // ----------------------- Complément Extrudeuse 4 ---------------------

        // ----------------------- Consigne vitesse de rotation-----------------

        string e4_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W166";  // 158 - 15.8

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
        string banc_tirage_consigne = "S7:[Liaison_S7_1],DB31,W168";  // 1850 - 18.50


        // ----------------------- Banc de coupe--------------------------------

        const string longueur_coupe_instant = "S7:[Liaison_S7_1],DB70,W2"; // 6700 - 67
        const string longueur_coupe_consigne = "S7:[Liaison_S7_1],DB30,W86";  // 6700 - 67
        const string cadence_coupe = "S7:[Liaison_S7_1],DB30,W84";  // 1850 - 185.0
        const string compteur_piece = "S7:[Liaison_S7_1],DB30,W94";  // 0 - 0


        //PROG
        const string activeProg = "S7:[Liaison_S7_1],DB31,W2";  // 0 - 0
        const string nextProg = "S7:[Liaison_S7_1],DB31,W6";  // 0 - 0
        const string timer = "S7:[Liaison_S7_1],DB31,W0";  // 0 - 0




        // ----------------------- Zumbach---------------------------------------

        // ----------------------- Regulation diametre---------------------------

        const string numero_prog_zumbach = "Usys_200_Coex/Product/c1100";  // 
        const string activer_prog = "Usys_200_Coex/Product/c1040";  // W
        const string nom_prog_zumbach = "Usys_200_Coex/Product/c1110";
        const string matiere_tube = "Usys_200_Coex/Product/c1115";
        const string diametre_mesure = "Usys_200_Coex/Data/d1100";
        const string diametre_cible = "Usys_200_Coex/Product/c2100.1";
        const string diametre_max = "Usys_200_Coex/Product/c2120.1";
        const string diametre_min = "Usys_200_Coex/Product/c2140.1";
        const string numero_of = "Usys_200_Coex/Data/d8000";
        const string numero_processus = "Usys_200_Coex/Data/d8000";
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
                lblConnexion.Text = "Connecté au Zumbach";

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
                lblConnexion2.Text = "Connecté au PLC";
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
                lblConnexion3.Text = "Connecté au Eurotherm";
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
                // delete all subscriptions
                stopMonitorItemsZumbach();
                stopMonitorItemsPLC();
                stopMonitorItemsEurotherm();

                // Change GUI settings
                lblConnexion.Text = "Déconnecté de l'Eurotherm";
                lblConnexion2.Text = "Déconnecté du PLC";
                lblConnexion3.Text = "Déconnecté du Zumbach";

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
        private void ConnectOPC_Eurotherm()
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
        private void ConnectOPC_Zumbach()
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
        public Dash()
        {
            InitializeComponent();
            // ----------------------- CONNEXION ---------------------------------
            ConnectOPC_Eurotherm(); //Connect OPC EUROTHERM
            ConnectOPC_PLC(); //Connect OPC PLC
            ConnectOPC_Zumbach(); //Connect OPC Zumbach
            // ----------------------- CONNEXION ---------------------------------

            ReadDBVitesse();

            //// ----------------------- MONITORING ---------------------------------
            MonitorZumbach();
            MonitorPLC();
            //// ----------------------- MONITORING ---------------------------------
            ///
            txtNewVitesse.Text = "";
            
        }
        #endregion

        #region Fonction d'évenement

        void ReadDBVitesse()
        {
            string numPorg = ReadPLC(activeProg);
            int numeroProg;
            bool IsNumeric = int.TryParse(numPorg, out numeroProg);

            if (IsNumeric == true) {
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
            try
            {
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
            }
            catch (Exception exception)
            {
                return exception.Message;
            }
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
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
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

        private void MonitorZumbach()
        {
            // Check if we have a subscription
            //  - No  -> Create a new subscription and create monitored items
            //  - Yes -> Delete Subcription
            if (m_Subscription == null)
            {
                startMonitorItemsZumbach();
            }
            else
            {
                stopMonitorItemsZumbach();
            }
        }
        private void MonitorPLC()
        {
            // Check if we have a subscription
            //  - No  -> Create a new subscription and create monitored items
            //  - Yes -> Delete Subcription
            if (m_Subscription2 == null)
            {
                startMonitorItemsPLC();
            }
            else
            {
                stopMonitorItemsPLC();
            }
        }
        private void MonitorEuroterm()
        {
            // Check if we have a subscription
            //  - No  -> Create a new subscription and create monitored items
            //  - Yes -> Delete Subcription
            if (m_Subscription3 == null)
            {
                startMonitorItemsEurotherm();
            }
            else
            {
                stopMonitorItemsEurotherm();
            }
        }

        #endregion

        #region Fonction Monitoring
        void startMonitorItemsZumbach()
        {
            // Check if we have a subscription. If not - create a new subscription.
            if (m_Subscription == null)
            {
                try
                {
                    m_Subscription = servZumbach.CreateSubscription("Subscription1", OnDataChangeZumbach);
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Create subscription failed:\n\n" + exception.Message);
                    return;
                }
            }
            // Add item 1 of
            try
            {
                m_Subscription.AddItem(code_operateur, (int)ClientHandles.Item1);

            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }
            // Add item 2 pross
            try
            {
                m_Subscription.AddItem(numero_processus, (int)ClientHandles.Item2);

            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }
            // Add item 3 diam tube  
            try
            {
                m_Subscription.AddItem(diametre_cible, (int)ClientHandles.Item3);

            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }
            // Add item 4 matiere tube
            try
            {
                m_Subscription.AddItem(matiere_tube, (int)ClientHandles.Item4);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }
            // Add item 15 matiere tube
            try
            {
                m_Subscription.AddItem(nom_prog_zumbach, (int)ClientHandles.Item15);


            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }
        }
        void startMonitorItemsPLC()
        {
            // Check if we have a subscription. If not - create a new subscription.
            if (m_Subscription2 == null)
            {
                try
                {
                    // Create subscription
                    m_Subscription2 = servPLC.CreateSubscription("Subscription2", OnDataChangePLC);
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Create subscription failed:\n\n" + exception.Message);
                    return;
                }
            }

            // Add item 5
            try
            {
                    m_Subscription2.AddItem(e1_consigne_vitesse_rotation, (int)ClientHandles.Item5);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

            // Add item 6
            try
            {
                m_Subscription2.AddItem(e2_consigne_vitesse_rotation, (int)ClientHandles.Item6);


            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

            // Add item 7
            try
            {
                m_Subscription2.AddItem(e3_consigne_vitesse_rotation, (int)ClientHandles.Item7);


            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

            // Add item 8
            try
            {
                m_Subscription2.AddItem(e4_consigne_vitesse_rotation, (int)ClientHandles.Item8);


            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

            // Add item 9
            try
            {
                m_Subscription2.AddItem(banc_tirage_consigne, (int)ClientHandles.Item9);


            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

            // Add item 10
            try
            {
                m_Subscription2.AddItem(longueur_coupe_consigne, (int)ClientHandles.Item10);


            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

            // Add item 11
            try
            {
                m_Subscription2.AddItem(cadence_coupe, (int)ClientHandles.Item11);


            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }
        }
        void startMonitorItemsEurotherm()
        {
            // Check if we have a subscription. If not - create a new subscription.
            if (m_Subscription3 == null)
            {
                try
                {
                    // Create subscription
                    m_Subscription3 = servEurotherm.CreateSubscription("Subscription3", OnDataChangeEurotherm);
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Create subscription failed:\n\n" + exception.Message);
                    return;
                }
            }

            // Add item 5
            try
            {
                m_Subscription3.AddItem(videbac1_consigne, (int)ClientHandles.Item13);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Create AddItem13 failed:\n\n" + exception.Message);
            }

            // Add item 6
            try
            {
                m_Subscription3.AddItem(videbac1_instant, (int)ClientHandles.Item14);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Create AddItem13 failed:\n\n" + exception.Message);
            }
        }
        void stopMonitorItemsZumbach()
        {
            if (m_Subscription != null)
            {
                try
                {
                    servZumbach.DeleteSubscription(m_Subscription);
                    m_Subscription = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Stopping data monitoring failed:\n\n" + ex.Message);
                }
            }
        }
        void stopMonitorItemsPLC()
        {
            if (m_Subscription2 != null)
            {
                try
                {
                    servPLC.DeleteSubscription(m_Subscription2);
                    m_Subscription2 = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Stopping data monitoring2 failed:\n\n" + ex.Message);
                }
            }
        }
        void stopMonitorItemsEurotherm()
        {
            if (m_Subscription3 != null)
            {
                try
                {
                    servEurotherm.DeleteSubscription(m_Subscription3);
                    m_Subscription3 = null;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Stopping data monitoring failed:\n\n" + ex.Message);
                }
            }
        }
        private void OnDataChangeZumbach(IList<DataValue> DataValues)
        {
            try
            {
                // We have to call an invoke method 
                if (this.InvokeRequired)
                {
                    // Asynchronous execution of the valueChanged delegate
                    this.BeginInvoke(new DataChange(OnDataChangeZumbach), DataValues);
                    return;
                }

                foreach (DataValue value in DataValues)
                {
                    // 1 is Item1, 2 is Item2, 3 is ItemBlockRead
                    switch (value.ClientHandle)
                    {
                        case (int)ClientHandles.Item1:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblNumOF.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                lblNumOF.Text = value.Value.ToString();
                            }
                            break;
                        case (int)ClientHandles.Item2:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblNumPross.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                lblNumPross.Text = value.Value.ToString();
                            }
                            break;
                        case (int)ClientHandles.Item3:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblDiamTube.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                lblDiamTube.Text = value.Value.ToString();
                            }
                            break;
                        case (int)ClientHandles.Item4:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblMatiereTube.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                lblMatiereTube.Text = value.Value.ToString();
                            }
                            break;
                        case (int)ClientHandles.Item15:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblDestination.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                lblDestination.Text = value.Value.ToString();
                            }
                            break;
                        default:
                            // error
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + ex.Message);
            }
        }
        private void OnDataChangePLC(IList<DataValue> DataValues)
        {
            try
            {
                // We have to call an invoke method 
                if (this.InvokeRequired)
                {
                    // Asynchronous execution of the valueChanged delegate
                    this.BeginInvoke(new DataChange(OnDataChangePLC), DataValues);
                    return;
                }

                foreach (DataValue value in DataValues)
                {
                    // 1 is Item1, 2 is Item2, 3 is ItemBlockRead
                    switch (value.ClientHandle)
                    {
                        case (int)ClientHandles.Item5:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblConsigneVis1.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConvertToFloat = (float)Convert.ToDouble(value.Value.ToString());
                                lblConsigneVis1.Text = (ConvertToFloat / 10).ToString();
                            }
                            break;
                        case (int)ClientHandles.Item6:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblConsigneVis2.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConvertToFloat = (float)Convert.ToDouble(value.Value.ToString());
                                lblConsigneVis2.Text = (ConvertToFloat / 10).ToString();
                            }
                            break;
                        case (int)ClientHandles.Item7:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblConsigneVis3.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConvertToFloat = (float)Convert.ToDouble(value.Value.ToString());
                                lblConsigneVis3.Text = (ConvertToFloat / 10).ToString();
                            }
                            break;
                        case (int)ClientHandles.Item8:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblConsigneVis4.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConvertToFloat = (float)Convert.ToDouble(value.Value.ToString());
                                lblConsigneVis4.Text = (ConvertToFloat / 10).ToString();
                            }
                            break;
                        case (int)ClientHandles.Item9:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblConsigneBancTirage.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConvertToFloat = (float)Convert.ToDouble(value.Value.ToString());
                                lblConsigneBancTirage.Text = (ConvertToFloat / 100).ToString();
                            }
                            break;
                        case (int)ClientHandles.Item10:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                lblLongTube.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConvertToFloat = (float)Convert.ToDouble(value.Value.ToString());
                                lblLongTube.Text = (ConvertToFloat / 100).ToString();
                            }
                            break;
                        case (int)ClientHandles.Item11:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                txtVitesse.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConvertToFloat = (float)Convert.ToDouble(value.Value.ToString());
                                txtVitesse.Text = (ConvertToFloat / 10).ToString();
                            }
                            break;
                        default:
                            // error
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error in the data change callback:\n\n" + ex.Message);
            }
        }
        private void OnDataChangeEurotherm(IList<DataValue> DataValues)
        {
            try
            {
                // We have to call an invoke method 
                if (this.InvokeRequired)
                {
                    // Asynchronous execution of the valueChanged delegate
                    this.BeginInvoke(new DataChange(OnDataChangeEurotherm), DataValues);
                    return;
                }

                foreach (DataValue value in DataValues)
                {
                    // 1 is Item1, 2 is Item2, 3 is ItemBlockRead
                    switch (value.ClientHandle)
                    {
                        case (int)ClientHandles.Item13:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                //lblConsigneVideBac1.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                //lblConsigneVideBac1.Text = value.Value.ToString();
                            }
                            break;
                        case (int)ClientHandles.Item14:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                //textBox6.Text = "Error: 0x" + value.Error.ToString("X");
                                //MessageBox.Show("Error: 0x" + value.Error.ToString("X"));
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                //textBox6.Text = value.Value.ToString();
                                //string ValueInstantVB1 = value.Value.ToString();
                                //aGauge1.Value = float.Parse(ValueInstantVB1);   //POSSIBILITE D'ERREUR
                            }
                            break;
                        default:
                            // error
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + ex.Message);
            }
        }

        private void searchVitesse()
        {
            bool b = false;
            for (ushort i = 1; i <= BDDVitesse.Rows.LastRow; i++)
            {
                lblDiamTube.Text.Replace(".", ",");
                int Diam = (int)Math.Round(Convert.ToDouble(lblDiamTube.Text), 0);
                if (Diam.ToString().Contains(BDDVitesse.Rows[i].Cells[0].Value.ToString()))
                {
                    lblVitesseObjectif.Text = BDDVitesse.Rows[i].Cells[3].Value.ToString();
                    b = true;
                    //Console.WriteLine("{0}\t{1}\t{2}\t{3}", worksheet.Rows[i].Cells[0].Value.ToString(), worksheet.Rows[i].Cells[1].Value.ToString(), worksheet.Rows[i].Cells[2].Value.ToString(), worksheet.Rows[i].Cells[6].Value.ToString());
                }
                else if (b == false)
                {
                    lblVitesseObjectif.Text = "";
                } 
            }
        }

        #endregion

        #region EvenementPanel

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }
        private void label5_Click(object sender, EventArgs e)
        {
              
        }
        private void panel21_Paint(object sender, PaintEventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (txtNewVitesse.Text != "" && txtNewVitesse.Text != "0" && txtVitesse.Text != "" && txtVitesse.Text != "0")
            {
                int NewConsigneVis1 = (int)((float)Convert.ToDouble(lblNewConsigneVis1.Text) * 10);
                int NewConsigneVis2 = (int)((float)Convert.ToDouble(lblNewConsigneVis2.Text) * 10);
                int NewConsigneVis3 = (int)((float)Convert.ToDouble(lblNewConsigneVis3.Text) * 10);
                int NewConsigneVis4 = (int)((float)Convert.ToDouble(lblNewConsigneVis4.Text) * 10);
                int NewConsigneBancTirage = (int)((float)Convert.ToDouble(lblNewConsigneBancTirage.Text) * 100);
                int NewViteseProd = (int)((float)Convert.ToDouble(txtNewVitesse.Text) * 10);
                //MessageBox.Show("NewConsigneVis1: "+ NewConsigneVis1+ "\nNewConsigneVis2: "+ NewConsigneVis2+ "\nNewConsigneVis3: "+ NewConsigneVis3 + "\nNewConsigneVis4: " + NewConsigneVis4 + "\nNewConsigneBancTirage: " + NewConsigneBancTirage+ "\n\nNewViteseProd: "+ NewViteseProd, "");
                WritePLC(e1_consigne_vitesse_rotation, NewConsigneVis1.ToString());
                WritePLC(e2_consigne_vitesse_rotation, NewConsigneVis2.ToString());
                WritePLC(e3_consigne_vitesse_rotation, NewConsigneVis3.ToString());
                WritePLC(e4_consigne_vitesse_rotation, NewConsigneVis4.ToString());
                WritePLC(banc_tirage_consigne, NewConsigneBancTirage.ToString());
                WritePLC(cadence_coupe, NewViteseProd.ToString());
                lblNewConsigneVis1.Text = "0";
                lblNewConsigneVis2.Text = "0";
                lblNewConsigneVis3.Text = "0";
                lblNewConsigneVis4.Text = "0";
                lblNewConsigneBancTirage.Text = "0";
                txtNewVitesse.Text = "";
            } else
            {
                MessageBox.Show("Ce programme fonctionne uniquement quand la ligne tourne.");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(txtNewVitesse.Text, "[^0-9]"))
            {
                MessageBox.Show("Merci de ne rentrer que des nombres entiers.","ERREUR");
                txtNewVitesse.Text = txtNewVitesse.Text.Remove(txtNewVitesse.Text.Length - 1);
                lblNewConsigneVis1.Text = "0";
                lblNewConsigneVis2.Text = "0";
                lblNewConsigneVis3.Text = "0";
                lblNewConsigneVis4.Text = "0";
                lblNewConsigneBancTirage.Text = "0";
                txtNewVitesse.Text = "";
            } else if(txtNewVitesse.Text != "" && txtNewVitesse.Text != "0")// && txtVitesse.Text != "" && txtVitesse.Text != "0")
            {
                float numVal = float.Parse(txtNewVitesse.Text);
                if(numVal > 400)
                {
                    MessageBox.Show("Plage de vitesse entre 10 et 400 t/min.");
                    txtNewVitesse.Text = "";
                    lblNewConsigneVis1.Text = "0";
                    lblNewConsigneVis2.Text = "0";
                    lblNewConsigneVis3.Text = "0";
                    lblNewConsigneVis4.Text = "0";
                    lblNewConsigneBancTirage.Text = "0";
                }
                else
                {
                    float VitesseVis1 = (float)Convert.ToDouble(lblConsigneVis1.Text);
                    float VitesseVis2 = (float)Convert.ToDouble(lblConsigneVis2.Text);
                    float VitesseVis3 = (float)Convert.ToDouble(lblConsigneVis3.Text);
                    float VitesseVis4 = (float)Convert.ToDouble(lblConsigneVis4.Text);
                    float VitesseBancTirage = (float)Convert.ToDouble(lblConsigneBancTirage.Text);
                    //float Vitesse = (float)Convert.ToDouble(txtVitesse.Text);

                    float Vitesse = VitesseBancTirage / ((float)Convert.ToDouble(lblLongTube.Text) / 1000);
                    float NewVitesse = (float)Convert.ToDouble(txtNewVitesse.Text);
                    float CalculNewVitesse1 = (float)Math.Round(((NewVitesse * VitesseVis1) / Vitesse),1);
                    float CalculNewVitesse2 = (float)Math.Round(((NewVitesse * VitesseVis2) / Vitesse),1);
                    float CalculNewVitesse3 = (float)Math.Round(((NewVitesse * VitesseVis3) / Vitesse),1);
                    float CalculNewVitesse4 = (float)Math.Round(((NewVitesse * VitesseVis4) / Vitesse),1);
                    float CalculNewVitesseBT = (float)Math.Round(((NewVitesse * VitesseBancTirage) / Vitesse),1);
                    lblNewConsigneVis1.Text = CalculNewVitesse1.ToString();
                    lblNewConsigneVis2.Text = CalculNewVitesse2.ToString();
                    lblNewConsigneVis3.Text = CalculNewVitesse3.ToString();
                    lblNewConsigneVis4.Text = CalculNewVitesse4.ToString();
                    lblNewConsigneBancTirage.Text = CalculNewVitesseBT.ToString();
                }
            } else if (txtNewVitesse.Text == "")
            {
                lblNewConsigneVis1.Text = "0";
                lblNewConsigneVis2.Text = "0";
                lblNewConsigneVis3.Text = "0";
                lblNewConsigneVis4.Text = "0";
                lblNewConsigneBancTirage.Text = "0";
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            txtVitesse.Text = "1";
        }

        #endregion

        private void txtNewVitesse_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) // && (e.KeyChar != '/'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            //if ((e.KeyChar == '/') && ((sender as TextBox).Text.IndexOf('/') > -1))
            //{
            //    e.Handled = true;
            //}
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblNewConsigneVis1_Click(object sender, EventArgs e)
        {

        }

        private void lblDiamTube_Click(object sender, EventArgs e)
        {

        }

        private void lblDiamTube_TextChanged(object sender, EventArgs e)
        {
            searchVitesse();
        }

        private void lblLongTube_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblVitesseObjectif_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {
            searchVitesse();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Dash_Load(object sender, EventArgs e)
        {

        }
    }
}
