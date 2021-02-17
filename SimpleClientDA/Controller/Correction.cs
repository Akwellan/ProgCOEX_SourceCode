using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows.Forms;
using Siemens.Opc.Da;
using LiveCharts.Geared;
using LiveCharts.Wpf;
using System.Windows.Media;

namespace Dashboard
{
    public partial class Correction : UserControl
    {
        private Boolean Button2Centieme = true;
        private Boolean CorrectionDimatre = true;
        private SpeedTestVm ChartDeviation = new SpeedTestVm();

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

        const string e1_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W48";  // 158 - 15.8

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

        const string e2_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W50";  // 158 - 15.8

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

        const string e3_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W52";  // 158 - 15.8

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

        const string e4_consigne_vitesse_rotation = "S7:[Liaison_S7_1],DB31,W54";  // 158 - 15.8

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
        const string banc_tirage_consigne = "S7:[Liaison_S7_1],DB31,W56";  // 1850 - 18.50


        // ----------------------- Banc de coupe--------------------------------

        const string longueur_coupe_instant = "S7:[Liaison_S7_1],DB70,W2"; // 6700 - 67
        const string longueur_coupe_consigne = "S7:[Liaison_S7_1],DB30,W86";  // 6700 - 67
        const string cadence_coupe = "S7:[Liaison_S7_1],DB30,W84";  // 1850 - 185.0
        const string compteur_piece = "S7:[Liaison_S7_1],DB30,W94";  // 0 - 0





        // ----------------------- Zumbach---------------------------------------

        // ----------------------- Regulation diametre---------------------------

        const string numero_prog_zumbach = "Usys_200_Coex/Product/c1100";  // 
        const string activer_prog = "Usys_200_Coex/Product/c1040";  // W
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

                lblConnexion3.Text = "Déconnecté du Eurotherm";
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
                lblConnexion3.Text = "Déconnecté de l'Eurotherm";
                lblConnexion2.Text = "Déconnecté du PLC";
                lblConnexion.Text = "Déconnecté du Zumbach";

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

        private void OnDisconnectEuroterm()
        {
            if (servEurotherm == null)
            {
                return;
            }

            try
            {
                // delete all subscriptions
                stopMonitorItemsEurotherm();

                // Change GUI settings
                lblConnexion3.Text = "Déconnecté de l'Eurotherm";

                // Disconnect
                servEurotherm.Disconnect();

                // Cleanup
                servEurotherm = null;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Disconnect failed");
            }
        }

        private void OnDisconnectZumbach()
        {
            if (servZumbach == null)
            {
                return;
            }

            try
            {
                // delete all subscriptions
                stopMonitorItemsZumbach();

                // Change GUI settings
                lblConnexion.Text = "Déconnecté du Zumbach";

                // Disconnect
                servZumbach.Disconnect();

                // Cleanup
                servZumbach = null;
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
                OnDisconnectZumbach();
            }
            else
            {
                OnDisconnectZumbach();
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
                OnDisconnectEuroterm();
            }
            else
            {
                OnDisconnectEuroterm();
            }
        }

        #endregion

        #region Construction
        public Correction()
        {
            InitializeComponent();
            aGauge1.MinValue = -5000;
            aGauge1.MaxValue = -4000;
            aGauge1.ScaleLinesMajorStepValue = 100;
            aGauge1.Value = -4500;

            // ----------------------- CONNEXION ---------------------------------
            ConnectOPC_Eurotherm(); //Connect OPC EUROTHERM
            ConnectOPC_PLC(); //Connect OPC PLC
            ConnectOPC_Zumbach(); //Connect OPC Zumbach
            // ----------------------- CONNEXION ---------------------------------


            //// ----------------------- MONITORING ---------------------------------
            MonitorZumbach(); 
            MonitorEuroterm();
            //// ----------------------- MONITORING ---------------------------------
            ///

            
            //StartTimer(); //Demarrage du timer
            DeviationChart();
            ChartDeviation.Read();
        }
        #endregion

        #region Fonction BUG

        private void BUG_Euroterm()
        {
            OnDisconnectEuroterm();
            ConnectOPC_Eurotherm();
            MonitorEuroterm();
        }

        private void BUG_Zumbach()
        {
            OnDisconnectZumbach();
            ConnectOPC_Zumbach();
            MonitorZumbach();
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
                BUG_Euroterm();
                //MessageBox.Show(exception.Message, "Erreur READ EUROTERM");
                return "E";
            }
        }
        private string ReadPLC(string NomItem)
        {
            //try
            //{
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
            //}
            //catch (Exception exception)
            //{
            //    return exception.Message;
            //}
        }
        private string ReadZumbach(string NomItem)
        {
            try
            {
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
            }
            catch (Exception exception)
            {
                BUG_Zumbach();
                //MessageBox.Show(exception.Message, "Erreur READ ZUMBACH");
                return "E";
            }
        }

        private void WriteEurotherm(string NomItemID, string Valeur)
        {
            try
            {
                this.servEurotherm.Write(NomItemID, Valeur);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Unexpected error in the data change callback:\n\n" + exception.Message);
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

        #region Fonction d'Evenement
        /// <summary>
        /// Show shutdown message
        /// When receiving a shutdown message just disconnect.
        /// </summary>
        /// 

        private void DeviationChart()
        {
            cartesianChart1.Series.Add(new GLineSeries
            {
                Values = ChartDeviation.Values,
                StrokeThickness = 4, //epaisseur
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255)),
                Fill = System.Windows.Media.Brushes.Transparent,
                LineSmoothness = 1, //Lissage
                PointGeometrySize = 15,

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
                        Stroke = new SolidColorBrush(System.Windows.Media.Color.FromArgb(10,248, 213, 72))
                    },
                    new AxisSection

                    {
                        Value = -0.03,
                        SectionWidth = 0.06,
                        Fill = new SolidColorBrush
                        {
                            Color = System.Windows.Media.Color.FromArgb(10,0,204,0),
                            Opacity = 80
                        }
                    },
                    new AxisSection
                    {
                        Value = -0.05,
                        SectionWidth = 0.02,
                        Fill = new SolidColorBrush
                        {
                            Color = System.Windows.Media.Color.FromArgb(10,254,132,132),
                            Opacity = 50
                        }
                    },
                    new AxisSection
                    {
                        Value = 0.03,
                        SectionWidth = 0.02,
                        Fill = new SolidColorBrush
                        {
                            Color = System.Windows.Media.Color.FromArgb(10,254,132,132),
                            Opacity = 50
                        }
                    }
                }
            });

            
        }

        private void CorrectGauge()
        {
            if ((aGauge1.MinValue == aGauge1.Value || aGauge1.Value == aGauge1.MaxValue))
            {
                
                aGauge1.MaxValue = Int32.Parse(lblConsigneVideBac1.Text) + 3;
                aGauge1.MinValue = Int32.Parse(lblConsigneVideBac1.Text) - 3;
                aGauge1.ScaleLinesMajorStepValue = 1;
            }
        }

        /// <summary>
        /// callback to receive datachanges
        /// </summary>
        /// <param name="clientHandle"></param>
        /// <param name="value"></param>
        private void OnDataChange(IList<DataValue> DataValues)
        {
            try
            {
                // We have to call an invoke method 
                if (this.InvokeRequired)
                {
                    // Asynchronous execution of the valueChanged delegate
                    this.BeginInvoke(new DataChange(OnDataChange), DataValues);
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
                                lblDeviation.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                lblDeviation.Text = value.Value.ToString();
                                float Deviation = (float)Convert.ToDouble(value.Value.ToString());
                                ChartDeviation._trend = Deviation;
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
        private void OnDataChange2(IList<DataValue> DataValues)
        {
            try
            {
                // We have to call an invoke method 
                if (this.InvokeRequired)
                {
                    // Asynchronous execution of the valueChanged delegate
                    this.BeginInvoke(new DataChange(OnDataChange2), DataValues);
                    return;
                }

                foreach (DataValue value in DataValues)
                {
                    // 1 is Item1, 2 is Item2, 3 is ItemBlockRead
                    switch (value.ClientHandle)
                    {
                        case (int)ClientHandles.Item3:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                //                               txtMonitor3.Text = "Error: 0x" + value.Error.ToString("X");
                                //                               txtMonitor3.BackColor = Color.Red;
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                //                                txtMonitor3.Text = value.Value.ToString();
                                //                                txtMonitor3.BackColor = Color.White;
                            }
                            break;
                        case (int)ClientHandles.Item4:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                //                                txtMonitor4.Text = "Error: 0x" + value.Error.ToString("X");
                                //                                txtMonitor4.BackColor = Color.Red;
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                //                                txtMonitor4.Text = value.Value.ToString();
                                //                                txtMonitor4.BackColor = Color.White;
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
        private void OnDataChange3(IList<DataValue> DataValues)
        {
            try
            {
                // We have to call an invoke method 
                if (this.InvokeRequired)
                {
                    // Asynchronous execution of the valueChanged delegate
                    this.BeginInvoke(new DataChange(OnDataChange3), DataValues);
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
                                lblConsigneVideBac1.Text = "Error: 0x" + value.Error.ToString("X");
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                float ConsigneVideBac1 = (float)Convert.ToDouble(value.Value.ToString());
                                float NewConsigneVideBac1 = (float)Math.Round(ConsigneVideBac1, 0);     //POSSIBILITE DERREUR
                                lblConsigneVideBac1.Text = NewConsigneVideBac1.ToString();
                            }
                            break;
                        case (int)ClientHandles.Item6:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                //textBox6.Text = "Error: 0x" + value.Error.ToString("X");
                                MessageBox.Show("Error: 0x" + value.Error.ToString("X"));
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                //textBox6.Text = value.Value.ToString();
                                string ValueInstantVB1 = value.Value.ToString();
                                aGauge1.Value = float.Parse(ValueInstantVB1);   //POSSIBILITE D'ERREUR
                                CorrectGauge();
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
        #endregion

        #region Fonction Monitoring
        void startMonitorItemsZumbach()
        {
            // Check if we have a subscription. If not - create a new subscription.
            if (m_Subscription == null)
            {
                try
                {
                    m_Subscription = servZumbach.CreateSubscription("Subscription1", OnDataChange);
                }
                catch (Exception exception)
                {
                    lblDeviation.Text = ("Create subscription failed:\n\n" + exception.Message);
                    return;
                }
            }
            // Add item 1
            try
            {
                m_Subscription.AddItem(deviation, (int)ClientHandles.Item1);

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
                    m_Subscription2 = servPLC.CreateSubscription("Subscription2", OnDataChange2);
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Create subscription failed:\n\n" + exception.Message);
                    return;
                }
            }

            // Add item 3
            try
            {
                //                m_Subscription2.AddItem(txtItemID3.Text, (int)ClientHandles.Item3);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Unexpected error in the data change callback:\n\n" + exception.Message);
            }

            // Add item 4
            try
            {
                //               m_Subscription2.AddItem(txtItemID4.Text, (int)ClientHandles.Item4);


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
                    m_Subscription3 = servEurotherm.CreateSubscription("Subscription3", OnDataChange3);
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
                m_Subscription3.AddItem(videbac1_consigne, (int)ClientHandles.Item5);
            }
            catch (Exception exception)
            {
                lblConsigneVideBac1.Text = ("Create AddItem5 failed:\n\n" + exception.Message);
            }

            // Add item 6
            try
            {
                m_Subscription3.AddItem(videbac1_instant, (int)ClientHandles.Item6);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Create AddItem6 failed:\n\n" + exception.Message);
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

        #endregion

        #region Fonction de Correction diametre

        void CorrectionDiamViaVide()
        {
            try
            {
                string TestEuroterm = ReadEurotherm(videbac1_consigne);
                string TestZumbach = ReadZumbach(deviation); 
            
                //int TestZumbachNumeric;
                //int TestEurotermNumeric;
                //bool IsNumericEuroterm = int.TryParse(TestEuroterm, out TestEurotermNumeric);
                //bool IsNumericZumbach = int.TryParse(TestZumbach, out TestZumbachNumeric);
                //MessageBox.Show(TestEuroterm + " " + IsNumericEuroterm + "\n" + TestZumbach + " " + IsNumericZumbach);
                //if (IsNumericEuroterm == true && IsNumericZumbach == true)
                //{
                    if (TestEuroterm != "E" && TestZumbach != "E")
                    {
                        //AJOUTER RAJOUT ISNUMERIC ANTI BUG "INPUT STRING WAS NOT IN A CORRECT FORMAT"
                        int Val_VideBac1 = (int)Math.Round(Convert.ToDouble(ReadEurotherm(videbac1_consigne)), 0); //RAJOUT IS NUMERIC
                        double Val_Deviation = (double)Math.Round(Convert.ToDouble(ReadZumbach(deviation)), 2);
                        int CorrectionVide = 0;

                        switch (Val_Deviation)
                        {
                            case 0.1:
                                CorrectionVide = 5;
                                break;
                            case 0.09:
                                CorrectionVide = 4;
                                break;
                            case 0.08:
                                CorrectionVide = 4;
                                break;
                            case 0.07:
                                CorrectionVide = 3;
                                break;
                            case 0.06:
                                CorrectionVide = 3;
                                break;
                            case 0.05:
                                CorrectionVide = 2;
                                break;
                            case 0.04:
                                CorrectionVide = 2;
                                break;
                            case 0.03:
                                CorrectionVide = 1;
                                break;
                            case 0.02:
                                if (Button2Centieme == true)
                                {
                                    CorrectionVide = 1;
                                }
                                else if (Button2Centieme == false)
                                {
                                    CorrectionVide = 0;
                                }
                                break;
                            case 0.01:
                                CorrectionVide = 0;
                                break;
                            case 0.00:
                                CorrectionVide = 0;
                                break;
                            case -0.01:
                                CorrectionVide = 0;
                                break;
                            case -0.1:
                                CorrectionVide = -5;
                                break;
                            case -0.09:
                                CorrectionVide = -4;
                                break;
                            case -0.08:
                                CorrectionVide = -4;
                                break;
                            case -0.07:
                                CorrectionVide = -3;
                                break;
                            case -0.06:
                                CorrectionVide = -3;
                                break;
                            case -0.05:
                                CorrectionVide = -2;
                                break;
                            case -0.04:
                                CorrectionVide = -2;
                                break;
                            case -0.03:
                                CorrectionVide = -1;
                                break;
                            case -0.02:
                                if (Button2Centieme == true)
                                {
                                    CorrectionVide = -1;
                                }
                                else if (Button2Centieme == false)
                                {
                                    CorrectionVide = 0;
                                }
                                break;
                        }

                        lblConsigneVideBac1Avant.Text = Val_VideBac1.ToString();
                        lblDeviationCorrect.Text = Val_Deviation.ToString();
                        lblCorrectVide.Text = CorrectionVide.ToString();
                        lblCorrectVideBac1.Text = (Val_VideBac1 + CorrectionVide).ToString();

                        if (CorrectionDimatre == true)
                        {
                            WriteEurotherm(videbac1_consigne, lblCorrectVideBac1.Text);
                        }
                    }
                    else
                    {
                        //MessageBox.Show("Probleme de communication au Automate\n" +
                        // "\nTestEuroterm = " + TestEuroterm +
                        //  "\nTestZumbach = " + TestZumbach);
                    }
                //}
            } catch
            {
                lblConsigneVideBac1Avant.Text = "ERREUR";
                lblDeviationCorrect.Text = "ERREUR";
                lblCorrectVide.Text = "ERREUR";
                lblCorrectVideBac1.Text = "ERREUR";
            }
            
        }

        #endregion

        #region EvenementPanel
        private void correctionOn_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = correctionOn.Height;
            pnlBtn.Top = correctionOn.Top;
            pnlBtn.Left = correctionOn.Left;
            correctionOn.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            correctionOff.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            CorrectionDimatre = true;
        }

        private void correctionOff_Click(object sender, EventArgs e)
        {
            pnlBtn.Height = correctionOff.Height;
            pnlBtn.Top = correctionOff.Top;
            pnlBtn.Left = correctionOff.Left;
            correctionOff.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            correctionOn.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            CorrectionDimatre = false;
        }

        private void mode2ct_Click(object sender, EventArgs e)
        {
            PnlMode.Height = mode2ct.Height;
            PnlMode.Top = mode2ct.Top;
            PnlMode.Left = mode2ct.Left;
            mode2ct.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            mode3ct.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            Button2Centieme = true;
        }

        private void mode3ct_Click(object sender, EventArgs e)
        {
            PnlMode.Height = mode3ct.Height;
            PnlMode.Top = mode3ct.Top;
            PnlMode.Left = mode3ct.Left;
            mode3ct.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            mode2ct.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            Button2Centieme = false;
        }

        private void correctionOff_Leave(object sender, EventArgs e)
        {
            correctionOff.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
        }

        private void correctionOn_Leave(object sender, EventArgs e)
        {
            correctionOn.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
        }

        private void Correction_Load(object sender, EventArgs e)
        {
            correctionOff_Click(sender, e);
            mode2ct_Click(sender, e);
            btn1s_Click(sender, e);
        }

        /// <summary>
        /// POSSIBILITE DERREUR
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void btn1s_Click(object sender, EventArgs e) //30s
        {
            pnlBtnTimer.Height = btn30s.Height;
            pnlBtnTimer.Top = btn30s.Top;
            pnlBtnTimer.Left = btn30s.Left;
            btn30s.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btn40s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn60s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn50s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            this.TimerCorrection.Interval = 30000;
            //StartTimer();
        }

        private void btn10s_Click(object sender, EventArgs e) //40s
        {
            pnlBtnTimer.Height = btn40s.Height;
            pnlBtnTimer.Top = btn40s.Top;
            pnlBtnTimer.Left = btn40s.Left;
            btn40s.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btn30s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn60s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn50s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            this.TimerCorrection.Interval = 40000;
            //StartTimer();
        }

        private void btn30s_Click(object sender, EventArgs e) //60s
        {
            pnlBtnTimer.Height = btn60s.Height;
            pnlBtnTimer.Top = btn60s.Top;
            pnlBtnTimer.Left = btn60s.Left;
            btn60s.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btn40s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn30s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn50s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            this.TimerCorrection.Interval = 60000;
           // StartTimer();
        }

        private void btn15s_Click(object sender, EventArgs e) //50s
        {
            pnlBtnTimer.Height = btn50s.Height;
            pnlBtnTimer.Top = btn50s.Top;
            pnlBtnTimer.Left = btn50s.Left;
            btn50s.BackColor = System.Drawing.Color.FromArgb(46, 51, 73);
            btn40s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn60s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            btn30s.BackColor = System.Drawing.Color.FromArgb(37, 42, 64);
            this.TimerCorrection.Interval = 50000;

           // StartTimer();
        }

        private void EtatTimerHeure_Click(object sender, EventArgs e)
        {

        }

        private void aGauge1_ValueInRangeChanged(object sender, ValueInRangeChangedEventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            TimerCorrection.Enabled = false;

        }

        private void lblConsigneVideBac1Avant_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

            CorrectionDiamViaVide();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            TimerCorrection.Enabled = true;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            TimerCorrection.Interval = 1000;
        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void lblDeviation_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void TimerCorrection_Tick(object sender, EventArgs e)
        {
            //string EtatTimer = e.SignalTime.ToString();
            string Date = DateTime.Now.ToString("dd/MM/yyyy");
            string Heure = DateTime.Now.ToString("HH:mm:ss");
            //Console.WriteLine("Date :" + Date + "Heure :" + Heure);
            EtatTimerDate.Text = Date;
            EtatTimerHeure.Text = Heure;
            //CorrectGauge();
            CorrectionDiamViaVide();
        }

        #endregion

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

    }
}
