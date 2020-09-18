using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Siemens.Opc;
using Siemens.Opc.Da;

namespace Siemens.Opc.DaClient
{
    public partial class SimpleClientDA : Form
    {
        enum ClientHandles
        {
            Item1 = 0,
            Item2,
            ItemBlockRead,
            ItemBlockWrite
        };

        #region Connect and Disconnect Server
        /// <summary>
        /// Handles connect procedure
        /// </summary>
        private void OnConnect()
        {
            if (m_Server == null)
            {
                // Create a server object
                m_Server = new Server();
            }
            if (m_Server2 == null)
            {
                // Create a server object
                m_Server2 = new Server();
            }
            if (m_Server3 == null)
            {
                // Create a server object
                m_Server3 = new Server();
            }

            try
            {
                // connect to the server
                m_Server.Connect(txtServerUrl.Text);
                m_Server2.Connect(txtServerUrl.Text);
                m_Server2.Connect(txtServerUrl.Text);

                // Change GUI settings
                btnConnect.Text = "Disconnect";
                txtServerUrl.Enabled = false;
                txtServerUrl2.Enabled = false;
                txtServerUrl3.Enabled = false;

                // enable buttons
                btnMonitor.Enabled = true;
                btnMonitor2.Enabled = true;
                btnMonitor3.Enabled = true;
                btnRead.Enabled = true;
                btnRead2.Enabled = true;
                btnRead3.Enabled = true;
                btnWrite.Enabled = true;
                btnWrite2.Enabled = true;
                btnWrite3.Enabled = true;
                btnMonitorBlock.Enabled = true;
                btnWriteBlock1.Enabled = true;
                btnWriteBlock2.Enabled = true;
            }
            catch (Exception exception)
            {
                // Cleanup
                m_Server = null;

                MessageBox.Show(exception.Message, "Connect failed");
            }
        }

        /// <summary>
        /// Handles disconnect procedure
        /// </summary>
        private void OnDisconnect()
        {
            if (m_Server == null)
            {
                return;
            }

            try
            {
                // delete all subscriptions
                stopMonitorItems();
                stopMonitorBlock();

                // Change GUI settings
                btnConnect.Text = "Connect";
                txtServerUrl.Enabled = true;
                txtServerUrl2.Enabled = true;
                txtServerUrl3.Enabled = true;

                // disable buttons
                btnMonitor.Enabled = false;
                btnMonitor2.Enabled = false;
                btnMonitor3.Enabled = false;
                btnRead.Enabled = false;
                btnRead2.Enabled = false;
                btnRead3.Enabled = false;
                btnWrite.Enabled = false;
                btnWrite2.Enabled = false;
                btnWrite3.Enabled = false;
                btnMonitorBlock.Enabled = false;
                btnWriteBlock1.Enabled = false;
                btnWriteBlock2.Enabled = false;

                // clear text fields
                txtRead1.Clear();
                txtRead2.Clear();
                txtRead3.Clear();
                txtRead4.Clear();
                txtRead5.Clear();
                txtRead6.Clear();
                txtWrite1.Clear();
                txtWrite2.Clear();
                txtWrite4.Clear();
                txtWrite3.Clear();
                txtWrite5.Clear();
                txtWrite6.Clear();

                // reset all text field colors
                txtRead1.BackColor = Color.White;
                txtRead2.BackColor = Color.White;
                txtRead3.BackColor = Color.White;
                txtRead4.BackColor = Color.White;
                txtRead5.BackColor = Color.White;
                txtRead6.BackColor = Color.White;
                txtWrite1.BackColor = Color.White;
                txtWrite2.BackColor = Color.White;
                txtWrite3.BackColor = Color.White;
                txtWrite4.BackColor = Color.White;
                txtWrite5.BackColor = Color.White;
                txtWrite6.BackColor = Color.White;
                txtMonitor1.BackColor = Color.White;
                txtMonitor2.BackColor = Color.White;
                txtMonitor3.BackColor = Color.White;
                txtMonitor4.BackColor = Color.White;
                txtMonitor5.BackColor = Color.White;
                txtMonitor6.BackColor = Color.White;
                txtMonitorBlock.BackColor = Color.White;
                txtItemIDBlockWrite.BackColor = Color.White;
                txtWriteBlockLength.BackColor = Color.White;

                // Disconnect
                m_Server.Disconnect();

                // Cleanup
                m_Server = null;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Disconnect failed");
            }
        }
        #endregion

        #region Construction
        public SimpleClientDA()
        {
            InitializeComponent();

            // set the sever we want to connet to
            txtServerUrl.Text = serverUrl;

            // update item ids
            radioBtn_CheckedChanged(null, null);
        }
        #endregion

        #region User Actions
        /// <summary>
        /// Handle action when connect button was clicked
        /// </summary>
        private void btnConnect_Click(object sender, EventArgs e)
        {
            if (m_Server == null)
            {
                OnConnect();
            }
            else if (m_Server.IsConnected)
            {
                OnDisconnect();
            }
            else
            {
                OnDisconnect();
            }
        }

        /// <summary>
        /// Handle action when read button was clicked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRead_Click(object sender, EventArgs e)
        {
            try
            {
                bool bResult;
                StringCollection itemIds = new StringCollection();
                object[] values;
                int[] pErrors;

                // ItemIds to read
                itemIds.Add(txtItemID1.Text);
                itemIds.Add(txtItemID2.Text);

                bResult = m_Server.Read(itemIds, out values, out pErrors);

                // everything went fine
                if (bResult)
                {
                    txtRead1.Text = values[0].ToString();
                    txtRead1.BackColor = Color.White;

                    txtRead2.Text = values[1].ToString();
                    txtRead2.BackColor = Color.White;
                }
                // we have to check the individual result codes
                else
                {
                    if (pErrors[0] == 0)
                    {
                        txtRead1.Text = values[0].ToString();
                        txtRead1.BackColor = Color.White;
                    }
                    else
                    {
                        string errorString;
                        m_Server.GetErrorString(pErrors[0], 0, out errorString);
                        txtRead1.Text = errorString;
                        txtRead1.BackColor = Color.Red;
                    }

                    if (pErrors[1] == 0)
                    {
                        txtRead2.Text = values[1].ToString();
                        txtRead2.BackColor = Color.White;
                    }
                    else
                    {
                        string errorString;
                        m_Server.GetErrorString(pErrors[1], 0, out errorString);
                        txtRead2.Text = errorString;
                        txtRead2.BackColor = Color.Red;
                    }
                }
            }
            catch (Exception exception)
            {
                txtRead1.BackColor = Color.Red;
                txtRead1.Text = exception.Message;
                txtRead2.BackColor = Color.Red;
                txtRead2.Text = exception.Message;
            }
        }

        private void btnRead2_Click_1(object sender, EventArgs e)
        {
            try
            {
                bool bResult;
                StringCollection itemIds = new StringCollection();
                object[] values;
                int[] pErrors;

                // ItemIds to read
                itemIds.Add(txtItemID3.Text);
                itemIds.Add(txtItemID4.Text);

                bResult = m_Server2.Read(itemIds, out values, out pErrors);

                // everything went fine
                if (bResult)
                {
                    txtRead3.Text = values[0].ToString();
                    txtRead3.BackColor = Color.White;

                    txtRead4.Text = values[1].ToString();
                    txtRead4.BackColor = Color.White;
                }
                // we have to check the individual result codes
                else
                {
                    if (pErrors[0] == 0)
                    {
                        txtRead3.Text = values[0].ToString();
                        txtRead3.BackColor = Color.White;
                    }
                    else
                    {
                        string errorString;
                        m_Server.GetErrorString(pErrors[0], 0, out errorString);
                        txtRead3.Text = errorString;
                        txtRead3.BackColor = Color.Red;
                    }

                    if (pErrors[1] == 0)
                    {
                        txtRead4.Text = values[1].ToString();
                        txtRead4.BackColor = Color.White;
                    }
                    else
                    {
                        string errorString;
                        m_Server.GetErrorString(pErrors[1], 0, out errorString);
                        txtRead4.Text = errorString;
                        txtRead4.BackColor = Color.Red;
                    }
                }
            }
            catch (Exception exception)
            {
                txtRead3.BackColor = Color.Red;
                txtRead3.Text = exception.Message;
                txtRead4.BackColor = Color.Red;
                txtRead4.Text = exception.Message;
            }
        }
        private void btnRead3_Click(object sender, EventArgs e)
        {
            try
            {
                bool bResult;
                StringCollection itemIds = new StringCollection();
                object[] values;
                int[] pErrors;

                // ItemIds to read
                itemIds.Add(txtItemID5.Text);
                itemIds.Add(txtItemID6.Text);

                bResult = m_Server3.Read(itemIds, out values, out pErrors);

                // everything went fine
                if (bResult)
                {
                    txtRead5.Text = values[0].ToString();
                    txtRead5.BackColor = Color.White;

                    txtRead6.Text = values[1].ToString();
                    txtRead6.BackColor = Color.White;
                }
                // we have to check the individual result codes
                else
                {
                    if (pErrors[0] == 0)
                    {
                        txtRead5.Text = values[0].ToString();
                        txtRead5.BackColor = Color.White;
                    }
                    else
                    {
                        string errorString;
                        m_Server.GetErrorString(pErrors[0], 0, out errorString);
                        txtRead5.Text = errorString;
                        txtRead5.BackColor = Color.Red;
                    }

                    if (pErrors[1] == 0)
                    {
                        txtRead6.Text = values[1].ToString();
                        txtRead6.BackColor = Color.White;
                    }
                    else
                    {
                        string errorString;
                        m_Server.GetErrorString(pErrors[1], 0, out errorString);
                        txtRead6.Text = errorString;
                        txtRead6.BackColor = Color.Red;
                    }
                }
            }
            catch (Exception exception)
            {
                txtRead5.BackColor = Color.Red;
                txtRead5.Text = exception.Message;
                txtRead6.BackColor = Color.Red;
                txtRead6.Text = exception.Message;
            }
        }

        /// <summary>
        /// Handle action when write button was clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWrite_Click(object sender, EventArgs e)
        {
            try
            {
                m_Server.Write(txtItemID1.Text, txtWrite1.Text);
                txtWrite1.BackColor = Color.White;
            }
            catch (Exception exception)
            {
                txtWrite1.BackColor = Color.Red;
                txtWrite1.Text = exception.Message;
            }

            try
            {
                m_Server.Write(txtItemID2.Text, txtWrite2.Text);
                txtWrite2.BackColor = Color.White;
            }
            catch (Exception exception)
            {
                txtWrite2.BackColor = Color.Red;
                txtWrite2.Text = exception.Message;
            }
        }

        /// <summary>
        /// Handle action when WriteBlock1 button was clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWriteBlock1_Click(object sender, EventArgs e)
        {
            int writeLength = 0;
            try
            {
                writeLength = (int)Convert.ChangeType(txtWriteBlockLength.Text, typeof(int));
            }
            catch (Exception)
            { }

            if (writeLength <= 0)
            {
                txtWriteBlockLength.BackColor = Color.Red;
                return;
            }

            byte[] rawValue = new byte[writeLength];
            byte currentValue = 0;

            for (int i = 0; i < rawValue.Length; i++)
            {
                rawValue[i] = currentValue;
                currentValue++;
            }

            try
            {
                m_Server.Write(txtItemIDBlockWrite.Text, rawValue);
                txtItemIDBlockWrite.BackColor = Color.White;
            }
            catch (Exception)
            {
                txtItemIDBlockWrite.BackColor = Color.Red;
            }
        }

        /// <summary>
        /// Handle action when WriteBlock2 button was clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWriteBlock2_Click(object sender, EventArgs e)
        {
            int writeLength = 0;
            try
            {
                writeLength = (int)Convert.ChangeType(txtWriteBlockLength.Text, typeof(int));
            }
            catch (Exception)
            { }

            if (writeLength <= 0)
            {
                txtWriteBlockLength.BackColor = Color.Red;
                return;
            }

            byte[] rawValue = new byte[writeLength];
            byte currentValue = 255;

            for (int i = 0; i < rawValue.Length; i++)
            {
                rawValue[i] = currentValue;
                currentValue--;
            }

            try
            {
                m_Server.Write(txtItemIDBlockWrite.Text, rawValue);
                txtItemIDBlockWrite.BackColor = Color.White;
            }
            catch (Exception)
            {
                txtItemIDBlockWrite.BackColor = Color.Red;
            }
        }

        /// <summary>
        /// Handle action when Monitor button was clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMonitor_Click(object sender, EventArgs e)
        {
            // Check if we have a subscription
            //  - No  -> Create a new subscription and create monitored items
            //  - Yes -> Delete Subcription
            if (m_Subscription == null)
            {
                startMonitorItems();
            }
            else
            {
                stopMonitorItems();
            }
        }

        /// <summary>
        /// Handle action when MonitorBlock button was clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMonitorBlock_Click(object sender, EventArgs e)
        {
            // Check if we have a subscription
            //  - No  -> Create a new subscription and create monitored items
            //  - Yes -> Delete Subcription
            if (m_SubscriptionBlock == null)
            {
                startMonitorBlock();
            }
            else
            {
                stopMonitorBlock();
            }
        }

        /// <summary>
        /// Enable or disable GUI elements depending on the selected device
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioBtn_CheckedChanged(object sender, EventArgs e)
        {
            // stop all monitoring
            stopMonitorItems();
            stopMonitorBlock();

            // enable block read / write
            if (radioBtn300.Checked)
            {
                grpBoxBlockRead.Enabled = true;
                grpBoxBlockWrite.Enabled = true;
                txtItemID1.Text = itemID1_300;
                txtItemID2.Text = itemID2_300;
                txtItemIDBlockRead.Text = itemIDBlockRead;
                txtItemIDBlockWrite.Text = itemIDBlockWrite;
            }
            // disable block read / write
            else
            {
                grpBoxBlockRead.Enabled = false;
                grpBoxBlockWrite.Enabled = false;
                txtItemID1.Text = itemID1_1200;
                txtItemID2.Text = itemID2_1200;
                txtItemIDBlockRead.Text = "";
                txtItemIDBlockWrite.Text = "";
            }
        }
        #endregion

        #region Event Handlers
        /// <summary>
        /// Show shutdown message
        /// When receiving a shutdown message just disconnect.
        /// </summary>
        public void ShutDownRequest(string reason)
        {
            OnDisconnect();
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
                                txtMonitor1.Text = "Error: 0x" + value.Error.ToString("X");
                                txtMonitor1.BackColor = Color.Red;
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                txtMonitor1.Text = value.Value.ToString();
                                txtMonitor1.BackColor = Color.White;
                            }
                            break;
                        case (int)ClientHandles.Item2:
                            // Print data change information for variable - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                txtMonitor2.Text = "Error: 0x" + value.Error.ToString("X");
                                txtMonitor2.BackColor = Color.Red;
                            }
                            else
                            {
                                // The node succeeded - print the value as string
                                txtMonitor2.Text = value.Value.ToString();
                                txtMonitor2.BackColor = Color.White;
                            }
                            break;
                        case (int)ClientHandles.ItemBlockRead:
                            // Print result for block - check first the result code
                            if (value.Error != 0)
                            {
                                // The node failed - print the symbolic name of the status code
                                txtMonitorBlock.Text = "Error: 0x" + value.Error.ToString("X");
                                txtMonitorBlock.BackColor = Color.Red;
                            }
                            else
                            {
                                if (value.Value.GetType() != typeof(byte[]))
                                {
                                    throw new Exception("Value change for block did not send byte array");
                                }

                                byte[] rawValue = (byte[])value.Value;
                                string stringValue = "";
                                int lineLength = 0;

                                for (int i = 0; i < rawValue.Length; i++)
                                {
                                    stringValue += string.Format("{0:X2} ", rawValue[i]);
                                    lineLength++;
                                    if (lineLength > 25)
                                    {
                                        stringValue += "\n";
                                        lineLength = 0;
                                    }
                                }

                                // The node succeeded - print the value as string
                                txtMonitorBlock.Text = stringValue;
                                txtMonitorBlock.BackColor = Color.White;
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

        #region Internal Helper Methods
        void startMonitorItems()
        {
            // Check if we have a subscription. If not - create a new subscription.
            if (m_Subscription == null)
            {
                try
                {
                    // Create subscription
                    m_Subscription = m_Server.CreateSubscription("Subscription1", OnDataChange);
                    btnMonitor.Text = "Stop";

                    // disable changing the itemID
                    txtItemID1.Enabled = false;
                    txtItemID2.Enabled = false;
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Create subscription failed:\n\n" + exception.Message);
                    return;
                }
            }

            // Add item 1
            try
            {
                m_Subscription.AddItem(
                    txtItemID1.Text,
                    (int)ClientHandles.Item1);

                txtMonitor1.BackColor = Color.White;
            }
            catch (Exception exception)
            {
                txtMonitor1.BackColor = Color.Red;
                txtMonitor1.Text = exception.Message;
            }

            // Add item 2
            try
            {
                m_Subscription.AddItem(
                    txtItemID2.Text,
                    (int)ClientHandles.Item2);

                txtMonitor2.BackColor = Color.White;
            }
            catch (Exception exception)
            {
                txtMonitor2.BackColor = Color.Red;
                txtMonitor2.Text = exception.Message;
            }
        }

        void stopMonitorItems()
        {
            if (m_Subscription != null)
            {
                try
                {
                    m_Server.DeleteSubscription(m_Subscription);
                    m_Subscription = null;

                    btnMonitor.Text = "Monitor";
                    txtMonitor1.Clear();
                    txtMonitor1.BackColor = Color.White;
                    txtMonitor2.Clear();
                    txtMonitor2.BackColor = Color.White;

                    // enable changing the itemID
                    txtItemID1.Enabled = true;
                    txtItemID2.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Stopping data monitoring failed:\n\n" + ex.Message);
                }
            }
        }

        void startMonitorBlock()
        {
            try
            {
                if (m_SubscriptionBlock == null)
                {
                    // Create subscription
                    m_SubscriptionBlock = m_Server.CreateSubscription("SubscriptionBlock", OnDataChange);
                    btnMonitorBlock.Text = "Stop";
                }

                // Add the item
                m_SubscriptionBlock.AddItem(
                    txtItemIDBlockRead.Text,
                    (int)ClientHandles.ItemBlockRead);

                txtMonitorBlock.BackColor = Color.White;

                // disable changing the itemID
                txtItemIDBlockRead.Enabled = false;
            }
            catch (Exception exception)
            {
                txtMonitorBlock.BackColor = Color.Red;
                txtMonitorBlock.Text = exception.Message;
            }
        }

        void stopMonitorBlock()
        {
            if (m_SubscriptionBlock != null)
            {
                try
                {
                    m_Server.DeleteSubscription(m_SubscriptionBlock);
                    m_SubscriptionBlock = null;

                    btnMonitorBlock.Text = "Monitor Block";
                    txtMonitorBlock.Clear();
                    txtMonitorBlock.BackColor = Color.White;

                    // enable changing the itemID
                    txtItemIDBlockRead.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Stopping data monitoring failed:\n\n" + ex.Message);
                }
            }
        }
        #endregion

        #region Private Members
        private Server m_Server = null;
        private Subscription m_Subscription = null;
        private Subscription m_SubscriptionBlock = null;
        private Server m_Server2 = null;
        private Subscription m_Subscription2 = null;
        private Subscription m_SubscriptionBlock2 = null;
        private Server m_Server3 = null;
        private Subscription m_Subscription3 = null;
        private Subscription m_SubscriptionBlock3 = null;

        const string serverUrl = "opcda://localhost/OPC.ZUMBACH";
        const string itemID1_300 = "Usys_200_Coex/Data/d1100";
        const string itemID2_300 = "Usys_200_Coex/Product/c1100";
        const string itemID1_1200 = "Dynamic/Analog Types/Int";
        const string itemID2_1200 = "Static/Simple Types/Double";
        const string itemIDBlockRead = "Static/ArrayTypes/Byte[]";
        const string itemIDBlockWrite = "Static/ArrayTypes/Byte[]";
        #endregion

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtServerUrl_TextChanged(object sender, EventArgs e)
        {

        }


    }
}