/*
Created by: John Walker
Last modified: 1/31/2018
Description: FORM1.CS is the main display form for AdminRDP
*/


using MSTSCLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO;
using AxMSTSCLib;
using System.Diagnostics;
using System.Xml;


namespace AdminRDP
{
    public partial class Form1 : Form
    {
        public AxMsTscAxNotSafeForScripting rdp;
        public AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();

        public Form1()
        {
            InitializeComponent();            
        }

        private void Load_XML()
        {
            try
            {
                //declare xml doc
                XmlDocument xmldoc = new XmlDocument();
                XmlNode xmlnode;

                //load xml file 
                using (FileStream fs = new FileStream(@"C:\Program Files\AdminRDP\SavedServers\ServerList.xml", FileMode.Open, FileAccess.Read))
                {
                    xmldoc.Load(fs);
                    xmlnode = xmldoc.ChildNodes[1];
                    //populate treeview from xml file
                    treeView1.Nodes.Clear();
                    treeView1.Nodes.Add(new TreeNode(xmldoc.DocumentElement.Name));
                    TreeNode tNode;
                    tNode = treeView1.Nodes[0];
                    AddNode(xmlnode, tNode);

                }
                treeView1.ExpandAll();
            }catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        //function to populate treeview control from XML file
        private void AddNode(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            XmlNode xNode;
            TreeNode tNode;
            XmlNodeList nodeList;
            int i = 0;
            if (inXmlNode.HasChildNodes)
            {
                nodeList = inXmlNode.ChildNodes;
                for (i = 0; i <= nodeList.Count - 1; i++)
                {
                    xNode = inXmlNode.ChildNodes[i];
                    inTreeNode.Nodes.Add(new TreeNode(xNode.Name));
                    tNode = inTreeNode.Nodes[i];
                    AddNode(xNode, tNode);
                }
            }
        }

        //run Powershell script to log out
        private static void StartPowershell(string ps, string username, string password, string domain, string computername)
        {
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "Powershell.exe";
            myProcess.StartInfo.Arguments = "-ExecutionPolicy Bypass -File " + ps + " -Username " + username + " -Password " + password + " -Domain " + domain + " -ComputerName " + computername;
            myProcess.StartInfo.UseShellExecute = true;
            myProcess.StartInfo.CreateNoWindow = false;
            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            myProcess.StartInfo.Verb = "runas";
            myProcess.Start();
        }

        //onDisconnected event handler
        private void rdp_OnDisconnected(object sender, AxMSTSCLib.IMsTscAxEvents_OnDisconnectedEvent e)
        {
            TabControl.TabPageCollection pages = RDPTabControl.TabPages;
            string disconnectedReason = null;
            if(pages != null)
            {
                foreach(TabPage page in pages)
                {
                    foreach (AxMsTscAxNotSafeForScripting connection in page.Controls)
                    {
                        if (connection is AxMsTscAxNotSafeForScripting)
                        {
                            if (connection.Connected == 0)
                            {
                                switch (e.discReason)
                                {
                                    case 0: disconnectedReason = "No information is available."; break;
                                    case 1: RDPTabControl.TabPages.Remove(page); break;
                                    case 2: RDPTabControl.TabPages.Remove(page) ; break;
                                    case 3: disconnectedReason = "Remote disconnection by server."; break;
                                    case 260: disconnectedReason = "Remote Desktop can't find the computer. This might mean that does not belong to the specified network.  Verify the computer name and domain that you are trying to connect to."; break;
                                    case 262: disconnectedReason = "This computer can't connect to the remote computer.  Your computer does not have enough virtual memory available. Close your other programs, and then try connecting again. If the problem continues, contact your network administrator or technical support."; break;
                                    case 264: disconnectedReason = "This computer can't connect to the remote computer.  The two computers couldn't connect in the amount of time allotted. Try connecting again. If the problem continues, contact your network administrator or technical support."; break;
                                    case 266: disconnectedReason = "The smart card service is not running. Please start the smart card service and try again."; break;
                                    case 516: disconnectedReason = "Remote Desktop can’t connect to the remote computer for one of these reasons:  1) Remote access to the server is not enabled 2) The remote computer is turned off 3) The remote computer is not available on the network  Make sure the remote computer is turned on and connected to the network, and that remote access is enabled."; break;
                                    case 522: disconnectedReason = "A smart card reader was not detected. Please attach a smart card reader and try again."; break;
                                    case 772: disconnectedReason = "This computer can't connect to the remote computer.  The connection was lost due to a network error. Try connecting again. If the problem continues, contact your network administrator or technical support."; break;
                                    case 778: disconnectedReason = "There is no card inserted in the smart card reader. Please insert your smart card and try again."; break;
                                    case 1030: disconnectedReason = "Because of a security error, the client could not connect to the remote computer. Verify that you are logged on to the network, and then try connecting again."; break;
                                    case 1032: disconnectedReason = "The specified computer name contains invalid characters. Please verify the name and try again."; break;
                                    case 1034: disconnectedReason = "An error has occurred in the smart card subsystem. Please contact your helpdesk about this error."; break;
                                    case 1796: disconnectedReason = "This computer can't connect to the remote computer.  Try connecting again. If the problem continues, contact the owner of the remote computer or your network administrator."; break;
                                    case 1800: disconnectedReason = "Your computer could not connect to another console session on the remote computer because you already have a console session in progress."; break;
                                    case 2056: disconnectedReason = "The remote computer disconnected the session because of an error in the licensing protocol. Please try connecting to the remote computer again or contact your server administrator."; break;
                                    case 2308: disconnectedReason = "Your Remote Desktop Services session has ended.  The connection to the remote computer was lost, possibly due to network connectivity problems. Try connecting to the remote computer again. If the problem continues, contact your network administrator or technical support."; break;
                                    case 2311: disconnectedReason = "The connection has been terminated because an unexpected server authentication certificate was received from the remote computer. Try connecting again. If the problem continues, contact the owner of the remote computer or your network administrator."; break;
                                    case 2312: disconnectedReason = "A licensing error occurred while the client was attempting to connect (Licensing timed out). Please try connecting to the remote computer again."; break;
                                    case 2567: disconnectedReason = "The specified user name does not exist. Verify the user name and try logging in again. If the problem continues, contact your system administrator or technical support."; break;
                                    case 2820: disconnectedReason = "This computer can’t connect to the remote computer.  An error occurred that prevented the connection. Try connecting again. If the problem continues, contact the owner of the remote computer or your network administrator."; break;
                                    case 2822: disconnectedReason = "Because of an error in data encryption, this session will end. Please try connecting to the remote computer again."; break;
                                    case 2823: disconnectedReason = "The user account is currently disabled and cannot be used. For assistance, contact your system administrator or technical support."; break;
                                    case 2825: disconnectedReason = "The remote computer requires Network Level Authentication, which your computer does not support. For assistance, contact your system administrator or technical support."; break;
                                    case 3079: disconnectedReason = "A user account restriction (for example, a time-of-day restriction) is preventing you from logging on. For assistance, contact your system administrator or technical support."; break;
                                    case 3080: disconnectedReason = "The remote session was disconnected because of a decompression failure at the client side. Please try connecting to the remote computer again."; break;
                                    case 3335: disconnectedReason = "As a security precaution, the user account has been locked because there were too many logon attempts or password change attempts. Wait a while before trying again, or contact your system administrator or technical support."; break;
                                    case 3337: disconnectedReason = "The security policy of your computer requires you to type a password on the Windows Security dialog box. However, the remote computer you want to connect to cannot recognize credentials supplied using the Windows Security dialog box. For assistance, contact your system administrator or technical support."; break;
                                    case 3590: disconnectedReason = "The client can't connect because it doesn't support FIPS encryption level.  Please lower the server side required security level Policy, or contact your network administrator for assistance"; break;
                                    case 3591: disconnectedReason = "This user account has expired. For assistance, contact your system administrator or technical support."; break;
                                    case 3592: disconnectedReason = "Failed to reconnect to your remote session. Please try to connect again."; break;
                                    case 3593: disconnectedReason = "The remote PC doesn't support Restricted Administration mode."; break;
                                    case 3847: disconnectedReason = "This user account's password has expired. The password must change in order to logon. Please update the password or contact your system administrator or technical support."; break;
                                    case 3848: disconnectedReason = "A connection will not be made because credentials may not be sent to the remote computer. For assistance, contact your system administrator."; break;
                                    case 4103: disconnectedReason = "The system administrator has restricted the times during which you may log in. Try logging in later. If the problem continues, contact your system administrator or technical support."; break;
                                    case 4104: disconnectedReason = "The remote session was disconnected because your computer is running low on video resources.  Close your other programs, and then try connecting again. If the problem continues, contact your network administrator or technical support."; break;
                                    case 4359: disconnectedReason = "The system administrator has limited the computers you can log on with. Try logging on at a different computer. If the problem continues, contact your system administrator or technical support."; break;
                                    case 4615: disconnectedReason = "You must change your password before logging on the first time. Please update your password or contact your system administrator or technical support."; break;
                                    case 4871: disconnectedReason = "The system administrator has restricted the types of logon (network or interactive) that you may use. For assistance, contact your system administrator or technical support."; break;
                                    case 5127: disconnectedReason = "The Kerberos sub-protocol User2User is required. For assistance, contact your system administrator or technical support."; break;
                                    case 6919: disconnectedReason = "Remote Desktop cannot connect to the remote computer because the authentication certificate received from the remote computer is expired or invalid.  In some cases, this error might also be caused by a large time discrepancy between the client and server computers."; break;
                                    case 7431: disconnectedReason = "Remote Desktop cannot verify the identity of the remote computer because there is a time or date difference between your computer and the remote computer. Make sure your computer’s clock is set to the correct time, and then try connecting again. If the problem occurs again, contact your network administrator or the owner of the remote computer."; break;
                                    case 8711: disconnectedReason = "Your computer can't connect to the remote computer because your smart card is locked out. Contact your network administrator about unlocking your smart card or resetting your PIN."; break;
                                    case 9479: disconnectedReason = "Could not auto-reconnect to your applications,please re-launch your applications"; break;
                                    case 9732: disconnectedReason = "Client and server versions do not match. Please upgrade your client software and then try connecting again."; break;
                                    case 33554433: disconnectedReason = "Failed to reconnect to the remote program. Please restart the remote program."; break;
                                    case 33554434: disconnectedReason = "The remote computer does not support RemoteApp. For assistance, contact your system administrator."; break;
                                    case 50331649: disconnectedReason = "Your computer can't connect to the remote computer because the user name or password is not valid. Type a valid user name and password."; break;
                                    case 50331650: disconnectedReason = "Your computer can't connect to the remote computer because it can’t verify the certificate revocation list. Contact your network administrator for assistance."; break;
                                    case 50331651: disconnectedReason = "Your computer can't connect to the remote computer due to one of the following reasons:  1) The requested Remote Desktop Gateway server address and the server SSL certificate subject name do not match. 2) The certificate is expired or revoked. 3) The certificate root authority does not trust the certificate.  Contact your network administrator for assistance."; break;
                                    case 50331652: disconnectedReason = "Your computer can't connect to the remote computer because the SSL certificate was revoked by the certification authority. Contact your network administrator for assistance."; break;
                                    case 50331653: disconnectedReason = "This computer can't verify the identity of the RD Gateway. It's not safe to connect to servers that can't be identified. Contact your network administrator for assistance."; break;
                                    case 50331654: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server address requested and the certificate subject name do not match. Contact your network administrator for assistance."; break;
                                    case 50331655: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server's certificate has expired or has been revoked. Contact your network administrator for assistance."; break;
                                    case 50331656: disconnectedReason = "Your computer can't connect to the remote computer because an error occurred on the remote computer that you want to connect to. Contact your network administrator for assistance."; break;
                                    case 50331657: disconnectedReason = "An error occurred while sending data to the Remote Desktop Gateway server. The server is temporarily unavailable or a network connection is down. Try again later, or contact your network administrator for assistance."; break;
                                    case 50331658: disconnectedReason = "An error occurred while receiving data from the Remote Desktop Gateway server. Either the server is temporarily unavailable or a network connection is down. Try again later, or contact your network administrator for assistance."; break;
                                    case 50331659: disconnectedReason = "Your computer can't connect to the remote computer because an alternate logon method is required. Contact your network administrator for assistance."; break;
                                    case 50331660: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server address is unreachable or incorrect. Type a valid Remote Desktop Gateway server address."; break;
                                    case 50331661: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server is temporarily unavailable. Try reconnecting later or contact your network administrator for assistance."; break;
                                    case 50331662: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Services client component is missing or is an incorrect version. Verify that setup was completed successfully, and then try reconnecting later."; break;
                                    case 50331663: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server is running low on server resources and is temporarily unavailable. Try reconnecting later or contact your network administrator for assistance."; break;
                                    case 50331664: disconnectedReason = "Your computer can't connect to the remote computer because an incorrect version of rpcrt4.dll has been detected. Verify that all components for Remote Desktop Gateway client were installed correctly."; break;
                                    case 50331665: disconnectedReason = "Your computer can't connect to the remote computer because no smart card service is installed. Install a smart card service and then try again, or contact your network administrator for assistance."; break;
                                    case 50331666: disconnectedReason = "Your computer can't stay connected to the remote computer because the smart card has been removed. Try again using a valid smart card, or contact your network administrator for assistance."; break;
                                    case 50331667: disconnectedReason = "Your computer can't connect to the remote computer because no smart card is available. Try again using a smart card."; break;
                                    case 50331668: disconnectedReason = "Your computer can't stay connected to the remote computer because the smart card has been removed. Reinsert the smart card and then try again."; break;
                                    case 50331669: disconnectedReason = "Your computer can't connect to the remote computer because the user name or password is not valid. Please type a valid user name and password."; break;
                                    case 50331671: disconnectedReason = "Your computer can't connect to the remote computer because a security package error occurred in the transport layer. Retry the connection or contact your network administrator for assistance."; break;
                                    case 50331672: disconnectedReason = "The Remote Desktop Gateway server has ended the connection. Try reconnecting later or contact your network administrator for assistance."; break;
                                    case 50331673: disconnectedReason = "The Remote Desktop Gateway server administrator has ended the connection. Try reconnecting later or contact your network administrator for assistance."; break;
                                    case 50331674: disconnectedReason = "Your computer can't connect to the remote computer due to one of the following reasons:   1) Your credentials (the combination of user name, domain, and password) were incorrect. 2) Your smart card was not recognized."; break;
                                    case 50331675: disconnectedReason = "Remote Desktop can’t connect to the remote computer for one of these reasons:  1) Your user account is not listed in the RD Gateway’s permission list 2) You might have specified the remote computer in NetBIOS format (for example, computer1), but the RD Gateway is expecting an FQDN or IP address format (for example, computer1.fabrikam.com or 157.60.0.1).  Contact your network administrator for assistance."; break;
                                    case 50331676: disconnectedReason = "Remote Desktop can’t connect to the remote computer for one of these reasons:  1) Your user account is not authorized to access the RD Gateway 2) Your computer is not authorized to access the RD Gateway 3) You are using an incompatible authentication method (for example, the RD Gateway might be expecting a smart card but you provided a password)  Contact your network administrator for assistance."; break;
                                    case 50331679: disconnectedReason = "Your computer can't connect to the remote computer because your network administrator has restricted access to this RD Gateway server. Contact your network administrator for assistance."; break;
                                    case 50331680: disconnectedReason = "Your computer can't connect to the remote computer because the web proxy server requires authentication. To allow unauthenticated traffic to an RD Gateway server through your web proxy server, contact your network administrator."; break;
                                    case 50331681: disconnectedReason = "Your computer can't connect to the remote computer because your password has expired or you must change the password. Please change the password or contact your network administrator or technical support for assistance."; break;
                                    case 50331682: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server reached its maximum allowed connections. Try reconnecting later or contact your network administrator for assistance."; break;
                                    case 50331683: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server does not support the request. Contact your network administrator for assistance."; break;
                                    case 50331684: disconnectedReason = "Your computer can't connect to the remote computer because the client does not support one of the Remote Desktop Gateway's capabilities. Contact your network administrator for assistance."; break;
                                    case 50331685: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server and this computer are incompatible. Contact your network administrator for assistance."; break;
                                    case 50331686: disconnectedReason = "Your computer can't connect to the remote computer because the credentials used are not valid. Insert a valid smart card and type a PIN or password, and then try connecting again."; break;
                                    case 50331687: disconnectedReason = "Your computer can't connect to the remote computer because your computer or device did not pass the Network Access Protection requirements set by your network administrator. Contact your network administrator for assistance."; break;
                                    case 50331688: disconnectedReason = "Your computer can’t connect to the remote computer because no certificate was configured to use at the Remote Desktop Gateway server. Contact your network administrator for assistance."; break;
                                    case 50331689: disconnectedReason = "Your computer can't connect to the remote computer because the RD Gateway server that you are trying to connect to is not allowed by your computer administrator. If you are the administrator, add this Remote Desktop Gateway server name to the trusted Remote Desktop Gateway server list on your computer and then try connecting again."; break;
                                    case 50331690: disconnectedReason = "Your computer can't connect to the remote computer because your computer or device did not meet the Network Access Protection requirements set by your network administrator, for one of the following reasons:  1) The Remote Desktop Gateway server name and the server's public key certificate subject name do not match. 2) The certificate has expired or has been revoked. 3) The certificate root authority does not trust the certificate. 4) The certificate key ext  ension does not support encryption. 5) Your computer cannot verify the certificate revocation list.  Contact your network administrator for assistance."; break;
                                    case 50331691: disconnectedReason = "Your computer can't connect to the remote computer because a user name and password are required to authenticate to the Remote Desktop Gateway server instead of smart card credentials."; break;
                                    case 50331692: disconnectedReason = "Your computer can't connect to the remote computer because smart card credentials are required to authenticate to the Remote Desktop Gateway server instead of a user name and password."; break;
                                    case 50331693: disconnectedReason = "Your computer can't connect to the remote computer because no smart card reader is detected. Connect a smart card reader and then try again, or contact your network administrator for assistance."; break;
                                    case 50331695: disconnectedReason = "Your computer can't connect to the remote computer because authentication to the firewall failed due to missing firewall credentials. To resolve the issue, go to the firewall website that your network administrator recommends, and then try the connection again, or contact your network administrator for assistance."; break;
                                    case 50331696: disconnectedReason = "Your computer can't connect to the remote computer because authentication to the firewall failed due to invalid firewall credentials. To resolve the issue, go to the firewall website that your network administrator recommends, and then try the connection again, or contact your network administrator for assistance."; break;
                                    case 50331698: disconnectedReason = "Your Remote Desktop Services session ended because the remote computer didn't receive any input from you."; break;
                                    case 50331699: disconnectedReason = "The connection has been disconnected because the session timeout limit was reached."; break;
                                    case 50331700: disconnectedReason = "Your computer can't connect to the remote computer because an invalid cookie was sent to the Remote Desktop Gateway server. Contact your network administrator for assistance."; break;
                                    case 50331701: disconnectedReason = "Your computer can't connect to the remote computer because the cookie was rejected by the Remote Desktop Gateway server. Contact your network administrator for assistance."; break;
                                    case 50331703: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway server is expecting an authentication method different from the one attempted. Contact your network administrator for assistance."; break;
                                    case 50331704: disconnectedReason = "The RD Gateway connection ended because periodic user authentication failed. Try reconnecting with a correct user name and password. If the reconnection fails, contact your network administrator for further assistance."; break;
                                    case 50331705: disconnectedReason = "The RD Gateway connection ended because periodic user authorization failed. Try reconnecting with a correct user name and password. If the reconnection fails, contact your network administrator for further assistance."; break;
                                    case 50331707: disconnectedReason = "Your computer can't connect to the remote computer because the Remote Desktop Gateway and the remote computer are unable to exchange policies. This could happen due to one of the following reasons:     1. The remote computer is not capable of exchanging policies with the Remote Desktop Gateway.     2. The remote computer's configuration does not permit a new connection.     3. The connection between the Remote Desktop Gateway and the remote computer ended.    Contact your network administrator for assistance."; break;
                                    case 50331708: disconnectedReason = "Your computer can't connect to the remote computer, possibly because the smart card is not valid, the smart card certificate was not found in the certificate store, or the Certificate Propagation service is not running. Contact your network administrator for assistance."; break;
                                    case 50331709: disconnectedReason = "To use this program or computer, first log on to the website"; break;
                                    case 50331710: disconnectedReason = "To use this program or computer, you must first log on to an authentication website. Contact your network administrator for assistance."; break;
                                    case 50331711: disconnectedReason = "Your session has ended. To continue using the program or computer, first log on to the website."; break;
                                    case 50331712: disconnectedReason = "Your session has ended. To continue using the program or computer, you must first log on to an authentication website. Contact your network administrator for assistance."; break;
                                    case 50331713: disconnectedReason = "The RD Gateway connection ended because periodic user authorization failed. Your computer or device didn't pass the Network Access Protection (NAP) requirements set by your network administrator. Contact your network administrator for assistance."; break;
                                    case 50331714: disconnectedReason = "Your computer can't connect to the remote computer because the size of the cookie exceeded the supported size. Contact your network administrator for assistance."; break;
                                    case 50331716: disconnectedReason = "Your computer can't connect to the remote computer using the specified forward proxy configuration. Contact your network administrator for assistance."; break;
                                    case 50331717: disconnectedReason = "This computer cannot connect to the remote resource because you do not have permission to this resource. Contact your network administrator for assistance."; break;
                                    case 50331718: disconnectedReason = "There are currently no resources available to connect to. Retry the connection or contact your network administrator."; break;
                                    case 50331719: disconnectedReason = "An error occurred while Remote Desktop Connection was accessing this resource. Retry the connection or contact your system administrator."; break;
                                    case 50331721: disconnectedReason = "Your Remote Desktop Client needs to be updated to the newest version. Contact your system administrator for help installing the update, and then try again."; break;
                                    case 50331722: disconnectedReason = "Your network configuration doesn’t allow the necessary HTTPS ports. Contact your network administrator for help allowing those ports or disabling the web proxy, and then try connecting again."; break;
                                    case 50331723: disconnectedReason = "We’re setting up more resources, and it might take a few minutes. Please try again later."; break;
                                    case 50331724: disconnectedReason = "The user name you entered does not match the user name used to subscribe to your applications. If you wish to sign in as a different user please choose Sign Out from the Home menu."; break;
                                    case 50331725: disconnectedReason = "Looks like there are too many users trying out the Azure RemoteApp service at the moment. Please wait a few minutes and then try again."; break;
                                    case 50331726: disconnectedReason = "Maximum user limit has been reached. Please contact your administrator for further assistance."; break;
                                    case 50331727: disconnectedReason = "Your trial period for Azure RemoteApp has expired. Ask your admin or tech support for help."; break;
                                    case 50331728: disconnectedReason = "You no longer have access to Azure RemoteApp. Ask your admin or tech support for help."; break;
                                    default: disconnectedReason = string.Format("Unrecognized error: code {0}", e.discReason); break;
                                }

                                if(disconnectedReason != null)
                                {
                                    connection.DisconnectedText = disconnectedReason;
                                }
                            }
                        }
                    }

                }
                
            }

        }


        private void Load_ComboBox()
        {
            //clears combo box for refresh every time list is clicked
            comboBox1.Items.Clear();
            comboBox1.Refresh();

            //grab XML info from SavedCredentials folder
            string[] strXMLExists = Directory.GetFiles("C:\\Program Files\\AdminRDP\\SavedCredentials\\", "*.xml", SearchOption.AllDirectories);

            //populdate dropbox data from xml if there are xml files present
            if (strXMLExists != null)
            {
                XmlSerializer sr = new XmlSerializer(typeof(Credential_Manager_Save));

                //loop through directory and grab info from saved credential xml files
                string[] dir = Directory.GetFiles("C:\\Program Files\\AdminRDP\\SavedCredentials\\", "*.xml");

                foreach (string d in dir)
                {
                    using (FileStream read = new FileStream(d, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        Credential_Manager_Save credman = (Credential_Manager_Save)sr.Deserialize(read);

                        //add items to combo box
                        comboBox1.Items.Add(credman.txtUsername + "_" + credman.txtDomain);
                        read.Close();
                    }

                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //maximize form on startup
            this.WindowState = FormWindowState.Maximized;
            System.Drawing.Size size = new System.Drawing.Size();

            //minimize panel on startup
            size.Height = panel1.Height;
            size.Width = 20;
            panel1.Size = new Size(size.Width, size.Height);

            //check for default credentials
            Load_ComboBox();
            string strFilePath = @"C:\Program Files\AdminRDP\SavedCredentials\Default.txt";

            //check to see if text file exists
            //if it does exist then open the text file to find default credentials
            try
            {

                if (File.Exists(strFilePath))
                {
                    //import text file to text box
                    string readLine = null;

                    //open and read file
                    using (System.IO.StreamReader reader = new System.IO.StreamReader(strFilePath))
                    {
                        //while readline is not equal to null read file entries
                        while ((readLine = reader.ReadLine()) != null)
                        {
                            //loop through  combobox until it finds an entry that matches
                            //the readline variable then set the selected index
                            for (int i = 0; i < comboBox1.Items.Count; i++)
                            {
                                //set variable for combobox text
                                string strUserName = comboBox1.GetItemText(comboBox1.Items[i]);

                                //it it finds a match then set the selected index
                                if (strUserName == readLine)
                                {
                                    comboBox1.SelectedIndex = i;
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
        }


        private void btnSubmit_Click(object sender, EventArgs e)
        {
            //form validation
            if (txtServer.Text == "")
            {
                MessageBox.Show("Please enter a server name!");
                return;
            }

            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Please select which credentials you would like to use!");
                return;
            }

            //add entry to auto complete
            autoComplete.Add(txtServer.Text);

            txtServer.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtServer.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //auto.Add(textBox1.Text);
            txtServer.AutoCompleteCustomSource = autoComplete;

            //add MSTSC control to new tab            
            string strCurrentDir = Directory.GetCurrentDirectory();
            XmlSerializer sr = new XmlSerializer(typeof(Credential_Manager_Save));
            using (FileStream read = new FileStream("C:\\Program Files\\AdminRDP\\SavedCredentials\\" + comboBox1.SelectedItem.ToString() + ".xml", FileMode.Open, FileAccess.Read, FileShare.Read))
            {
               // try
              //  {
                    //validate server entries
                    if (ValidateString.isValidServerName (txtServer.Text) == true)
                    {
                        //create AxMsTscAxNotSafeForScripting control
                        rdp = new AxMsTscAxNotSafeForScripting();
                        Credential_Manager_Save credman = (Credential_Manager_Save)sr.Deserialize(read);
                        string title = txtServer.Text; //create tab with server title
                        TabPage myTabPage = new TabPage(title); //creat new tab page
                        RDPTabControl.TabPages.Add(myTabPage); //add tab page to tab control
                        myTabPage.Controls.Add(rdp); //add control to tab page
                        RDPTabControl.SelectedTab = myTabPage; //make selected tab the current tab
                        rdp.Name = txtServer.Text;
                        rdp.Server =txtServer.Text;
                        rdp.UserName = credman.txtUsername.ToString();
                        rdp.Domain = credman.txtDomain.ToString();
                        rdp.ConnectingText = "Connecting...";
                        rdp.DisconnectedText = "Disconnected or could not connect to host.";
                        rdp.AdvancedSettings.BitmapPeristence = 0;                                                
                        IMsTscNonScriptable secured = (MSTSCLib.IMsTscNonScriptable)rdp.GetOcx();
                        IMsRdpClientNonScriptable4 credsupport = (MSTSCLib.IMsRdpClientNonScriptable4)rdp.GetOcx();
                        credsupport.EnableCredSspSupport = true; //enable for NLA
                        secured.ClearTextPassword = Encryption.DecryptString(credman.txtPassword);                        
                        read.Close();
                        rdp.Dock = DockStyle.Fill;
                        rdp.Connect();

                        //create event handler
                        ((AxMSTSCLib.AxMsTscAxNotSafeForScripting)this.rdp).OnDisconnected += new AxMSTSCLib.IMsTscAxEvents_OnDisconnectedEventHandler(this.rdp_OnDisconnected);
                    }
                    else
                   {
                        MessageBox.Show("Entries cannot contain any special characters or spaces!");
                   }

              //  }
               // catch (Exception error)
             //   {
               //     MessageBox.Show(error.Message);
               // }
            }

        }

        //disconnect
        private void button1_Click(object sender, EventArgs e)
        {
            TabPage myTabPage = new TabPage();
            myTabPage = RDPTabControl.SelectedTab;

            if (myTabPage != null)
            {
                foreach (AxMsTscAxNotSafeForScripting connection in myTabPage.Controls)
                {
                    if (connection is AxMsTscAxNotSafeForScripting)
                    {
                        if (connection.Connected == 1)
                        {
                            connection.Disconnect();
                        }
                    }
                }

                RDPTabControl.TabPages.Remove(myTabPage);
            }
        }

        //handle reconnect event
        private void btnReconnect_Click(object sender, EventArgs e)
        {
            TabPage myTabPage = new TabPage();
            myTabPage = RDPTabControl.SelectedTab;

            if (myTabPage != null)
            {
                foreach (AxMsTscAxNotSafeForScripting connection in myTabPage.Controls)
                {
                    if (connection is AxMsTscAxNotSafeForScripting)
                    {
                        if (connection.Connected == 0)
                        {
                            connection.Dock = DockStyle.Fill;
                            connection.Refresh();
                            connection.Connect();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Cannot reconnect on an invalid server name!");
            }
        }

        //log user off
        private void btnLogoff_Click(object sender, EventArgs e)
        {
            TabPage myTabPage = new TabPage();
            myTabPage = RDPTabControl.SelectedTab;
            string strPassword = null;
            string strUsername = null;
            string strDomain = null;

            if (myTabPage != null)
            {
                foreach (AxMsTscAxNotSafeForScripting connection in myTabPage.Controls)
                {
                    if (connection is AxMsTscAxNotSafeForScripting)
                    {
                        if (connection.Connected == 1)
                        {
                            XmlSerializer sr = new XmlSerializer(typeof(Credential_Manager_Save));
                            FileStream read = new FileStream("C:\\Program Files\\AdminRDP\\SavedCredentials\\" + connection.UserName.ToString() + "_" + connection.Domain.ToString() + ".xml", FileMode.Open, FileAccess.Read, FileShare.Read);
                            Credential_Manager_Save credman = (Credential_Manager_Save)sr.Deserialize(read);
                            strPassword = Encryption.DecryptString(credman.txtPassword);
                            strUsername = credman.txtUsername;
                            strDomain = credman.txtDomain;
                            read.Close();
                            try
                            {
                                MessageBox.Show("Log out has been initiated!");
                                StartPowershell(@"C:\Program Files\AdminRDP\Scripts\Invoke_LogOff.ps1", strUsername, strPassword, strDomain, connection.Server.ToString());
                            }
                            catch(Exception error)
                            {
                                MessageBox.Show(error.Message);
                            }
                                                  
                        }
                    }
                }
            }
        }

        //copy password to clipbaord
        private void copyPassword_Click(object sender, EventArgs e)
        {
            if(comboBox1.SelectedItem != null)
            {
                 string strPassword = null;
                 string strUsername = null;
                 string strDomain = null;
                 XmlSerializer sr = new XmlSerializer(typeof(Credential_Manager_Save));
                 FileStream read = new FileStream("C:\\Program Files\\AdminRDP\\SavedCredentials\\" + comboBox1.SelectedItem.ToString() + ".xml", FileMode.Open, FileAccess.Read, FileShare.Read);
                 Credential_Manager_Save credman = (Credential_Manager_Save)sr.Deserialize(read);
                 strPassword = Encryption.DecryptString(credman.txtPassword);
                 strUsername = credman.txtUsername;
                 strDomain = credman.txtDomain;
                 read.Close();
                 Clipboard.SetText(strPassword);
            }
        }

        private void credentialManagerToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();
            frm2.Show();
        }

        private void serverListManagerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.Show();
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            Load_ComboBox();
        }

        //when double clicking on node connect to server
        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            if(treeView1.SelectedNode != null)
            {
                txtServer.Text = treeView1.SelectedNode.Text;
                btnSubmit_Click(sender, e);
            }
        }

        //focus on current selected tab for password paste
        private void button2_Click(object sender, EventArgs e)
        {
            TabPage myTabPage = new TabPage();
            myTabPage = RDPTabControl.SelectedTab;
            RDPTabControl.SelectedTab = myTabPage;

            if (myTabPage != null)
            {
                foreach (AxMsTscAxNotSafeForScripting connection in myTabPage.Controls)
                {
                    if (connection is AxMsTscAxNotSafeForScripting)
                    {
                        if(connection.Connected == 1)
                        {
                            connection.Disconnect();
                            txtServer.Text = connection.Name;
                            btnSubmit_Click(sender, e);
                        }

                    }
                }
            }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(treeView1.SelectedNode.Text);
        }

        //display or hide panel and load server list
        private void button3_Click(object sender, EventArgs e)
        {
            if(button3.Text == ">")
            {
                button3.Text = "<";
                System.Drawing.Size size = new System.Drawing.Size();
                size.Height = panel1.Height;
                size.Width = 255;
                panel1.Size = new Size(size.Width, size.Height);
                Load_XML();
            }
            else
            {
                button3.Text = ">";
                System.Drawing.Size size = new System.Drawing.Size();
                size.Height = panel1.Height;
                size.Width = 20;
                panel1.Size = new Size(size.Width, size.Height);
                treeView1.Nodes.Clear();
            }
        }

        //when tab changes update server form name
        private void RDPTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            TabPage myTabPage = new TabPage();
            myTabPage = RDPTabControl.SelectedTab;

            if (myTabPage != null)
            {
                foreach (AxMsTscAxNotSafeForScripting connection in myTabPage.Controls)
                {
                    if (connection is AxMsTscAxNotSafeForScripting)
                    {
                        if (connection.Connected == 1)
                        {
                            //chang server name in text field
                            txtServer.Text = connection.Server.ToString();

                            for (int i = 0; i < comboBox1.Items.Count; i++)
                            {
                                if (comboBox1.GetItemText(comboBox1.Items[i]) == (connection.UserName.ToString() + "_" + connection.Domain.ToString()).ToString())
                                {
                                    comboBox1.SelectedIndex = i;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        //Copy tab name
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(txtServer.Text);
        }

        //this function launches the powershell session with selected credentials        
        private static void StartPSRemoteSession(string username, string password, string domain, string computername)
        {
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "Powershell.exe";
            myProcess.StartInfo.Arguments = "-ExecutionPolicy Bypass -NoExit -File \"C:\\Program Files\\AdminRDP\\Scripts\\start_ps_remote_session.ps1\" -Username " + username + " -Password " + password + " -Domain " + domain + " -ComputerName " + computername;
            myProcess.StartInfo.UseShellExecute = true;
            myProcess.StartInfo.Verb = "runas";
            myProcess.Start();
        }

        //call  StartPSRemoteSession
        private void button4_Click(object sender, EventArgs e)
        {
            //form validation
            if (txtServer.Text == "")
            {
                MessageBox.Show("Please enter a server name!");
                return;
            }

            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Please select which credentials you would like to use!");
                return;
            }

            //grab credentials for connection ps remote session         
            string strCurrentDir = Directory.GetCurrentDirectory();
            XmlSerializer sr = new XmlSerializer(typeof(Credential_Manager_Save));
            using (FileStream read = new FileStream("C:\\Program Files\\AdminRDP\\SavedCredentials\\" + comboBox1.SelectedItem.ToString() + ".xml", FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                try
                {
                    //validate server entries
                    if (ValidateString.isValidServerName(txtServer.Text) == true)
                    {
                        //create credential manager object
                        Credential_Manager_Save credman = (Credential_Manager_Save)sr.Deserialize(read);

                        //close reader object
                        read.Close();

                        //start powershell remote session
                        StartPSRemoteSession(credman.txtUsername.ToString(), Encryption.DecryptString(credman.txtPassword).ToString(), credman.txtDomain.ToString(), txtServer.Text);
                    }
                    else
                    {
                        MessageBox.Show("Entries cannot contain any special characters or spaces!");
                    }

                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            txtServer.Clear();
        }
    }
}

