using System;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualBasic;
using mRemoteNC.App;
using mRemoteNC.Connection;
using My;
using mRemoteNC.Security;

namespace mRemoteNC.Config.Connections.Strategies
{
    public class SQLConfigConnectionStrategy : IConfigConnectionStrategy
    {

        public bool Export { get; set; }
        public string ConnectionFileName { get; set; }
        public TreeNode RootTreeNode { get; set; }
        public Container.List ContainerList { get; set; }
        public Security.Save SaveSecurity { get; set; }
        public Connection.List ConnectionList { get; set; }

        public string SQLHost { get; set; }

        public string SQLDatabaseName { get; set; }

        public string SQLUsername { get; set; }

        public string SQLPassword { get; set; }

        private SqlConnection _sqlConnection;
        private SqlCommand _sqlQuery;
        private string _password = (string)mRemoteNC.AppInfo.General.EncryptionKey;

        private int _currentNodeIndex = 0;
        private string _parentConstantId = "0";


        public void SaveConnections(bool update)
        {
            Runtime.SetMainFormText("SQL Server");

            bool tmrWasEnabled = false;

            if (Runtime.TimerSqlWatcher != null)
            {
                tmrWasEnabled = Runtime.TimerSqlWatcher.Enabled;
                if (Runtime.TimerSqlWatcher.Enabled == true)
                {
                    Runtime.TimerSqlWatcher.Stop();
                }
            }

            SaveToSQL();

            Runtime.LastSqlUpdate = DateTime.Now;
            if (tmrWasEnabled)
            {
                Runtime.TimerSqlWatcher.Start();
            }
        }

        private void SaveToSQL()
        {
            if (SQLUsername != "")
            {
                _sqlConnection =
                    new SqlConnection(
                        (string)
                        ("Data Source=" + SQLHost + ";Initial Catalog=" + SQLDatabaseName + ";User Id=" +
                         SQLUsername + ";Password=" + SQLPassword));
            }
            else
            {
                _sqlConnection =
                    new SqlConnection("Data Source=" + SQLHost + ";Initial Catalog=" + SQLDatabaseName +
                                      ";Integrated Security=True");
            }

            _sqlConnection.Open();

            if (!VerifyDatabaseVersion(_sqlConnection))
            {
                Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg,
                                                    Language.strErrorConnectionListSaveFailed);
                return;
            }

            TreeNode tN;
            tN = (TreeNode)RootTreeNode.Clone();

            string strProtected;
            if (tN.Tag != null)
            {
                if ((tN.Tag as Root.Info).Password == true)
                {
                    _password = (tN.Tag as Root.Info).PasswordString;
                    strProtected = Security.Crypt.Encrypt("ThisIsProtected", _password);
                }
                else
                {
                    strProtected = Security.Crypt.Encrypt("ThisIsNotProtected", _password);
                }
            }
            else
            {
                strProtected = Security.Crypt.Encrypt("ThisIsNotProtected", _password);
            }

            _sqlQuery = new SqlCommand("DELETE FROM tblRoot", _sqlConnection);
            _sqlQuery.ExecuteNonQuery();

            _sqlQuery =
                new SqlCommand(
                    "INSERT INTO tblRoot (Name, Export, Protected, ConfVersion) VALUES(\'" +
                    Tools.Misc.PrepareValueForDB(tN.Text) + "\', 0, \'" + strProtected + "\'," +
                    mRemoteNC.AppInfo.Connections.ConnectionFileVersion.ToString(CultureInfo.InvariantCulture) +
                    ")", _sqlConnection);
            _sqlQuery.ExecuteNonQuery();

            _sqlQuery = new SqlCommand("DELETE FROM tblCons", _sqlConnection);
            _sqlQuery.ExecuteNonQuery();

            TreeNodeCollection tNC;
            tNC = tN.Nodes;

            SaveNodesSQL(tNC);

            _sqlQuery = new SqlCommand("DELETE FROM tblUpdate", _sqlConnection);
            _sqlQuery.ExecuteNonQuery();
            _sqlQuery =
                new SqlCommand(
                    "INSERT INTO tblUpdate (LastUpdate) VALUES(\'" + Tools.Misc.DBDate(DateTime.Now) + "\')",
                    _sqlConnection);
            _sqlQuery.ExecuteNonQuery();

            _sqlConnection.Close();
        }

        private bool VerifyDatabaseVersion(SqlConnection sqlConnection)
        {
            bool isVerified = false;
            SqlDataReader sqlDataReader = null;
            System.Version databaseVersion = null;
            try
            {
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM tblRoot", sqlConnection);
                sqlDataReader = sqlCommand.ExecuteReader();
                if (!sqlDataReader.HasRows)
                {
                    return true; // assume new empty database
                }
                sqlDataReader.Read();

                databaseVersion =
                    new System.Version(
                        System.Convert.ToString(Convert.ToDouble(sqlDataReader["confVersion"],
                                                                 CultureInfo.InvariantCulture)));

                sqlDataReader.Close();

                if (databaseVersion.CompareTo(new System.Version(2, 2)) == 0) // 2.2
                {
                    Runtime.MessageCollector.AddMessage(Messages.MessageClass.InformationMsg,
                                                        string.Format(
                                                            "Upgrading database from version {0} to version {1}.",
                                                            databaseVersion.ToString(), "2.3"));
                    sqlCommand =
                        new SqlCommand(
                            "ALTER TABLE tblCons ADD EnableFontSmoothing bit NOT NULL DEFAULT 0, EnableDesktopComposition bit NOT NULL DEFAULT 0, InheritEnableFontSmoothing bit NOT NULL DEFAULT 0, InheritEnableDesktopComposition bit NOT NULL DEFAULT 0;",
                            sqlConnection);
                    sqlCommand.ExecuteNonQuery();
                    databaseVersion = new System.Version(2, 3);
                }

                if (databaseVersion.CompareTo(new System.Version(2, 3)) == 0) // 2.3
                {
                    Runtime.MessageCollector.AddMessage(Messages.MessageClass.InformationMsg,
                                                        string.Format(
                                                            "Upgrading database from version {0} to version {1}.",
                                                            databaseVersion.ToString(), "2.4"));
                    sqlCommand =
                        new SqlCommand(
                            "ALTER TABLE tblCons ADD UseCredSsp bit NOT NULL DEFAULT 1, InheritUseCredSsp bit NOT NULL DEFAULT 0;",
                            sqlConnection);
                    sqlCommand.ExecuteNonQuery();
                    databaseVersion = new Version(2, 4);
                }

                if (databaseVersion.CompareTo(new System.Version(2, 4)) == 0) // 2.4
                {
                    isVerified = true;
                }

                if (isVerified == false)
                {
                    Runtime.MessageCollector.AddMessage(Messages.MessageClass.WarningMsg,
                                                        string.Format(Language.strErrorBadDatabaseVersion,
                                                                      databaseVersion.ToString(),
                                                                      (new Microsoft.VisualBasic.ApplicationServices
                                                                          .WindowsFormsApplicationBase()).Info.
                                                                          ProductName));
                }
            }
            catch (Exception ex)
            {
                Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg,
                                                    string.Format(Language.strErrorVerifyDatabaseVersionFailed,
                                                                  ex.Message));
            }
            finally
            {
                if (sqlDataReader != null)
                {
                    if (!sqlDataReader.IsClosed)
                    {
                        sqlDataReader.Close();
                    }
                }
            }
            return isVerified;
        }

        private void SaveNodesSQL(TreeNodeCollection tnc)
        {
            foreach (TreeNode node in tnc)
            {
                _currentNodeIndex++;

                Connection.Info curConI;
                _sqlQuery =
                    new SqlCommand(
                        "INSERT INTO tblCons (Name, Type, Expanded, Description, Icon, Panel, Username, " +
                        "DomainName, Password, Hostname, Protocol, PuttySession, " +
                        "Port, ConnectToConsole, RenderingEngine, ICAEncryptionStrength, RDPAuthenticationLevel, Colors, Resolution, DisplayWallpaper, " +
                        "DisplayThemes, EnableFontSmoothing, EnableDesktopComposition, CacheBitmaps, RedirectDiskDrives, RedirectPorts, " +
                        "RedirectPrinters, RedirectSmartCards, RedirectSound, RedirectKeys, " +
                        "Connected, PreExtApp, PostExtApp, MacAddress, UserField, ExtApp, VNCCompression, VNCEncoding, VNCAuthMode, " +
                        "VNCProxyType, VNCProxyIP, VNCProxyPort, VNCProxyUsername, VNCProxyPassword, " +
                        "VNCColors, VNCSmartSizeMode, VNCViewOnly, " +
                        "RDGatewayUsageMethod, RDGatewayHostname, RDGatewayUseConnectionCredentials, RDGatewayUsername, RDGatewayPassword, RDGatewayDomain, " +
                        "UseCredSsp, " + "InheritCacheBitmaps, InheritColors, " +
                        "InheritDescription, InheritDisplayThemes, InheritDisplayWallpaper, InheritEnableFontSmoothing, InheritEnableDesktopComposition, InheritDomain, " +
                        "InheritIcon, InheritPanel, InheritPassword, InheritPort, " +
                        "InheritProtocol, InheritPuttySession, InheritRedirectDiskDrives, " +
                        "InheritRedirectKeys, InheritRedirectPorts, InheritRedirectPrinters, " +
                        "InheritRedirectSmartCards, InheritRedirectSound, InheritResolution, " +
                        "InheritUseConsoleSession, InheritRenderingEngine, InheritUsername, InheritICAEncryptionStrength, InheritRDPAuthenticationLevel, " +
                        "InheritPreExtApp, InheritPostExtApp, InheritMacAddress, InheritUserField, InheritExtApp, InheritVNCCompression, InheritVNCEncoding, " +
                        "InheritVNCAuthMode, InheritVNCProxyType, InheritVNCProxyIP, InheritVNCProxyPort, " +
                        "InheritVNCProxyUsername, InheritVNCProxyPassword, InheritVNCColors, " +
                        "InheritVNCSmartSizeMode, InheritVNCViewOnly, " +
                        "InheritRDGatewayUsageMethod, InheritRDGatewayHostname, InheritRDGatewayUseConnectionCredentials, InheritRDGatewayUsername, InheritRDGatewayPassword, InheritRDGatewayDomain, " +
                        "InheritUseCredSsp, "
                        + "PositionID, ParentID, ConstantID, LastChange)" + "VALUES (", _sqlConnection
                        );

                if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Connection ||
                    Tree.Node.GetNodeType(node) == Tree.Node.Type.Container)
                {
                    //_xmlTextWriter.WriteStartElement("Node")
                    _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB(node.Text) + "\',"; //Name
                    _sqlQuery.CommandText += "\'" + Tree.Node.GetNodeType(node).ToString() + "\',"; //Type
                }

                if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Container) //container
                {
                    _sqlQuery.CommandText += "\'" + this.ContainerList[node.Tag].IsExpanded + "\',"; //Expanded
                    curConI = this.ContainerList[node.Tag].ConnectionInfo;
                    SaveConnectionFieldsSQL(curConI);

                    _sqlQuery.CommandText = (string)(Tools.Misc.PrepareForDB(_sqlQuery.CommandText));
                    _sqlQuery.ExecuteNonQuery();
                    //_parentConstantId = _currentNodeIndex
                    SaveNodesSQL(node.Nodes);
                    //_xmlTextWriter.WriteEndElement()
                }

                if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Connection)
                {
                    _sqlQuery.CommandText += "\'" + false + "\',";
                    curConI = (Connection.Info)this.ConnectionList[node.Tag];
                    SaveConnectionFieldsSQL(curConI);
                    //_xmlTextWriter.WriteEndElement()
                    _sqlQuery.CommandText = (string)(Tools.Misc.PrepareForDB(_sqlQuery.CommandText));
                    _sqlQuery.ExecuteNonQuery();
                }

                //_parentConstantId = 0
            }
        }

        private void SaveConnectionFieldsSQL(Connection.Info curConI)
        {
            Connection.Info with_1 = curConI;
            _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB((string)with_1.Description) + "\',";
            _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB((string)with_1.Icon) + "\',";
            _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB((string)with_1.Panel) + "\',";

            if (this.SaveSecurity.Username == true)
            {
                _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB((string)with_1.Username) + "\',";
            }
            else
            {
                _sqlQuery.CommandText += "\'" + "" + "\',";
            }

            if (this.SaveSecurity.Domain == true)
            {
                _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB((string)with_1.Domain) + "\',";
            }
            else
            {
                _sqlQuery.CommandText += "\'" + "" + "\',";
            }

            if (this.SaveSecurity.Password == true)
            {
                _sqlQuery.CommandText += "\'" +
                                         Tools.Misc.PrepareValueForDB(
                                             Security.Crypt.Encrypt((string)with_1.Password, _password)) + "\',";
            }
            else
            {
                _sqlQuery.CommandText += "\'" + "" + "\',";
            }

            _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB((string)with_1.Hostname) + "\',";
            _sqlQuery.CommandText += "\'" + with_1.Protocol.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + Tools.Misc.PrepareValueForDB((string)with_1.PuttySession) + "\',";
            _sqlQuery.CommandText += "\'" + with_1.Port + "\',";
            _sqlQuery.CommandText += "\'" + with_1.UseConsoleSession + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RenderingEngine.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.ICAEncryption.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RDPAuthenticationLevel.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.Colors.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.Resolution.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.DisplayWallpaper + "\',";
            _sqlQuery.CommandText += "\'" + with_1.DisplayThemes + "\',";
            _sqlQuery.CommandText += "\'" + with_1.EnableFontSmoothing + "\',";
            _sqlQuery.CommandText += "\'" + with_1.EnableDesktopComposition + "\',";
            _sqlQuery.CommandText += "\'" + with_1.CacheBitmaps + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RedirectDiskDrives + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RedirectPorts + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RedirectPrinters + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RedirectSmartCards + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RedirectSound.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RedirectKeys + "\',";

            if (curConI.OpenConnections.Count > 0)
            {
                _sqlQuery.CommandText += 1 + ",";
            }
            else
            {
                _sqlQuery.CommandText += 0 + ",";
            }

            _sqlQuery.CommandText += "\'" + with_1.PreExtApp + "\',";
            _sqlQuery.CommandText += "\'" + with_1.PostExtApp + "\',";
            _sqlQuery.CommandText += "\'" + with_1.MacAddress + "\',";
            _sqlQuery.CommandText += "\'" + with_1.UserField + "\',";
            _sqlQuery.CommandText += "\'" + with_1.ExtApp + "\',";

            _sqlQuery.CommandText += "\'" + with_1.VNCCompression.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCEncoding.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCAuthMode.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCProxyType.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCProxyIP + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCProxyPort + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCProxyUsername + "\',";
            _sqlQuery.CommandText += "\'" + Security.Crypt.Encrypt((string)with_1.VNCProxyPassword, _password) +
                                     "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCColors.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCSmartSizeMode.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.VNCViewOnly + "\',";

            _sqlQuery.CommandText += "\'" + with_1.RDGatewayUsageMethod.ToString() + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RDGatewayHostname + "\',";
            _sqlQuery.CommandText += "\'" + with_1.RDGatewayUseConnectionCredentials.ToString() + "\',";

            if (this.SaveSecurity.Username == true)
            {
                _sqlQuery.CommandText += "\'" + with_1.RDGatewayUsername + "\',";
            }
            else
            {
                _sqlQuery.CommandText += "\'" + "" + "\',";
            }

            if (this.SaveSecurity.Password == true)
            {
                _sqlQuery.CommandText += "\'" + Crypt.Encrypt(with_1.RDGatewayPassword, _password) + "\',";
            }
            else
            {
                _sqlQuery.CommandText += "\'" + "" + "\',";
            }

            if (this.SaveSecurity.Domain == true)
            {
                _sqlQuery.CommandText += "\'" + with_1.RDGatewayDomain + "\',";
            }
            else
            {
                _sqlQuery.CommandText += "\'" + "" + "\',";
            }

            _sqlQuery.CommandText += "\'" + with_1.UseCredSsp + "\',";

            var with_2 = with_1.Inherit;
            if (this.SaveSecurity.Inheritance == true)
            {
                _sqlQuery.CommandText += "\'" + with_2.CacheBitmaps + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Colors + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Description + "\',";
                _sqlQuery.CommandText += "\'" + with_2.DisplayThemes + "\',";
                _sqlQuery.CommandText += "\'" + with_2.DisplayWallpaper + "\',";
                _sqlQuery.CommandText += "\'" + with_2.EnableFontSmoothing + "\',";
                _sqlQuery.CommandText += "\'" + with_2.EnableDesktopComposition + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Domain + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Icon + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Panel + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Password + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Port + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Protocol + "\',";
                _sqlQuery.CommandText += "\'" + with_2.PuttySession + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RedirectDiskDrives + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RedirectKeys + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RedirectPorts + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RedirectPrinters + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RedirectSmartCards + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RedirectSound + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Resolution + "\',";
                _sqlQuery.CommandText += "\'" + with_2.UseConsoleSession + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RenderingEngine + "\',";
                _sqlQuery.CommandText += "\'" + with_2.Username + "\',";
                _sqlQuery.CommandText += "\'" + with_2.ICAEncryption + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RDPAuthenticationLevel + "\',";
                _sqlQuery.CommandText += "\'" + with_2.PreExtApp + "\',";
                _sqlQuery.CommandText += "\'" + with_2.PostExtApp + "\',";
                _sqlQuery.CommandText += "\'" + with_2.MacAddress + "\',";
                _sqlQuery.CommandText += "\'" + with_2.UserField + "\',";
                _sqlQuery.CommandText += "\'" + with_2.ExtApp + "\',";

                _sqlQuery.CommandText += "\'" + with_2.VNCCompression + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCEncoding + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCAuthMode + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCProxyType + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCProxyIP + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCProxyPort + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCProxyUsername + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCProxyPassword + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCColors + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCSmartSizeMode + "\',";
                _sqlQuery.CommandText += "\'" + with_2.VNCViewOnly + "\',";

                _sqlQuery.CommandText += "\'" + with_2.RDGatewayUsageMethod + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RDGatewayHostname + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RDGatewayUseConnectionCredentials + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RDGatewayUsername + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RDGatewayPassword + "\',";
                _sqlQuery.CommandText += "\'" + with_2.RDGatewayDomain + "\',";

                _sqlQuery.CommandText += "\'" + with_2.UseCredSsp + "\',";
            }
            else
            {
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";

                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";
                _sqlQuery.CommandText += "\'" + false + "\',";

                _sqlQuery.CommandText += "\'" + false + "\',"; // .RDGatewayUsageMethod
                _sqlQuery.CommandText += "\'" + false + "\',"; // .RDGatewayHostname
                _sqlQuery.CommandText += "\'" + false + "\',"; // .RDGatewayUseConnectionCredentials
                _sqlQuery.CommandText += "\'" + false + "\',"; // .RDGatewayUsername
                _sqlQuery.CommandText += "\'" + false + "\',"; // .RDGatewayPassword
                _sqlQuery.CommandText += "\'" + false + "\',"; // .RDGatewayDomain

                _sqlQuery.CommandText += "\'" + false + "\',"; // .UseCredSsp
            }

            with_1.PositionID = _currentNodeIndex;

            if (with_1.IsContainer == false)
            {
                if (with_1.Parent != null)
                {
                    _parentConstantId = (with_1.Parent as Container.Info).ConnectionInfo.ConstantID;
                }
                else
                {
                    _parentConstantId = "0";
                }
            }
            else
            {
                if ((with_1.Parent as Container.Info).Parent != null)
                {
                    _parentConstantId =
                        ((with_1.Parent as Container.Info).Parent as Container.Info).ConnectionInfo.ConstantID;
                }
                else
                {
                    _parentConstantId = "0";
                }
            }

            _sqlQuery.CommandText +=
                System.Convert.ToString(_currentNodeIndex.ToString() + "," + _parentConstantId + "," +
                                        with_1.ConstantID + ",\'" + Tools.Misc.DBDate(DateTime.Now) + "\')");
        }
    }
}
