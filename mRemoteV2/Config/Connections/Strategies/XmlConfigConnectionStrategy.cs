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
    public class XmlConfigConnectionStrategy : IConfigConnectionStrategy
    {
        private XmlTextWriter _xmlTextWriter;

        private string _password = (string)mRemoteNC.AppInfo.General.EncryptionKey;

        public bool Export { get; set; }

        public string ConnectionFileName { get; set; }

        public TreeNode RootTreeNode { get; set; }

        public Container.List ContainerList { get; set; }

        public Security.Save SaveSecurity { get; set; }

        public Connection.List ConnectionList { get; set; }

        public void SaveConnections(bool update)
        {
            if (update)
            {
                return;
            }

            SaveToXml();

            if (Settings.Default.EncryptCompleteConnectionsFile)
            {
                EncryptCompleteFile();
            }
            if (!Export)
            {
                Runtime.SetMainFormText(ConnectionFileName);
            }
        }

        private void EncryptCompleteFile()
        {
            StreamReader streamReader = new StreamReader(ConnectionFileName);

            string fileContents;
            fileContents = streamReader.ReadToEnd();
            streamReader.Close();

            if (!string.IsNullOrEmpty(fileContents))
            {
                StreamWriter streamWriter = new StreamWriter(ConnectionFileName);
                streamWriter.Write(Security.Crypt.Encrypt(fileContents, _password));
                streamWriter.Close();
            }
        }

        private void SaveToXml()
        {
            try
            {
                if (Runtime.IsConnectionsFileLoaded == false)
                {
                    return;
                }

                TreeNode treeNode;
                bool isExport = false;

                if (Tree.Node.GetNodeType(RootTreeNode) == Tree.Node.Type.Root)
                {
                    treeNode = (TreeNode)RootTreeNode.Clone();
                }
                else
                {
                    treeNode = new TreeNode("mR|Export (" + Tools.Misc.DBDate(DateTime.Now) + ")");
                    treeNode.Nodes.Add((string)(RootTreeNode.Clone()));
                    isExport = true;
                }

                string tempFileName = Path.GetTempFileName();
                _xmlTextWriter = new XmlTextWriter(tempFileName, System.Text.Encoding.UTF8);

                _xmlTextWriter.Formatting = Formatting.Indented;
                _xmlTextWriter.Indentation = 4;

                _xmlTextWriter.WriteStartDocument();

                _xmlTextWriter.WriteStartElement("Connections"); // Do not localize
                _xmlTextWriter.WriteAttributeString("Name", "", treeNode.Text);
                _xmlTextWriter.WriteAttributeString("Export", "", isExport.ToString());

                if (isExport)
                {
                    _xmlTextWriter.WriteAttributeString("Protected", "",
                                                        Security.Crypt.Encrypt("ThisIsNotProtected", _password));
                }
                else
                {
                    if ((treeNode.Tag as Root.Info).Password == true)
                    {
                        _password = (treeNode.Tag as Root.Info).PasswordString;
                        _xmlTextWriter.WriteAttributeString("Protected", "",
                                                            Security.Crypt.Encrypt("ThisIsProtected", _password));
                    }
                    else
                    {
                        _xmlTextWriter.WriteAttributeString("Protected", "",
                                                            Security.Crypt.Encrypt("ThisIsNotProtected", _password));
                    }
                }

                _xmlTextWriter.WriteAttributeString("ConfVersion", "",
                                                    (string)
                                                    (mRemoteNC.AppInfo.Connections.ConnectionFileVersion.ToString(
                                                        CultureInfo.InvariantCulture)));

                TreeNodeCollection treeNodeCollection;
                treeNodeCollection = treeNode.Nodes;

                SaveNode(treeNodeCollection);

                _xmlTextWriter.WriteEndElement();
                _xmlTextWriter.Close();

                string backupFileName = ConnectionFileName + ".backup";
                File.Delete(backupFileName);
                if (File.Exists(ConnectionFileName))
                {
                    File.Move(ConnectionFileName, backupFileName);
                }
                File.Move(tempFileName, ConnectionFileName);
            }
            catch (Exception ex)
            {
                Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg,
                                                    (string)("SaveToXml failed" + Constants.vbNewLine + ex.Message),
                                                    true);
            }
        }

        private void SaveNode(TreeNodeCollection tNC)
        {
            try
            {
                foreach (TreeNode node in tNC)
                {
                    Connection.Info curConI;

                    if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Connection ||
                        Tree.Node.GetNodeType(node) == Tree.Node.Type.Container)
                    {
                        _xmlTextWriter.WriteStartElement("Node");
                        _xmlTextWriter.WriteAttributeString("Name", "", node.Text);
                        _xmlTextWriter.WriteAttributeString("Type", "", Tree.Node.GetNodeType(node).ToString());
                    }

                    if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Container) //container
                    {
                        _xmlTextWriter.WriteAttributeString("Expanded", "",
                                                            (string)
                                                            (this.ContainerList[node.Tag].TreeNode.IsExpanded).
                                                                ToString());
                        curConI = this.ContainerList[node.Tag].ConnectionInfo;
                        SaveConnectionFields(curConI);
                        SaveNode(node.Nodes);
                        _xmlTextWriter.WriteEndElement();
                    }

                    if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Connection)
                    {
                        curConI = (Connection.Info)this.ConnectionList[node.Tag];
                        SaveConnectionFields(curConI);
                        _xmlTextWriter.WriteEndElement();
                    }
                }
            }
            catch (Exception ex)
            {
                Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg,
                                                    (string)("SaveNode failed" + Constants.vbNewLine + ex.Message),
                                                    true);
            }
        }

        private void SaveConnectionFields(Connection.Info curConI)
        {
            try
            {
                _xmlTextWriter.WriteAttributeString("Descr", "", (string)curConI.Description);

                _xmlTextWriter.WriteAttributeString("Icon", "", (string)curConI.Icon);

                _xmlTextWriter.WriteAttributeString("Panel", "", (string)curConI.Panel);

                if (this.SaveSecurity.Username == true)
                {
                    _xmlTextWriter.WriteAttributeString("Username", "", (string)curConI.Username);
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("Username", "", "");
                }

                if (this.SaveSecurity.Domain == true)
                {
                    _xmlTextWriter.WriteAttributeString("Domain", "", (string)curConI.Domain);
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("Domain", "", "");
                }

                if (this.SaveSecurity.Password == true)
                {
                    _xmlTextWriter.WriteAttributeString("Password", "",
                                                        Security.Crypt.Encrypt((string)curConI.Password, _password));
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("Password", "", "");
                }

                _xmlTextWriter.WriteAttributeString("Hostname", "", (string)curConI.Hostname);

                _xmlTextWriter.WriteAttributeString("Protocol", "", (string)(curConI.Protocol.ToString()));

                _xmlTextWriter.WriteAttributeString("PuttySession", "", (string)curConI.PuttySession);

                _xmlTextWriter.WriteAttributeString("Port", "", (string)curConI.Port.ToString());

                _xmlTextWriter.WriteAttributeString("ConnectToConsole", "",
                                                    (string)curConI.UseConsoleSession.ToString());

                _xmlTextWriter.WriteAttributeString("UseCredSsp", "", (string)curConI.UseCredSsp.ToString());

                _xmlTextWriter.WriteAttributeString("RenderingEngine", "",
                                                    (string)(curConI.RenderingEngine.ToString()));

                _xmlTextWriter.WriteAttributeString("ICAEncryptionStrength", "",
                                                    (string)(curConI.ICAEncryption.ToString()));

                _xmlTextWriter.WriteAttributeString("RDPAuthenticationLevel", "",
                                                    (string)(curConI.RDPAuthenticationLevel.ToString()));

                _xmlTextWriter.WriteAttributeString("Colors", "", (string)(curConI.Colors.ToString()));

                _xmlTextWriter.WriteAttributeString("Resolution", "", (string)(curConI.Resolution.ToString()));

                _xmlTextWriter.WriteAttributeString("DisplayWallpaper", "",
                                                    (string)curConI.DisplayWallpaper.ToString());

                _xmlTextWriter.WriteAttributeString("DisplayThemes", "", (string)curConI.DisplayThemes.ToString());

                _xmlTextWriter.WriteAttributeString("EnableFontSmoothing", "",
                                                    (string)curConI.EnableFontSmoothing.ToString());

                _xmlTextWriter.WriteAttributeString("EnableDesktopComposition", "",
                                                    (string)curConI.EnableDesktopComposition.ToString());

                _xmlTextWriter.WriteAttributeString("CacheBitmaps", "", (string)curConI.CacheBitmaps.ToString());

                _xmlTextWriter.WriteAttributeString("RedirectDiskDrives", "",
                                                    (string)curConI.RedirectDiskDrives.ToString());

                _xmlTextWriter.WriteAttributeString("RedirectPorts", "", (string)curConI.RedirectPorts.ToString());

                _xmlTextWriter.WriteAttributeString("RedirectPrinters", "",
                                                    (string)curConI.RedirectPrinters.ToString());

                _xmlTextWriter.WriteAttributeString("RedirectSmartCards", "",
                                                    (string)curConI.RedirectSmartCards.ToString());

                _xmlTextWriter.WriteAttributeString("RedirectSound", "", (string)(curConI.RedirectSound.ToString()));

                _xmlTextWriter.WriteAttributeString("RedirectKeys", "", (string)curConI.RedirectKeys.ToString());

                if (curConI.OpenConnections.Count > 0)
                {
                    _xmlTextWriter.WriteAttributeString("Connected", "", true.ToString());
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("Connected", "", false.ToString());
                }

                _xmlTextWriter.WriteAttributeString("PreExtApp", "", (string)curConI.PreExtApp);
                _xmlTextWriter.WriteAttributeString("PostExtApp", "", (string)curConI.PostExtApp);
                _xmlTextWriter.WriteAttributeString("MacAddress", "", (string)curConI.MacAddress);
                _xmlTextWriter.WriteAttributeString("UserField", "", (string)curConI.UserField);
                _xmlTextWriter.WriteAttributeString("ExtApp", "", (string)curConI.ExtApp);

                _xmlTextWriter.WriteAttributeString("VNCCompression", "",
                                                    (string)(curConI.VNCCompression.ToString()));
                _xmlTextWriter.WriteAttributeString("VNCEncoding", "", (string)(curConI.VNCEncoding.ToString()));
                _xmlTextWriter.WriteAttributeString("VNCAuthMode", "", (string)(curConI.VNCAuthMode.ToString()));
                _xmlTextWriter.WriteAttributeString("VNCProxyType", "", (string)(curConI.VNCProxyType.ToString()));
                _xmlTextWriter.WriteAttributeString("VNCProxyIP", "", (string)curConI.VNCProxyIP);
                _xmlTextWriter.WriteAttributeString("VNCProxyPort", "", curConI.VNCProxyPort.ToString());
                _xmlTextWriter.WriteAttributeString("VNCProxyUsername", "", (string)curConI.VNCProxyUsername);
                _xmlTextWriter.WriteAttributeString("VNCProxyPassword", "",
                                                    Security.Crypt.Encrypt((string)curConI.VNCProxyPassword,
                                                                           _password));
                _xmlTextWriter.WriteAttributeString("VNCColors", "", (string)(curConI.VNCColors.ToString()));
                _xmlTextWriter.WriteAttributeString("VNCSmartSizeMode", "",
                                                    (string)(curConI.VNCSmartSizeMode.ToString()));
                _xmlTextWriter.WriteAttributeString("VNCViewOnly", "", (string)curConI.VNCViewOnly.ToString());

                _xmlTextWriter.WriteAttributeString("RDGatewayUsageMethod", "",
                                                    (string)(curConI.RDGatewayUsageMethod.ToString()));
                _xmlTextWriter.WriteAttributeString("RDGatewayHostname", "", (string)curConI.RDGatewayHostname);

                _xmlTextWriter.WriteAttributeString("RDGatewayUseConnectionCredentials", "",
                                                    (string)(curConI.RDGatewayUseConnectionCredentials.ToString()));

                _xmlTextWriter.WriteAttributeString("ConnectOnStartup", "", (string)curConI.ConnectOnStartup.ToString());

                if (this.SaveSecurity.Username == true)
                {
                    _xmlTextWriter.WriteAttributeString("RDGatewayUsername", "", (string)curConI.RDGatewayUsername);
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("RDGatewayUsername", "", "");
                }

                if (this.SaveSecurity.Password == true)
                {
                    _xmlTextWriter.WriteAttributeString("RDGatewayPassword", "", Crypt.Encrypt(curConI.RDGatewayPassword, _password));
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("RDGatewayPassword", "", "");
                }

                if (this.SaveSecurity.Domain == true)
                {
                    _xmlTextWriter.WriteAttributeString("RDGatewayDomain", "", (string)curConI.RDGatewayDomain);
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("RDGatewayDomain", "", "");
                }

                if (this.SaveSecurity.Inheritance == true)
                {
                    _xmlTextWriter.WriteAttributeString("InheritCacheBitmaps", "",
                                                        curConI.Inherit.CacheBitmaps.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritColors", "", curConI.Inherit.Colors.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDescription", "",
                                                        curConI.Inherit.Description.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDisplayThemes", "",
                                                        curConI.Inherit.DisplayThemes.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDisplayWallpaper", "",
                                                        curConI.Inherit.DisplayWallpaper.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritEnableFontSmoothing", "",
                                                        curConI.Inherit.EnableFontSmoothing.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritEnableDesktopComposition", "",
                                                        curConI.Inherit.EnableDesktopComposition.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDomain", "", curConI.Inherit.Domain.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritIcon", "", curConI.Inherit.Icon.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPanel", "", curConI.Inherit.Panel.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPassword", "", curConI.Inherit.Password.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPort", "", curConI.Inherit.Port.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritProtocol", "", curConI.Inherit.Protocol.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPuttySession", "",
                                                        curConI.Inherit.PuttySession.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectDiskDrives", "",
                                                        curConI.Inherit.RedirectDiskDrives.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectKeys", "",
                                                        curConI.Inherit.RedirectKeys.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectPorts", "",
                                                        curConI.Inherit.RedirectPorts.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectPrinters", "",
                                                        curConI.Inherit.RedirectPrinters.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectSmartCards", "",
                                                        curConI.Inherit.RedirectSmartCards.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectSound", "",
                                                        curConI.Inherit.RedirectSound.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritResolution", "",
                                                        curConI.Inherit.Resolution.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUseConsoleSession", "",
                                                        curConI.Inherit.UseConsoleSession.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUseCredSsp", "",
                                                        curConI.Inherit.UseCredSsp.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRenderingEngine", "",
                                                        curConI.Inherit.RenderingEngine.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUsername", "", curConI.Inherit.Username.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritICAEncryptionStrength", "",
                                                        curConI.Inherit.ICAEncryption.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDPAuthenticationLevel", "",
                                                        curConI.Inherit.RDPAuthenticationLevel.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPreExtApp", "", curConI.Inherit.PreExtApp.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPostExtApp", "",
                                                        curConI.Inherit.PostExtApp.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritMacAddress", "",
                                                        curConI.Inherit.MacAddress.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUserField", "", curConI.Inherit.UserField.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritExtApp", "", curConI.Inherit.ExtApp.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCCompression", "",
                                                        curConI.Inherit.VNCCompression.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCEncoding", "",
                                                        curConI.Inherit.VNCEncoding.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCAuthMode", "",
                                                        curConI.Inherit.VNCAuthMode.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyType", "",
                                                        curConI.Inherit.VNCProxyType.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyIP", "",
                                                        curConI.Inherit.VNCProxyIP.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyPort", "",
                                                        curConI.Inherit.VNCProxyPort.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyUsername", "",
                                                        curConI.Inherit.VNCProxyUsername.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyPassword", "",
                                                        curConI.Inherit.VNCProxyPassword.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCColors", "", curConI.Inherit.VNCColors.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCSmartSizeMode", "",
                                                        curConI.Inherit.VNCSmartSizeMode.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCViewOnly", "",
                                                        curConI.Inherit.VNCViewOnly.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayUsageMethod", "",
                                                        curConI.Inherit.RDGatewayUsageMethod.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayHostname", "",
                                                        curConI.Inherit.RDGatewayHostname.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayUseConnectionCredentials", "",
                                                        curConI.Inherit.RDGatewayUseConnectionCredentials.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayUsername", "",
                                                        curConI.Inherit.RDGatewayUsername.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayPassword", "",
                                                        curConI.Inherit.RDGatewayPassword.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayDomain", "",
                                                        curConI.Inherit.RDGatewayDomain.ToString());
                }
                else
                {
                    _xmlTextWriter.WriteAttributeString("InheritCacheBitmaps", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritColors", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDescription", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDisplayThemes", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDisplayWallpaper", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritEnableFontSmoothing", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritEnableDesktopComposition", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritDomain", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritIcon", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPanel", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPassword", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPort", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritProtocol", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPuttySession", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectDiskDrives", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectKeys", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectPorts", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectPrinters", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectSmartCards", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRedirectSound", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritResolution", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUseConsoleSession", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUseCredSsp", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRenderingEngine", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUsername", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritICAEncryptionStrength", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDPAuthenticationLevel", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPreExtApp", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritPostExtApp", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritMacAddress", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritUserField", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritExtApp", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCCompression", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCEncoding", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCAuthMode", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyType", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyIP", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyPort", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyUsername", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCProxyPassword", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCColors", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCSmartSizeMode", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritVNCViewOnly", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayHostname", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayUseConnectionCredentials", "",
                                                        false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayUsername", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayPassword", "", false.ToString());
                    _xmlTextWriter.WriteAttributeString("InheritRDGatewayDomain", "", false.ToString());
                }
            }
            catch (Exception ex)
            {
                Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg,
                                                    (string)
                                                    ("SaveConnectionFields failed" + Constants.vbNewLine +
                                                     ex.Message), true);
            }
        }
    }
}
