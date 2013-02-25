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
    public class VisionAppVreConfigConnectionStrategy : IConfigConnectionStrategy
    {
        private XmlTextWriter _xmlTextWriter;

        public bool Export { get; set; }

        public string ConnectionFileName { get; set; }

        public TreeNode RootTreeNode { get; set; }

        public Container.List ContainerList { get; set; }

        public Security.Save SaveSecurity { get; set; }

        public Connection.List ConnectionList { get; set; }

        public void SaveConnections(bool update)
        {
            SaveToVRE();
        }

        private void SaveToVRE()
        {
            if (Runtime.IsConnectionsFileLoaded == false)
            {
                return;
            }

            TreeNode tN;
            tN = (TreeNode)RootTreeNode.Clone();

            TreeNodeCollection tNC;
            tNC = tN.Nodes;

            _xmlTextWriter = new XmlTextWriter(ConnectionFileName, System.Text.Encoding.UTF8);
            _xmlTextWriter.Formatting = Formatting.Indented;
            _xmlTextWriter.Indentation = 4;

            _xmlTextWriter.WriteStartDocument();

            _xmlTextWriter.WriteStartElement("vRDConfig");
            _xmlTextWriter.WriteAttributeString("Version", "", "2.0");

            _xmlTextWriter.WriteStartElement("Connections");
            SaveNodeVRE(tNC);
            _xmlTextWriter.WriteEndElement();

            _xmlTextWriter.WriteEndElement();
            _xmlTextWriter.WriteEndDocument();
            _xmlTextWriter.Close();
        }

        private void SaveNodeVRE(TreeNodeCollection tNC)
        {
            foreach (TreeNode node in tNC)
            {
                if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Connection)
                {
                    Connection.Info curConI = (Connection.Info)node.Tag;

                    if (curConI.Protocol == mRemoteNC.Protocols.RDP)
                    {
                        _xmlTextWriter.WriteStartElement("Connection");
                        _xmlTextWriter.WriteAttributeString("Id", "", "");

                        WriteVREitem(curConI);

                        _xmlTextWriter.WriteEndElement();
                    }
                }
                else
                {
                    SaveNodeVRE(node.Nodes);
                }
            }
        }

        private void WriteVREitem(Connection.Info con)
        {
            //Name
            _xmlTextWriter.WriteStartElement("ConnectionName");
            _xmlTextWriter.WriteValue(con.Name);
            _xmlTextWriter.WriteEndElement();

            //Hostname
            _xmlTextWriter.WriteStartElement("ServerName");
            _xmlTextWriter.WriteValue(con.Hostname);
            _xmlTextWriter.WriteEndElement();

            //Mac Adress
            _xmlTextWriter.WriteStartElement("MACAddress");
            _xmlTextWriter.WriteValue(con.MacAddress);
            _xmlTextWriter.WriteEndElement();

            //Management Board URL
            _xmlTextWriter.WriteStartElement("MgmtBoardUrl");
            _xmlTextWriter.WriteValue("");
            _xmlTextWriter.WriteEndElement();

            //Description
            _xmlTextWriter.WriteStartElement("Description");
            _xmlTextWriter.WriteValue(con.Description);
            _xmlTextWriter.WriteEndElement();

            //Port
            _xmlTextWriter.WriteStartElement("Port");
            _xmlTextWriter.WriteValue(con.Port);
            _xmlTextWriter.WriteEndElement();

            //Console Session
            _xmlTextWriter.WriteStartElement("Console");
            _xmlTextWriter.WriteValue(con.UseConsoleSession);
            _xmlTextWriter.WriteEndElement();

            //Redirect Clipboard
            _xmlTextWriter.WriteStartElement("ClipBoard");
            _xmlTextWriter.WriteValue(true);
            _xmlTextWriter.WriteEndElement();

            //Redirect Printers
            _xmlTextWriter.WriteStartElement("Printer");
            _xmlTextWriter.WriteValue(con.RedirectPrinters);
            _xmlTextWriter.WriteEndElement();

            //Redirect Ports
            _xmlTextWriter.WriteStartElement("Serial");
            _xmlTextWriter.WriteValue(con.RedirectPorts);
            _xmlTextWriter.WriteEndElement();

            //Redirect Disks
            _xmlTextWriter.WriteStartElement("LocalDrives");
            _xmlTextWriter.WriteValue(con.RedirectDiskDrives);
            _xmlTextWriter.WriteEndElement();

            //Redirect Smartcards
            _xmlTextWriter.WriteStartElement("SmartCard");
            _xmlTextWriter.WriteValue(con.RedirectSmartCards);
            _xmlTextWriter.WriteEndElement();

            //Connection Place
            _xmlTextWriter.WriteStartElement("ConnectionPlace");
            _xmlTextWriter.WriteValue("2"); //----------------------------------------------------------
            _xmlTextWriter.WriteEndElement();

            //Smart Size
            _xmlTextWriter.WriteStartElement("AutoSize");
            _xmlTextWriter.WriteValue(con.Resolution == mRemoteNC.RDP.RDPResolutions.SmartSize
                                          ? true
                                          : false);
            _xmlTextWriter.WriteEndElement();

            //SeparateResolutionX
            _xmlTextWriter.WriteStartElement("SeparateResolutionX");
            _xmlTextWriter.WriteValue("1024");
            _xmlTextWriter.WriteEndElement();

            //SeparateResolutionY
            _xmlTextWriter.WriteStartElement("SeparateResolutionY");
            _xmlTextWriter.WriteValue("768");
            _xmlTextWriter.WriteEndElement();

            //TabResolutionX
            _xmlTextWriter.WriteStartElement("TabResolutionX");
            if (con.Resolution != mRemoteNC.RDP.RDPResolutions.FitToWindow &&
                con.Resolution != mRemoteNC.RDP.RDPResolutions.Fullscreen &&
                con.Resolution != mRemoteNC.RDP.RDPResolutions.SmartSize)
            {
                _xmlTextWriter.WriteValue(con.Resolution.ToString().Remove(con.Resolution.ToString().IndexOf("x")));
            }
            else
            {
                _xmlTextWriter.WriteValue("1024");
            }
            _xmlTextWriter.WriteEndElement();

            //TabResolutionY
            _xmlTextWriter.WriteStartElement("TabResolutionY");
            if (con.Resolution != mRemoteNC.RDP.RDPResolutions.FitToWindow &&
                con.Resolution != mRemoteNC.RDP.RDPResolutions.Fullscreen &&
                con.Resolution != mRemoteNC.RDP.RDPResolutions.SmartSize)
            {
                _xmlTextWriter.WriteValue(con.Resolution.ToString().Remove(0, con.Resolution.ToString().IndexOf("x")));
            }
            else
            {
                _xmlTextWriter.WriteValue("768");
            }
            _xmlTextWriter.WriteEndElement();

            //RDPColorDepth
            _xmlTextWriter.WriteStartElement("RDPColorDepth");
            _xmlTextWriter.WriteValue(con.Colors.ToString().Replace("Colors", "").Replace("Bit", ""));
            _xmlTextWriter.WriteEndElement();

            //Bitmap Caching
            _xmlTextWriter.WriteStartElement("BitmapCaching");
            _xmlTextWriter.WriteValue(con.CacheBitmaps);
            _xmlTextWriter.WriteEndElement();

            //Themes
            _xmlTextWriter.WriteStartElement("Themes");
            _xmlTextWriter.WriteValue(con.DisplayThemes);
            _xmlTextWriter.WriteEndElement();

            //Wallpaper
            _xmlTextWriter.WriteStartElement("Wallpaper");
            _xmlTextWriter.WriteValue(con.DisplayWallpaper);
            _xmlTextWriter.WriteEndElement();
        }

    }
}
