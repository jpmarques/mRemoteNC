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
    public class CsvConfigConnectionStrategy : IConfigConnectionStrategy
    {
        public bool Export { get; set; }

        public string ConnectionFileName { get; set; }

        public TreeNode RootTreeNode { get; set; }

        public Container.List ContainerList { get; set; }

        public Security.Save SaveSecurity { get; set; }

        public Connection.List ConnectionList { get; set; }

        public void SaveConnections(bool update)
        {
            SaveTomRCSV();
        }

        private StreamWriter csvWr;

        private void SaveTomRCSV()
        {
            if (Runtime.IsConnectionsFileLoaded == false)
            {
                return;
            }

            TreeNode tN;
            tN = (TreeNode)RootTreeNode.Clone();

            TreeNodeCollection tNC;
            tNC = tN.Nodes;

            csvWr = new StreamWriter(ConnectionFileName);

            string csvLn = string.Empty;

            csvLn += "Name;Folder;Description;Icon;Panel;";

            if (SaveSecurity.Username)
            {
                csvLn += "Username;";
            }

            if (SaveSecurity.Password)
            {
                csvLn += "Password;";
            }

            if (SaveSecurity.Domain)
            {
                csvLn += "Domain;";
            }

            csvLn +=
                "Hostname;Protocol;PuttySession;Port;ConnectToConsole;UseCredSsp;RenderingEngine;ICAEncryptionStrength;RDPAuthenticationLevel;Colors;Resolution;DisplayWallpaper;DisplayThemes;EnableFontSmoothing;EnableDesktopComposition;CacheBitmaps;RedirectDiskDrives;RedirectPorts;RedirectPrinters;RedirectSmartCards;RedirectSound;RedirectKeys;PreExtApp;PostExtApp;MacAddress;UserField;ExtApp;VNCCompression;VNCEncoding;VNCAuthMode;VNCProxyType;VNCProxyIP;VNCProxyPort;VNCProxyUsername;VNCProxyPassword;VNCColors;VNCSmartSizeMode;VNCViewOnly;RDGatewayUsageMethod;RDGatewayHostname;RDGatewayUseConnectionCredentials;RDGatewayUsername;RDGatewayPassword;RDGatewayDomain;";

            if (SaveSecurity.Inheritance)
            {
                csvLn +=
                    "InheritCacheBitmaps;InheritColors;InheritDescription;InheritDisplayThemes;InheritDisplayWallpaper;InheritEnableFontSmoothing;InheritEnableDesktopComposition;InheritDomain;InheritIcon;InheritPanel;InheritPassword;InheritPort;InheritProtocol;InheritPuttySession;InheritRedirectDiskDrives;InheritRedirectKeys;InheritRedirectPorts;InheritRedirectPrinters;InheritRedirectSmartCards;InheritRedirectSound;InheritResolution;InheritUseConsoleSession;InheritUseCredSsp;InheritRenderingEngine;InheritUsername;InheritICAEncryptionStrength;InheritRDPAuthenticationLevel;InheritPreExtApp;InheritPostExtApp;InheritMacAddress;InheritUserField;InheritExtApp;InheritVNCCompression;InheritVNCEncoding;InheritVNCAuthMode;InheritVNCProxyType;InheritVNCProxyIP;InheritVNCProxyPort;InheritVNCProxyUsername;InheritVNCProxyPassword;InheritVNCColors;InheritVNCSmartSizeMode;InheritVNCViewOnly;InheritRDGatewayUsageMethod;InheritRDGatewayHostname;InheritRDGatewayUseConnectionCredentials;InheritRDGatewayUsername;InheritRDGatewayPassword;InheritRDGatewayDomain";
            }

            csvWr.WriteLine(csvLn);

            SaveNodemRCSV(tNC);

            csvWr.Close();
        }

        private void SaveNodemRCSV(TreeNodeCollection tNC)
        {
            foreach (TreeNode node in tNC)
            {
                if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Connection)
                {
                    Connection.Info curConI = (Connection.Info)node.Tag;

                    WritemRCSVLine(curConI);
                }
                else if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Container)
                {
                    SaveNodemRCSV(node.Nodes);
                }
            }
        }

        private void WritemRCSVLine(Connection.Info con)
        {
            string nodePath = (string)con.TreeNode.FullPath;

            int firstSlash = nodePath.IndexOf("\\");
            nodePath = nodePath.Remove(0, firstSlash + 1);
            int lastSlash = nodePath.LastIndexOf("\\");

            if (lastSlash > 0)
            {
                nodePath = nodePath.Remove(lastSlash);
            }
            else
            {
                nodePath = "";
            }

            string csvLn = string.Empty;

            csvLn +=
                (string)
                (con.Name + ";" + nodePath + ";" + con.Description + ";" + con.Icon + ";" + con.Panel + ";");

            if (SaveSecurity.Username)
            {
                csvLn += con.Username + ";";
            }

            if (SaveSecurity.Password)
            {
                csvLn += con.Password + ";";
            }

            if (SaveSecurity.Domain)
            {
                csvLn += con.Domain + ";";
            }

            csvLn +=
                (string)
                (con.Hostname + ";" + con.Protocol.ToString() + ";" + con.PuttySession + ";" + con.Port + ";" +
                 con.UseConsoleSession + ";" + con.UseCredSsp + ";" + con.RenderingEngine.ToString() + ";" +
                 con.ICAEncryption.ToString() + ";" + con.RDPAuthenticationLevel.ToString() + ";" +
                 con.Colors.ToString() + ";" + con.Resolution.ToString() + ";" + con.DisplayWallpaper + ";" +
                 con.DisplayThemes + ";" + con.EnableFontSmoothing + ";" + con.EnableDesktopComposition + ";" +
                 con.CacheBitmaps + ";" + con.RedirectDiskDrives + ";" + con.RedirectPorts + ";" +
                 con.RedirectPrinters + ";" + con.RedirectSmartCards + ";" + con.RedirectSound.ToString() + ";" +
                 con.RedirectKeys + ";" + con.PreExtApp + ";" + con.PostExtApp + ";" + con.MacAddress + ";" +
                 con.UserField + ";" + con.ExtApp + ";" + con.VNCCompression.ToString() + ";" +
                 con.VNCEncoding.ToString() + ";" + con.VNCAuthMode.ToString() + ";" + con.VNCProxyType.ToString() +
                 ";" + con.VNCProxyIP + ";" + con.VNCProxyPort + ";" + con.VNCProxyUsername + ";" +
                 con.VNCProxyPassword + ";" + con.VNCColors.ToString() + ";" + con.VNCSmartSizeMode.ToString() + ";" +
                 con.VNCViewOnly + ";");

            if (SaveSecurity.Inheritance)
            {
                csvLn +=
                    (string)
                    (con.Inherit.CacheBitmaps + ";" + con.Inherit.Colors + ";" + con.Inherit.Description + ";" +
                     con.Inherit.DisplayThemes + ";" + con.Inherit.DisplayWallpaper + ";" +
                     con.Inherit.EnableFontSmoothing + ";" + con.Inherit.EnableDesktopComposition + ";" +
                     con.Inherit.Domain + ";" + con.Inherit.Icon + ";" + con.Inherit.Panel + ";" +
                     con.Inherit.Password + ";" + con.Inherit.Port + ";" + con.Inherit.Protocol + ";" +
                     con.Inherit.PuttySession + ";" + con.Inherit.RedirectDiskDrives + ";" +
                     con.Inherit.RedirectKeys + ";" + con.Inherit.RedirectPorts + ";" + con.Inherit.RedirectPrinters +
                     ";" + con.Inherit.RedirectSmartCards + ";" + con.Inherit.RedirectSound + ";" +
                     con.Inherit.Resolution + ";" + con.Inherit.UseConsoleSession + ";" + con.Inherit.UseCredSsp +
                     ";" + con.Inherit.RenderingEngine + ";" + con.Inherit.Username + ";" +
                     con.Inherit.ICAEncryption + ";" + con.Inherit.RDPAuthenticationLevel + ";" +
                     con.Inherit.PreExtApp + ";" + con.Inherit.PostExtApp + ";" + con.Inherit.MacAddress + ";" +
                     con.Inherit.UserField + ";" + con.Inherit.ExtApp + ";" + con.Inherit.VNCCompression + ";" +
                     con.Inherit.VNCEncoding + ";" + con.Inherit.VNCAuthMode + ";" + con.Inherit.VNCProxyType + ";" +
                     con.Inherit.VNCProxyIP + ";" + con.Inherit.VNCProxyPort + ";" + con.Inherit.VNCProxyUsername +
                     ";" + con.Inherit.VNCProxyPassword + ";" + con.Inherit.VNCColors + ";" +
                     con.Inherit.VNCSmartSizeMode + ";" + con.Inherit.VNCViewOnly);
            }

            csvWr.WriteLine(csvLn);
        }


    }
}
