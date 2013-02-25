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
    public class VisionAppCsvConfigConnectionStrategy : IConfigConnectionStrategy
    {
        public bool Export { get; set; }

        public string ConnectionFileName { get; set; }

        public TreeNode RootTreeNode { get; set; }

        public Container.List ContainerList { get; set; }

        public Security.Save SaveSecurity { get; set; }

        public Connection.List ConnectionList { get; set; }

        public void SaveConnections(bool update)
        {
            SaveTovRDCSV();
        }

        private StreamWriter csvWr;

        private void SaveTovRDCSV()
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

            SaveNodevRDCSV(tNC);

            csvWr.Close();
        }

        private void SaveNodevRDCSV(TreeNodeCollection tNC)
        {
            foreach (TreeNode node in tNC)
            {
                if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Connection)
                {
                    Connection.Info curConI = (Connection.Info)node.Tag;

                    if (curConI.Protocol == mRemoteNC.Protocols.RDP)
                    {
                        WritevRDCSVLine(curConI);
                    }
                }
                else if (Tree.Node.GetNodeType(node) == Tree.Node.Type.Container)
                {
                    SaveNodevRDCSV(node.Nodes);
                }
            }
        }

        private void WritevRDCSVLine(Connection.Info con)
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

            csvWr.WriteLine(con.Name + ";" + con.Hostname + ";" + con.MacAddress + ";;" + con.Port + ";" +
                            con.UseConsoleSession + ";" + nodePath);
        }

    }
}
