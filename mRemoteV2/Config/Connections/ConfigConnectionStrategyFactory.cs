using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using mRemoteNC.App;
using mRemoteNC.Config.Connections.Strategies;
using My;
using mRemoteNC.Security;
using System.Windows.Forms;

namespace mRemoteNC.Config.Connections
{
    public class ConfigConnectionStrategyFactory
    {
        public static IConfigConnectionStrategy GetConfigConnection()
        {

            IConfigConnectionStrategy strategy = null;

            if (Settings.Default.UseSQLServer == true)
            {
                strategy = new SQLConfigConnectionStrategy
                {
                    SQLHost = (string)Settings.Default.SQLHost,
                    SQLDatabaseName = (string)Settings.Default.SQLDatabaseName,
                    SQLUsername = (string)Settings.Default.SQLUser,
                    SQLPassword = Crypt.Decrypt((string)Settings.Default.SQLPass, (string)mRemoteNC.AppInfo.General.EncryptionKey)
                };
            }
            else
            {
                strategy = new XmlConfigConnectionStrategy();

                if (Settings.Default.LoadConsFromCustomLocation == false)
                {
                    strategy.ConnectionFileName = GetDefaultConfigFilePath();
                }
                else
                {
                    strategy.ConnectionFileName = (string)Settings.Default.CustomConsPath;
                }
            }

            strategy.ConnectionList = Runtime.ConnectionList;
            strategy.ContainerList = Runtime.ContainerList;
            strategy.Export = false;
            strategy.SaveSecurity = new mRemoteNC.Security.Save(false);
            strategy.RootTreeNode = mRemoteNC.App.Runtime.Windows.treeForm.tvConnections.Nodes[0];

            return strategy;
        }

        public static IConfigConnectionStrategy GetConfigConnection(Security.Save SaveSecurity, TreeNode RootNode, SaveFileDialog sD)
        {
            IConfigConnectionStrategy strategy = null;

            switch (sD.FilterIndex)
            {
                case 1:
                    strategy = new XmlConfigConnectionStrategy();
                    break;
                case 2:
                    strategy = new CsvConfigConnectionStrategy();
                    break;
                case 3:
                    strategy = new VisionAppVreConfigConnectionStrategy();
                    break;
                default:
                    strategy = new XmlConfigConnectionStrategy();
                    break;
            }

            strategy.ConnectionFileName = sD.FileName;

            if (RootNode == mRemoteNC.App.Runtime.Windows.treeForm.tvConnections.Nodes[0] && strategy is XmlConfigConnectionStrategy)
            {
                if (strategy.ConnectionFileName == GetDefaultConfigFilePath())
                {
                    Settings.Default.LoadConsFromCustomLocation = false;
                }
                else
                {
                    Settings.Default.LoadConsFromCustomLocation = true;
                    Settings.Default.CustomConsPath = strategy.ConnectionFileName;
                }
            }

            strategy.ConnectionList = Runtime.ConnectionList;
            strategy.ContainerList = Runtime.ContainerList;

            if (RootNode != mRemoteNC.App.Runtime.Windows.treeForm.tvConnections.Nodes[0])
            {
                strategy.Export = true;
            }

            strategy.SaveSecurity = SaveSecurity;
            strategy.RootTreeNode = RootNode;

            return strategy;
        }

        private static string GetDefaultConfigFilePath()
        {
            return mRemoteNC.AppInfo.Connections.DefaultConnectionsPath + "\\" + mRemoteNC.AppInfo.Connections.DefaultConnectionsFile;
        }
    }
}
