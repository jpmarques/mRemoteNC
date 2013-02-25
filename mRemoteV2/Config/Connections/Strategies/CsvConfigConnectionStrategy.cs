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
        }

    }
}
