using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace mRemoteNC.Config.Connections
{
    public interface IConfigConnectionStrategy
    {
        void SaveConnections(bool update);
        bool Export { get; set; }
        string ConnectionFileName { get; set; }
        TreeNode RootTreeNode { get; set; }
        Container.List ContainerList { get; set; }
        Security.Save SaveSecurity { get; set; }
        Connection.List ConnectionList { get; set; }
    }
}
