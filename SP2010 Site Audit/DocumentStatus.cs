using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint.Client;


namespace SP2010_Site_Audit
{
	public partial class DocumentStatus : UserControl
	{
		public DocumentStatus()
		{
			InitializeComponent();
		}

		public string URL = String.Empty;

		private void DocumentStatus_Load(object sender,EventArgs e)
		{

		}

		public void LoadControl()
		{
			ClientContext clientContext = new ClientContext(URL);
			Web currentWeb = clientContext.Web;

			clientContext.Load(currentWeb);
			clientContext.ExecuteQuery();

			WebCollection childWebs = currentWeb.Webs;
			clientContext.Load(childWebs);
			clientContext.ExecuteQuery();

			foreach(Web tmpWeb in childWebs)
			{
				TreeNode tmpNode = this.tvStructure.Nodes.Add(tmpWeb.Title);

				clientContext.Load(tmpWeb);
				clientContext.ExecuteQuery();

				//WebCollection tmpChildWebs = tmpWeb.Webs;

				//clientContext.Load(tmpChildWebs);
				//clientContext.ExecuteQuery();

				//if(tmpChildWebs.Count > 0)
				//{

				//}
			}
		}
	}
}
