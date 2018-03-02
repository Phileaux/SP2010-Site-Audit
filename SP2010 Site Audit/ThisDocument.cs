using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace SP2010_Site_Audit
{
    public partial class ThisDocument
    {
		private DocumentStatus actionPane = null; 

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
			actionPane = new DocumentStatus();
			this.ActionsPane.Controls.Add(actionPane);

        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
			this.btnLoadForm.Click += new System.EventHandler(this.btnLoadForm_Click);
			this.Startup += new System.EventHandler(this.ThisDocument_Startup);
			this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

		}

		#endregion

		private void btnLoadForm_Click(object sender,EventArgs e)
		{

			// ===================== Populate Date Generated ==============================
			Word.ContentControl ccDateGenerated = GetContentControl("txtDateGenerated");
			ccDateGenerated.Range.Text = DateTime.Now.ToString("MM/dd/yyyy");

			String Stage = "Page Load";

			WebCollection childWebs = null; 

			String txtURL = String.Empty;
			Word.ContentControl ccURL = GetContentControl("txtURL");

			// Clear errors
			SetContentControl("txtSmartArtError","");
			SetContentControl("txtURLError", "");

			if (ccURL != null)
			{
				Stage = "Load Action Pane";
				// http://oigportal.hhsoig.gov/sites/OAS/AATS/TMSIS
				txtURL = ccURL.Range.Text;

				txtURL = txtURL.Replace("SitePages/Home.aspx","");

				txtURL = txtURL.TrimEnd('/');

				actionPane.URL = txtURL;
				actionPane.LoadControl();



				try
				{
					Stage = "Load Web";
					ClientContext clientContext = new ClientContext(txtURL);
					Web currentWeb = clientContext.Web;

					clientContext.Load(currentWeb);
					clientContext.ExecuteQuery();

					// Get Web details
					Guid webId = currentWeb.Id;
					SetContentControl("txtTitle",currentWeb.Title);
					SetContentControl("txtCreated",currentWeb.Created.ToString("MM/dd/yyyy"));
					SetContentControl("txtModified",currentWeb.LastItemModifiedDate.ToString("MM/dd/yyyy"));

					long webSize = GetWebSize(currentWeb);
					SetContentControl("txtSize",webSize.ToString("N0"));

					// Set document properties
					Microsoft.Office.Core.DocumentProperties properties;
					properties = (Office.DocumentProperties)this.CustomDocumentProperties;

					// properties["Title"].Value = currentWeb.Title;

					#region Smart Art Population
					// ============ Smart Art =====================================================

					try
					{
						// Set up for diagram
						Stage = "Modify Smart Art";
						// for background color of the current web cell in the smartart. 
						const int OrangeCell = unchecked((int)0xED7D31);

						Site tempSite = clientContext.Site;
						clientContext.Load(tempSite);
						clientContext.ExecuteQuery();
						
						string siteUrl = tempSite.Url + "/";  // http://oigportal.hhsoig.gov/sites/OAS

						Web tmpRoot = tempSite.RootWeb;
						clientContext.Load(tmpRoot);
						clientContext.ExecuteQuery();

						string rootTitle = tmpRoot.Title;

						// Get site names by breaking down URL. 
						//  SharePoint 2010 client Web class doesn't have any way to get the parent web. 
						//  example: AATS/TMSIS
						string navTree = txtURL.Replace(siteUrl,"");
						string[] nodes = navTree.Split('/');

						// Find the diagram and get a reference to it. 
						Word.InlineShape treeShape = null;

						foreach (Word.InlineShape tmpShape in this.InlineShapes)
						{
							if(tmpShape.Type == Word.WdInlineShapeType.wdInlineShapeSmartArt)
							{
								treeShape = tmpShape;
							}
						}

						Office.SmartArt treeArt = treeShape.SmartArt;
						// treeShape.Height

						// clear out existing nodes
						foreach (Office.SmartArtNode tmpNode in treeArt.Nodes)
						{
							if (tmpNode != null)
							{
								tmpNode.Delete();
							}
						}

						Office.SmartArtNode rootNode = treeArt.Nodes.Add();
						rootNode.TextFrame2.TextRange.Text = rootTitle;

						// Nodes from root to current site
						foreach (string tmpNodeText in nodes)
						{
							Office.SmartArtNode tmpChildNode = treeArt.Nodes.Add();
							tmpChildNode.TextFrame2.TextRange.Text = tmpNodeText;
						}

						// Root node - add then node, then set the text.
						Office.SmartArtNode currentNode = treeArt.Nodes[treeArt.Nodes.Count];
						currentNode.TextFrame2.TextRange.Text = currentWeb.Title;
						// set root node color

						currentNode.Shapes.Fill.ForeColor.RGB = 0xED7D31; // OrangeCell;

						// Child webs for SmartArt
						childWebs = currentWeb.Webs;
						clientContext.Load(childWebs);
						clientContext.ExecuteQuery();

						foreach(Web tmpWeb in childWebs)
						{
							Office.SmartArtNode childNode = currentNode.AddNode(Office.MsoSmartArtNodePosition.msoSmartArtNodeBelow);
							childNode.TextFrame2.TextRange.Text = tmpWeb.Title;
						}

					}
					catch (Exception ex)
					{
						Word.ContentControl smartArtError = GetContentControl("txtSmartArtError");
						Word.Range tagRange = smartArtError.Range;
						tagRange.Text = String.Concat("ERROR: ",ex.Message);
						tagRange.Font.Color = Word.WdColor.wdColorRed;
					}

					#endregion

					#region Build Child Web Table
					// ============ Child Web Table ===============================================
					Stage = "Load Child Web Table";

					Word.Table webTable = GetTable("ChildWebs");

					if (webTable != null)
					{
						foreach (Web tmpWeb in childWebs)
						{
							Word.Row newRow = webTable.Rows.Add();

							newRow.Cells[1].Range.Text = tmpWeb.Title;
							newRow.Cells[2].Range.Text = tmpWeb.ServerRelativeUrl;
							// newRow.Cells[3].Range.Text = Owners
							newRow.Cells[4].Range.Text = tmpWeb.Created.ToString("MM/dd/yyyy");

							long WebSize = GetWebSize(tmpWeb);
							newRow.Cells[5].Range.Text = WebSize.ToString("N0");

						}
					}
					#endregion

					#region Build Child Object Table
					// ================== Child Object Table =========================================
					Microsoft.SharePoint.Client.ListCollection webLists = currentWeb.Lists;
					clientContext.Load(webLists);
					clientContext.ExecuteQuery();

					Word.Table objTable = GetTable("tblContentObjects");

					if (objTable != null)
					{
						foreach (List tmpList in webLists)
						{
							Word.Row newRow = objTable.Rows.Add();

							newRow.Cells[1].Range.Text = tmpList.BaseType.ToString();
							newRow.Cells[2].Range.Text = tmpList.Title;
							newRow.Cells[3].Range.Text = tmpList.ItemCount.ToString();
							newRow.Cells[4].Range.Text = tmpList.LastItemModifiedDate.ToString("MM/dd/yyyy");
						}
					}
					#endregion

					#region Build Permissions Table
					// =================== Permissions Table ==============================================
					Stage = "Load Permissions Table";
					Word.Table permTable = GetTable("tblPermissions");

					RoleAssignmentCollection roleAssignments = currentWeb.RoleAssignments;
					clientContext.Load(roleAssignments);
					clientContext.ExecuteQuery();

					Stage = "Role Assignments";
					foreach (RoleAssignment assign in roleAssignments)
					{
						clientContext.Load(assign);
						clientContext.ExecuteQuery();

						Stage = "Load Role Principal";
						Principal tmpMember = assign.Member;
						clientContext.Load(tmpMember);
						clientContext.ExecuteQuery();

						Word.Row newRow = permTable.Rows.Add();

						newRow.Cells[1].Range.Text = assign.Member.Title;
						newRow.Cells[2].Range.Text = assign.Member.PrincipalType.ToString();
						newRow.Cells[3].Range.Text = assign.Member.LoginName;

						Stage = "Role Collection";
						RoleDefinitionBindingCollection roles = assign.RoleDefinitionBindings;
						clientContext.Load(roles);
						clientContext.ExecuteQuery();

						Stage = "Role Definitions";
						foreach (RoleDefinition roleDef in roles)
						{
							clientContext.Load(roleDef);
							clientContext.ExecuteQuery();

							switch (roleDef.Name)
							{
								case "Full Control":
									newRow.Cells[4].Range.Text = "X";
									break;
								case "Design":
									newRow.Cells[5].Range.Text = "X";
									break;
								case "Contribute":
									newRow.Cells[6].Range.Text = "X";
									break;
								case "Read":
									newRow.Cells[7].Range.Text = "X";
									break;

							}
						}
					}
					#endregion

					#region Fill Workflow Table
					Stage = "Load Workflow Table";

					Word.Table workflowTable = GetTable("tblWorkflows");

					WorkflowAssociationCollection workflows = currentWeb.WorkflowAssociations;
					
					clientContext.Load(workflows);
					clientContext.ExecuteQuery();

					foreach(WorkflowAssociation workflow in workflows)
					{
						clientContext.Load(workflow);
						clientContext.ExecuteQuery();

						Word.Row newRow = workflowTable.Rows.Add();

						newRow.Cells[1].Range.Text = workflow.Name;
					}

					#endregion

				}
				catch (Exception ex)
				{
					Word.ContentControl urlError = GetContentControl("txtURLError");
					Word.Range rngError = urlError.Range;
					rngError.Text = String.Concat("ERROR at stage ", Stage, ": ", ex.Message);
					rngError.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed;
				}
			}
		}

		private long GetWebSize(Web currentWeb)
		{
			long tmpSize = 0;

			ClientRuntimeContext clientContext = currentWeb.Context;

			clientContext.Load(currentWeb);
			clientContext.ExecuteQuery();

			//ListCollection lists = currentWeb.Lists;
			//clientContext.Load(lists);
			//clientContext.ExecuteQuery();

			var Libraries = clientContext.LoadQuery(currentWeb.Lists.Where(l => l.BaseTemplate == 101));
			clientContext.ExecuteQuery();
			
			foreach(List tmpList in Libraries)
			{
				clientContext.Load(tmpList);
				clientContext.ExecuteQuery();

				string Query = String.Concat(
					"<View>",
						"<ViewFields>",
							"<FieldRef Name='Title'/>",
							"<FieldRef Name='ID'/>",
							"<FieldRef Name='File_x0020_Size'/>",
						"</ViewFields>",
					"</View>");

				CamlQuery oQuery = new CamlQuery();
				ListItemCollection collListItems = tmpList.GetItems(oQuery);
				clientContext.Load(collListItems);

				//FileCollection listFiles = tmpList.RootFolder.Files;
				//clientContext.Load(listFiles,
				//	files => files.Include(file => file.ETag),
				//	files => files.Include(file => file.ListItemAllFields["File_x0020_Size"]));

				clientContext.ExecuteQuery();

				foreach(ListItem oListItem in collListItems)
				{
					clientContext.Load(oListItem);
					clientContext.ExecuteQuery();

					var fileSize = (string)oListItem["File_x0020_Size"];

					int itemFileSize = String.IsNullOrEmpty(fileSize) ? 0 : int.Parse(fileSize);

					tmpSize += itemFileSize;
				}

			}

			return tmpSize; 
		}

		private Word.Table GetTable(string Title)
		{
			Word.Table resultTbl = null;

			foreach(Word.Table tmpTable in this.Tables)
			{
				if(tmpTable.Title == Title)
				{
					resultTbl = tmpTable;
				}
			}

			return resultTbl; 
		}

		private void SetContentControl(string txtTitle, string txtValue)
		{
			Word.ContentControl tmpCtrl = GetContentControl(txtTitle);

			if(tmpCtrl != null)
			{
				tmpCtrl.Range.Text = txtValue;
			}
		}

		private Word.ContentControl GetContentControl(string txtTitle)
		{
			Word.ContentControl result = null; 

			foreach(Word.ContentControl tmpCtl in this.ContentControls)
			{
				if(tmpCtl.Title == txtTitle)
				{
					result = tmpCtl;
				}
			}

			return result;
		}
	}
}
