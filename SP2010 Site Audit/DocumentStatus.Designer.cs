namespace SP2010_Site_Audit
{
	partial class DocumentStatus
	{
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tvStructure = new System.Windows.Forms.TreeView();
			this.lblTitle = new System.Windows.Forms.Label();
			this.chkStatus = new System.Windows.Forms.CheckedListBox();
			this.lblSubWebs = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// tvStructure
			// 
			this.tvStructure.Location = new System.Drawing.Point(3, 528);
			this.tvStructure.Name = "tvStructure";
			this.tvStructure.Size = new System.Drawing.Size(386, 606);
			this.tvStructure.TabIndex = 0;
			// 
			// lblTitle
			// 
			this.lblTitle.AutoSize = true;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitle.Location = new System.Drawing.Point(21, 23);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(229, 25);
			this.lblTitle.TabIndex = 1;
			this.lblTitle.Text = "SharePoint 2010 Audit";
			// 
			// chkStatus
			// 
			this.chkStatus.FormattingEnabled = true;
			this.chkStatus.Items.AddRange(new object[] {
            "Header",
            "Child Webs",
            "Content Objects",
            "Pages",
            "Permissions",
            "Workflows",
            "Custom Solutions"});
			this.chkStatus.Location = new System.Drawing.Point(26, 161);
			this.chkStatus.Name = "chkStatus";
			this.chkStatus.Size = new System.Drawing.Size(313, 140);
			this.chkStatus.TabIndex = 2;
			// 
			// lblSubWebs
			// 
			this.lblSubWebs.AutoSize = true;
			this.lblSubWebs.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblSubWebs.Location = new System.Drawing.Point(3, 507);
			this.lblSubWebs.Name = "lblSubWebs";
			this.lblSubWebs.Size = new System.Drawing.Size(159, 18);
			this.lblSubWebs.TabIndex = 3;
			this.lblSubWebs.Text = "Child Web Structure";
			// 
			// DocumentStatus
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.lblSubWebs);
			this.Controls.Add(this.chkStatus);
			this.Controls.Add(this.lblTitle);
			this.Controls.Add(this.tvStructure);
			this.Name = "DocumentStatus";
			this.Size = new System.Drawing.Size(392, 1137);
			this.Load += new System.EventHandler(this.DocumentStatus_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.TreeView tvStructure;
		private System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.CheckedListBox chkStatus;
		private System.Windows.Forms.Label lblSubWebs;
	}
}
