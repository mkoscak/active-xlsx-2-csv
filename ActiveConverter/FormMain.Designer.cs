namespace ActiveConverter
{
    partial class frmMain
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.btnTest = new System.Windows.Forms.Button();
            this.lblCatIds = new System.Windows.Forms.Label();
            this.txtCatIds = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnTest
            // 
            this.btnTest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnTest.Location = new System.Drawing.Point(12, 61);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(184, 62);
            this.btnTest.TabIndex = 0;
            this.btnTest.Text = "Process";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // lblCatIds
            // 
            this.lblCatIds.AutoSize = true;
            this.lblCatIds.Location = new System.Drawing.Point(12, 9);
            this.lblCatIds.Name = "lblCatIds";
            this.lblCatIds.Size = new System.Drawing.Size(65, 13);
            this.lblCatIds.TabIndex = 1;
            this.lblCatIds.Text = "Category ids";
            // 
            // txtCatIds
            // 
            this.txtCatIds.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCatIds.Location = new System.Drawing.Point(83, 6);
            this.txtCatIds.Name = "txtCatIds";
            this.txtCatIds.Size = new System.Drawing.Size(274, 20);
            this.txtCatIds.TabIndex = 2;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(369, 135);
            this.Controls.Add(this.txtCatIds);
            this.Controls.Add(this.lblCatIds);
            this.Controls.Add(this.btnTest);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMain";
            this.Text = "Active Converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Label lblCatIds;
        private System.Windows.Forms.TextBox txtCatIds;
    }
}

