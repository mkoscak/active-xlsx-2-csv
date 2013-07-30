namespace ActiveConverter
{
    partial class FrmSetCatIds
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSetCatIds));
            this.btnOk = new System.Windows.Forms.Button();
            this.gridData = new ActiveConverter.DoubleBufferedGrid();
            this.chkUseDefault = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.gridData)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOk
            // 
            this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnOk.Location = new System.Drawing.Point(12, 496);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(184, 44);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // gridData
            // 
            this.gridData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.gridData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.gridData.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.gridData.BackgroundColor = System.Drawing.SystemColors.ControlLight;
            this.gridData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridData.Location = new System.Drawing.Point(12, 12);
            this.gridData.Name = "gridData";
            this.gridData.RowHeadersVisible = false;
            this.gridData.Size = new System.Drawing.Size(519, 478);
            this.gridData.TabIndex = 2;
            // 
            // chkUseDefault
            // 
            this.chkUseDefault.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkUseDefault.AutoSize = true;
            this.chkUseDefault.Location = new System.Drawing.Point(202, 511);
            this.chkUseDefault.Name = "chkUseDefault";
            this.chkUseDefault.Size = new System.Drawing.Size(230, 17);
            this.chkUseDefault.TabIndex = 3;
            this.chkUseDefault.Text = "Use common category id from main window";
            this.chkUseDefault.UseVisualStyleBackColor = true;
            // 
            // FrmSetCatIds
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(543, 552);
            this.Controls.Add(this.chkUseDefault);
            this.Controls.Add(this.gridData);
            this.Controls.Add(this.btnOk);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmSetCatIds";
            this.Text = "Set category ids";
            ((System.ComponentModel.ISupportInitialize)(this.gridData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOk;
        private DoubleBufferedGrid gridData;
        private System.Windows.Forms.CheckBox chkUseDefault;

    }
}