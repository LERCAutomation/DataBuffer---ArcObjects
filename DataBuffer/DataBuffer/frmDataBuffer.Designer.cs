namespace DataBuffer
{
    partial class frmDataBuffer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDataBuffer));
            this.lstInput = new System.Windows.Forms.ListBox();
            this.lblInput = new System.Windows.Forms.Label();
            this.chkClearLog = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lstInput
            // 
            this.lstInput.FormattingEnabled = true;
            this.lstInput.Location = new System.Drawing.Point(15, 26);
            this.lstInput.Name = "lstInput";
            this.lstInput.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstInput.Size = new System.Drawing.Size(193, 212);
            this.lstInput.TabIndex = 0;
            // 
            // lblInput
            // 
            this.lblInput.AutoSize = true;
            this.lblInput.Location = new System.Drawing.Point(12, 9);
            this.lblInput.Name = "lblInput";
            this.lblInput.Size = new System.Drawing.Size(68, 13);
            this.lblInput.TabIndex = 1;
            this.lblInput.Text = "Input Layers:";
            // 
            // chkClearLog
            // 
            this.chkClearLog.AutoSize = true;
            this.chkClearLog.Location = new System.Drawing.Point(15, 244);
            this.chkClearLog.Name = "chkClearLog";
            this.chkClearLog.Size = new System.Drawing.Size(89, 17);
            this.chkClearLog.TabIndex = 2;
            this.chkClearLog.Text = "Clear log file?";
            this.chkClearLog.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(153, 263);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(54, 23);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(93, 263);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(54, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmDataBuffer
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(218, 298);
            this.Controls.Add(this.lstInput);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.chkClearLog);
            this.Controls.Add(this.lblInput);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmDataBuffer";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Data Buffer 1.0";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lstInput;
        private System.Windows.Forms.Label lblInput;
        private System.Windows.Forms.CheckBox chkClearLog;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
      
    }
}