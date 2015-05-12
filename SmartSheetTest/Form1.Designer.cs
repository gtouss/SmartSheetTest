namespace SmartSheetTest
{
    partial class Form1
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
            this.ListSheets = new System.Windows.Forms.Button();
            this.Output = new System.Windows.Forms.TextBox();
            this.UpdateSheet = new System.Windows.Forms.Button();
            this.ListColumns = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.syncButton = new System.Windows.Forms.Button();
            this.updateDiscusssions = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // ListSheets
            // 
            this.ListSheets.Location = new System.Drawing.Point(12, 12);
            this.ListSheets.Name = "ListSheets";
            this.ListSheets.Size = new System.Drawing.Size(87, 23);
            this.ListSheets.TabIndex = 0;
            this.ListSheets.Text = "ListSheets";
            this.ListSheets.UseVisualStyleBackColor = true;
            this.ListSheets.Click += new System.EventHandler(this.button1_Click);
            // 
            // Output
            // 
            this.Output.Location = new System.Drawing.Point(12, 103);
            this.Output.Multiline = true;
            this.Output.Name = "Output";
            this.Output.Size = new System.Drawing.Size(392, 236);
            this.Output.TabIndex = 1;
            // 
            // UpdateSheet
            // 
            this.UpdateSheet.Location = new System.Drawing.Point(12, 74);
            this.UpdateSheet.Name = "UpdateSheet";
            this.UpdateSheet.Size = new System.Drawing.Size(87, 23);
            this.UpdateSheet.TabIndex = 2;
            this.UpdateSheet.Text = "Update Sheet";
            this.UpdateSheet.UseVisualStyleBackColor = true;
            this.UpdateSheet.Click += new System.EventHandler(this.UpdateSheet_Click);
            // 
            // ListColumns
            // 
            this.ListColumns.Location = new System.Drawing.Point(12, 45);
            this.ListColumns.Name = "ListColumns";
            this.ListColumns.Size = new System.Drawing.Size(87, 23);
            this.ListColumns.TabIndex = 3;
            this.ListColumns.Text = "List Columns";
            this.ListColumns.UseVisualStyleBackColor = true;
            this.ListColumns.Click += new System.EventHandler(this.ListColumns_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(410, 103);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(240, 150);
            this.dataGridView1.TabIndex = 4;
            // 
            // syncButton
            // 
            this.syncButton.Location = new System.Drawing.Point(121, 11);
            this.syncButton.Name = "syncButton";
            this.syncButton.Size = new System.Drawing.Size(136, 23);
            this.syncButton.TabIndex = 5;
            this.syncButton.Text = "Sync";
            this.syncButton.UseVisualStyleBackColor = true;
            this.syncButton.Click += new System.EventHandler(this.syncButton_Click);
            // 
            // updateDiscusssions
            // 
            this.updateDiscusssions.Location = new System.Drawing.Point(121, 44);
            this.updateDiscusssions.Name = "updateDiscusssions";
            this.updateDiscusssions.Size = new System.Drawing.Size(136, 23);
            this.updateDiscusssions.TabIndex = 6;
            this.updateDiscusssions.Text = "Update Discussions";
            this.updateDiscusssions.UseVisualStyleBackColor = true;
            this.updateDiscusssions.Click += new System.EventHandler(this.updateDiscusssions_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(910, 350);
            this.Controls.Add(this.updateDiscusssions);
            this.Controls.Add(this.syncButton);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.ListColumns);
            this.Controls.Add(this.UpdateSheet);
            this.Controls.Add(this.Output);
            this.Controls.Add(this.ListSheets);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ListSheets;
        private System.Windows.Forms.TextBox Output;
        private System.Windows.Forms.Button UpdateSheet;
        private System.Windows.Forms.Button ListColumns;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button syncButton;
        private System.Windows.Forms.Button updateDiscusssions;
    }
}

