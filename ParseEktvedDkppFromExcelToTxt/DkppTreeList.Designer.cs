namespace ParserExcelUKTVEDtoBunariList
{
    partial class DkppTreeList
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
            this.treeList = new DevExpress.XtraTreeList.TreeList();
            this.numberCol = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.decreeDateCol = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            ((System.ComponentModel.ISupportInitialize)(this.treeList)).BeginInit();
            this.SuspendLayout();
            // 
            // treeList
            // 
            this.treeList.Appearance.FooterPanel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.treeList.Appearance.FooterPanel.Options.UseFont = true;
            this.treeList.Appearance.FooterPanel.Options.UseTextOptions = true;
            this.treeList.Appearance.FooterPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.treeList.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
            this.numberCol,
            this.decreeDateCol});
            this.treeList.Cursor = System.Windows.Forms.Cursors.Default;
            this.treeList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeList.KeyFieldName = "Id";
            this.treeList.Location = new System.Drawing.Point(0, 0);
            this.treeList.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.treeList.Name = "treeList";
            this.treeList.OptionsBehavior.EnableFiltering = true;
            this.treeList.OptionsFind.AlwaysVisible = true;
            this.treeList.OptionsView.ShowAutoFilterRow = true;
            this.treeList.Size = new System.Drawing.Size(1318, 469);
            this.treeList.TabIndex = 2;
            // 
            // numberCol
            // 
            this.numberCol.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.numberCol.AppearanceHeader.ForeColor = System.Drawing.Color.Navy;
            this.numberCol.AppearanceHeader.Options.UseFont = true;
            this.numberCol.AppearanceHeader.Options.UseForeColor = true;
            this.numberCol.AppearanceHeader.Options.UseTextOptions = true;
            this.numberCol.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.numberCol.Caption = "Код товару";
            this.numberCol.FieldName = "Code";
            this.numberCol.MinWidth = 34;
            this.numberCol.Name = "numberCol";
            this.numberCol.OptionsColumn.AllowEdit = false;
            this.numberCol.OptionsColumn.AllowFocus = false;
            this.numberCol.OptionsFilter.AutoFilterCondition = DevExpress.XtraTreeList.Columns.AutoFilterCondition.Contains;
            this.numberCol.Visible = true;
            this.numberCol.VisibleIndex = 0;
            // 
            // decreeDateCol
            // 
            this.decreeDateCol.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.decreeDateCol.AppearanceHeader.ForeColor = System.Drawing.Color.Navy;
            this.decreeDateCol.AppearanceHeader.Options.UseFont = true;
            this.decreeDateCol.AppearanceHeader.Options.UseForeColor = true;
            this.decreeDateCol.AppearanceHeader.Options.UseTextOptions = true;
            this.decreeDateCol.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.decreeDateCol.Caption = "Опис";
            this.decreeDateCol.FieldName = "Name";
            this.decreeDateCol.Name = "decreeDateCol";
            this.decreeDateCol.OptionsColumn.AllowEdit = false;
            this.decreeDateCol.OptionsColumn.AllowFocus = false;
            this.decreeDateCol.OptionsFilter.AutoFilterCondition = DevExpress.XtraTreeList.Columns.AutoFilterCondition.Contains;
            this.decreeDateCol.Visible = true;
            this.decreeDateCol.VisibleIndex = 1;
            // 
            // DkppTreeList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1318, 469);
            this.Controls.Add(this.treeList);
            this.Name = "DkppTreeList";
            this.Text = "DkppTreeList";
            ((System.ComponentModel.ISupportInitialize)(this.treeList)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraTreeList.TreeList treeList;
        private DevExpress.XtraTreeList.Columns.TreeListColumn numberCol;
        private DevExpress.XtraTreeList.Columns.TreeListColumn decreeDateCol;
    }
}