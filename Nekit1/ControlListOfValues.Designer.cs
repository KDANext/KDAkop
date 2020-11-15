namespace ClassLibraryControlSelected
{
    partial class ControlListOfValues
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
			this.listBoxValues = new System.Windows.Forms.ListBox();
			this.SuspendLayout();
			// 
			// listBoxValues
			// 
			this.listBoxValues.FormattingEnabled = true;
			this.listBoxValues.Location = new System.Drawing.Point(18, 20);
			this.listBoxValues.Name = "listBoxValues";
			this.listBoxValues.Size = new System.Drawing.Size(233, 108);
			this.listBoxValues.TabIndex = 0;
			// 
			// ControlListOfValues
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.listBoxValues);
			this.Name = "ControlListOfValues";
			this.Size = new System.Drawing.Size(254, 192);
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxValues;
    }
}
