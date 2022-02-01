
namespace Journal
{
    partial class Form1
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

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.управлениеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.загрузкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обновитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выходToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Nomer = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FIO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Predmet = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Ocenka = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Data = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.управлениеToolStripMenuItem,
            this.выходToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(804, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // управлениеToolStripMenuItem
            // 
            this.управлениеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.загрузкаToolStripMenuItem,
            this.добавитьToolStripMenuItem,
            this.удалитьToolStripMenuItem,
            this.обновитьToolStripMenuItem});
            this.управлениеToolStripMenuItem.Name = "управлениеToolStripMenuItem";
            this.управлениеToolStripMenuItem.Size = new System.Drawing.Size(85, 20);
            this.управлениеToolStripMenuItem.Text = "Управление";
            // 
            // загрузкаToolStripMenuItem
            // 
            this.загрузкаToolStripMenuItem.Name = "загрузкаToolStripMenuItem";
            this.загрузкаToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.загрузкаToolStripMenuItem.Text = "Загрузка";
            this.загрузкаToolStripMenuItem.Click += new System.EventHandler(this.загрузкаToolStripMenuItem_Click);
            // 
            // добавитьToolStripMenuItem
            // 
            this.добавитьToolStripMenuItem.Name = "добавитьToolStripMenuItem";
            this.добавитьToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.добавитьToolStripMenuItem.Text = "Добавить";
            this.добавитьToolStripMenuItem.Click += new System.EventHandler(this.добавитьToolStripMenuItem_Click);
            // 
            // удалитьToolStripMenuItem
            // 
            this.удалитьToolStripMenuItem.Name = "удалитьToolStripMenuItem";
            this.удалитьToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.удалитьToolStripMenuItem.Text = "Удалить";
            this.удалитьToolStripMenuItem.Click += new System.EventHandler(this.удалитьToolStripMenuItem_Click);
            // 
            // обновитьToolStripMenuItem
            // 
            this.обновитьToolStripMenuItem.Name = "обновитьToolStripMenuItem";
            this.обновитьToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.обновитьToolStripMenuItem.Text = "Обновить";
            this.обновитьToolStripMenuItem.Click += new System.EventHandler(this.обновитьToolStripMenuItem_Click);
            // 
            // выходToolStripMenuItem
            // 
            this.выходToolStripMenuItem.Name = "выходToolStripMenuItem";
            this.выходToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.выходToolStripMenuItem.Text = "Выход";
            this.выходToolStripMenuItem.Click += new System.EventHandler(this.выходToolStripMenuItem_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Nomer,
            this.FIO,
            this.Predmet,
            this.Ocenka,
            this.Data});
            this.dataGridView1.Location = new System.Drawing.Point(12, 27);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(780, 268);
            this.dataGridView1.TabIndex = 1;
            // 
            // Nomer
            // 
            this.Nomer.HeaderText = "Номер_в_журнале";
            this.Nomer.Name = "Nomer";
            // 
            // FIO
            // 
            this.FIO.HeaderText = "ФИО_учащегося";
            this.FIO.Name = "FIO";
            // 
            // Predmet
            // 
            this.Predmet.HeaderText = "Предмет";
            this.Predmet.Name = "Predmet";
            // 
            // Ocenka
            // 
            this.Ocenka.HeaderText = "Оценка";
            this.Ocenka.Name = "Ocenka";
            // 
            // Data
            // 
            this.Data.HeaderText = "Дата";
            this.Data.Name = "Data";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(804, 451);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Журнал";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem управлениеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem загрузкаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem добавитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem удалитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem обновитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem выходToolStripMenuItem;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Nomer;
        private System.Windows.Forms.DataGridViewTextBoxColumn FIO;
        private System.Windows.Forms.DataGridViewTextBoxColumn Predmet;
        private System.Windows.Forms.DataGridViewTextBoxColumn Ocenka;
        private System.Windows.Forms.DataGridViewTextBoxColumn Data;
    }
}

