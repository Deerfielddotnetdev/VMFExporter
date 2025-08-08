// MailFlowEmlExporter - Updated with Unlock Code Prompt
// Kevin Fortune II & GPT

using System;
using System.Windows.Forms;

namespace MailFlowEmlExporter
{
    static class Program
    {
        // 💡 CHANGE THIS CODE PER BUILD IF DESIRED
        private const string UnlockCode = "letmein123";

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (var unlockForm = new UnlockForm())
            {
                if (unlockForm.ShowDialog() != DialogResult.OK ||
                    unlockForm.UnlockCode != UnlockCode)
                {
                    MessageBox.Show("Invalid unlock code. Exiting.");
                    return;
                }
            }

            // ✅ Unlock successful — proceed with normal execution
            Console.WriteLine("Unlock successful. Starting export...");

            // Your existing CLI logic starts here...
            // Example: RunExport(args);
        }
    }

    public partial class UnlockForm : Form
    {
        private TextBox txtCode;
        private Button btnOK;
        private Button btnCancel;

        public UnlockForm()
        {
            InitializeComponent();
        }

        public string UnlockCode => txtCode.Text;

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void InitializeComponent()
        {
            this.txtCode = new TextBox();
            this.btnOK = new Button();
            this.btnCancel = new Button();
            this.SuspendLayout();

            // txtCode
            this.txtCode.Location = new System.Drawing.Point(12, 12);
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(260, 23);
            this.txtCode.TabIndex = 0;
            this.txtCode.UseSystemPasswordChar = true;

            // btnOK
            this.btnOK.Location = new System.Drawing.Point(116, 50);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new EventHandler(this.btnOK_Click);

            // btnCancel
            this.btnCancel.Location = new System.Drawing.Point(197, 50);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new EventHandler(this.btnCancel_Click);

            // UnlockForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 85);
            this.Controls.Add(this.txtCode);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UnlockForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Enter Unlock Code";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
