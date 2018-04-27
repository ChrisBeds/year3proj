namespace WindowsFormsApp2
{
    partial class form1
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
            this.storePromptLabel = new MetroFramework.Controls.MetroLabel();
            this.TabControl2 = new MetroFramework.Controls.MetroTabControl();
            this.storeCodeInput = new MetroFramework.Controls.MetroTextBox();
            this.loginButton = new MetroFramework.Controls.MetroButton();
            this.employeesTableAdapter1 = new WindowsFormsApp2.dashboardDataSet1TableAdapters.EmployeesTableAdapter();
            this.regionLabel = new MetroFramework.Controls.MetroLabel();
            this.regionLoginButton = new MetroFramework.Controls.MetroButton();
            this.regionCodeBox = new MetroFramework.Controls.MetroTextBox();
            this.SuspendLayout();
            // 
            // storePromptLabel
            // 
            this.storePromptLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.storePromptLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.storePromptLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.storePromptLabel.Location = new System.Drawing.Point(415, 89);
            this.storePromptLabel.Name = "storePromptLabel";
            this.storePromptLabel.Size = new System.Drawing.Size(196, 32);
            this.storePromptLabel.TabIndex = 5;
            this.storePromptLabel.Text = "Enter your store code";
            this.storePromptLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.storePromptLabel.Click += new System.EventHandler(this.metroLabel1_Click);
            // 
            // TabControl2
            // 
            this.TabControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TabControl2.FontSize = MetroFramework.MetroTabControlSize.Tall;
            this.TabControl2.FontWeight = MetroFramework.MetroTabControlWeight.Bold;
            this.TabControl2.Location = new System.Drawing.Point(1, 35);
            this.TabControl2.Name = "TabControl2";
            this.TabControl2.Size = new System.Drawing.Size(1401, 481);
            this.TabControl2.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
            this.TabControl2.Style = MetroFramework.MetroColorStyle.Red;
            this.TabControl2.TabIndex = 20;
            this.TabControl2.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.TabControl2.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.TabControl2.UseStyleColors = true;
            // 
            // storeCodeInput
            // 
            this.storeCodeInput.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.storeCodeInput.FontSize = MetroFramework.MetroTextBoxSize.Tall;
            this.storeCodeInput.Location = new System.Drawing.Point(415, 144);
            this.storeCodeInput.Name = "storeCodeInput";
            this.storeCodeInput.Size = new System.Drawing.Size(196, 30);
            this.storeCodeInput.TabIndex = 6;
            this.storeCodeInput.Text = "Store code";
            this.storeCodeInput.Theme = MetroFramework.MetroThemeStyle.Dark;
            // 
            // loginButton
            // 
            this.loginButton.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.loginButton.Location = new System.Drawing.Point(415, 199);
            this.loginButton.Name = "loginButton";
            this.loginButton.Size = new System.Drawing.Size(196, 30);
            this.loginButton.TabIndex = 7;
            this.loginButton.Text = "Login";
            this.loginButton.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.loginButton.Click += new System.EventHandler(this.loginButton_Click);
            // 
            // employeesTableAdapter1
            // 
            this.employeesTableAdapter1.ClearBeforeFill = true;
            // 
            // regionLabel
            // 
            this.regionLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.regionLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.regionLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.regionLabel.Location = new System.Drawing.Point(807, 89);
            this.regionLabel.Name = "regionLabel";
            this.regionLabel.Size = new System.Drawing.Size(213, 32);
            this.regionLabel.TabIndex = 21;
            this.regionLabel.Text = "Enter your region code";
            this.regionLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.regionLabel.Click += new System.EventHandler(this.metroLabel1_Click_1);
            // 
            // regionLoginButton
            // 
            this.regionLoginButton.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.regionLoginButton.Location = new System.Drawing.Point(807, 199);
            this.regionLoginButton.Name = "regionLoginButton";
            this.regionLoginButton.Size = new System.Drawing.Size(213, 30);
            this.regionLoginButton.TabIndex = 22;
            this.regionLoginButton.Text = "Login";
            this.regionLoginButton.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.regionLoginButton.Click += new System.EventHandler(this.regionLoginButton_Click);
            // 
            // regionCodeBox
            // 
            this.regionCodeBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.regionCodeBox.FontSize = MetroFramework.MetroTextBoxSize.Tall;
            this.regionCodeBox.Location = new System.Drawing.Point(807, 144);
            this.regionCodeBox.Name = "regionCodeBox";
            this.regionCodeBox.Size = new System.Drawing.Size(213, 30);
            this.regionCodeBox.TabIndex = 23;
            this.regionCodeBox.Text = "Region Code";
            this.regionCodeBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            // 
            // form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BorderStyle = MetroFramework.Drawing.MetroBorderStyle.FixedSingle;
            this.ClientSize = new System.Drawing.Size(1403, 516);
            this.Controls.Add(this.regionCodeBox);
            this.Controls.Add(this.regionLoginButton);
            this.Controls.Add(this.regionLabel);
            this.Controls.Add(this.loginButton);
            this.Controls.Add(this.storeCodeInput);
            this.Controls.Add(this.storePromptLabel);
            this.Controls.Add(this.TabControl2);
            this.Name = "form1";
            this.Style = MetroFramework.MetroColorStyle.Red;
            this.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.TransparencyKey = System.Drawing.Color.LightGoldenrodYellow;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private dashboardDataSet1TableAdapters.EmployeesTableAdapter employeesTableAdapter1;
        internal MetroFramework.Controls.MetroTextBox regionCodeBox;
        internal MetroFramework.Controls.MetroTextBox storeCodeInput;
        internal MetroFramework.Controls.MetroButton loginButton;
        internal MetroFramework.Controls.MetroLabel storePromptLabel;
        internal MetroFramework.Controls.MetroLabel regionLabel;
        internal MetroFramework.Controls.MetroButton regionLoginButton;
        internal MetroFramework.Controls.MetroTabControl TabControl2;
    }
}

