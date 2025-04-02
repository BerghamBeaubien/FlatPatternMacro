using System;
using System.Windows.Forms;

public partial class FlatPatternPromptForm : Form
{
    public bool IsAutomatic { get; private set; }

    public FlatPatternPromptForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.Text = "Sélection du mode";
        this.Width = 300;
        this.Height = 200;
        this.StartPosition = FormStartPosition.CenterParent;

        Label promptLabel = new Label()
        {
            Text = "Quelle méthode souhaitez-vous utiliser pour choisir la face du déplié ?",
            AutoSize = true,
            Location = new System.Drawing.Point(10, 10),
            Width = 280
        };

        CheckBox manualCheckBox = new CheckBox()
        {
            Text = "Mode manuel (choisir la face)",
            Location = new System.Drawing.Point(20, 50),
            AutoSize = true
        };

        CheckBox automaticCheckBox = new CheckBox()
        {
            Text = "Mode automatique (plus grande/smart)",
            Location = new System.Drawing.Point(20, 80),
            AutoSize = true
        };

        Button confirmButton = new Button()
        {
            Text = "Confirmer",
            Location = new System.Drawing.Point(100, 120),
            DialogResult = DialogResult.OK
        };

        manualCheckBox.CheckedChanged += (s, e) =>
        {
            if (manualCheckBox.Checked)
                automaticCheckBox.Checked = false;
        };

        automaticCheckBox.CheckedChanged += (s, e) =>
        {
            if (automaticCheckBox.Checked)
                manualCheckBox.Checked = false;
        };

        confirmButton.Click += (s, e) =>
        {
            if (!manualCheckBox.Checked && !automaticCheckBox.Checked)
            {
                MessageBox.Show("Veuillez sélectionner une option.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            IsAutomatic = automaticCheckBox.Checked;
            this.DialogResult = DialogResult.OK;
            this.Close();
        };

        this.Controls.Add(promptLabel);
        this.Controls.Add(manualCheckBox);
        this.Controls.Add(automaticCheckBox);
        this.Controls.Add(confirmButton);
    }
}
