using System;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using Application_Cyrell.Utils;

public partial class FlatPatternPromptForm : Form
{
    public bool IsAutomatic { get; private set; }
    public bool CloseDocument { get; private set; }

    private BouttonToggle modeToggle;
    private Label manualLabel;
    private Label automaticLabel;
    private Label titleLabel;
    private Label descriptionLabel;
    private CheckBox closeDocumentCheckBox;
    private Button confirmButton;
    private Panel separatorPanel;
    private Label hoverTextLabel;
    private ToolTip manTip;
    private ToolTip autoTip;
    private bool _closeOption = true;

    public FlatPatternPromptForm()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        // Form properties
        this.Text = "Selection Mode";
        this.Width = 400;
        this.Height = _closeOption ? 280 : 220; // Adjust height based on closeOption
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;
        this.Font = new Font("Segoe UI", 9F);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // Title label
        titleLabel = new Label()
        {
            Text = "Flat Pattern Settings",
            Font = new Font("Segoe UI", 14F, FontStyle.Bold),
            ForeColor = Color.MediumSlateBlue,
            Location = new Point(20, 15),
            AutoSize = true
        };

        // Description label
        descriptionLabel = new Label()
        {
            Text = "Choose the method with which you want the face and the edge of the flat pattern to be determined",
            Font = new Font("Segoe UI", 9F),
            Location = new Point(20, 50),
            Size = new Size(350, 40)
        };

        // Mode selection section
        hoverTextLabel = new Label()
        {
            Text = "",
            Font = new Font("Segoe UI", 9F),
            Location = new Point(50, 120),
            AutoSize = true
        };

        manualLabel = new Label()
        {
            Text = "Manual",
            Font = new Font("Segoe UI", 9F),
            Location = new Point(50, 100),
            AutoSize = true,
        };

        modeToggle = new BouttonToggle()
        {
            Location = new Point(160, 97),
            Size = new Size(55, 25),
            OnBackColor = Color.MediumSlateBlue,
            OffBackColor = Color.MediumSlateBlue,
            Checked = false // Default to manual mode (off position)
        };

        automaticLabel = new Label()
        {
            Text = "Automatic",
            Font = new Font("Segoe UI", 9F),
            Location = new Point(230, 100),
            AutoSize = true
        };

        manTip = new ToolTip()
        {
            IsBalloon = true,
            ShowAlways = true
        };

        autoTip = new ToolTip()
        {
            IsBalloon = true,
            ShowAlways = true
        };

        manTip.SetToolTip(manualLabel, "The program will ask you to choose face and edge.");
        autoTip.SetToolTip(automaticLabel, "The program will choose the farthest face from the center of coordinates system.");

        // Only show these controls if closeOption is true
        if (_closeOption)
        {
            // Separator
            separatorPanel = new Panel()
            {
                Location = new Point(20, 140),
                Size = new Size(350, 1),
                BackColor = Color.LightGray
            };

            // Close document checkbox
            closeDocumentCheckBox = new CheckBox()
            {
                Text = "Close the document after the operation",
                Location = new Point(50, 160),
                Size = new Size(280, 24),
                Font = new Font("Segoe UI", 9F)
            };

            this.Controls.Add(separatorPanel);
            this.Controls.Add(closeDocumentCheckBox);
        }

        // Confirm button position depends on closeOption
        int confirmButtonY = _closeOption ? 200 : 140;

        confirmButton = new Button()
        {
            Text = "Confirm",
            Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
            Size = new Size(100, 35),
            Location = new Point(150, confirmButtonY),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.MediumSlateBlue,
            ForeColor = Color.White,
            DialogResult = DialogResult.OK
        };
        confirmButton.FlatAppearance.BorderSize = 0;

        // Event handlers
        confirmButton.Click += (s, e) =>
        {
            IsAutomatic = modeToggle.Checked;
            CloseDocument = _closeOption && closeDocumentCheckBox.Checked;
            this.DialogResult = DialogResult.OK;
            this.Close();
        };

        // Add common controls to form
        this.Controls.Add(titleLabel);
        this.Controls.Add(descriptionLabel);
        this.Controls.Add(manualLabel);
        this.Controls.Add(modeToggle);
        this.Controls.Add(automaticLabel);
        this.Controls.Add(hoverTextLabel);
        this.Controls.Add(confirmButton);
    }
}