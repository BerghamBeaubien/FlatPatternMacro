using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SolidEdgeFramework;
using SolidEdgePart;
using SolidEdgeGeometry;
using System.Windows.Forms.VisualStyles;

public class FaceAndEdgeSelectionForm : Form
{
    private SolidEdgeFramework.Application _application;
    private SolidEdgeFramework.SolidEdgeDocument _document;
    private SolidEdgePart.Model _model;
    private SolidEdgeGeometry.Body _body;

    // UI Controls
    private TabControl _tabControl;
    private TabPage _faceTabPage;
    private TabPage _edgeTabPage;
    private ListBox _faceListBox;
    private ListBox _edgeListBox;
    private Button _selectFaceButton;
    private Button _selectEdgeButton;
    private Button _cancelButton;
    private Button _backButton;
    private Button _completeButton;
    Label edgeInstructionLabel = new Label();

    // Data
    private List<SolidEdgeGeometry.Face> _faces = new List<SolidEdgeGeometry.Face>();
    private List<SolidEdgeGeometry.Edge> _edges = new List<SolidEdgeGeometry.Edge>();
    private SolidEdgeGeometry.Face _selectedFace = null;
    private SolidEdgeGeometry.Edge _selectedEdge = null;

    // Properties
    public SolidEdgeGeometry.Face SelectedFace => _selectedFace;
    public SolidEdgeGeometry.Edge SelectedEdge => _selectedEdge;

    public FaceAndEdgeSelectionForm(SolidEdgeFramework.Application application, SolidEdgeFramework.SolidEdgeDocument document, SolidEdgePart.Model model)
    {
        _application = application;
        _document = document;
        _model = model;
        _body = (Body)model.Body;

        InitializeComponent();
        LoadFaces();
    }

    private void InitializeComponent()
    {
        this.Text = "Sélection de Face et Arête pour Développé";
        this.Size = new Size(600, 500);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.BackColor = Color.White;

        // Create tab control
        _tabControl = new TabControl();
        _tabControl.Dock = DockStyle.Fill;
        _tabControl.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
        _tabControl.ItemSize = new Size(120, 30);

        // Create face tab
        _faceTabPage = new TabPage("Sélection de Face");
        _faceTabPage.Padding = new Padding(15);
        _faceTabPage.BackColor = Color.White;

        // Create edge tab
        _edgeTabPage = new TabPage("Sélection d'Arête");
        _edgeTabPage.Padding = new Padding(15);
        _edgeTabPage.BackColor = Color.White;
        _edgeTabPage.Enabled = false; // Disable until face is selected

        // Setup face listbox
        _faceListBox = new ListBox();
        _faceListBox.Dock = DockStyle.Fill;
        _faceListBox.Font = new Font("Segoe UI", 10F);
        _faceListBox.BorderStyle = BorderStyle.FixedSingle;
        _faceListBox.DisplayMember = "DisplayName";
        _faceListBox.SelectedIndexChanged += FaceListBox_SelectedIndexChanged;
        _faceListBox.DoubleClick += (s, e) => SelectFace();

        // Setup edge listbox
        _edgeListBox = new ListBox();
        _edgeListBox.Dock = DockStyle.Fill;
        _edgeListBox.Font = new Font("Segoe UI", 10F);
        _edgeListBox.BorderStyle = BorderStyle.FixedSingle;
        _edgeListBox.DisplayMember = "DisplayName";
        _edgeListBox.SelectedIndexChanged += EdgeListBox_SelectedIndexChanged;
        _edgeListBox.DoubleClick += (s, e) => SelectFinalEdge();

        // Create face tab instruction label
        Label faceInstructionLabel = new Label();
        faceInstructionLabel.Text = "Veuillez sélectionner une face. Les faces sont triées par surface (de la plus grande à la plus petite).";
        faceInstructionLabel.AutoSize = false;
        faceInstructionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        faceInstructionLabel.Dock = DockStyle.Top;
        faceInstructionLabel.Height = 50;
        faceInstructionLabel.BackColor = Color.FromArgb(240, 240, 240);
        faceInstructionLabel.Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);
        faceInstructionLabel.Padding = new Padding(10, 0, 0, 0);
        faceInstructionLabel.Margin = new Padding(0, 0, 0, 10);

        // Create edge tab instruction label
        edgeInstructionLabel.Text = "Arêtes de la face XXX : Veuillez en sélectionner une seule.";
        edgeInstructionLabel.AutoSize = false;
        edgeInstructionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        edgeInstructionLabel.Dock = DockStyle.Top;
        edgeInstructionLabel.Height = 50;
        edgeInstructionLabel.BackColor = Color.FromArgb(240, 240, 240);
        edgeInstructionLabel.Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);
        edgeInstructionLabel.Padding = new Padding(10, 0, 0, 0);
        edgeInstructionLabel.Margin = new Padding(0, 0, 0, 10);

        // Create buttons for face tab
        _selectFaceButton = new Button();
        _selectFaceButton.Text = "Suivant";
        _selectFaceButton.Enabled = false;
        _selectFaceButton.Size = new Size(100, 35);
        _selectFaceButton.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
        _selectFaceButton.BackColor = Color.FromArgb(0, 114, 198);
        _selectFaceButton.ForeColor = Color.White;
        _selectFaceButton.FlatStyle = FlatStyle.Flat;
        _selectFaceButton.FlatAppearance.BorderSize = 0;
        _selectFaceButton.Click += (s, e) => SelectFace();

        _cancelButton = new Button();
        _cancelButton.Text = "Annuler";
        _cancelButton.Size = new Size(100, 35);
        _cancelButton.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
        _cancelButton.BackColor = Color.FromArgb(240, 240, 240);
        _cancelButton.FlatStyle = FlatStyle.Flat;
        _cancelButton.FlatAppearance.BorderSize = 1;
        _cancelButton.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

        // Create buttons for edge tab
        _backButton = new Button();
        _backButton.Text = "Retour";
        _backButton.Size = new Size(100, 35);
        _backButton.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
        _backButton.BackColor = Color.FromArgb(240, 240, 240);
        _backButton.FlatStyle = FlatStyle.Flat;
        _backButton.FlatAppearance.BorderSize = 1;
        _backButton.Click += (s, e) => GoBackToFaceSelection();

        _selectEdgeButton = new Button();
        _selectEdgeButton.Text = "Sélectionner";
        _selectEdgeButton.Enabled = false;
        _selectEdgeButton.Size = new Size(100, 35);
        _selectEdgeButton.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
        _selectEdgeButton.BackColor = Color.FromArgb(0, 114, 198);
        _selectEdgeButton.ForeColor = Color.White;
        _selectEdgeButton.FlatStyle = FlatStyle.Flat;
        _selectEdgeButton.FlatAppearance.BorderSize = 0;
        _selectEdgeButton.Click += (s, e) => SelectFinalEdge();

        _completeButton = new Button();
        _completeButton.Text = "Annuler";
        _completeButton.Size = new Size(100, 35);
        _completeButton.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
        _completeButton.BackColor = Color.FromArgb(240, 240, 240);
        _completeButton.FlatStyle = FlatStyle.Flat;
        _completeButton.FlatAppearance.BorderSize = 1;
        _completeButton.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

        // Create button panels with flow layout
        Panel faceButtonPanel = new Panel();
        faceButtonPanel.Height = 60;
        faceButtonPanel.Dock = DockStyle.Bottom;
        faceButtonPanel.Padding = new Padding(15, 10, 15, 10);
        faceButtonPanel.BackColor = Color.White;

        FlowLayoutPanel faceFlowPanel = new FlowLayoutPanel();
        faceFlowPanel.Dock = DockStyle.Fill;
        faceFlowPanel.FlowDirection = FlowDirection.RightToLeft;
        faceFlowPanel.Controls.Add(_cancelButton);
        faceFlowPanel.Controls.Add(_selectFaceButton);
        faceFlowPanel.Controls.Add(new Panel() { Width = 10 }); // Spacer
        faceButtonPanel.Controls.Add(faceFlowPanel);

        Panel edgeButtonPanel = new Panel();
        edgeButtonPanel.Height = 60;
        edgeButtonPanel.Dock = DockStyle.Bottom;
        edgeButtonPanel.Padding = new Padding(15, 10, 15, 10);
        edgeButtonPanel.BackColor = Color.White;

        FlowLayoutPanel edgeFlowPanel = new FlowLayoutPanel();
        edgeFlowPanel.Dock = DockStyle.Fill;
        edgeFlowPanel.FlowDirection = FlowDirection.RightToLeft;
        edgeFlowPanel.Controls.Add(_completeButton);
        edgeFlowPanel.Controls.Add(_selectEdgeButton);
        edgeFlowPanel.Controls.Add(new Panel() { Width = 10 }); // Spacer
        edgeFlowPanel.Controls.Add(_backButton);
        edgeButtonPanel.Controls.Add(edgeFlowPanel);

        // Create content panels
        Panel faceContentPanel = new Panel();
        faceContentPanel.Dock = DockStyle.Fill;
        faceContentPanel.Padding = new Padding(0, 0, 0, 10);
        faceContentPanel.Controls.Add(_faceListBox);
        faceContentPanel.Controls.Add(faceInstructionLabel);

        Panel edgeContentPanel = new Panel();
        edgeContentPanel.Dock = DockStyle.Fill;
        edgeContentPanel.Padding = new Padding(0, 0, 0, 10);
        edgeContentPanel.Controls.Add(_edgeListBox);
        edgeContentPanel.Controls.Add(edgeInstructionLabel);

        // Add controls to tabs
        _faceTabPage.Controls.Add(faceContentPanel);
        _faceTabPage.Controls.Add(faceButtonPanel);

        _edgeTabPage.Controls.Add(edgeContentPanel);
        _edgeTabPage.Controls.Add(edgeButtonPanel);

        // Add tabs to tab control
        _tabControl.TabPages.Add(_faceTabPage);
        _tabControl.TabPages.Add(_edgeTabPage);

        // Add tab control to form
        this.Controls.Add(_tabControl);
    }

    #region Face Selection

    private class FaceItem
    {
        public SolidEdgeGeometry.Face Face { get; set; }
        public string DisplayName { get; set; }
        public double AreaInSqMm { get; set; }
    }

    private void LoadFaces()
    {
        try
        {
            SolidEdgeGeometry.Faces faces = (SolidEdgeGeometry.Faces)_body.Faces[FeatureTopologyQueryTypeConstants.igQueryPlane];
            List<FaceItem> faceItems = new List<FaceItem>();

            for (int i = 1; i <= faces.Count; i++)
            {
                SolidEdgeGeometry.Face face = (SolidEdgeGeometry.Face)faces.Item(i);
                _faces.Add(face);

                // Get face area in mm²
                double areaInSqMm = face.Area * 1550.0031;

                FaceItem item = new FaceItem
                {
                    Face = face,
                    AreaInSqMm = areaInSqMm,
                    DisplayName = $"Face ID: {face.ID} - Area: {areaInSqMm:F2} mm²"
                };

                faceItems.Add(item);
            }

            // Sort faces by area (largest first)
            faceItems.Sort((a, b) => b.AreaInSqMm.CompareTo(a.AreaInSqMm));

            _faceListBox.DataSource = faceItems;

            if (faceItems.Count > 0)
            {
                _faceListBox.SelectedIndex = 0;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading faces: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void FaceListBox_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (_faceListBox.SelectedItem != null)
        {
            _selectFaceButton.Enabled = true;

            // Highlight the selected face in Solid Edge
            FaceItem selectedItem = (FaceItem)_faceListBox.SelectedItem;
            HighlightFace(selectedItem.Face);
        }
        else
        {
            _selectFaceButton.Enabled = false;
        }
    }

    private void HighlightFace(SolidEdgeGeometry.Face face)
    {
        try
        {
            // Clear any existing selection
            _application.StartCommand(SolidEdgeCommandConstants.sePartSelectCommand);

            // Select the face
            SolidEdgeFramework.SelectSet selectSet = _application.ActiveSelectSet;
            selectSet.RemoveAll();
            selectSet.Add(face);

            // Ensure the face is visible in the viewport
            _application.DoIdle();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error highlighting face: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SelectFace()
    {
        if (_faceListBox.SelectedItem != null)
        {
            FaceItem selectedItem = (FaceItem)_faceListBox.SelectedItem;
            _selectedFace = selectedItem.Face;

            // Load edges for the selected face
            LoadEdges(_selectedFace);
            this.edgeInstructionLabel.Text = $"Arêtes de la Face {_selectedFace.ID} : Veuillez en choisir une seule.";

            // Switch to edge tab
            _edgeTabPage.Enabled = true;
            _tabControl.SelectedTab = _edgeTabPage;
        }
    }

    #endregion

    #region Edge Selection

    private class EdgeItem
    {
        public SolidEdgeGeometry.Edge Edge { get; set; }
        public string DisplayName { get; set; }
        public double LengthInPo { get; set; }
    }

    private void LoadEdges(SolidEdgeGeometry.Face face)
    {
        try
        {
            _edges.Clear();
            SolidEdgeGeometry.Edges edges = (SolidEdgeGeometry.Edges)face.Edges;
            List<EdgeItem> edgeItems = new List<EdgeItem>();

            this._edgeTabPage.Text = $"Arêtes de la Face {face.ID}";

            for (int i = 1; i <= edges.Count; i++)
            {
                try
                {
                    SolidEdgeGeometry.Edge edge = (SolidEdgeGeometry.Edge)edges.Item(i);
                    _edges.Add(edge);
                    EdgeItem item = new EdgeItem();

                    // Get vertices of the edge
                    Vertex startVertex = (Vertex)edge.StartVertex;
                    Vertex endVertex = (Vertex)edge.EndVertex;

                    // Skip edge if vertices are null
                    if (startVertex == null || endVertex == null)
                    {
                        item.Edge = edge;
                        item.LengthInPo = 0;
                        item.DisplayName = $"Edge ID: {edge.ID} - Unable to calculate length";
                        edgeItems.Add(item);
                        continue;
                    }

                    // Extract coordinates
                    Array startPointArray = Array.CreateInstance(typeof(double), 3);
                    Array endPointArray = Array.CreateInstance(typeof(double), 3);
                    startVertex.GetPointData(ref startPointArray);
                    endVertex.GetPointData(ref endPointArray);

                    double[] startPoint = new double[3];
                    double[] endPoint = new double[3];

                    for (int j = 0; j < 3; j++)
                    {
                        startPoint[j] = (double)startPointArray.GetValue(j);
                        endPoint[j] = (double)endPointArray.GetValue(j);
                    }

                    // Calculate edge length (Euclidean distance)
                    double length = Math.Sqrt(
                        Math.Pow(endPoint[0] - startPoint[0], 2) +
                        Math.Pow(endPoint[1] - startPoint[1], 2) +
                        Math.Pow(endPoint[2] - startPoint[2], 2)
                    );

                    // Convert to inches (adjust conversion factor if needed)
                    double lengthInPo = length * 39.3701;

                    item.Edge = edge;
                    item.LengthInPo = lengthInPo;
                    item.DisplayName = $"Edge ID: {edge.ID} - Length: {lengthInPo:F2} po";


                    edgeItems.Add(item);
                }
                catch (Exception edgeEx)
                {
                    // Handle individual edge errors gracefully
                    try
                    {
                        SolidEdgeGeometry.Edge edge = (SolidEdgeGeometry.Edge)edges.Item(i);
                        _edges.Add(edge);

                        EdgeItem item = new EdgeItem
                        {
                            Edge = edge,
                            LengthInPo = 0,
                            DisplayName = $"Edge ID: {edge.ID} - Error: {edgeEx.Message}"
                        };
                        edgeItems.Add(item);
                    }
                    catch
                    {
                        // If we can't even get the edge, just add a placeholder
                        Console.WriteLine($"Error processing edge at index {i}: {edgeEx.Message}");
                    }
                }
            }

            // Sort edges by length (longest first), put error edges at the end
            edgeItems.Sort((a, b) =>
            {
                // If either item has an error (length 0), sort it to the end
                if (a.LengthInPo == 0 && b.LengthInPo == 0) return 0;
                if (a.LengthInPo == 0) return 1;
                if (b.LengthInPo == 0) return -1;

                // Otherwise sort by length (longest first)
                return b.LengthInPo.CompareTo(a.LengthInPo);
            });

            // Update display
            _edgeListBox.DataSource = null;
            _edgeListBox.DisplayMember = "DisplayName";
            _edgeListBox.ValueMember = "Edge";
            _edgeListBox.DataSource = edgeItems;

            if (edgeItems.Count > 0)
            {
                _edgeListBox.SelectedIndex = 0;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading edges: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void EdgeListBox_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (_edgeListBox.SelectedItem != null)
        {
            _selectEdgeButton.Enabled = true;

            // Highlight the selected edge in Solid Edge
            EdgeItem selectedItem = (EdgeItem)_edgeListBox.SelectedItem;
            HighlightEdge(selectedItem.Edge);
        }
        else
        {
            _selectEdgeButton.Enabled = false;
        }
    }

    private void HighlightEdge(SolidEdgeGeometry.Edge edge)
    {
        try
        {
            // Clear any existing selection
            _application.StartCommand(SolidEdgeCommandConstants.sePartSelectCommand);

            // Select the edge
            SolidEdgeFramework.SelectSet selectSet = _application.ActiveSelectSet;
            selectSet.RemoveAll();
            selectSet.Add(edge);

            // Ensure the edge is visible in the viewport
            _application.DoIdle();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error highlighting edge: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SelectFinalEdge()
    {
        if (_edgeListBox.SelectedItem != null)
        {
            EdgeItem selectedItem = (EdgeItem)_edgeListBox.SelectedItem;
            _selectedEdge = selectedItem.Edge;
            this.DialogResult = DialogResult.OK;
        }
    }

    private void GoBackToFaceSelection()
    {
        // Switch back to face tab
        _tabControl.SelectedTab = _faceTabPage;
    }

    #endregion
}