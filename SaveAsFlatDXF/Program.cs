using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Linq;
using System.Windows.Forms.VisualStyles;

namespace SaveAsFlatDXF
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            dynamic dynamicDoc = null;

            SolidEdgePart.FlatPatternModel flatPatternModel = null;
            SolidEdgeGeometry.Body body = null;
            SolidEdgeGeometry.Faces faces = null;
            SolidEdgeGeometry.Face face = null;
            SolidEdgeGeometry.Edges edges = null;
            SolidEdgeGeometry.Edge edge = null;
            SolidEdgeGeometry.Vertex vertex = null;

            SolidEdgePart.Models models = null;
            SolidEdgePart.Model model = null;
            bool useForm = true;

            using (var form = new FlatPatternPromptForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    if (form.IsAutomatic)
                    {
                        useForm = false;
                    }
                }
                else if (form.DialogResult != DialogResult.OK)
                {
                    Console.WriteLine("Program stopped by user");
                    System.Environment.Exit(0);
                }

                try
                {
                    // Register with OLE to handle concurrency issues on the current thread.
                    SolidEdgeCommunity.OleMessageFilter.Register();

                    // Connect to Solid Edge,
                    application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);

                    // Get a reference to the active document.
                    dynamicDoc = (SolidEdgeDocument)application.ActiveDocument;

                    // Get a reference to the active document.
                    if (dynamicDoc is SolidEdgePart.PartDocument)
                    {
                        dynamicDoc = (SolidEdgeDocument)application.GetActiveDocument<SolidEdgePart.PartDocument>(false);
                        Console.WriteLine("PartDocument");
                    }
                    else if (dynamicDoc is SolidEdgePart.SheetMetalDocument)
                    {
                        dynamicDoc = (SolidEdgeDocument)application.GetActiveDocument<SolidEdgePart.SheetMetalDocument>(false);
                        Console.WriteLine("SheetMetalDocument");
                    }
                    else
                    {
                        MessageBox.Show("Active document is not a PartDocument or SheetMetalDocument.", "Erreur");
                        return;
                    }

                    //Transform the BodyFeature so that it can be used in the FlatPattern
                    try
                    {
                        models = dynamicDoc.Models;
                        model = models.Item(1);

                        Console.WriteLine($"Features number: {model.Features.Count.ToString()}");

                        if (model.ConvToSMs.Count == 0 && model.Features.Count == 1)
                        {
                            model.HealAndOptimizeBody(true, true);
                            body = (Body)model.Body;
                            faces = (Faces)body.Faces[FeatureTopologyQueryTypeConstants.igQueryPlane];
                            face = (Face)faces.Item(1);
                            for (int i = 2; i <= faces.Count; i++)
                            {
                                SolidEdgeGeometry.Face currentFace = (Face)faces.Item(i);

                                if (currentFace.Area > face.Area) face = currentFace;
                            }
                            Console.WriteLine($"User selected Face {face.ID} - Area: {face.Area * 1550.0031} po²");

                            edges = (SolidEdgeGeometry.Edges)face.Edges;
                            Array edgesArray = Array.CreateInstance(typeof(object), edges.Count);

                            for (int i = 1; i <= edges.Count; i++)
                            {
                                edgesArray.SetValue(edges.Item(i), i - 1);
                            }

                            model.ConvToSMs.AddEx(face, 0, edgesArray, 0, 0, 0);
                            model.ConvToSMs.Item(1).ShowDimensions = true;
                        }
                        else
                        {
                            Console.WriteLine("The part is already transformed.", "Transformation existant");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Erreur during the transformation : {ex.Message}");
                        DialogResult result = MessageBox.Show(
                            "Please Transfrom the Part manually.\n" +
                            "When you are done, press OK to continue or cancel to end the program.",
                            "Transform Problem",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Warning
                        );

                        if (result == DialogResult.Cancel)
                        {
                            Console.WriteLine("Programme stopped by user.");
                            System.Environment.Exit(0);
                        }
                    }

                    if (dynamicDoc.FlatPatternModels.Count == 0)
                    {
                        flatPatternModel = dynamicDoc.FlatPatternModels.Add(dynamicDoc.Models.Item(1));
                    }
                    else
                    {
                        flatPatternModel = dynamicDoc.FlatPatternModels.Item(1);
                    }

                    if (flatPatternModel.FlatPatterns.Count != 0)
                    {
                        DialogResult result = MessageBox.Show(
                                "The part is already has a flat pattern.\n" +
                                "Do you want to create a new one?",
                                "Flat Pattern Existant",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Warning
                            );
                        if (result != DialogResult.OK) return;
                    }

                    models = dynamicDoc.Models;
                    model = models.Item(1);
                    body = (Body)model.Body;
                    faces = (Faces)body.Faces[FeatureTopologyQueryTypeConstants.igQueryPlane];

                    if (useForm)
                    {
                        using (FaceAndEdgeSelectionForm selectionForm = new FaceAndEdgeSelectionForm(application, dynamicDoc, model))
                        {
                            if (selectionForm.ShowDialog() == DialogResult.OK)
                            {
                                face = selectionForm.SelectedFace;
                                edge = selectionForm.SelectedEdge;
                                Console.WriteLine($"User selected Edge {edge.ID} on Face {face.ID}");
                            }
                            else
                            {
                                Console.WriteLine("User canceled selection for flat pattern.");
                                return;
                            }
                        }
                    }
                    else
                    {

                        face = GetFaceFurthestFromCenter(body, faces);

                        //// Manual face selection replace XXXX by face ID (Only for debugging)
                        //face = (Face)faces.Item(1);
                        //for (int i =1; i <= faces.Count; i++)
                        //{f
                        //    SolidEdgeGeometry.Face currentFace = (Face)faces.Item(i);
                        //    if (currentFace.ID == XXXX) face = currentFace;
                        //}
                        //Console.WriteLine($"User selected Face {face.ID} - Area: {face.Area * 1550.0031} mm²");

                        //Automatic edge selections
                        edge = GetEdgeAlignedWithCoordinatesSystem(face);
                    }

                    Console.WriteLine($"Chosen Edge: {edge.ID}, Chosen Face: {face.ID}");
                    vertex = (SolidEdgeGeometry.Vertex)edge.StartVertex;
                    flatPatternModel.FlatPatterns.Add(edge, face, vertex, SolidEdgeConstants.FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything);
                    Console.WriteLine("Flat pattern created successfully.");
                    //PartDocument partDocument = (PartDocument)dynamicDoc;
                    //System.Threading.Thread.Sleep(2000);
                    //partDocument.Save();
                    //partDocument.Close();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    SolidEdgeCommunity.OleMessageFilter.Unregister();
                }
            }
        }

        #region Face and Edge retrieval
        /// <summary>
        /// Finds the face that is furthest from the center of the body
        /// </summary>
        private static SolidEdgeGeometry.Face GetFaceFurthestFromCenter(SolidEdgeGeometry.Body body, SolidEdgeGeometry.Faces faces)
        {
            if (faces.Count == 0)
                return null;

            // List to store face areas and IDs
            List<(int Id, double Area, SolidEdgeGeometry.Face Face)> faceInfo = new List<(int, double, SolidEdgeGeometry.Face)>();

            // Initialize bounding box variables
            double[] minPoint = new double[3] { double.MaxValue, double.MaxValue, double.MaxValue };
            double[] maxPoint = new double[3] { double.MinValue, double.MinValue, double.MinValue };

            // Iterate through all faces to find the bounding box
            for (int i = 1; i <= faces.Count; i++)
            {
                SolidEdgeGeometry.Face currentFace = (SolidEdgeGeometry.Face)faces.Item(i);
                Array minRangePoint = Array.CreateInstance(typeof(double), 3);
                Array maxRangePoint = Array.CreateInstance(typeof(double), 3);
                currentFace.GetRange(ref minRangePoint, ref maxRangePoint);
                Array faceRange = Array.CreateInstance(typeof(double), 6);
                minRangePoint.CopyTo(faceRange, 0);
                maxRangePoint.CopyTo(faceRange, 3);

                // Store face area information
                double area = currentFace.Area;
                faceInfo.Add((i, area, currentFace));

                if (faceRange != null && faceRange.Length >= 6)
                {
                    minPoint[0] = Math.Min(minPoint[0], (double)faceRange.GetValue(0));
                    minPoint[1] = Math.Min(minPoint[1], (double)faceRange.GetValue(1));
                    minPoint[2] = Math.Min(minPoint[2], (double)faceRange.GetValue(2));

                    maxPoint[0] = Math.Max(maxPoint[0], (double)faceRange.GetValue(3));
                    maxPoint[1] = Math.Max(maxPoint[1], (double)faceRange.GetValue(4));
                    maxPoint[2] = Math.Max(maxPoint[2], (double)faceRange.GetValue(5));
                }
            }

            // Calculate the center of the bounding box
            double[] boxCenter = new double[3];
            boxCenter[0] = (minPoint[0] + maxPoint[0]) / 2.0;
            boxCenter[1] = (minPoint[1] + maxPoint[1]) / 2.0;
            boxCenter[2] = (minPoint[2] + maxPoint[2]) / 2.0;

            SolidEdgeGeometry.Face furthestFace = null;
            double maxDistance = -1;

            // Define minimum face area threshold (adjust this value as needed)
            const double MIN_FACE_AREA_THRESHOLD = 0.05; // in square meters

            // Loop through all faces to find the one furthest from center
            for (int i = 1; i <= faces.Count; i++)
            {
                SolidEdgeGeometry.Face currentFace = (SolidEdgeGeometry.Face)faces.Item(i);

                // Skip faces with area below the threshold
                if (currentFace.Area < MIN_FACE_AREA_THRESHOLD)
                    continue;

                Array minRangePoint = Array.CreateInstance(typeof(double), 3);
                Array maxRangePoint = Array.CreateInstance(typeof(double), 3);
                currentFace.GetRange(ref minRangePoint, ref maxRangePoint);
                Array faceRange = Array.CreateInstance(typeof(double), 6);
                minRangePoint.CopyTo(faceRange, 0);
                maxRangePoint.CopyTo(faceRange, 3);

                if (faceRange != null && faceRange.Length >= 6)
                {
                    // Calculate face center from its range
                    double[] faceCenter = new double[3];
                    faceCenter[0] = ((double)faceRange.GetValue(0) + (double)faceRange.GetValue(3)) / 2.0;
                    faceCenter[1] = ((double)faceRange.GetValue(1) + (double)faceRange.GetValue(4)) / 2.0;
                    faceCenter[2] = ((double)faceRange.GetValue(2) + (double)faceRange.GetValue(5)) / 2.0;

                    // Calculate distance from body center to face center
                    double distance = Math.Sqrt(
                        Math.Pow(faceCenter[0] - boxCenter[0], 2) +
                        Math.Pow(faceCenter[1] - boxCenter[1], 2) +
                        Math.Pow(faceCenter[2] - boxCenter[2], 2));

                    // Update if this face is further
                    if (distance > maxDistance)
                    {
                        maxDistance = distance;
                        furthestFace = currentFace;
                    }
                }
            }

            // Display information about the selected face
            if (furthestFace != null)
            {
                Console.WriteLine($"Selected furthest face area: {furthestFace.Area* 1550.0031:F6}sqin");
            }

            return furthestFace;
        }

        private static SolidEdgeGeometry.Edge GetEdgeAlignedWithCoordinatesSystem(SolidEdgeGeometry.Face face)
        {
            if (face == null)
                throw new System.Exception("No face selected.");

            SolidEdgeGeometry.Edges edges = (SolidEdgeGeometry.Edges)face.Edges;
            Console.WriteLine($"{edges.Count} edges found on face {face.ID}");

            if (edges.Count == 0)
                throw new System.Exception("Selected face has no edges.");

            const double METERS_TO_INCHES = 39.3701;
            SolidEdgeGeometry.Edge firstEdge = null;
            SolidEdgeGeometry.Edge selectedEdge = null;

            for (int i = 1; i <= edges.Count; i++)
            {
                SolidEdgeGeometry.Edge edge = (SolidEdgeGeometry.Edge)edges.Item(i);
                if (firstEdge == null) firstEdge = edge; // Store first edge found

                SolidEdgeGeometry.Vertex startVertex = (SolidEdgeGeometry.Vertex)edge.StartVertex;
                SolidEdgeGeometry.Vertex endVertex = (SolidEdgeGeometry.Vertex)edge.EndVertex;

                if (startVertex == null || endVertex == null)
                    continue;

                Array startPointArray = Array.CreateInstance(typeof(double), 3);
                Array endPointArray = Array.CreateInstance(typeof(double), 3);

                startVertex.GetPointData(ref startPointArray);
                endVertex.GetPointData(ref endPointArray);

                double[] startPoint = { (double)startPointArray.GetValue(0), (double)startPointArray.GetValue(1), (double)startPointArray.GetValue(2) };
                double[] endPoint = { (double)endPointArray.GetValue(0), (double)endPointArray.GetValue(1), (double)endPointArray.GetValue(2) };

                double startX = startPoint[0] * METERS_TO_INCHES;
                double startY = startPoint[1] * METERS_TO_INCHES;
                double startZ = startPoint[2] * METERS_TO_INCHES;
                double endX = endPoint[0] * METERS_TO_INCHES;
                double endY = endPoint[1] * METERS_TO_INCHES;
                double endZ = endPoint[2] * METERS_TO_INCHES;

                int sameCoordinates = 0;
                if (Math.Abs(startX - endX) < 0.001) sameCoordinates++;
                if (Math.Abs(startY - endY) < 0.001) sameCoordinates++;
                if (Math.Abs(startZ - endZ) < 0.001) sameCoordinates++;

                if (sameCoordinates >= 2)
                {
                    selectedEdge = edge;
                    break;
                }
            }

            return selectedEdge ?? firstEdge;
        }

        #endregion
    }
}