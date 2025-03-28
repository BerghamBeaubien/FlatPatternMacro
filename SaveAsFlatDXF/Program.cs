using SolidEdgeCommunity.Extensions;
using SolidEdgeGeometry;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SaveAsFlatDXF
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            //SolidEdgeFramework.Documents documents = null;

            dynamic dynamicDoc = null;
            //SolidEdgePart.PartDocument document = null;
            //SolidEdgePart.SheetMetalDocument sheetMetalDocument = null;

            SolidEdgePart.FlatPatternModel flatPatternModel = null;
            SolidEdgeGeometry.Body body = null;
            SolidEdgeGeometry.Faces faces = null;
            SolidEdgeGeometry.Face face = null;
            SolidEdgeGeometry.Edges edges = null;
            SolidEdgeGeometry.Edge edge = null;
            SolidEdgeGeometry.Vertex vertex = null;

            try
            {
                // Register with OLE to handle concurrency issues on the current thread.
                SolidEdgeCommunity.OleMessageFilter.Register();

                // Connect to Solid Edge,
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);

                // Get a reference to the Documents collection.
                //documents = application.Documents;
                dynamicDoc = application.ActiveDocument;

                // Get a refernce to the active sheetmetal document.
                if (dynamicDoc is SolidEdgePart.PartDocument)
                {
                    dynamicDoc = application.GetActiveDocument<SolidEdgePart.PartDocument>(false);
                    Console.WriteLine("PartDocument");
                }
                else if (dynamicDoc is SolidEdgePart.SheetMetalDocument)
                {
                    dynamicDoc = application.GetActiveDocument<SolidEdgePart.SheetMetalDocument>(false);
                    Console.WriteLine("SheetMetalDocument");
                }
                else
                {
                    throw new System.Exception("Active document is not a PartDocument or SheetMetalDocument.");
                }
                

                if (dynamicDoc == null)
                {
                    throw new System.Exception("No active document.");
                }

                if (dynamicDoc.FlatPatternModels.Count == 0)
                {
                    flatPatternModel = dynamicDoc.FlatPatternModels.Add(dynamicDoc.Models.Item(1));
                }
                else
                {
                    flatPatternModel = dynamicDoc.FlatPatternModels.Item(1);
                }

                // Get a reference to the model's body.
                body = (Body)flatPatternModel.Body;

                // Query for all faces in the body.
                faces = (SolidEdgeGeometry.Faces)body.Faces[SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll];

                // Show the number of faces for debugging
                //MessageBox.Show($"Number of faces in the part: {faces.Count}", "Face Count");

                // Find the face furthest from the center
                face = GetFaceFurthestFromCenter(body, faces);

                MessageBox.Show($"Selected face: {face.ID} + Aire = {face.Area.ToString()}", "Selected Face");

                if (face != null)
                {
                    // Get edges from the selected face
                    edges = (SolidEdgeGeometry.Edges)face.Edges;
                    MessageBox.Show($"Number of edges in the face: {edges.Count}", "Edge Count");

                    if (edges.Count > 0)
                    {
                        // Select the first edge as a default
                        edge = (SolidEdgeGeometry.Edge)edges.Item(1);
                    }
                    else
                    {
                        throw new System.Exception("Selected face has no edges.");
                    }
                }
                else
                {
                    throw new System.Exception("Failed to select a face.");
                }

                vertex = (SolidEdgeGeometry.Vertex)edge.StartVertex;

                if (dynamicDoc is SolidEdgePart.PartDocument partDocument)
                {
                    if (faces == null || faces.Count == 0)
                    {
                        Console.WriteLine("La collection faces est vide ou nulle.");
                        return;
                    }

                    for (int i = 1; i <= faces.Count; i++)
                    {
                        try
                        {
                            object faz = faces.Item(i);

                            if (faz == null)
                            {
                                Console.WriteLine($"Face {i} est null, passage à la suivante.");
                                continue;
                            }

                            // Vérifier si l'objet est bien un COM Object
                            if (!Marshal.IsComObject(faz))
                            {
                                Console.WriteLine($"Face {i} n'est pas un objet COM valide.");
                                continue;
                            }

                            // Vérifier si l'objet supporte une méthode typique d'une Face
                            var faceType = faz.GetType();
                            if (faceType.InvokeMember("GetType", System.Reflection.BindingFlags.InvokeMethod, null, faz, null) == null)
                            {
                                Console.WriteLine($"faz {i} ne semble pas être un type reconnu.");
                                continue;
                            }

                            // Si toutes les vérifications passent, tenter la transformation
                            partDocument.TransformToSynchronousSheetmetal(faz);
                            Console.WriteLine($"Transformation réussie avec la face {i}");
                            break; // Sortie de la boucle si la transformation réussit
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erreur sur la face {i}: {ex.Message}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Le document n'est pas un PartDocument.");
                }



                flatPatternModel.FlatPatterns.Add(edge, face, vertex, SolidEdgeConstants.FlattenPatternModelTypeConstants.igFlattenPatternModelTypeFlattenAnything);

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
            const double MIN_FACE_AREA_THRESHOLD = 0.005; // in square meters

            // Display face areas information
            string faceInfoMessage = "Face Information:\n";
            foreach (var info in faceInfo)
            {
                faceInfoMessage += $"Face ID: {info.Id}, Area: {info.Area:F6}\n";
            }
            MessageBox.Show(faceInfoMessage);

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
                MessageBox.Show($"Selected furthest face area: {furthestFace.Area:F6}");
            }

            return furthestFace;
        }
    }
}