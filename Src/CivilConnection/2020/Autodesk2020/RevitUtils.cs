// Copyright (c) 2016 Autodesk, Inc. All rights reserved.
// Author: paolo.serra@autodesk.com
// 
// Licensed under the Apache License, Version 2.0 (the "License").
// You may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//    http://www.apache.org/licenses/LICENSE-2.0
// 
//  Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
// implied.  See the License for the specific language governing
// permissions and limitations under the License.
using Autodesk.AECC.Interop.Land;
using Autodesk.AECC.Interop.Roadway;
using Autodesk.AECC.Interop.UiRoadway;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using Revit.Elements;
using Revit.Elements.Views;
using Revit.GeometryConversion;
using RevitServices.Persistence;
using RevitServices.Transactions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ADSK_Parameters = CivilConnection.UtilsObjectsLocation.ADSK_Parameters;


namespace CivilConnection
{
    /// <summary>
    /// Collection of utilities for the integration with Revit.
    /// </summary>
    public class RevitUtils
    {
        #region PRIVATE PROPERTIES


        #endregion

        #region PUBLIC PROPERTIES


        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Initializes a new instance of the <see cref="RevitUtils"/> class.
        /// </summary>
        internal RevitUtils()
        { }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Return element id of crop box for a given view.
        /// The built-in parameter ID_PARAM of the crop box
        /// contains the element id of the view it is used in;
        /// e.g., the crop box 'points' to the view using it
        /// via ID_PARAM. Therefore, we can use a parameter
        /// filter to retrieve all crop boxes with the
        /// view's element id in that parameter.
        /// 
        /// source:
        /// http://thebuildingcoder.typepad.com/blog/2018/02/efficiently-retrieve-crop-box-for-given-view.html
        /// </summary>
        private static Autodesk.Revit.DB.ElementId GetCropBoxFor(Autodesk.Revit.DB.View view)
        {
            Utils.Log(string.Format("RevitUtils.GetCropBoxFor started...", ""));

            Autodesk.Revit.DB.ParameterValueProvider provider
                = new Autodesk.Revit.DB.ParameterValueProvider(new Autodesk.Revit.DB.ElementId(
                    (int)Autodesk.Revit.DB.BuiltInParameter.ID_PARAM));

            Autodesk.Revit.DB.FilterElementIdRule rule
                = new Autodesk.Revit.DB.FilterElementIdRule(provider,
                                          new Autodesk.Revit.DB.FilterNumericEquals(), view.Id);

            Autodesk.Revit.DB.ElementParameterFilter filter
                = new Autodesk.Revit.DB.ElementParameterFilter(rule);

            Utils.Log(string.Format("RevitUtils.GetCropBoxFor completed.", ""));

            return new Autodesk.Revit.DB.FilteredElementCollector(view.Document)
                .WherePasses(filter)
                .ToElementIds()
                .Where<Autodesk.Revit.DB.ElementId>(a => a.IntegerValue
                                  != view.Id.IntegerValue)
                .FirstOrDefault<Autodesk.Revit.DB.ElementId>();
        }

        /// <summary>
        /// Rotate the FamilyInstance around the insertion point to match the local Z-Axis with the provided vector.
        /// </summary>
        /// <param name="familyInstance">The FamilyInstance to rotate.</param>
        /// <param name="vector">The Vector used to align hte FmailyInstance local Z-Axis.</param>
        /// <returns></returns>
        /// 
        [IsVisibleInDynamoLibrary(false)]
        private static FamilyInstance SetZAxisByVector(FamilyInstance familyInstance, Vector vector)
        {
            Utils.Log(string.Format("RevitUtils.SetZAxisByVector started...", ""));

            Autodesk.Revit.DB.XYZ zaxis = vector.ToXyz().Normalize();

            Autodesk.Revit.DB.FamilyInstance element = familyInstance.InternalElement as Autodesk.Revit.DB.FamilyInstance;

            Autodesk.Revit.DB.XYZ xaxis = element.GetTransform().BasisY.CrossProduct(zaxis);

            Autodesk.Revit.DB.XYZ normal = element.GetTransform().BasisZ.CrossProduct(xaxis);

            Autodesk.Revit.DB.FamilyInstance copy = null;

            if (!normal.IsZeroLength())
            {
                Autodesk.Revit.DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

                TransactionManager.Instance.ForceCloseTransaction();

                TransactionManager.Instance.EnsureInTransaction(doc);

                Autodesk.Revit.DB.LocationPoint lp = element.Location as Autodesk.Revit.DB.LocationPoint;

                Autodesk.Revit.DB.FamilySymbol fs = element.Symbol;

                if (!fs.IsActive)
                {
                    fs.Activate();
                }

                copy = doc.Create.NewFamilyInstance(lp.Point, fs, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(lp.Point, normal + lp.Point);

                try
                {
                    Autodesk.Revit.DB.ElementTransformUtils.RotateElement(doc, copy.Id, line, element.GetTransform().BasisX.AngleTo(xaxis));
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: RevitUtils.SetZAxisByVector {0}", ex.Message));

                    throw new Exception(string.Format("Rotation Failed.\n{0}", ex.Message));
                }

                TransactionManager.Instance.TransactionTaskDone();
            }

            Utils.Log(string.Format("RevitUtils.SetZAxisByVector completed.", ""));

            return copy.ToDSType(false) as FamilyInstance;
        }

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Sections the view by station.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="station">The station.</param>
        /// <param name="lengthLeft">The length left.</param>
        /// <param name="lengthRight">The length right.</param>
        /// <param name="elevationMin">The elevation minimum.</param>
        /// <param name="elevationMax">The elevation maximum.</param>
        /// <param name="depth">The depth.</param>
        /// <returns></returns>
        public static SectionView SectionViewByStation(Baseline baseline, double station, double lengthLeft = 20, double lengthRight = 20, double elevationMin = -30, double elevationMax = 30, double depth = 1)
        {
            Utils.Log(string.Format("RevitUtils.SectionViewByStation started...", ""));

            CoordinateSystem cs = baseline.CoordinateSystemByStation(station);
            CoordinateSystem docCs = cs.Transform(DocumentTotalTransform());
            cs = CoordinateSystem.ByOriginVectors(docCs.Origin, docCs.XAxis.Reverse(), docCs.ZAxis);

            if (elevationMax < elevationMin)
            {
                var t = elevationMin;
                elevationMin = elevationMax;
                elevationMax = t;
            }

            if (elevationMin == elevationMax)
            {
                elevationMax += 10;
            }

            // The X Axis in the CoordinateSystem is reversed so Left is Right and Right is Left
            SectionView section = SectionView.ByCoordinateSystemMinPointMaxPoint(cs, Point.ByCoordinates(-lengthRight, elevationMin, 0), Point.ByCoordinates(lengthLeft, elevationMax, depth));

            string name = string.Format("{0}_BL {1}_{2:0+000.00}", baseline.CorridorName, baseline.Index, station);

            var doc = DocumentManager.Instance.CurrentDBDocument;  // 1.1.0

            var existing = new Autodesk.Revit.DB.FilteredElementCollector(doc)  // 1.1.0
                 .OfClass(typeof(Autodesk.Revit.DB.ViewSection))
                 .WhereElementIsNotElementType()
                 .Where(x => x.Name.Contains(name));

            if (existing.Count() > 0)  // 1.1.0
            {
                name += string.Format("_{0}", existing.Count() - 1);
            }

            section.SetParameterByName("View Name", name);  // 1.1.0

            cs.Dispose();

            Utils.Log(string.Format("RevitUtils.SectionViewByStation complted.", ""));
            return section;
        }

        /// <summary>
        /// Updates the section view by Coordinate System, Min and Max points. The view depth is controlled by the the difference of the Z coordinates of min and max points.
        /// </summary>
        /// <param name="section">The section view.</param>
        /// <param name="cs">The coordinate system with vertical Z axis.</param>
        /// <param name="minPoint">The 2D point in the view coordinates that represents the bottom left corner in the view.</param>
        /// <param name="maxPoint">The 2D point in hte view coordinates that represents the top right corner in the view.</param>
        /// <returns></returns>
        public static SectionView UpdateSectionViewByCoordinateSystem(SectionView section, CoordinateSystem cs, Point minPoint = null, Point maxPoint = null)
        {
            Utils.Log(string.Format("RevitUtils.UpdateSectionViewByCoordinateSystem started...", ""));

            if (Math.Abs(cs.ZAxis.Dot(Vector.ZAxis()) - 1) > 0.001)
            {
                var message = "The coordinate system must have a vertical Z axis.";

                Utils.Log(string.Format("ERROR: RevitUtils.UpdateSectionViewByCoordinateSystem {0}", message));

                throw new Exception(message);
            }

            Autodesk.Revit.DB.ViewSection internalSection = section.InternalElement as Autodesk.Revit.DB.ViewSection;

            Autodesk.Revit.DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

            // The view and the element that represents the section are different objects!
            Autodesk.Revit.DB.Element view = new Autodesk.Revit.DB.FilteredElementCollector(doc)
                .OfCategory(Autodesk.Revit.DB.BuiltInCategory.OST_Viewers)
                .WhereElementIsNotElementType()
                .First(x => x.Name == section.Name);

            if (view != null)
            {
                Autodesk.Revit.DB.Transform tr = Autodesk.Revit.DB.Transform.Identity;

                tr.Origin = cs.Origin.ToXyz();
                tr.BasisX = cs.XAxis.ToXyz();
                tr.BasisY = Vector.ZAxis().ToXyz();
                tr.BasisZ = tr.BasisX.CrossProduct(tr.BasisY);

                TransactionManager.Instance.EnsureInTransaction(doc);

                Autodesk.Revit.DB.ElementTransformUtils.RotateElement(doc,
                    view.Id,
                    Autodesk.Revit.DB.Line.CreateBound(internalSection.Origin, internalSection.Origin + Autodesk.Revit.DB.XYZ.BasisZ),
                    internalSection.RightDirection.AngleOnPlaneTo(tr.BasisX, Autodesk.Revit.DB.XYZ.BasisZ));

                Autodesk.Revit.DB.ElementTransformUtils.MoveElement(doc, view.Id, tr.Origin - internalSection.Origin);

                if (minPoint == null)
                {
                    minPoint = internalSection.CropBox.Min.ToPoint();
                }

                if (maxPoint == null)
                {
                    maxPoint = internalSection.CropBox.Max.ToPoint();
                }

                Autodesk.Revit.DB.BoundingBoxXYZ bb = new Autodesk.Revit.DB.BoundingBoxXYZ();

                bb.Transform = tr;
                Autodesk.Revit.DB.XYZ min = minPoint.ToXyz();
                Autodesk.Revit.DB.XYZ max = maxPoint.ToXyz();
                double depth = Math.Abs(max.Z - min.Z);

                bb.Min = new Autodesk.Revit.DB.XYZ(min.X, min.Y, 0);
                bb.Max = new Autodesk.Revit.DB.XYZ(max.X, max.Y, 0);

                internalSection.CropBox = bb;

                if (depth > 0)
                {
                    internalSection.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                        .First(x => x.Id.IntegerValue.Equals(
                            (int)Autodesk.Revit.DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR))
                        .Set(depth);
                }

                TransactionManager.Instance.TransactionTaskDone();
            }

            Utils.Log(string.Format("RevitUtils.UpdateSectionViewByCoordinateSystem completed.", ""));

            return section;
        }

        /// <summary>
        /// Updates the section view by Coordinate System, Min and Max points. The view depth is controlled by the the difference of the Z coordinates of min and max points.
        /// </summary>
        /// <param name="plan">The plan view.</param>
        /// <param name="cs">The coordinate system with vertical Z axis.</param>
        /// <param name="minPoint">The 2D point in the view coordinates that represents the bottom left corner in the view.</param>
        /// <param name="maxPoint">The 2D point in tHe view coordinates that represents the top right corner in the view.</param>
        /// <returns></returns>
        public static FloorPlanView UpdatePlanViewByCoordinateSystem(FloorPlanView plan, CoordinateSystem cs, Point minPoint, Point maxPoint)
        {
            Utils.Log(string.Format("RevitUtils.UpdatePlanViewByCoordinateSystem started...", ""));

            if (Math.Abs(cs.ZAxis.Dot(Vector.ZAxis()) - 1) > 0.001)
            {
                var message = "The coordinate system must have a vertical Z axis.";

                Utils.Log(string.Format("ERROR: RevitUtils.UpdatePlanViewByCoordinateSystem {0}", message));

                throw new Exception(message);
            }

            Autodesk.Revit.DB.ViewPlan internalPlan = plan.InternalElement as Autodesk.Revit.DB.ViewPlan;

            Autodesk.Revit.DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

            // The view and the element that represents the view are different objects!
            Autodesk.Revit.DB.Element view = doc.GetElement(GetCropBoxFor(internalPlan));

            if (view != null)
            {
                Autodesk.Revit.DB.Transform tr = Autodesk.Revit.DB.Transform.Identity;

                tr.Origin = cs.Origin.ToXyz();
                tr.BasisX = cs.XAxis.ToXyz();
                tr.BasisY = cs.YAxis.ToXyz();
                tr.BasisZ = cs.ZAxis.ToXyz();

                TransactionManager.Instance.EnsureInTransaction(doc);

                Autodesk.Revit.DB.ElementTransformUtils.RotateElement(doc,
                    view.Id,
                    Autodesk.Revit.DB.Line.CreateBound(internalPlan.Origin, internalPlan.Origin + Autodesk.Revit.DB.XYZ.BasisZ),
                    internalPlan.RightDirection.AngleOnPlaneTo(tr.BasisX, Autodesk.Revit.DB.XYZ.BasisZ));

                internalPlan.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                        .First(x => x.Id.IntegerValue.Equals(
                            (int)Autodesk.Revit.DB.BuiltInParameter.PLAN_VIEW_NORTH))
                        .Set(0);  // Project North

                Autodesk.Revit.DB.ElementTransformUtils.MoveElement(doc, view.Id, tr.Origin - internalPlan.Origin);

                if (minPoint == null)
                {
                    minPoint = internalPlan.CropBox.Min.ToPoint();
                }

                if (maxPoint == null)
                {
                    maxPoint = internalPlan.CropBox.Max.ToPoint();
                }

                Autodesk.Revit.DB.BoundingBoxXYZ bb = new Autodesk.Revit.DB.BoundingBoxXYZ();

                bb.Transform = tr;
                Autodesk.Revit.DB.XYZ min = minPoint.ToXyz();
                Autodesk.Revit.DB.XYZ max = maxPoint.ToXyz();
                double depth = Math.Abs(max.Z - min.Z);

                bb.Min = new Autodesk.Revit.DB.XYZ(min.X, min.Y, 0);
                bb.Max = new Autodesk.Revit.DB.XYZ(max.X, max.Y, 0);

                internalPlan.CropBox = bb;

                if (depth > 0)
                {
                    internalPlan.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                        .First(x => x.Id.IntegerValue.Equals(
                            (int)Autodesk.Revit.DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR))
                        .Set(depth);
                }

                TransactionManager.Instance.TransactionTaskDone();
            }

            Utils.Log(string.Format("RevitUtils.UpdatePlanViewByCoordinateSystem completed.", ""));

            return plan;
        }

        /// <summary>
        /// Returns the sample lines parameters associated to an alignment
        /// </summary>
        /// <param name="alignment">The alignment</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<Dictionary<string, object>> AlignmentSampleLinesParameters(AeccAlignment alignment)
        {
            Utils.Log(string.Format("RevitUtils.AlignmentSampleLinesParameters started...", ""));

            Utils.Log(string.Format("Alignment: {0}", alignment.Name));

            var output = new List<Dictionary<string, object>>();

            if (alignment.SampleLineGroups.Count == 0)
            {
                string message = "The alignment does not contain any sample line groups.";

                Utils.Log(string.Format("ERROR: {0}", message));

                return output;
            }

            foreach (AeccSampleLineGroup group in alignment.SampleLineGroups)
            {
                IList<double[]> sectionList = new List<double[]>();

                double[] sectionValues = new double[5];

                var slg = new Dictionary<string, object> { { "station", null }, { "lengthLeft", null }, { "lengthRight", null }, { "elevationMin", null }, { "elevationMax", null } };

                Utils.Log(string.Format("SampleLineGroup: {0}", group.Name));

                try
                {
                    if (group.SampleLines.Count == 0)
                    {
                        string message = "The sample line group does not contain any sample lines.";

                        Utils.Log(string.Format("ERROR: {0}", message));
                    }

                    foreach (AeccSampleLine line in group.SampleLines)
                    {

                        try
                        {
                            Utils.Log(string.Format("SampleLine: {0} : {1}", line.Name, line.Station));

                            if (line.Sections.Count == 0)
                            {
                                string message = "The sample line does not contain any sections.";

                                Utils.Log(string.Format("ERROR: {0}", message));

                                // throw new Exception(message);
                            }

                            foreach (AeccSection sec in line.Sections)
                            {

                                Utils.Log(string.Format("Station: {0}, Left: {1}, Right: {2}, Elev. Max: {3}, Elev.Min: {4}", sec.Station, sec.LengthLeft, sec.LengthRight, sec.ElevationMin, sec.ElevationMax));

                                try
                                {
                                    sectionList.Add(new double[] { sec.Station, sec.LengthLeft, sec.LengthRight, sec.ElevationMin, sec.ElevationMax });
                                }
                                catch (Exception ex)
                                {
                                    sectionList.Add(new double[] { sec.Station, 0, 0, 0, 0 });

                                    Utils.Log(string.Format("ERROR: {0}", ex.Message));
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            sectionList.Add(new double[] { line.Station, 0, 0, 0, 0 });

                            Utils.Log(string.Format("ERROR: {0}", ex.Message));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: {0}", ex.Message));
                }

                var groups = sectionList.GroupBy(x => x.First()).OrderBy(g => g.Key);
                var station = groups.Select(g => g.Key);
                var left = groups.Select(g => g.Min(q => -q[1]));
                var right = groups.Select(g => g.Max(q => q[2]));
                var min = groups.Select(g => g.Min(q => q[3]));
                var max = groups.Select(g => g.Max(q => q[4]));

                slg["station"] = station;
                slg["lengthLeft"] = left;
                slg["lengthRight"] = right;
                slg["elevationMin"] = min;
                slg["elevationMax"] = max;

                output.Add(slg);
            }

            Utils.Log(string.Format("RevitUtils.AlignmentSampleLinesParameters completed.", ""));

            return output;
        }

        /// <summary>
        /// Returns the sections lines associated to the sample line.
        /// </summary>
        /// <param name="line">The sample line.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<IList<Line>> GetSampleLinesSections(AeccSampleLine line)
        {
            Utils.Log(string.Format("RevitUtils.GetSampleLinesSections started...", ""));

            var output = new List<IList<Line>>();  // Dictionary<string, object>();

            Utils.Log(string.Format("SampleLine: {0} : {1}", line.Name, line.Station));     

            if (line.Sections.Count == 0)
            {
                string message = "The sample line does not contain any sections.";

                Utils.Log(string.Format("{0}", message));

                return null;
            }

            foreach (AeccSection sec in line.Sections)
            {
                Utils.Log(string.Format("Section: {0}", sec.Name));

                IList<Line> lines = new List<Line>();

                foreach (AeccSectionLink link in sec.Links)
                {
                    Point start = Point.ByCoordinates(link.StartPointX, link.StartPointY, link.StartPointZ);

                    Point end = Point.ByCoordinates(link.EndPointX, link.EndPointY, link.EndPointZ);

                    Utils.Log(string.Format("Start: {0} End: {1}", start, end));

                    if (start.DistanceTo(end) > 0.0001)
                    {
                        Line l = Line.ByStartPointEndPoint(start, end);

                        lines.Add(l);
                    }
                }

                if (lines.Count > 0)
                {
                    output.Add(lines);
                }
            }

            Utils.Log(string.Format("RevitUtils.GetSampleLinesSections completed.", ""));

            return output;
        }

        /// <summary>
        /// Returns the section lines associated to the sample lines in an alignment
        /// </summary>
        /// <param name="alignment">The alignment</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<Dictionary<double, IList<IList<Line>>>> AlignmentSectionsLines(Alignment alignment)
        {

            Utils.Log(string.Format("RevitUtils.AlignmentSectionsLines started...", ""));

            Utils.Log(string.Format("Alignment: {0}", alignment.Name));

            var output = new List<Dictionary<double, IList<IList<Line>>>>();

            var alg = (AeccAlignment)alignment.InternalElement;

            if (alg.SampleLineGroups.Count == 0)
            {
                string message = "The alignment does not contain any sample line groups.";

                Utils.Log(string.Format("ERROR: {0}", message));

                return null;
            }

            foreach (AeccSampleLineGroup group in alg.SampleLineGroups)
            {
                var slg = new Dictionary<double, IList<IList<Line>>>();

                Utils.Log(string.Format("SampleLineGroup: {0}", group.Name));

                try
                {
                    if (group.SampleLines.Count == 0)
                    {
                        string message = "The sample line group does not contain any sample lines.";

                        Utils.Log(string.Format("ERROR: {0}", message));
                    }

                    foreach (AeccSampleLine line in group.SampleLines)
                    {
                        var sectionsLines = GetSampleLinesSections(line);

                        if (sectionsLines != null)
                        {
                            slg.Add(line.Station, sectionsLines);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: {0}", ex.Message));
                }

                if (slg.Keys.Count > 0)
                {
                    output.Add(slg);
                }
            }

            Utils.Log(string.Format("RevitUtils.AlignmentSectionsLines completed.", ""));

            return output;
        }


        /// <summary>
        /// Returns the Sample Lines parameters associated with the alignment.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "station", "lengthLeft", "lengthRight", "elevationMin", "elevationMax" })]
        public static IList<Dictionary<string, object>> SampleLinesParameters(Baseline baseline)
        {
            AeccAlignment alignment = baseline.Alignment.InternalElement as AeccAlignment;

            return AlignmentSampleLinesParameters(alignment);
        }

        /// <summary>
        /// Detail group by section view.
        /// </summary>
        /// <param name="section">The section.</param>
        /// <returns></returns>
        public static SectionView DetailGroupBySectionView(SectionView section)
        {
            UtilsSectionView.CutLines(section.InternalElement.Id);

            return section;
        }

        /// <summary>
        /// Gets the CoordinateSystem of the Revit document total transform.
        /// </summary>
        /// <returns></returns>
        public static CoordinateSystem DocumentTotalTransform()
        {
            // TODO Consider creating a session variable so it gets calculated only once
            // The risk is that if the user changes the shared coordinates in the session it will not update
            Utils.Log(string.Format("RevitUtils.DocumentTotalTransform started...", ""));

            CoordinateSystem cs = null;

            if (SessionVariables.DocumentTotalTransform != null)
            {
                cs = SessionVariables.DocumentTotalTransform;
            }
            else
            {
                var doc = DocumentManager.Instance.CurrentDBDocument;

                if (!doc.IsFamilyDocument)
                {
                    var location = doc.ActiveProjectLocation;

                    var transform = location.GetTotalTransform();

                    cs = CoordinateSystem.ByOriginVectors(transform.Origin.ToPoint(), transform.BasisX.ToVector(), transform.BasisY.ToVector(), transform.BasisZ.ToVector());
                }

                SessionVariables.DocumentTotalTransform = cs;
            }

            SessionVariables.DocumentTotalTransformInverse = cs.Inverse();

            Utils.Log(string.Format("{0}", cs));

            Utils.Log(string.Format("RevitUtils.DocumentTotalTransform completed.", ""));

            return cs;
        }

        /// <summary>
        /// Gets the Inverse CoordinateSystem of the Revit document total transform.
        /// </summary>
        /// <returns></returns>
        public static CoordinateSystem DocumentTotalTransformInverse()
        {
            // TODO Consider creating a session variable so it gets calculated only once
            // The risk is that if the user changes the shared coordinates in the session it will not update
            Utils.Log(string.Format("RevitUtils.DocumentTotalTransformInverse started...", ""));

            CoordinateSystem cs = null;

            if (SessionVariables.DocumentTotalTransformInverse != null)
            {
                cs = SessionVariables.DocumentTotalTransformInverse;
            }
            else
            {
                var doc = DocumentManager.Instance.CurrentDBDocument;

                if (!doc.IsFamilyDocument)
                {
                    var location = doc.ActiveProjectLocation;

                    var transform = location.GetTotalTransform();

                    cs = CoordinateSystem.ByOriginVectors(transform.Origin.ToPoint(), transform.BasisX.ToVector(), transform.BasisY.ToVector(), transform.BasisZ.ToVector());
                }

                SessionVariables.DocumentTotalTransform = cs;
                SessionVariables.DocumentTotalTransformInverse = cs.Inverse();
            }

            Utils.Log(string.Format("{0}", cs));

            Utils.Log(string.Format("RevitUtils.DocumentTotalTransformInverse completed.", ""));

            return cs;
        }


        /// <summary>
        /// Extracts the location paramaters by category.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="category">The category.</param>
        /// <returns></returns>
        public static object[][] ExtractParamatersByCategory(CivilDocument civilDocument, Category category)
        {
            return UtilsObjectsLocation.ReadFamilyInstancesPointBased(DocumentManager.Instance.CurrentDBDocument, civilDocument, category.Id);
        }

        /// <summary>
        /// Returns the object location parameters.
        /// </summary>
        /// <param name="update">if set to <c>true</c> [update].</param>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        public static object[] ObjectLocationParameters(bool update, CivilDocument civilDocument, Revit.Elements.Element element)
        {
            if (element.InternalElement is Autodesk.Revit.DB.FamilyInstance)
            {
                var familyInstance = element as FamilyInstance;
                return UtilsObjectsLocation.ObjectLocationParameters(familyInstance);
            }
            if (element.InternalElement is Autodesk.Revit.DB.MEPCurve)
            {
                var curve = element as AbstractMEPCurve;
                return UtilsObjectsLocation.LinearObjectLocationParameters(curve);
            }

            return null;
        }

        /// <summary>
        /// Captures the Revit Elements with linear coordinate system parameters.
        /// </summary>
        /// <param name="run">if set to <c>true</c> [run].</param>
        /// <returns></returns>
        public static string ExportXML(bool run)
        {
            if (run)
            {
                var doc = UtilsObjectsLocation.LocationXML();

                if (doc != null)
                {
                    return doc;
                }
            }

            return null;
        }

        /// <summary>
        /// Updates the location and the linear coordinate system parameters of all the Revit Elements captured in the project XML.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="elements">The elements.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public static IList<Revit.Elements.Element> UpdateDocument(CivilDocument civilDocument, Revit.Elements.Element[] elements = null)
        {
            return UtilsObjectsLocation.UdpateDocumentFromXML(civilDocument, elements);
        }

        /// <summary>
        /// Updates the location and the linear coordinate system parameters of a collection of Revit Elements.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="elements">A collection of elements.</param>
        /// <param name="normalized">If true it will keep the same proportion along the featureline rather than the exact station.</param>
        /// <remarks>At the end of the update the Station values may be different.</remarks>
        /// <returns></returns>
        public static IList<Revit.Elements.Element> UpdateObjects(CivilDocument civilDocument, IList<Revit.Elements.Element> elements, bool normalized = false)
        {
            ExportXML(true);

            return UtilsObjectsLocation.OptimizedUdpateObjectFromXML(civilDocument, elements, normalized);
        }

        /// <summary>
        /// Uses a featureline to assign or recalculate the linear coordinate system parameters of a Revit Element.
        /// In case of Adaptive Components it calculates the parameters for the first adaptive point only.
        /// In case of Floors it calculates the parameters for the first point of the top face only.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns></returns>
        public static object AssignFeatureline(Revit.Elements.Element element, Featureline featureline)
        {
            Utils.Log(string.Format("RevitUtils.AssignFeatureline started...", ""));

            try
            {
                Utils.Log(string.Format("Element {0} Featureline {1}", element.InternalElement.Id.IntegerValue, featureline.Code));

                Autodesk.Revit.DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

                if (!SessionVariables.ParametersCreated)
                {
                    UtilsObjectsLocation.CheckParameters(doc); 
                }

                Point lp = null;
                Point lpe = null;
                Vector localX = null;
                Vector localZ = null;
                CoordinateSystem cs = null;

                var totalTransform = RevitUtils.DocumentTotalTransform();
                var totalTransformInverse = totalTransform.Inverse();

                if (element.InternalElement.Category.Id.IntegerValue.Equals((int)Autodesk.Revit.DB.BuiltInCategory.OST_Floors))
                {
                    // extract vertices
                    // create multipoint object with featureline
                    // store serialization to ADSK_Multipoint parameter
                    foreach (Autodesk.Revit.DB.Face face in Autodesk.Revit.DB.HostObjectUtils.GetTopFaces((Autodesk.Revit.DB.HostObject)element.InternalElement)
                        .Select(x => element.InternalElement.GetGeometryObjectFromReference(x)).OrderBy(s => ((Autodesk.Revit.DB.Face)s).Area).ToList())
                    {

                        foreach (Autodesk.Revit.DB.CurveLoop cl in face.GetEdgesAsCurveLoops())
                        {
                            // Create the multipoint for the first curveloop
                            ShapePointArray sps = new ShapePointArray();

                            foreach (Autodesk.Revit.DB.Curve c in cl)
                            {
                                Point p = c.GetEndPoint(0).ToPoint().Transform(totalTransformInverse) as Point;

                                try
                                {
                                    sps.Add(ShapePoint.ByPointFeatureline(p, featureline));
                                }
                                catch (Exception ex)
                                {
                                    Utils.Log(string.Format("ERROR: RevitUtils.AssignFeatureline {0}", ex.Message));
                                }
                            }

                            if (sps.Count > 2)
                            {
                                if ((string)element.GetParameterValueByName(ADSK_Parameters.Instance.MultiPoint.Name) == "")
                                {
                                    MultiPoint mp = MultiPoint.ByShapePointArray(sps.Renumber());

                                    string serialized = mp.SerializeJSON();

                                    element.SetParameterByName(ADSK_Parameters.Instance.MultiPoint.Name, serialized);
                                    element.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                                    element.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);

                                    Utils.Log(string.Format("ADSK_MultiPoint: {0}", serialized));
                                }
                            }

                            break;  // first loop
                        }

                        break;  // first face
                        //}
                    }

                    Utils.Log(string.Format("RevitUtils.AssignFeatureline completed.", ""));

                    return element;
                }

                if (element.InternalElement is Autodesk.Revit.DB.FamilyInstance)
                {
                    if (Autodesk.Revit.DB.AdaptiveComponentInstanceUtils.IsAdaptiveComponentInstance(element.InternalElement as Autodesk.Revit.DB.FamilyInstance))
                    {

                        // extract vertices
                        // create multipoint object with featureline
                        // store serialization to ADSK_Multipoint parameter

                        Autodesk.Revit.DB.FamilyInstance ac = element.InternalElement as Autodesk.Revit.DB.FamilyInstance;

                        IList<ShapePoint> spList = new List<ShapePoint>();

                        foreach (var pid in Autodesk.Revit.DB.AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(ac))
                        {
                            Autodesk.Revit.DB.ReferencePoint refPoint = DocumentManager.Instance.CurrentDBDocument.GetElement(pid) as Autodesk.Revit.DB.ReferencePoint;

                            lp = refPoint.Position.ToPoint().Transform(totalTransformInverse) as Point;

                            ShapePoint sp = ShapePoint.ByPointFeatureline(lp, featureline);

                            spList.Add(sp);
                        }

                        ShapePointArray sps = ShapePointArray.ByShapePointList(spList);

                        MultiPoint mp = MultiPoint.ByShapePointArray(sps);

                        string serialized = mp.SerializeJSON();

                        element.SetParameterByName(ADSK_Parameters.Instance.MultiPoint.Name, serialized);
                        element.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                        element.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);

                        Utils.Log(string.Format("ADSK_MultiPoint: {0}", serialized));

                        Utils.Log(string.Format("RevitUtils.AssignFeatureline completed.", ""));

                        return element;
                    }
                    else if (element.InternalElement.Location is Autodesk.Revit.DB.LocationPoint)
                    {
                        var locPoint = element.InternalElement.Location as Autodesk.Revit.DB.LocationPoint;

                        lp = locPoint.Point.ToPoint().Transform(totalTransformInverse) as Point;
                    }
                    else if (element.InternalElement.Location is Autodesk.Revit.DB.LocationCurve)
                    {
                        var locCurve = element.InternalElement.Location as Autodesk.Revit.DB.LocationCurve;

                        lp = locCurve.Curve.GetEndPoint(0).ToPoint().Transform(totalTransformInverse) as Point;
                        lpe = locCurve.Curve.GetEndPoint(1).ToPoint().Transform(totalTransformInverse) as Point;
                    }

                    var soe = featureline.GetStationOffsetElevationByPoint(lp);
                    double station = (double)soe["Station"];
                    double offset = (double)soe["Offset"];
                    double elevation = (double)soe["Elevation"];

                    element.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);
                    Utils.Log(string.Format("ADSK_Corridor: {0}", featureline.Baseline.CorridorName));

                    element.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);
                    Utils.Log(string.Format("ADSK_BaselineIndex: {0}", featureline.Baseline.Index));

                    element.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, featureline.BaselineRegionIndex);  // 1.1.0
                    Utils.Log(string.Format("ADSK_RegionIndex: {0}", featureline.BaselineRegionIndex));

                    element.SetParameterByName(ADSK_Parameters.Instance.RegionRelative.Name, station - featureline.Start);  // 1.1.0
                    Utils.Log(string.Format("ADSK_RegionRelative: {0}", station - featureline.Start));

                    element.SetParameterByName(ADSK_Parameters.Instance.RegionNormalized.Name, (station - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0
                    Utils.Log(string.Format("ADSK_RegionNormalized: {0}", (station - featureline.Start) / (featureline.End - featureline.Start)));

                    element.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
                    Utils.Log(string.Format("ADSK_Code: {0}", featureline.Code));

                    element.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
                    Utils.Log(string.Format("ADSK_Side: {0}", featureline.Side));

                    element.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(lp.X, 3));
                    Utils.Log(string.Format("ADSK_X: {0}", Math.Round(lp.X, 3)));

                    element.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(lp.Y, 3));
                    Utils.Log(string.Format("ADSK_Y: {0}", Math.Round(lp.Y, 3)));

                    element.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(lp.Z, 3));
                    Utils.Log(string.Format("ADSK_Z: {0}", Math.Round(lp.Z, 3)));

                    element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round(station, 3));
                    Utils.Log(string.Format("ADSK_Station: {0}", Math.Round(station, 3)));

                    element.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round(offset, 3));
                    Utils.Log(string.Format("ADSK_Offset: {0}", Math.Round(offset, 3)));

                    element.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round(elevation, 3));
                    Utils.Log(string.Format("ADSK_Elevation: {0}", Math.Round(elevation, 3)));

                    element.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                    Utils.Log(string.Format("ADSK_Update: {0}", true));

                    element.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);
                    Utils.Log(string.Format("ADSK_Delete: {0}", false));


                    // TODO: check Revit length units and perform the conversion accordingly
                    if (element.InternalElement.Category.Id.IntegerValue.Equals((int)Autodesk.Revit.DB.BuiltInCategory.OST_StructuralColumns) ||
                            element.InternalElement.Category.Id.IntegerValue.Equals((int)Autodesk.Revit.DB.BuiltInCategory.OST_Columns))
                    {
                        Autodesk.Revit.DB.FamilyInstance column = element.InternalElement as Autodesk.Revit.DB.FamilyInstance;
                        if (!column.IsSlantedColumn)
                        {
                            var height = Convert.ToDouble(column.Parameters.Cast<Autodesk.Revit.DB.Parameter>()
                                .First(x => x.Id.IntegerValue.Equals((int)Autodesk.Revit.DB.BuiltInParameter.INSTANCE_LENGTH_PARAM))
                                .AsDouble());
                            lpe = lp.Translate(0, 0, Utils.FeetToM(height)) as Point;
                        }
                    }

                    if (lpe != null)
                    {
                        soe = featureline.GetStationOffsetElevationByPoint(lpe);
                        station = (double)soe["Station"];
                        offset = (double)soe["Offset"];
                        elevation = (double)soe["Elevation"];

                        element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, Math.Round(station, 3));
                        Utils.Log(string.Format("ADSK_EndStation: {0}", Math.Round(station, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndOffset.Name, Math.Round(offset, 3));
                        Utils.Log(string.Format("ADSK_EndOffset: {0}", Math.Round(offset, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndElevation.Name, Math.Round(elevation, 3));
                        Utils.Log(string.Format("ADSK_EndElevation: {0}", Math.Round(elevation, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndRegionRelative.Name, station - featureline.Start);  // 1.1.0
                        Utils.Log(string.Format("ADSK_EndRegionRelative: {0}", station - featureline.Start));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndRegionNormalized.Name, (station - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0
                        Utils.Log(string.Format("ADSK_EndRegionNormalized: {0}", (station - featureline.Start) / (featureline.End - featureline.Start)));

                    }
                    else
                    {
                        FamilyInstance fi = element as FamilyInstance;

                        cs = featureline.CoordinateSystemByStation(station);

                        if (cs != null)
                        {
                            localX = cs.XAxis;
                            localZ = cs.ZAxis;

                            element.SetParameterByName(ADSK_Parameters.Instance.AngleZ.Name, Math.Round(-localX.AngleAboutAxis(fi.FacingOrientation, localZ), 3));
                            Utils.Log(string.Format("ADSK_AngleZ: {0}", Math.Round(-localX.AngleAboutAxis(fi.FacingOrientation, localZ), 3)));
                        }
                        else
                        {
                            Utils.Log(string.Format("ERROR: Cannot calculate the Angle Z value", ""));
                        }
                    }
                }
                else
                {
                    // System Families such as walls or MEP Curves
                    if (element.InternalElement.Location is Autodesk.Revit.DB.LocationCurve)
                    {
                        Autodesk.Revit.DB.LocationCurve locCurve = element.InternalElement.Location as Autodesk.Revit.DB.LocationCurve;

                        lp = locCurve.Curve.GetEndPoint(0).ToPoint().Transform(totalTransformInverse) as Point;

                        var soe = featureline.GetStationOffsetElevationByPoint(lp);
                        double station = (double)soe["Station"];
                        double offset = (double)soe["Offset"];
                        double elevation = (double)soe["Elevation"];
                        cs = featureline.CoordinateSystemByStation(station);

                        element.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);
                        Utils.Log(string.Format("ADSK_Corridor: {0}", featureline.Baseline.CorridorName));

                        element.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);
                        Utils.Log(string.Format("ADSK_BaselineIndex: {0}", featureline.Baseline.Index));

                        element.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, featureline.BaselineRegionIndex);  // 1.1.0
                        Utils.Log(string.Format("ADSK_RegionIndex: {0}", featureline.BaselineRegionIndex));

                        element.SetParameterByName(ADSK_Parameters.Instance.RegionRelative.Name, station - featureline.Start);  // 1.1.0
                        Utils.Log(string.Format("ADSK_RegionRelative: {0}", station - featureline.Start));

                        element.SetParameterByName(ADSK_Parameters.Instance.RegionNormalized.Name, (station - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0
                        Utils.Log(string.Format("ADSK_RegionNormalized: {0}", (station - featureline.Start) / (featureline.End - featureline.Start)));

                        element.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
                        Utils.Log(string.Format("ADSK_Code: {0}", featureline.Code));

                        element.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
                        Utils.Log(string.Format("ADSK_Side: {0}", featureline.Side));

                        element.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(lp.X, 3));
                        Utils.Log(string.Format("ADSK_X: {0}", Math.Round(lp.X, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(lp.Y, 3));
                        Utils.Log(string.Format("ADSK_Y: {0}", Math.Round(lp.Y, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(lp.Z, 3));
                        Utils.Log(string.Format("ADSK_Z: {0}", Math.Round(lp.Z, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round(station, 3));
                        Utils.Log(string.Format("ADSK_Station: {0}", Math.Round(station, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round(offset, 3));
                        Utils.Log(string.Format("ADSK_Offset: {0}", Math.Round(offset, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round(elevation, 3));
                        Utils.Log(string.Format("ADSK_Elevation: {0}", Math.Round(elevation, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                        Utils.Log(string.Format("ADSK_Update: {0}", true));

                        element.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);
                        Utils.Log(string.Format("ADSK_Delete: {0}", false));

                        lpe = locCurve.Curve.GetEndPoint(1).ToPoint().Transform(totalTransformInverse) as Point;

                        soe = featureline.GetStationOffsetElevationByPoint(lpe);
                        station = (double)soe["Station"];
                        offset = (double)soe["Offset"];
                        elevation = (double)soe["Elevation"];

                        element.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, Math.Round(station, 3));
                        Utils.Log(string.Format("ADSK_EndStation: {0}", Math.Round(station, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndOffset.Name, Math.Round(offset, 3));
                        Utils.Log(string.Format("ADSK_EndOffset: {0}", Math.Round(offset, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndElevation.Name, Math.Round(elevation, 3));
                        Utils.Log(string.Format("ADSK_EndElevation: {0}", Math.Round(elevation, 3)));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndRegionRelative.Name, station - featureline.Start);  // 1.1.0
                        Utils.Log(string.Format("ADSK_EndRegionRelative: {0}", station - featureline.Start));

                        element.SetParameterByName(ADSK_Parameters.Instance.EndRegionNormalized.Name, (station - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0
                        Utils.Log(string.Format("ADSK_EndRegionNormalized: {0}", (station - featureline.Start) / (featureline.End - featureline.Start)));
                    }
                }

                Utils.Log(string.Format("RevitUtils.AssignFeatureline completed.", ""));

                return element;
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: RevitUtils.AssignFeatureline {0}", ex.Message));
                throw ex;
            }
        }

        /// <summary>
        /// Given a Revit Element it returns the first Featureline that meets the arguments.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="corridor">The corridor.</param>
        /// <param name="baselineIndex">Index of the baseline.</param>
        /// <param name="code">The code.</param>
        /// <param name="side">The side.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Featureline" })]
        public static Dictionary<string, object> GetFeaturelineByElementCodeSide(Revit.Elements.Element element, Corridor corridor, int baselineIndex, string code, string side)
        {
            return new Dictionary<string, object>() { { "Featureline", UtilsObjectsLocation.ClosestFeaturelineByElement(element, corridor, baselineIndex, code, side) } };
        }

        /// <summary>
        /// Processes the point based family instances by data.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        [MultiReturn(new string[] { "Created", "Updated", "Deleted" })]
        public static Dictionary<string, object> ProcessPointBasedFamilyInstancesByData(CivilDocument civilDocument, object[][] data)
        {
            return UtilsObjectsLocation.FamilyInstancesPointBased(DocumentManager.Instance.CurrentDBDocument, civilDocument, data);
        }

        /// <summary>
        /// Returns the FamilieInstance location parameters for update.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        [MultiReturn(new string[] { "FamilyInstance", "FamilyType", "Mark", "Featureline", "UseBaseline", "Station", "Offset", "Elevation", "AngleZ" })]
        public static Dictionary<string, object> FamilyInstanceLocationParametersForUpdate(CivilDocument civilDocument, object[][] data)
        {
            return UtilsObjectsLocation.ReadFamilyInstanceLocationParametersForUpdate(DocumentManager.Instance.CurrentDBDocument, civilDocument, data);
        }

        /// <summary>
        /// Returns the FamilieInstance location parameters for creation.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        [MultiReturn(new string[] { "FamilyType", "Mark", "Featureline", "UseBaseline", "Station", "Offset", "Elevation", "AngleZ" })]
        public static Dictionary<string, object> FamilyInstanceLocationParametersForCreate(CivilDocument civilDocument, object[][] data)
        {
            return UtilsObjectsLocation.ReadFamilyInstanceLocationParametersForCreate(DocumentManager.Instance.CurrentDBDocument, civilDocument, data);
        }

        /// <summary>
        /// Creates FamilyInstances using the featurelines to define linear coordinate systems and assign the parameters for the update.
        /// </summary>
        /// <param name="run">if set to <c>true</c> [run].</param>
        /// <param name="familyType">Type of the family.</param>
        /// <param name="featureline">The featureline.</param>
        /// <param name="useBaseline">if set to <c>true</c> [use baseline].</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <param name="angleZ">The angle z.</param>
        /// <returns></returns>
        public static Revit.Elements.FamilyInstance CreateFamilyInstance(bool run, Revit.Elements.FamilyType familyType, Featureline featureline, bool useBaseline = false, double station = 0, double offset = 0, double elevation = 0, double angleZ = 0)
        {
            if (run)
            {
                var output = UtilsObjectsLocation.CreateFamilyInstance(familyType, featureline, useBaseline, station, offset, elevation, angleZ);
                return output;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Insert the Revit Link Instances of a give Revit Link Type in the host file.
        /// </summary>
        /// <param name="revitLinkType">Type of the revit link.</param>
        /// <param name="featureline">The featureline.</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <param name="rotate">if set to <c>true</c> [rotate].</param>
        /// <param name="rotation">The rotation.</param>
        /// <returns></returns>
        public static string RevitLinkByStationOffsetElevation(Revit.Elements.Element revitLinkType, Featureline featureline, double station, double offset = 0, double elevation = 0, bool rotate = true, double rotation = 0)
        {
            return UtilsObjectsLocation.RevitLinkByStationOffsetElevation(revitLinkType, featureline, station, offset, elevation, rotate, rotation);
        }

        /// <summary>
        /// Defines the Named site in the Revit Link.
        /// </summary>
        /// <param name="featureline">The featureline.</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <param name="rotate">if set to <c>true</c> [rotate].</param>
        /// <param name="rotation">The rotation.</param>
        /// <returns></returns>
        public static string NamedSiteByStationOffsetElevation(Featureline featureline, double station, double offset = 0, double elevation = 0, bool rotate = true, double rotation = 0)
        {
            return UtilsObjectsLocation.NamedSiteByStationOffsetElevation(featureline, station, offset, elevation, rotate, rotation);
        }

        /// <summary>
        /// Exports the location parameters of Revit Link Instances.
        /// </summary>
        /// <returns></returns>
        public static IList<IList<object>> RevitLinkParameters()
        {
            Utils.Log(string.Format("RevitUtils.RevitLinkParameters started...", ""));

            IList<IList<object>> output = new List<IList<object>>();

            Autodesk.Revit.DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

            output.Add(new List<object>() { "LINK TYPE", "CORRIDOR", "BASELINE", "CODE", "SIDE", "STATION", "OFFSET", "ELEVATION", "ROTATION" });

            foreach (Autodesk.Revit.DB.RevitLinkInstance rli in new Autodesk.Revit.DB.FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.RevitLinkInstance))
                .Cast<Autodesk.Revit.DB.RevitLinkInstance>())
            {
                var typeId = rli.Parameters.Cast<Autodesk.Revit.DB.Parameter>().First(x => x.Id.IntegerValue.Equals((int)Autodesk.Revit.DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)).AsElementId();
                try
                {
                    string typeName = Revit.Elements.ElementSelector.ByElementId(typeId.IntegerValue).Name;
                    string name = rli.Parameters.Cast<Autodesk.Revit.DB.Parameter>().First(x => x.Id.IntegerValue.Equals((int)Autodesk.Revit.DB.BuiltInParameter.RVT_LINK_INSTANCE_NAME)).AsString();
                    string corridor = name.Split(new string[] { "_" }, StringSplitOptions.None)[0];
                    string baseline = name.Split(new string[] { "_" }, StringSplitOptions.None)[1];
                    string code = name.Split(new string[] { "_" }, StringSplitOptions.None)[2];
                    string side = name.Split(new string[] { "_" }, StringSplitOptions.None)[3];
                    string station = name.Split(new string[] { "_" }, StringSplitOptions.None)[4];
                    string offset = name.Split(new string[] { "_" }, StringSplitOptions.None)[5];
                    string elevation = name.Split(new string[] { "_" }, StringSplitOptions.None)[6];
                    string rotation = name.Split(new string[] { "_" }, StringSplitOptions.None)[7];

                    output.Add(new List<object>() { typeName, corridor, baseline, code, side, station, offset, elevation, rotation });
                }
                catch (Exception ex)
                {
                    Utils.Log(string.Format("ERROR: RevitUtils.RevitLinkParameters {0]", ex.Message));

                    continue;
                }
            }

            Utils.Log(string.Format("RevitUtils.RevitLinkParameters completed.", ""));

            return output;
        }

        /// <summary>
        /// Exports the IFC file of the DWG in the folder of the Revit document with in local coordinates.
        /// </summary>
        /// <param name="civilDocument">The civil document.</param>
        /// <param name="desktopConnectorFolder"> The Autodesk Desktop Connector folder for the project on the cloud environment (BIM 360, BIM 360 Team, Fusion 360).</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "IFCOrigin" })]
        public static Dictionary<string, object> ExportIFC(CivilDocument civilDocument, string desktopConnectorFolder = "")
        {
            Utils.Log(string.Format("RevitUtils.ExportIFC started...", ""));

            string folderRVT = "";  // 1.1.0

            try
            {
                folderRVT = Path.GetDirectoryName(DocumentManager.Instance.CurrentDBDocument.PathName);  // 1.1.0
            }
            catch (Exception ex)
            {
                var message = "Save the Revit file first to a local folder.";

                Utils.Log(string.Format("ERROR: RevitUtils.ExportIFC {0} {1}", message, ex.Message));

                throw new Exception(string.Format("{0}\n{1}", message, ex.Message));  // 1.1.0
            }

            if (folderRVT == "")
            {
                var message = "The Revit file path is invalid, save the file first to a local folder.";

                Utils.Log(string.Format("ERROR: RevitUtils.ExportIFC {0}", message));

                throw new Exception(message);  // 1.1.0
            }

            if (desktopConnectorFolder != "")
            {
                desktopConnectorFolder = Path.GetDirectoryName(desktopConnectorFolder);

                folderRVT = desktopConnectorFolder;
            }

            AeccRoadwayDocument mDoc = civilDocument._document;

            string original = mDoc.FullName;

            string ifcOrigin = "";

            IList<AcadEntity> cSolids = new List<AcadEntity>();

            // Civil 3D 2020 Reference to type 'AcadModelSpace' claims it is defined in 'Autodesk.AutoCAD.Interop', but it could not be found

            AcadDatabase db = mDoc as AcadDatabase;

            AcadModelSpace ams = db.ModelSpace;

            if (ams != null)
            {
                for (int i = 0; i < ams.Count; ++i)
                {
                    if (ams.Item(i).EntityName.Contains("Solid") ||
                        ams.Item(i).EntityName.Contains("Body") ||
                        ams.Item(i).EntityName.Contains("Surface") ||
                        ams.Item(i).EntityName.Contains("Face") ||
                        ams.Item(i).EntityName.Contains("MassElement"))
                    {
                        cSolids.Add(ams.Item(i));
                    }
                }
            }
            else
            {
                Utils.Log(string.Format("ERROR: AcadModelSpace is null.", ""));
            }

            var totalTransform = RevitUtils.DocumentTotalTransform();

            if (cSolids.Count > 0)
            {
                Point origin = totalTransform.Origin;

                Autodesk.Revit.DB.ProjectLocation location = DocumentManager.Instance.CurrentDBDocument.ActiveProjectLocation;

                Autodesk.Revit.DB.ProjectPosition position = ProjectPositionUtils.Instance.ProjectPosition;

                double[] end = new double[] { origin.X, origin.Y, origin.Z };

                foreach (AcadEntity cs in cSolids)
                {
                    cs.Rotate(new double[] { 0, 0, 0 }, -position.Angle);

                    cs.Move(new double[] { 0, 0, 0 }, end);
                }

                var name = Path.GetFileNameWithoutExtension(original);

                ifcOrigin = Path.Combine(folderRVT, name + "_Origin.ifc");

                if (File.Exists(ifcOrigin))
                {
                    File.Delete(ifcOrigin);
                }

                mDoc.SendCommand("-IFCEXPORT\nNumber\n\n" + ifcOrigin + "\ne\n");

                end = new double[] { -origin.X, -origin.Y, -origin.Z };

                foreach (AcadEntity cs in cSolids)
                {
                    cs.Move(new double[] { 0, 0, 0 }, end);

                    cs.Rotate(new double[] { 0, 0, 0 }, position.Angle);
                }

                // mDoc.Save();  20200622 avoid to save when the DWG is open in read only mode
            }

            Utils.Log(string.Format("RevitUtils.ExportIFC completed.", ""));

            return new Dictionary<string, object>() { { "IFCOrigin", ifcOrigin } };
        }

        /// <summary>
        /// Replaces the IFC Link with the intermediate RVT document
        /// </summary>
        /// <returns></returns>
        public static bool ReplaceIFCLink(string ifcOrigin, bool keepIFC = true)
        {
            Utils.Log(string.Format("RevitUtils.ReplaceIFCLink started...", ""));

            try
            {
                Autodesk.Revit.DB.RevitLinkType rlt = null;

                Autodesk.Revit.DB.RevitLinkInstance rli = null;

                Autodesk.Revit.DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

                Autodesk.Revit.DB.Document ifcDocument = null;

                string folderRVT = Path.GetDirectoryName(doc.PathName);

                if (folderRVT.StartsWith("BIM 360:\\"))
                {
                    folderRVT = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                }

                string rvtIfcOrigin = ifcOrigin + ".RVT";

                Autodesk.Revit.DB.ElementId rltid = Autodesk.Revit.DB.ElementId.InvalidElementId;

                if (!keepIFC)
                {
                    rltid = Autodesk.Revit.DB.RevitLinkType.GetTopLevelLink(doc, Autodesk.Revit.DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(ifcOrigin));
                }
                else
                {
                    rltid = Autodesk.Revit.DB.RevitLinkType.GetTopLevelLink(doc, Autodesk.Revit.DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(rvtIfcOrigin));
                }

                if (keepIFC && rltid != Autodesk.Revit.DB.ElementId.InvalidElementId)
                {
                    if (File.Exists(rvtIfcOrigin))
                    {
                        File.Delete(rvtIfcOrigin);
                    }

                    rlt = doc.GetElement(rltid) as Autodesk.Revit.DB.RevitLinkType;

                    rlt.Unload(null);

                    Autodesk.Revit.ApplicationServices.Application revitApp = DocumentManager.Instance.CurrentDBDocument.Application;

                    string template = revitApp.DefaultProjectTemplate;

                    var ifcTemplate = revitApp.DefaultIFCProjectTemplate;

                    if (File.Exists(rvtIfcOrigin))
                    {
                        try
                        {
                            ifcDocument = doc.Application.OpenDocumentFile(rvtIfcOrigin);
                        }
                        catch (Exception)
                        {
                            ifcDocument = doc.Application.NewProjectDocument(ifcTemplate == "" ? template : ifcTemplate);
                        }
                    }
                    else
                    {
                        ifcDocument = doc.Application.NewProjectDocument(ifcTemplate == "" ? template : ifcTemplate);

                        ifcDocument.SaveAs(rvtIfcOrigin, new Autodesk.Revit.DB.SaveAsOptions() { MaximumBackups = 1, OverwriteExistingFile = true, Compact = true });
                    }

                    ifcDocument.Close(false);

                    // Check if the intermediate Revit file has been loaded already
                    rltid = Autodesk.Revit.DB.RevitLinkType.GetTopLevelLink(doc, Autodesk.Revit.DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(rvtIfcOrigin));

                    // Check if the IFC file is still loaded instead
                    if (rltid == Autodesk.Revit.DB.ElementId.InvalidElementId)
                    {
                        rltid = Autodesk.Revit.DB.RevitLinkType.GetTopLevelLink(doc, Autodesk.Revit.DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(ifcOrigin));
                    }

                    // If the rltId is still Invalid it means that it is the setup cycle
                    TransactionManager.Instance.EnsureInTransaction(doc);

                    if (rltid == Autodesk.Revit.DB.ElementId.InvalidElementId)
                    {
                        var res = Autodesk.Revit.DB.RevitLinkType.CreateFromIFC(doc, ifcOrigin, rvtIfcOrigin, true, new Autodesk.Revit.DB.RevitLinkOptions(true));

                        rltid = res.ElementId;
                    }
                    else
                    {
                        rlt = doc.GetElement(rltid) as Autodesk.Revit.DB.RevitLinkType;

                        rlt.UpdateFromIFC(doc, ifcOrigin, rvtIfcOrigin, true);
                    }

                    if (rltid != Autodesk.Revit.DB.ElementId.InvalidElementId)
                    {

                        try
                        {
                            rli = new Autodesk.Revit.DB.FilteredElementCollector(doc)
                                .OfClass(typeof(Autodesk.Revit.DB.RevitLinkInstance))
                                .WhereElementIsNotElementType()
                                .Cast<Autodesk.Revit.DB.RevitLinkInstance>()
                                .First(x => x.GetTypeId().Equals(rltid));
                        }
                        catch { }

                        if (rli == null)
                        {
                            rli = Autodesk.Revit.DB.RevitLinkInstance.Create(doc, rltid);
                        }
                    }

                    TransactionManager.Instance.TransactionTaskDone();

                    TransactionManager.Instance.ForceCloseTransaction();

                    rlt = doc.GetElement(rltid) as Autodesk.Revit.DB.RevitLinkType;

                    rlt.Reload();
                }

                else
                {
                    TransactionManager.Instance.EnsureInTransaction(doc);

                    if (rltid != Autodesk.Revit.DB.ElementId.InvalidElementId)
                    {

                        try
                        {
                            rli = new Autodesk.Revit.DB.FilteredElementCollector(doc)
                                .OfClass(typeof(Autodesk.Revit.DB.RevitLinkInstance))
                                .WhereElementIsNotElementType()
                                .Cast<Autodesk.Revit.DB.RevitLinkInstance>()
                                .First(x => x.GetTypeId().Equals(rltid));
                        }
                        catch { }

                        if (rli != null)
                        {
                            doc.Delete(rli.Id);
                        }

                        doc.Delete(rltid);
                    }

                    rltid = Autodesk.Revit.DB.ElementId.InvalidElementId;

                    Autodesk.Revit.DB.ModelPath mp = Autodesk.Revit.DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(rvtIfcOrigin);

                    var result2 = Autodesk.Revit.DB.RevitLinkType.Create(doc, mp, new Autodesk.Revit.DB.RevitLinkOptions(true));

                    rltid = result2.ElementId;

                    if (rltid != Autodesk.Revit.DB.ElementId.InvalidElementId)
                    {
                        rli = Autodesk.Revit.DB.RevitLinkInstance.Create(doc, rltid);
                    }

                    TransactionManager.Instance.TransactionTaskDone();
                }
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: RevitUtils.ReplaceIFCLink {0}", ex.Message));

                throw new Exception(string.Format("The IFC Replacement failed\n\n{0}", ex.Message));
            }

            Utils.Log(string.Format("RevitUtils.ReplaceIFCLink completed.", ""));

            return true;
        }

        /// <summary>
        /// Creates a wall from a Dynamo surface.
        /// The wall is recreated but not updated. The input surface must be planar and its normal must be orthogonal to the world Z Axis.
        /// </summary>
        /// <param name="surface">The surface.</param>
        /// <param name="wallType">Type of the wall.</param>
        /// <param name="structural">if set to <c>true</c> [structural].</param>
        /// <returns></returns>
        public static Wall WallBySurface(Surface surface, WallType wallType, bool structural = true)
        {
            return UtilsObjectsLocation.WallBySurface(surface, wallType, structural);
        }

        #endregion
    }

}
