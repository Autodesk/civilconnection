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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using RevitServices.Persistence;
using Autodesk.DesignScript.Runtime;
// TODO: Add Dynamo references to extract the profile via Dynamo Solids / intersection with the view plane

namespace CivilConnection
{
    /// <summary>
    /// Collection of utilities for SectionViews.
    /// </summary>
    [SupressImportIntoVM()]
    public class UtilsSectionView
    {
        /// <summary>
        /// Returns the cut lines.
        /// </summary>
        /// <param name="viewId">The view identifier.</param>
        public static void CutLines(ElementId viewId)
        {
            Utils.Log(string.Format("UtilsSectionView.CutLines started...", ""));

            var uidoc = DocumentManager.Instance.CurrentUIDocument;
            var doc = uidoc.Document;
            var view = doc.GetElement(viewId) as View;
            var app = doc.Application;

            Options opt = new Options();
            opt.View = view;

            BoundingBoxXYZ crop = view.CropBox;

            IList<DetailCurve> detailCurves = new List<DetailCurve>();
            IList<Curve> linkedCurves = new List<Curve>();

            Transform tr = view.CropBox.Transform;

            string path = Path.Combine(Path.GetTempPath(), "Copy.rvt");

            var rliCollection = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .WhereElementIsNotElementType()
                .Cast<RevitLinkInstance>();

            Document link = null;

            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            CloseCopy();

            Utils.Log(string.Format("Preparing Revit Links...", ""));

            if (rliCollection.Count() > 0)
            {
                #region RevitLinkInstances
                foreach (RevitLinkInstance rli in rliCollection)
                {
                    ElementId efrId = rli.GetTypeId();

                    var status = ExternalFileUtils.GetExternalFileReference(doc, efrId).GetLinkedFileStatus();

                    Utils.Log(string.Format("Status {0}...", status));

                    if (status == LinkedFileStatus.InClosedWorkset ||
                        status == LinkedFileStatus.Invalid ||
                        status == LinkedFileStatus.LocallyUnloaded ||
                        status == LinkedFileStatus.NotFound ||
                        status == LinkedFileStatus.Unloaded)
                    {
                        continue;
                    }
                    app.CopyModel(ModelPathUtils.ConvertUserVisiblePathToModelPath(rli.GetLinkDocument().PathName), path, true);

                    link = app.OpenDocumentFile(path);
                    // access projectlocations
                    // test if the the project location in the host file exists in the linked document
                    // if exists retrieve the transform object via project position
                    // apply the transform to the total transform of the revit link instance

                    Utils.Log(string.Format("Processing {0}...", path));

                    Transform rliTransform = rli.GetTotalTransform();

                    // TODO: include angle in the calculation
                    // TODO: include different project location from the Link Instance
                    // TODO: test for different locations of the same instance

                    tr.Origin = rliTransform.Inverse.OfPoint(view.Origin);
                    tr.BasisX = rliTransform.Inverse.OfVector(-view.RightDirection);
                    tr.BasisY = XYZ.BasisZ;
                    tr.BasisZ = tr.BasisX.CrossProduct(tr.BasisY);

                    BoundingBoxXYZ bb = new BoundingBoxXYZ();
                    bb.Transform = tr;
                    XYZ min = view.CropBox.Min;
                    XYZ max = view.CropBox.Max;
                    bb.Min = new XYZ(min.X, min.Y, 0);
                    bb.Max = new XYZ(max.X, max.Y, -min.Z);

                    IList<ElementId> detailCurveIds = new List<ElementId>();

                    var linkedViewId = ElementId.InvalidElementId;

                    // TODO: create other view family types, currently works only for sections

                    using (Transaction q = new Transaction(link, "Cut1"))
                    {
                        q.Start();

                        ElementId vftId = new FilteredElementCollector(link)
                            .OfClass(typeof(ViewFamilyType))
                            .WhereElementIsElementType()
                            .Cast<ViewFamilyType>()
                            .First(x => x.ViewFamily == ViewFamily.Section)
                            .Id;

                        View linkedView = ViewSection.CreateSection(link, vftId, bb);

                        link.Regenerate();

                        foreach (Curve c in CutCurvesInView(link, linkedView.Id))
                        {
                            linkedCurves.Add(c.CreateTransformed(rliTransform));
                        }

                        q.RollBack();
                    }

                    link.Close(false);

                    Utils.Log(string.Format("Completed", ""));
                }
                #endregion
            }

           

            Utils.Log(string.Format("Processing current document...", ""));

            using (Transaction t = new Transaction(doc, "Cut"))
            {
                t.Start();

                #region Current Document Detail Curves

                IList<ElementId> toDelete = new List<ElementId>();

                foreach (Group g in new FilteredElementCollector(doc, view.Id)
                        .OfClass(typeof(Group))
                        .WhereElementIsNotElementType()
                        .Cast<Group>()
                        .Where(x => x.GroupType.Name == view.Name))
                {
                    g.Pinned = false;

                    toDelete.Add(g.Id);
                }

                if (toDelete.Count > 0)
                {
                    doc.Delete(toDelete);
                    toDelete.Clear();
                }

                Utils.Log(string.Format("Deleted {0} groups.", toDelete.Count));

                GraphicsStyle gs = new FilteredElementCollector(doc)
                    .OfClass(typeof(GraphicsStyle))
                    .WhereElementIsNotElementType()
                    .Cast<GraphicsStyle>()
                    .First(x => x.GraphicsStyleCategory.Name == "Medium Lines");  // TODO make it usable for different languages as well

                Utils.Log(string.Format("Preparing for extracting curves...", ""));

                foreach (Curve c in CutCurvesInView(doc, view.Id))
                {
                    try
                    {
                        var dc = doc.Create.NewDetailCurve(view, c);
                        if (null != dc)
                        {
                            detailCurves.Add(dc);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (linkedCurves.Count > 0)
                {
                    foreach (Curve c in linkedCurves.Where(x => x.Length > doc.Application.ShortCurveTolerance))
                    {
                        try
                        {
                            var dc = doc.Create.NewDetailCurve(view, c);
                            if (null != dc)
                            {
                                detailCurves.Add(dc);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                foreach (DetailCurve dc in detailCurves)
                {
                    dc.LineStyle = gs;
                }

                // this is necessary otherwise a warning pops up for editing a group outside of scope
                doc.Regenerate();
                #endregion

                #region Group
                foreach (GroupType gt in new FilteredElementCollector(doc)
                        .OfClass(typeof(GroupType))
                        .WhereElementIsElementType()
                        .Where(x => x.Name == view.Name))
                {
                    toDelete.Add(gt.Id);
                }

                if (toDelete.Count > 0)
                {
                    doc.Delete(toDelete);
                }

                if (detailCurves.Count > 0)
                {
                    Group g = doc.Create.NewGroup(detailCurves.Select(x => x.Id).ToList());
                    g.GroupType.Name = view.Name;
                    g.Pinned = true;
                }
                #endregion

                t.Commit();
            }

            CloseCopy();

            Utils.Log(string.Format("UtilsSectionView.CutLines completed.", ""));
        }

        /// <summary>
        /// Closes the copy of the auxilliaary Revit linked file.
        /// </summary>
        private static void CloseCopy()
        {
            Utils.Log(string.Format("UtilsSectionView.CloseCopy started...", ""));

            var app = DocumentManager.Instance.CurrentUIApplication.Application;

            string path = Path.Combine(Path.GetTempPath(), "Copy.rvt");

            if (File.Exists(path))
            {
                foreach (Document doc in app.Documents)
                {
                    if (doc.PathName == path)
                    {
                        doc.Close(false);
                        break;
                    }
                }

                File.Delete(path);
            }

            Utils.Log(string.Format("UtilsSectionView.CloseCopy completed.", ""));
        }

        /// <summary>
        /// Optimizes the curves.
        /// </summary>
        /// <param name="ca">The ca.</param>
        /// <param name="cb">The cb.</param>
        /// <returns></returns>
        private static Curve SketchOptimizer(Curve ca, Curve cb)
        {
            Utils.Log(string.Format("UtilsSectionView.SketchOptimizer started...", ""));

            Curve cc = null;
            Dictionary<double, Curve> pairs = new Dictionary<double, Curve>();
            if (!ca.IsCyclic && !cb.IsCyclic &&
               ca.Intersect(cb) != SetComparisonResult.Disjoint &&
               (ca.GetEndPoint(1) - ca.GetEndPoint(0))
               .CrossProduct(cb.GetEndPoint(1) - cb.GetEndPoint(0))
               .IsAlmostEqualTo(XYZ.Zero)
               &&
               Math.Round(ca.Length, 5).Equals(Math.Round(ca.GetEndPoint(1).DistanceTo(ca.GetEndPoint(0)), 5)) &&
               Math.Round(cb.Length, 5).Equals(Math.Round(cb.GetEndPoint(1).DistanceTo(cb.GetEndPoint(0)), 5))
              )
            {
                List<Line> ll = new List<Line>();
                try
                {
                    Line l1 = Line.CreateBound(ca.GetEndPoint(0), ca.GetEndPoint(1));
                    ll.Add(l1);
                }
                catch { }
                try
                {
                    Line l2 = Line.CreateBound(ca.GetEndPoint(0), cb.GetEndPoint(0));
                    ll.Add(l2);
                }
                catch { }
                try
                {
                    Line l3 = Line.CreateBound(ca.GetEndPoint(0), cb.GetEndPoint(1));
                    ll.Add(l3);
                }
                catch { }
                try
                {
                    Line l4 = Line.CreateBound(ca.GetEndPoint(1), cb.GetEndPoint(0));
                    ll.Add(l4);
                }
                catch { }
                try
                {
                    Line l5 = Line.CreateBound(ca.GetEndPoint(1), cb.GetEndPoint(1));
                    ll.Add(l5);
                }
                catch { }
                try
                {
                    Line l6 = Line.CreateBound(cb.GetEndPoint(0), cb.GetEndPoint(1));
                    ll.Add(l6);
                }
                catch { }
                ll = ll.GroupBy(x => x.Length).Select(g => g.First()).ToList();
                foreach (Line l in ll)
                {
                    pairs.Add(l.Length, l as Curve);
                }
                cc = pairs.Values.First(x => x.Length == pairs.Keys.Max());
            }
            else if (ca.IsCyclic &&
                     cb.IsCyclic &&
                     ((Math.Round(ca.Project(cb.GetEndPoint(0)).XYZPoint.DistanceTo(cb.GetEndPoint(0)), 5).Equals(0) ||
                       Math.Round(ca.Project(cb.GetEndPoint(1)).XYZPoint.DistanceTo(cb.GetEndPoint(1)), 5).Equals(0)) ||
                      (Math.Round(cb.Project(ca.GetEndPoint(0)).XYZPoint.DistanceTo(ca.GetEndPoint(0)), 5).Equals(0) ||
                       Math.Round(cb.Project(ca.GetEndPoint(1)).XYZPoint.DistanceTo(ca.GetEndPoint(1)), 5).Equals(0))))
            {

                Arc aa = ca as Arc;
                Arc ab = cb as Arc;
                List<Arc> arcs = new List<Arc>();
                if (ab.Normal.Negate().IsAlmostEqualTo(aa.Normal))
                {
                    if (aa.Length > ab.Length)
                    {
                        aa = Arc.Create(ca.GetEndPoint(1), ca.GetEndPoint(0), ca.Evaluate(.5, true));
                    }
                    else
                    {
                        ab = Arc.Create(cb.GetEndPoint(1), cb.GetEndPoint(0), cb.Evaluate(.5, true));
                    }
                }
                if (aa.Center.IsAlmostEqualTo(ab.Center) &&
                   Math.Round(aa.Radius, 5).Equals(Math.Round(ab.Radius, 5)))
                {

                    foreach (Arc arc in ArcsList(aa, ab))
                    {
                        arcs.Add(arc);
                    }
                    foreach (Arc arc in ArcsList(ab, aa))
                    {
                        arcs.Add(arc);
                    }
                }
                if (arcs.Count < 1)
                {
                    aa = Arc.Create(ca.GetEndPoint(0), ca.Evaluate(.5, true), ca.Evaluate(.25, true));
                    foreach (Arc arc in ArcsList(aa, ab))
                    {
                        arcs.Add(arc);
                    }
                    foreach (Arc arc in ArcsList(ab, aa))
                    {
                        arcs.Add(arc);
                    }
                }
                arcs = arcs.GroupBy(x => x.Length).Select(g => g.First()).ToList();
                foreach (Arc l in arcs)
                {
                    pairs.Add(l.Length, l as Curve);
                }
                cc = pairs.Values.First(x => x.Length == pairs.Keys.Min());
            }

            Utils.Log(string.Format("UtilsSectionView.SketchOptimizer completed.", ""));

            return cc;
        }

        /// <summary>
        /// Returns a list of optimized arcs.
        /// </summary>
        /// <param name="aa">The aa.</param>
        /// <param name="ab">The ab.</param>
        /// <returns></returns>
        private static List<Arc> ArcsList(Arc aa, Arc ab)
        {
            Utils.Log(string.Format("UtilsSectionView.ArcsList started...", ""));

            List<Arc> arcs = new List<Arc>();
            XYZ Saa = aa.GetEndPoint(0);
            XYZ Eaa = aa.GetEndPoint(1);
            XYZ Sab = ab.GetEndPoint(0);
            XYZ Eab = ab.GetEndPoint(1);
            try
            {
                Arc a1 = Arc.Create(Saa, Eaa, Sab);
                if (Math.Round(a1.Project(Eab).XYZPoint.DistanceTo(Eab), 5) == 0)
                    arcs.Add(a1);
            }
            catch { }
            try
            {
                Arc a2 = Arc.Create(Saa, Sab, Eaa);
                if (Math.Round(a2.Project(Eab).XYZPoint.DistanceTo(Eab), 5) == 0)
                    arcs.Add(a2);
            }
            catch { }
            try
            {
                Arc a3 = Arc.Create(Saa, Eab, Eaa);
                if (Math.Round(a3.Project(Sab).XYZPoint.DistanceTo(Sab), 5) == 0)
                    arcs.Add(a3);
            }
            catch { }

            Utils.Log(string.Format("UtilsSectionView.ArcsList completed.", ""));

            return arcs;
        }

        /// <summary>
        /// Returns the cutting curves in the view.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <param name="viewId">The view identifier.</param>
        /// <returns></returns>
        private static IList<Curve> CutCurvesInView(Document doc, ElementId viewId)
        {
            Utils.Log(string.Format("UtilsSectionView.CutCurvesInView started...", ""));

            IList<Curve> curves = new List<Curve>();

            if (!doc.IsFamilyDocument)
            {
                View view = doc.GetElement(viewId) as View;

                Options opt = new Options();
                opt.View = view;
                opt.ComputeReferences = true;
                opt.IncludeNonVisibleObjects = true;

                foreach (Category cat in doc.Settings.Categories)
                {
                    if (cat.IsCuttable)
                    {
                        foreach (Element e in new FilteredElementCollector(doc, viewId).OfCategoryId(cat.Id))
                        {

                            GeometryElement ge = e.get_Geometry(opt);

                            foreach (GeometryObject go in ge)
                            {
                                if (go is Solid)
                                {
                                    Solid s = go as Solid;
                                    foreach (Curve c in SolidEdgesToCurve(doc, s, view, cat))
                                    {
                                        curves.Add(c);
                                    }
                                }
                                else if (go is GeometryInstance)
                                {
                                    GeometryInstance gi = go as GeometryInstance;

                                    foreach (GeometryObject gio in gi.GetInstanceGeometry())
                                    {
                                        if (gio is Solid)
                                        {
                                            Solid si = gio as Solid;
                                            foreach (Curve c in SolidEdgesToCurve(doc, si, view, cat))
                                            {
                                                curves.Add(c);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            IList<Curve> output = new List<Curve>();

            while (curves.Count > 0)
            {
                Curve f = curves.ElementAt(0);
                curves.RemoveAt(0);
                for (int i = 0; i < curves.Count; ++i)
                {
                    Curve q = null;
                    try
                    {
                        q = SketchOptimizer(f, curves[i]);
                    }
                    catch { }
                    if (null != q)
                    {
                        curves.RemoveAt(i);
                        f = q;
                    }
                }
                output.Add(f);
            }

            Utils.Log(string.Format("UtilsSectionView.CutCurvesInView completed.", ""));

            return output;
        }


        /// <summary>
        /// Returns the Solid edges in the View plane.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="s"></param>
        /// <param name="view"></param>
        /// <param name="cat"></param>
        /// <returns></returns>
        private static IList<Curve> SolidEdgesToCurve(Document doc, Solid s, View view, Category cat)
        {
            Utils.Log(string.Format("UtilsSectionView.SolidEdgesToCurve started...", ""));

            IList<Curve> curves = new List<Curve>();

            if (s.Volume <= 0)
            {
                return curves;
            }

            Options opt = new Options();
            opt.View = view;
            opt.ComputeReferences = true;
            opt.IncludeNonVisibleObjects = true;

            Transform tr = Transform.Identity;
            tr.Origin = view.Origin;
            tr.BasisX = view.RightDirection;
            tr.BasisY = view.UpDirection;
            tr.BasisZ = tr.BasisX.CrossProduct(tr.BasisY);

            foreach (Edge edge in s.Edges)
            {
                GraphicsStyle gse = doc.GetElement(edge.GraphicsStyleId) as GraphicsStyle;

                Curve c = edge.AsCurve();
                XYZ p = c.Evaluate(0.5, true);

                XYZ g = tr.Inverse.OfPoint(p);

                if (Math.Abs(g.Z) < 0.00001)
                {

                    if (c is HermiteSpline)
                    {
                        var hs = c as HermiteSpline;
                        IList<XYZ> pts = c.Tessellate();
                        for (int i = 0; i < pts.Count - 1; ++i)
                        {
                            curves.Add(Line.CreateBound(pts[i], pts[i + 1]));
                        }
                    }
                    else
                    {
                        curves.Add(c);
                    }
                }
            }

            Utils.Log(string.Format("UtilsSectionView.SolidEdgesToCurve completed.", ""));

            return curves;
        }
    }
}
