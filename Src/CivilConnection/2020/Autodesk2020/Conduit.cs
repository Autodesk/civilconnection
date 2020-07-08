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
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Revit.GeometryConversion;
using RevitServices.Persistence;
using RevitServices.Transactions;
using System;
using System.Collections.Generic;
using System.Linq;
using ADSK_Parameters = CivilConnection.UtilsObjectsLocation.ADSK_Parameters;


namespace CivilConnection.MEP
{
    /// <summary>
    /// Conduit object type.
    /// </summary>
    /// <seealso cref="CivilConnection.AbstractMEPCurve" />
    [DynamoServices.RegisterForTrace()]
    public class Conduit : AbstractMEPCurve
    {

        #region PUBLIC PROPERTIES

        /// <summary>
        /// Gets the diameter.
        /// </summary>
        /// <value>
        /// The diameter.
        /// </value>
        public double Diameter
        {
            get
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                double d = InternalMEPCurve.Diameter;
                TransactionManager.Instance.TransactionTaskDone();
                return UtilsObjectsLocation.FeetToMm(d);
            }
        }

        /// <summary>
        /// Gets the run.
        /// </summary>
        /// <value>
        /// The run.
        /// </value>
        public string Run
        {
            get
            {
                TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
                DocumentManager.Regenerate();
                var c = InternalMEPCurve as Autodesk.Revit.DB.Electrical.Conduit;
                var run = DocumentManager.Instance.CurrentDBDocument.GetElement(c.RunId) as ConduitRun;
                TransactionManager.Instance.TransactionTaskDone();
                return run.Name;
            }
        }

        

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="Conduit"/> class.
        /// </summary>
        /// <param name="instance">The instance.</param>
        protected Conduit(Autodesk.Revit.DB.Electrical.Conduit instance)
        {
            SafeInit(() => InitObject(instance));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Conduit"/> class.
        /// </summary>
        /// <param name="conduitType">Type of the conduit.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        internal Conduit(Autodesk.Revit.DB.Electrical.ConduitType conduitType, XYZ start, XYZ end)
        {
            InitObject(conduitType, start, end);
        }

        /// <summary>
        /// Returns an empty Conduit
        /// </summary>
        internal Conduit()
        { }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Initialize a Conduit element
        /// </summary>
        /// <param name="instance">The instance.</param>
        private void InitObject(Autodesk.Revit.DB.Electrical.Conduit instance)
        {
            Autodesk.Revit.DB.MEPCurve fi = instance as Autodesk.Revit.DB.MEPCurve;
            InternalSetMEPCurve(fi);
        }

        /// <summary>
        /// Initialize a Conduit element
        /// </summary>
        /// <param name="conduitType">Type of the conduit.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        private void InitObject(Autodesk.Revit.DB.Electrical.ConduitType conduitType, XYZ start, XYZ end)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.MEPCurve>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetMEPCurve(oldFam);
                InternalSetMEPCurveType(conduitType);
                InternalSetPosition(start, end);
                return;
            }

            //Phase 2- There was no existing point, create one
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.Revit.DB.MEPCurve fi;

            if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
            {
                fi = null;
            }
            else
            {
                fi = Autodesk.Revit.DB.Electrical.Conduit.Create(DocumentManager.Instance.CurrentDBDocument, conduitType.Id, start, end, ElementId.InvalidElementId);
            }

            InternalSetMEPCurve(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Gets the conduit by ids.
        /// </summary>
        /// <param name="ids">The ids.</param>
        /// <returns></returns>
        private static Conduit[] GetConduitByIds(ICollection<ElementId> ids)
        {
            var pipes = new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                  .OfClass(typeof(Autodesk.Revit.DB.Electrical.Conduit))
                  .WhereElementIsNotElementType()
                  .Cast<Autodesk.Revit.DB.Electrical.Conduit>()
                  .Where(x => ids.Contains(x.Id))
                  .ToList();

            Conduit[] output = new Conduit[pipes.Count];

            int c = 0;

            foreach (var p in pipes)
            {
                output[c] = new Conduit(p);

                c = c + 1;
            }

            return output;
        }

        #region NOT WORKING
        private static Conduit[] ByPolyCurve(Revit.Elements.Element conduitType, PolyCurve polyCurve, double maxLength)
        {
            Utils.Log(string.Format("Conduit.ByPolyCurve started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            var oType = conduitType.InternalElement as Autodesk.Revit.DB.Electrical.ConduitType;

            double length = polyCurve.Length;

            double subdivisions = Math.Ceiling(length / maxLength);
            double increment = 1 / subdivisions;

            IList<double> parameters = new List<double>();

            double parameter = 0;

            IList<Autodesk.DesignScript.Geometry.Point> points = new List<Autodesk.DesignScript.Geometry.Point>();

            while (parameter <= 1)
            {
                points.Add(polyCurve.PointAtParameter(parameter));
                parameter = parameter + increment;
            }

            points.Add(polyCurve.EndPoint);

            points = Autodesk.DesignScript.Geometry.Point.PruneDuplicates(points);  // this is slow

            IList<ElementId> ids = new List<ElementId>();

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            Autodesk.DesignScript.Geometry.Point start = null;
            Autodesk.DesignScript.Geometry.Point end = null;

            for (int i = 0; i < points.Count - 1; ++i)
            {
                start = points[i].Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                var s = start.ToXyz();
                end = points[i + 1].Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                var e = end.ToXyz();

                Autodesk.Revit.DB.Electrical.Conduit p = Autodesk.Revit.DB.Electrical.Conduit.Create(DocumentManager.Instance.CurrentDBDocument, oType.Id, s, e, ElementId.InvalidElementId);
                ids.Add(p.Id);
            }

            var res = GetConduitByIds(ids);

            for (int i = 0; i < GetConduitByIds(ids).Length - 1; ++i)
            {
                Conduit ct1 = res[i];
                Conduit ct2 = res[i + 1];
                Fitting.Elbow(ct1, ct2);
            }

            TransactionManager.Instance.TransactionTaskDone();

            if (start != null)
            {
                start.Dispose();
            }
            if (end != null)
            {
                end.Dispose();
            }

            foreach(var pt in points)
            {
                if (pt != null)
                {
                    pt.Dispose();
                }
            }

            points.Clear();

            Utils.Log(string.Format("Conduit.ByPolyCurve completed.", ""));

            return res;
        }

        /// <summary>
        /// Create Conduits following a polycurve
        /// </summary>
        /// <param name="conduitType">The conduit type.</param>
        /// <param name="polyCurve">the Polycurve to follow.</param>
        /// <returns></returns>
        private static Conduit[] ByPolyCurve(Revit.Elements.Element conduitType, PolyCurve polyCurve)
        {
            Utils.Log(string.Format("Conduit.ByPolyCurve started...", ""));

            var oType = conduitType.InternalElement as Autodesk.Revit.DB.Electrical.ConduitType;

            IList<ElementId> ids = new List<ElementId>();

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            var curves = polyCurve.Curves().ToList();

            foreach (Autodesk.DesignScript.Geometry.Curve c in curves)
            {
                Conduit ct = Conduit.ByCurve(conduitType, c);
                ids.Add(ct.InternalMEPCurve.Id);
            }

            for (int i = 0; i < GetConduitByIds(ids).Length - 1; ++i)
            {
                Conduit ct1 = GetConduitByIds(ids)[i];
                Conduit ct2 = GetConduitByIds(ids)[i + 1];
                Fitting.Elbow(ct1, ct2);
            }

            TransactionManager.Instance.TransactionTaskDone();

            foreach (var c in curves)
            {
                if (c != null)
                {
                    c.Dispose();
                }
            }

            curves.Clear();

            Utils.Log(string.Format("Conduit.ByPolyCurve completed.", ""));

            return GetConduitByIds(ids);
        }

        /// <summary>
        /// Creates a list of Conduits from a PolyCurve.
        /// </summary>
        /// <param name="conduitType">Type of the conduit.</param>
        /// <param name="polyCurve">The poly curvein WCS.</param>
        /// <param name="maxLength">The maximum length.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Conduit", "Fittings" })]
        public static Dictionary<string, object> ByPolyCurve(Revit.Elements.Element conduitType, Autodesk.DesignScript.Geometry.PolyCurve polyCurve, double maxLength, Featureline featureline)
        {
            Utils.Log(string.Format("Conduit.ByPolyCurve started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument);
            }

            var totalTransform = RevitUtils.DocumentTotalTransform();

            var oType = conduitType.InternalElement as Autodesk.Revit.DB.Electrical.ConduitType;
            IList<Conduit> output = new List<Conduit>();
            IList<Fitting> fittings = new List<Fitting>();

            double length = polyCurve.Length;

            int subdivisions = Convert.ToInt32(Math.Ceiling(length / maxLength));

            IList<Autodesk.DesignScript.Geometry.Point> points = new List<Autodesk.DesignScript.Geometry.Point>();

            points.Add(polyCurve.StartPoint);

            foreach (Autodesk.DesignScript.Geometry.Point p in polyCurve.PointsAtEqualChordLength(subdivisions))
            {
                points.Add(p);
            }

            points.Add(polyCurve.EndPoint);

            points = Autodesk.DesignScript.Geometry.Point.PruneDuplicates(points);

            Autodesk.DesignScript.Geometry.Point start = null;
            Autodesk.DesignScript.Geometry.Point end = null;
            Autodesk.DesignScript.Geometry.Point sp = null;
            Autodesk.DesignScript.Geometry.Point ep = null;
            Autodesk.DesignScript.Geometry.Curve curve = null;

            for (int i = 0; i < points.Count - 1; ++i)
            {
                start = points[i];
                end = points[i + 1];

                curve = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(start, end);

                sp = start.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                ep = end.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;

                var pipe = new Conduit(oType, sp.ToXyz(), ep.ToXyz());

                pipe.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);
                pipe.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
                pipe.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(start.X, 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(start.Y, 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(start.Z, 3));
                var soe = featureline.GetStationOffsetElevationByPoint(start);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round((double)soe["Station"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round((double)soe["Offset"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round((double)soe["Elevation"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);
                soe = featureline.GetStationOffsetElevationByPoint(end);
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, Math.Round((double)soe["Station"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndOffset.Name, Math.Round((double)soe["Offset"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndElevation.Name, Math.Round((double)soe["Elevation"], 3));

                output.Add(pipe);

                if (start != null)
                {
                    start.Dispose();
                }
                if (end != null)
                {
                    end.Dispose();
                }
                if (sp != null)
                {
                    sp.Dispose();
                }
                if (ep != null)
                {
                    ep.Dispose();
                }
                if (curve != null)
                {
                    curve.Dispose();
                }
            }

            for (int i = 0; i < output.Count - 1; ++i)
            {
                Fitting fitting = null;
                try
                {
                    fitting = Fitting.Elbow(output[i], output[i + 1]);
                }
                catch { }

                fittings.Add(fitting);
            }

            foreach (var pt in points)
            {
                if (pt != null)
                {
                    pt.Dispose();
                }
            }

            points.Clear();

            Utils.Log(string.Format("Conduit.ByPolyCurve completed.", ""));

            return new Dictionary<string, object>() { { "Conduit", output }, { "Fittings", fittings } };

        }
        #endregion

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Creates a Conduit by two points.
        /// </summary>
        /// <param name="conduitType">Type of the conduit.</param>
        /// <param name="start">The start point.</param>
        /// <param name="end">The end point.</param>
        /// <returns></returns>
        public static Conduit ByPoints(Revit.Elements.Element conduitType, Autodesk.DesignScript.Geometry.Point start, Autodesk.DesignScript.Geometry.Point end)
        {
            Utils.Log(string.Format("Conduit.ByPoints started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }
            var oType = conduitType.InternalElement as Autodesk.Revit.DB.Electrical.ConduitType;
            var nstart = start.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = nstart.ToXyz();
            var nend = end.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = nend.ToXyz();

            if (nstart != null)
            {
                nstart.Dispose();
            }
            if (nend != null)
            {
                nend.Dispose();
            }

            Utils.Log(string.Format("Conduit.ByPoints completed.", ""));

            return new Conduit(oType, s, e);
        }

        /// <summary>
        /// Creates a Conduit by a curve.
        /// </summary>
        /// <param name="conduitType">Type of the conduit.</param>
        /// <param name="curve">The curve.</param>
        /// <returns></returns>
        public static Conduit ByCurve(Revit.Elements.Element conduitType, Autodesk.DesignScript.Geometry.Curve curve)
        {
            Utils.Log(string.Format("Conduit.ByCurve started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }
            var oType = conduitType.InternalElement as Autodesk.Revit.DB.Electrical.ConduitType;
            Autodesk.DesignScript.Geometry.Point start = curve.StartPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = start.ToXyz();
            Autodesk.DesignScript.Geometry.Point end = curve.EndPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = end.ToXyz();

            if (start != null)
            {
                start.Dispose();
            }
            if (end != null)
            {
                end.Dispose();
            }

            Utils.Log(string.Format("Conduit.ByCurve completed.", ""));

            return new Conduit(oType, s, e);
        }

        /// <summary>
        /// Creates a Conduit by a curve.
        /// </summary>
        /// <param name="conduitType">Type of the conduit.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns></returns>
        public static Conduit ByCurveFeatureline(Revit.Elements.Element conduitType, Autodesk.DesignScript.Geometry.Curve curve, Featureline featureline)
        {
            Utils.Log(string.Format("Conduit.ByCurveFeatureline started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }

            var pipe = Conduit.ByCurve(conduitType, curve);  //  new Conduit(oType, s, e);

            var start = curve.StartPoint;
            var end = curve.EndPoint;

            var startSOE = featureline.GetStationOffsetElevationByPoint(start);
            var endSOE = featureline.GetStationOffsetElevationByPoint(end);

            double startStation = (double)startSOE["Station"];
            double startOffset = (double)startSOE["Offset"];
            double startElevation = (double)startSOE["Elevation"];
            double endStation = (double)endSOE["Station"];
            double endOffset = (double)endSOE["Offset"];
            double endElevation = (double)endSOE["Elevation"];

            pipe.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);
            pipe.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);
            pipe.SetParameterByName(ADSK_Parameters.Instance.RegionIndex.Name, featureline.BaselineRegionIndex);  // 1.1.0
            pipe.SetParameterByName(ADSK_Parameters.Instance.RegionRelative.Name, startStation - featureline.Start);  // 1.1.0
            pipe.SetParameterByName(ADSK_Parameters.Instance.RegionNormalized.Name, (startStation - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0
            pipe.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
            pipe.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
            pipe.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(curve.StartPoint.X, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(curve.StartPoint.Y, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(curve.StartPoint.Z, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round(startStation, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round(startOffset, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round(startElevation, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
            pipe.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);
            pipe.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, Math.Round(endStation, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.EndOffset.Name, Math.Round(endOffset, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.EndElevation.Name, Math.Round(endElevation, 3));
            pipe.SetParameterByName(ADSK_Parameters.Instance.EndRegionRelative.Name, endStation - featureline.Start);  // 1.1.0
            pipe.SetParameterByName(ADSK_Parameters.Instance.EndRegionNormalized.Name, (endStation - featureline.Start) / (featureline.End - featureline.Start));  // 1.1.0

            if (start != null)
            {
                start.Dispose();
            }
            if (end != null)
            {
                end.Dispose();
            }

            Utils.Log(string.Format("Conduit.ByCurveFeatureline completed.", ""));

            return pipe;
        }

        /// <summary>
        /// Creates a Conduit by revit Conduit.
        /// </summary>
        /// <param name="element">The MEP Curve from Revit</param>
        /// <returns></returns>
        public static Conduit ByRevitElement(Revit.Elements.Element element)
        {
            if (element.InternalElement is Autodesk.Revit.DB.Electrical.Conduit)
            {
                var c = element.InternalElement as Autodesk.Revit.DB.Electrical.Conduit;
                var conduit = new Conduit();
                conduit.InternalSetMEPCurve(c);
                return conduit;
            }

            return null;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("{0}", InternalElement.Name);
        }

        #endregion
    }
}
