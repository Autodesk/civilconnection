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
    /// CableTray obejct type.
    /// </summary>
    /// <seealso cref="CivilConnection.AbstractMEPCurve" />
    [DynamoServices.RegisterForTrace()]
    public class CableTray : AbstractMEPCurve
    {

        #region PUBLIC PROPERTIES

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
        /// Initializes a new instance of the <see cref="CableTray"/> class.
        /// </summary>
        /// <param name="instance">The instance.</param>
        protected CableTray(Autodesk.Revit.DB.Electrical.CableTray instance)
        {
            SafeInit(() => InitObject(instance));
        }

        internal CableTray() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CableTray"/> class.
        /// </summary>
        /// <param name="cableTrayType">Type of the cable tray.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        internal CableTray(Autodesk.Revit.DB.Electrical.CableTrayType cableTrayType, XYZ start, XYZ end)
        {
            InitObject(cableTrayType, start, end);
        }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Initialize a CableTray element.
        /// </summary>
        /// <param name="instance"></param>
        private void InitObject(Autodesk.Revit.DB.Electrical.CableTray instance)
        {
            Autodesk.Revit.DB.MEPCurve fi = instance as Autodesk.Revit.DB.MEPCurve;
            InternalSetMEPCurve(fi);
        }

        /// <summary>
        /// Initialize a CableTray element.
        /// </summary>
        private void InitObject(Autodesk.Revit.DB.Electrical.CableTrayType cableTrayType, XYZ start, XYZ end)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.MEPCurve>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetMEPCurve(oldFam);
                InternalSetMEPCurveType(cableTrayType);
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
                fi = Autodesk.Revit.DB.Electrical.CableTray.Create(DocumentManager.Instance.CurrentDBDocument, cableTrayType.Id, start, end, ElementId.InvalidElementId) as Autodesk.Revit.DB.MEPCurve;
            }

            InternalSetMEPCurve(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Private method to get the CableTrays by ElementId
        /// </summary>
        /// <param name="ids"></param>
        /// <returns></returns>
        private static CableTray[] GetCableTrayByIds(ICollection<ElementId> ids)
        {
            Utils.Log(string.Format("CableTray.GetCableTrayByIds started...", ""));

            var pipes = new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                  .OfClass(typeof(Autodesk.Revit.DB.Electrical.CableTray))
                  .WhereElementIsNotElementType()
                  .Cast<Autodesk.Revit.DB.Electrical.CableTray>()
                  .Where(x => ids.Contains(x.Id))
                  .ToList();

            CableTray[] output = new CableTray[pipes.Count];

            int c = 0;

            foreach (var p in pipes)
            {
                output[c] = new CableTray(p);

                c = c + 1;
            }

            Utils.Log(string.Format("CableTray.GetCableTrayByIds completed...", ""));

            return output;
        }

        /// <summary>
        /// CableTray by curve.
        /// </summary>
        /// <param name="cableTrayType">Type of the cable tray.</param>
        /// <param name="curve">The curve.</param>
        /// <returns>It Uses the start and end Points of the curve to create the CableTray</returns>
        private static CableTray CableTrayByCurve(Revit.Elements.Element cableTrayType, Autodesk.DesignScript.Geometry.Curve curve)
        {
            Utils.Log(string.Format("CableTray.CableTrayByCurve started...", ""));

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
            var oType = cableTrayType.InternalElement as Autodesk.Revit.DB.Electrical.CableTrayType;
            var totalTransform = RevitUtils.DocumentTotalTransform();
            Autodesk.DesignScript.Geometry.Point start = curve.StartPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = start.ToXyz();
            Autodesk.DesignScript.Geometry.Point end = curve.EndPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = end.ToXyz();
            TransactionManager.Instance.TransactionTaskDone();

            start.Dispose();
            end.Dispose();

            Utils.Log(string.Format("CableTray.CableTrayByCurve completed.", ""));

            return new CableTray(oType, s, e);
        }

        /// <summary>
        /// Creates a set of CableTrays following a PolyCurve specifying a maximum length.
        /// </summary>
        /// <param name="CableTrayType">The CableTray type.</param>
        /// <param name="polyCurve">The PolyCurve to follow in WCS.</param>
        /// <param name="maxLength">The maximum length of the CableTrays following the PolyCurve.</param>
        /// <returns></returns>
        private CableTray[] ByPolyCurve_(Revit.Elements.Element CableTrayType, PolyCurve polyCurve, double maxLength)
        {
            Utils.Log(string.Format("CableTray.ByPolyCurve started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument);
            }

            var oType = CableTrayType.InternalElement as Autodesk.Revit.DB.Electrical.CableTrayType;

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

            points = Autodesk.DesignScript.Geometry.Point.PruneDuplicates(points);  // TODO this is slow

            IList<ElementId> ids = new List<ElementId>();

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            var totalTransform = RevitUtils.DocumentTotalTransform();

            Autodesk.DesignScript.Geometry.Point start = null;
            Autodesk.DesignScript.Geometry.Point end = null;

            for (int i = 0; i < points.Count - 1; ++i)
            {
                start = points[i].Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                var s = start.ToXyz();
                end = points[i + 1].Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                var e = end.ToXyz();

                Autodesk.Revit.DB.Electrical.CableTray p = Autodesk.Revit.DB.Electrical.CableTray.Create(DocumentManager.Instance.CurrentDBDocument, oType.Id, s, e, ElementId.InvalidElementId);
                ids.Add(p.Id);
            }

            for (int i = 0; i < GetCableTrayByIds(ids).Length - 1; ++i)
            {
                CableTray ct1 = GetCableTrayByIds(ids)[i];
                CableTray ct2 = GetCableTrayByIds(ids)[i + 1];
                Fitting.Elbow(ct1, ct2);
            }

            TransactionManager.Instance.TransactionTaskDone();
            start.Dispose();
            end.Dispose();

            foreach (var pt in points)
            {
                if (pt != null)
                {
                    pt.Dispose();
                }
            }

            points.Clear();


            Utils.Log(string.Format("CableTray.ByPolyCurve completed.", ""));

            return GetCableTrayByIds(ids);
        }

        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Creates a CableTray by two Points.
        /// </summary>
        /// <param name="cableTrayType">Type of the cable tray.</param>
        /// <param name="start">The start Point in WCS.</param>
        /// <param name="end">The end Point in WCS.</param>
        /// <returns></returns>
        public static CableTray ByPoints(Revit.Elements.Element cableTrayType, Autodesk.DesignScript.Geometry.Point start, Autodesk.DesignScript.Geometry.Point end)
        {
            Utils.Log(string.Format("CableTray.ByPoints started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }
            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);
            var oType = cableTrayType.InternalElement as Autodesk.Revit.DB.Electrical.CableTrayType;
            var totalTransform = RevitUtils.DocumentTotalTransform();
            var nstart = start.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = nstart.ToXyz();
            var nend = end.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = nend.ToXyz();
            TransactionManager.Instance.TransactionTaskDone();

            if (nstart != null)
            {
                nstart.Dispose();
            }
            if (nend != null)
            {
                nend.Dispose();
            }

            Utils.Log(string.Format("CableTray.ByPoints completed.", ""));

            return new CableTray(oType, s, e);
        }


        /// <summary>
        /// Creates a Conduit by revit Conduit.
        /// </summary>
        /// <param name="element">The MEP Curve from Revit</param>
        /// <returns></returns>
        public static CableTray ByRevitElement(Revit.Elements.Element element)
        {
            Utils.Log(string.Format("CableTray.ByRevitElement started...", ""));

            if (element.InternalElement is Autodesk.Revit.DB.Electrical.CableTray)
            {
                var c = element.InternalElement as Autodesk.Revit.DB.Electrical.CableTray;
                var conduit = new CableTray();
                conduit.InternalSetMEPCurve(c);
                return conduit;
            }

            Utils.Log(string.Format("CableTray.ByRevitElement completed.", ""));

            return null;
        }

        /// <summary>
        /// Creates a CableTray using the start and end points of a curve.
        /// </summary>
        /// <param name="cableTrayType">The CableTray Type.</param>
        /// <param name="curve">The Curve</param>
        /// <returns></returns>
        public static CableTray ByCurve(Revit.Elements.Element cableTrayType, Autodesk.DesignScript.Geometry.Curve curve)
        {
            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument);
            }
            return CableTrayByCurve(cableTrayType, curve);
        }

        /// <summary>
        /// CableTray by curve.
        /// </summary>
        /// <param name="cableTrayType">Type of the cable tray.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns>Associates the CableTray to a Featureline.</returns>
        public static CableTray ByCurveFeatureline(Revit.Elements.Element cableTrayType, Autodesk.DesignScript.Geometry.Curve curve, Featureline featureline)
        {
            Utils.Log(string.Format("CableTray.ByCurveFeatureline started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument);
            }

            var pipe = CableTrayByCurve(cableTrayType, curve);

            var startSOE = featureline.GetStationOffsetElevationByPoint(curve.StartPoint);
            var endSOE = featureline.GetStationOffsetElevationByPoint(curve.EndPoint);

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

            Utils.Log(string.Format("CableTray.ByCurveFeatureline completed.", ""));

            return pipe;
        }

       
        /// <summary>
        /// Creates a list of CableTrays from a PolyCurve.
        /// </summary>
        /// <param name="cableTrayType">Type of the cable tray.</param>
        /// <param name="polyCurve">The poly curve.</param>
        /// <param name="maxLength">The maximum length.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "CableTray", "Fittings" })]
        public static Dictionary<string, object> ByPolyCurve(Revit.Elements.Element cableTrayType, Autodesk.DesignScript.Geometry.PolyCurve polyCurve, double maxLength, Featureline featureline)
        {
            Utils.Log(string.Format("CableTray.ByPolyCurve started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument);
            }

            double length = polyCurve.Length;

            var oType = cableTrayType.InternalElement as Autodesk.Revit.DB.Electrical.CableTrayType;
            IList<CableTray> output = new List<CableTray>();
            IList<Fitting> fittings = new List<Fitting>();

            int subdivisions = Convert.ToInt32(Math.Ceiling(length / maxLength));

            Utils.Log(string.Format("subdivisions {0}", subdivisions));

            IList<Autodesk.DesignScript.Geometry.Point> points = new List<Autodesk.DesignScript.Geometry.Point>();

            try
            {
                points.Add(polyCurve.StartPoint);

                foreach (Autodesk.DesignScript.Geometry.Point p in polyCurve.PointsAtEqualChordLength(subdivisions))
                {
                    points.Add(p);
                }

                points.Add(polyCurve.EndPoint);

                points = Autodesk.DesignScript.Geometry.Point.PruneDuplicates(points);  // This is slow
            }
            catch
            {
                points = Featureline.PointsByChord(polyCurve, maxLength);  // This is slow
            }

            Utils.Log(string.Format("Points {0}", points.Count));

            var totalTransform = RevitUtils.DocumentTotalTransform();

            Autodesk.DesignScript.Geometry.Point s = null;
            Autodesk.DesignScript.Geometry.Point e = null;
            Autodesk.DesignScript.Geometry.Point sp = null;
            Autodesk.DesignScript.Geometry.Point ep = null;
            Autodesk.DesignScript.Geometry.Curve curve = null;

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            for (int i = 0; i < points.Count - 1; ++i)
            {
                s = points[i];
                e = points[i + 1];
                curve = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(s, e);

                sp = s.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                ep = e.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;

                var pipe = new CableTray();

                Autodesk.Revit.DB.MEPCurve fi;

                if (DocumentManager.Instance.CurrentDBDocument.IsFamilyDocument)
                {
                    fi = null;
                }
                else
                {
                    fi = Autodesk.Revit.DB.Electrical.CableTray.Create(DocumentManager.Instance.CurrentDBDocument, oType.Id, sp.ToXyz(), ep.ToXyz(), ElementId.InvalidElementId) as Autodesk.Revit.DB.MEPCurve;
                }

                pipe.InitObject((Autodesk.Revit.DB.Electrical.CableTray)fi);

                pipe.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);
                pipe.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
                pipe.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(s.X, 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(s.Y, 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(s.Z, 3));
                var soe = featureline.GetStationOffsetElevationByPoint(s);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round((double)soe["Station"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round((double)soe["Offset"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round((double)soe["Elevation"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);
                soe = featureline.GetStationOffsetElevationByPoint(e);
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, Math.Round((double)soe["Station"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndOffset.Name, Math.Round((double)soe["Offset"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndElevation.Name, Math.Round((double)soe["Elevation"], 3));
                output.Add(pipe);

                Utils.Log(string.Format("Pipe {0}", pipe.Id));
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

            TransactionManager.Instance.TransactionTaskDone();
         
            Utils.Log(string.Format("CableTray.ByPolyCurve completed.", ""));

            return new Dictionary<string, object>() { { "CableTray", output }, { "Fittings", fittings } };

        }

        /// <summary>
        /// Public textual representation of the Dynamo node preview
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("CableTray(Name={0})", InternalElement.Name);
        }

        #endregion
    }
}
