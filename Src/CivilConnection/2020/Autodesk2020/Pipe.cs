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
    /// Pipe obejct type.
    /// </summary>
    /// <seealso cref="CivilConnection.AbstractMEPCurve" />
    [DynamoServices.RegisterForTrace()]
    public class Pipe : AbstractMEPCurve
    {
        #region PRIVATE PROPERTIES


        #endregion

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

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="Pipe"/> class.
        /// </summary>
        /// <param name="instance">The instance.</param>
        protected Pipe(Autodesk.Revit.DB.Plumbing.Pipe instance)
        {
            SafeInit(() => InitObject(instance));
        }

        internal Pipe() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="Pipe"/> class.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="systemType">Type of the system.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <param name="level">The level.</param>
        internal Pipe(Autodesk.Revit.DB.Plumbing.PipeType pipeType, Autodesk.Revit.DB.Plumbing.PipingSystemType systemType, XYZ start, XYZ end,
            Autodesk.Revit.DB.Level level)
        {
            InitObject(pipeType, systemType, start, end, level);
        }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Initialize a Pipe element
        /// </summary>
        /// <param name="instance">The instance.</param>
        private void InitObject(Autodesk.Revit.DB.Plumbing.Pipe instance)
        {
            Autodesk.Revit.DB.MEPCurve fi = instance as Autodesk.Revit.DB.MEPCurve;
            InternalSetMEPCurve(fi);
        }

        /// <summary>
        /// Initialize a Pipe element
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="pipingSystemType">Type of the piping system.</param>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <param name="level">The level.</param>
        private void InitObject(Autodesk.Revit.DB.Plumbing.PipeType pipeType, Autodesk.Revit.DB.Plumbing.PipingSystemType pipingSystemType, XYZ start, XYZ end,
            Autodesk.Revit.DB.Level level)
        {
            //Phase 1 - Check to see if the object exists and should be rebound
            var oldFam =
                ElementBinder.GetElementFromTrace<Autodesk.Revit.DB.MEPCurve>(DocumentManager.Instance.CurrentDBDocument);

            //There was a point, rebind to that, and adjust its position
            if (oldFam != null)
            {
                InternalSetMEPCurve(oldFam);
                InternalSetMEPCurveType(pipeType);
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
                fi = Autodesk.Revit.DB.Plumbing.Pipe.Create(DocumentManager.Instance.CurrentDBDocument, pipingSystemType.Id, pipeType.Id, level.Id, start, end);
            }

            InternalSetMEPCurve(fi);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(InternalElement);
        }

        /// <summary>
        /// Internals the type of the set piping system.
        /// </summary>
        /// <param name="type">The type.</param>
        private void InternalSetPipingSystemType(Autodesk.Revit.DB.Plumbing.PipingSystemType type)
        {
            if (InternalMEPCurve.MEPSystem.GetTypeId().IntegerValue.Equals(type.Id.IntegerValue))
                return;

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);

            InternalMEPCurve.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).Set(type.Id);

            TransactionManager.Instance.TransactionTaskDone();
        }

        /// <summary>
        /// Gets the pipes by ids.
        /// </summary>
        /// <param name="ids">The ids.</param>
        /// <returns></returns>
        private static Pipe[] GetPipesByIds(ICollection<ElementId> ids)
        {
            Utils.Log(string.Format("Pipe.GetPipesByIds started...", ""));

            var pipes = new FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument)
                  .OfClass(typeof(Autodesk.Revit.DB.Plumbing.Pipe))
                  .WhereElementIsNotElementType()
                  .Cast<Autodesk.Revit.DB.Plumbing.Pipe>()
                  .Where(x => ids.Contains(x.Id))
                  .ToList();

            Pipe[] output = new Pipe[pipes.Count];

            int c = 0;

            foreach (var p in pipes)
            {
                output[c] = new Pipe(p);

                c = c + 1;
            }

            Utils.Log(string.Format("Pipe.GetPipesByIds completed.", ""));

            return output;
        }

        #endregion

        #region PUBLIC METHODS
       
        /// <summary>
        /// Creates a pipe by two points.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="pipingSystemType">Type of the piping system.</param>
        /// <param name="start">The start point.</param>
        /// <param name="end">The end point.</param>
        /// <param name="level">The level.</param>
        /// <returns></returns>
        public static Pipe ByPoints(Revit.Elements.Element pipeType, Revit.Elements.Element pipingSystemType, Autodesk.DesignScript.Geometry.Point start, Autodesk.DesignScript.Geometry.Point end, Revit.Elements.Level level)
        {
            Utils.Log(string.Format("Pipe.ByPoints started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }

            var oType = pipeType.InternalElement as Autodesk.Revit.DB.Plumbing.PipeType;
            var oSystemType = pipingSystemType.InternalElement as Autodesk.Revit.DB.Plumbing.PipingSystemType;
            var totalTransform = RevitUtils.DocumentTotalTransform();
            var nstart = start.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = nstart.ToXyz();
            var nend = end.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = nend.ToXyz();
            var l = level.InternalElement as Autodesk.Revit.DB.Level;

            if (nstart != null)
            {
                nstart.Dispose();
            }
            if (nend != null)
            {
                nend.Dispose();
            }

            Utils.Log(string.Format("Pipe.ByPoints completed.", ""));

            return new Pipe(oType, oSystemType, s, e, l);
        }

        /// <summary>
        /// Creates a Conduit by revit Conduit.
        /// </summary>
        /// <param name="element">The MEP Curve from Revit</param>
        /// <returns></returns>
        public static Pipe ByRevitElement(Revit.Elements.Element element)
        {
            if (element.InternalElement is Autodesk.Revit.DB.Plumbing.Pipe)
            {
                var c = element.InternalElement as Autodesk.Revit.DB.Plumbing.Pipe;
                var conduit = new Pipe();
                conduit.InternalSetMEPCurve(c);
                return conduit;
            }

            return null;
        }

        /// <summary>
        /// Creates a pipe by curve.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="pipingSystemType">Type of the piping system.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="level">The level.</param>
        /// <returns></returns>
        public static Pipe ByCurve(Revit.Elements.Element pipeType, Revit.Elements.Element pipingSystemType, Autodesk.DesignScript.Geometry.Curve curve, Revit.Elements.Level level)
        {
            Utils.Log(string.Format("Pipe.ByCurve started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }
            var oType = pipeType.InternalElement as Autodesk.Revit.DB.Plumbing.PipeType;
            var oSystemType = pipingSystemType.InternalElement as Autodesk.Revit.DB.Plumbing.PipingSystemType;
            var totalTransform = RevitUtils.DocumentTotalTransform();
            var start = curve.StartPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = start.ToXyz();
            var end = curve.EndPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = end.ToXyz();
            var l = level.InternalElement as Autodesk.Revit.DB.Level;

            if (start != null)
            {
                start.Dispose();
            }
            if (end != null)
            {
                end.Dispose();
            }

            Utils.Log(string.Format("Pipe.ByCurve completed.", ""));

            return new Pipe(oType, oSystemType, s, e, l);
        }

        /// <summary>
        /// Creates a pipe by curve.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="pipingSystemType">Type of the piping system.</param>
        /// <param name="level">The level.</param>
        /// <param name="curve">The curve.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns></returns>
        public static Pipe ByCurveFeatureline(Revit.Elements.Element pipeType, Revit.Elements.Element pipingSystemType, Revit.Elements.Level level, Autodesk.DesignScript.Geometry.Curve curve, Featureline featureline)
        {
            Utils.Log(string.Format("Pipe.ByCurveFeatureline started...", ""));

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }

            Autodesk.DesignScript.Geometry.Point start = null;
            Autodesk.DesignScript.Geometry.Point end = null;

            var oType = pipeType.InternalElement as Autodesk.Revit.DB.Plumbing.PipeType;
            var oSystemType = pipingSystemType.InternalElement as Autodesk.Revit.DB.Plumbing.PipingSystemType;
            var totalTransform = RevitUtils.DocumentTotalTransform();
            start = curve.StartPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var s = start.ToXyz();
            end = curve.EndPoint.Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
            var e = end.ToXyz();
            var l = level.InternalElement as Autodesk.Revit.DB.Level;

            var pipe = new Pipe(oType, oSystemType, s, e, l);

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

            if (start != null)
            {
                start.Dispose();
            }
            if (end != null)
            {
                end.Dispose();
            }

            Utils.Log(string.Format("Pipe.ByCurveFeatureline completed.", ""));

            return pipe;
        }

        /// <summary>
        /// Creates a pipe by PolyCurve.
        /// </summary>
        /// <param name="pipeType">Type of the pipe.</param>
        /// <param name="pipingSystemType">Type of the piping system.</param>
        /// <param name="polyCurve">The poly curve.</param>
        /// <param name="level">The level.</param>
        /// <param name="maxLength">The maximum length.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Pipes", "Fittings" })]
        public static Dictionary<string, object> ByPolyCurve(Revit.Elements.Element pipeType, Revit.Elements.Element pipingSystemType, PolyCurve polyCurve, Revit.Elements.Level level, double maxLength, Featureline featureline)
        {
            Utils.Log(string.Format("Pipe.ByPolyCurve started...", ""));

            var totalTransform = RevitUtils.DocumentTotalTransform();
            var totalTransformInverse = totalTransform.Inverse();

            if (!SessionVariables.ParametersCreated)
            {
                UtilsObjectsLocation.CheckParameters(DocumentManager.Instance.CurrentDBDocument); 
            }
            var oType = pipeType.InternalElement as Autodesk.Revit.DB.Plumbing.PipeType;
            var oSystemType = pipingSystemType.InternalElement as Autodesk.Revit.DB.Plumbing.PipingSystemType;
            var l = level.InternalElement as Autodesk.Revit.DB.Level;

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

            IList<ElementId> ids = new List<ElementId>();

            TransactionManager.Instance.EnsureInTransaction(DocumentManager.Instance.CurrentDBDocument);


            Autodesk.DesignScript.Geometry.Point start = null;
            Autodesk.DesignScript.Geometry.Point end = null;

            Autodesk.DesignScript.Geometry.Point sp = null;
            Autodesk.DesignScript.Geometry.Point ep = null;
            Autodesk.DesignScript.Geometry.Curve curve = null;

            for (int i = 0; i < points.Count - 1; ++i)
            {
                start = points[i].Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                var s = start.ToXyz();
                end = points[i + 1].Transform(totalTransform) as Autodesk.DesignScript.Geometry.Point;
                var e = end.ToXyz();

                Autodesk.Revit.DB.Plumbing.Pipe p = Autodesk.Revit.DB.Plumbing.Pipe.CreatePlaceholder(DocumentManager.Instance.CurrentDBDocument, oSystemType.Id, oType.Id, l.Id, s, e);
                ids.Add(p.Id);
            }

            var pipeIds = Autodesk.Revit.DB.Plumbing.PlumbingUtils.ConvertPipePlaceholders(DocumentManager.Instance.CurrentDBDocument, ids);

            TransactionManager.Instance.TransactionTaskDone();

            DocumentManager.Instance.CurrentDBDocument.Regenerate();

            Pipe[] pipes = GetPipesByIds(pipeIds);

            foreach (Pipe pipe in pipes)
            {
                curve = pipe.Location.Transform(totalTransformInverse) as Autodesk.DesignScript.Geometry.Curve;
                sp = curve.StartPoint;
                ep = curve.EndPoint;

                pipe.SetParameterByName(ADSK_Parameters.Instance.Corridor.Name, featureline.Baseline.CorridorName);
                pipe.SetParameterByName(ADSK_Parameters.Instance.BaselineIndex.Name, featureline.Baseline.Index);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Code.Name, featureline.Code);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Side.Name, featureline.Side.ToString());
                pipe.SetParameterByName(ADSK_Parameters.Instance.X.Name, Math.Round(sp.X, 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Y.Name, Math.Round(sp.Y, 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Z.Name, Math.Round(sp.Z, 3));
                var soe = featureline.GetStationOffsetElevationByPoint(sp);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Station.Name, Math.Round((double)soe["Station"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Offset.Name, Math.Round((double)soe["Offset"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Elevation.Name, Math.Round((double)soe["Elevation"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.Update.Name, 1);
                pipe.SetParameterByName(ADSK_Parameters.Instance.Delete.Name, 0);
                soe = featureline.GetStationOffsetElevationByPoint(ep);
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndStation.Name, Math.Round((double)soe["Station"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndOffset.Name, Math.Round((double)soe["Offset"], 3));
                pipe.SetParameterByName(ADSK_Parameters.Instance.EndElevation.Name, Math.Round((double)soe["Elevation"], 3));
            }

            IList<Fitting> fittings = new List<Fitting>();

            for (int i = 0; i < pipes.Length - 1; ++i)
            {
                Fitting fitting = null;
                try
                {
                    fitting = Fitting.Elbow(pipes[i], pipes[i + 1]);
                }
                catch { }

                fittings.Add(fitting);
            }

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

            foreach (var item in points)
            {
                if (item != null)
                {
                    item.Dispose();
                }
            }

            points.Clear();

            Utils.Log(string.Format("Pipe.ByPolyCurve completed.", ""));

            return new Dictionary<string, object>() { { "Pipes", pipes }, { "Fittings", fittings } };
        }

        #endregion
    }
}
