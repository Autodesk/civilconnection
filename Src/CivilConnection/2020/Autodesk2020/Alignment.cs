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

using System.Runtime;
using System.Runtime.InteropServices;

using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AECC.Interop.UiRoadway;
using Autodesk.AECC.Interop.Roadway;
using Autodesk.AECC.Interop.Land;
using Autodesk.AECC.Interop.UiLand;
using System.Reflection;

using Autodesk.DesignScript.Runtime;
using Autodesk.DesignScript.Geometry;

namespace CivilConnection
{
    /// <summary>
    /// Alignment object type.
    /// </summary>
    public class Alignment
    {
        #region PRIVATE PROPERTIES
        private AeccAlignment _alignment;
        private AeccAlignmentEntities _entities;
        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return _alignment.DisplayName; } }
        /// <summary>
        /// Gets the length.
        /// </summary>
        /// <value>
        /// The length.
        /// </value>
        public double Length { get { return _alignment.Length; } }
        /// <summary>
        /// Gets the start.
        /// </summary>
        /// <value>
        /// The start.
        /// </value>
        public double Start { get { return _alignment.StartingStation; } }
        /// <summary>
        /// Gets the end.
        /// </summary>
        /// <value>
        /// The end.
        /// </value>
        public double End { get { return _alignment.EndingStation; } }

        /// <summary>
        /// Gets the stations of the geometry points.
        /// </summary>
        /// <value>
        /// The GeometryStations.
        /// </value>
        public double[] GeometryStations { get { return _alignment.GetStations(AeccStationType.aeccGeometryPoint, this.Start, this.End).Cast<AeccAlignmentStation>().Select(x => x.Station).ToArray(); } }

        /// <summary>
        /// Gets the stations of the points of intersection.
        /// </summary>
        /// <value>
        /// The PIStations.
        /// </value>
        public double[] PIStations { get { return _alignment.GetStations(AeccStationType.aeccPIPoint, this.Start, this.End).Cast<AeccAlignmentStation>().Select(x => x.Station).ToArray(); } }

        /// <summary>
        /// Gets the stations of the points of superelevation transition.
        /// </summary>
        /// <value>
        /// The SuperTransStations.
        /// </value>
        public double[] SuperTransStations { get { return _alignment.GetStations(AeccStationType.aeccSuperTransPoint, this.Start, this.End).Cast<AeccAlignmentStation>().Select(x => x.Station).ToArray(); } }

        #endregion

        #region INTERNAL CONSTRUCTORS
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this._alignment; } }

        /// <summary>
        /// Internal constructor.
        /// </summary>
        /// <param name="alignment">The internal AeccAlignment.</param>
        internal Alignment(AeccAlignment alignment)
        {
            this._alignment = alignment;
            this._entities = alignment.Entities;
        }
        #endregion

        #region PUBLIC CONSTRUCTORS
        /// <summary>
        /// Creates an Alignment in the Civil Document starting from a Dynamo polygonal PolyCurve.
        /// </summary>
        /// <param name="civilDocument">The CivilDocument.</param>
        /// <param name="name">The name of the alignment.</param>
        /// <param name="polyCurve">The source PolyCurve.</param>
        /// <param name="layer">The layer.</param>
        /// <returns>
        /// The new Alignment
        /// </returns>
        public static Alignment ByPolygonal(CivilDocument civilDocument, string name, PolyCurve polyCurve, string layer)
        {
            var pl = civilDocument._document.HandleToObject(Utils.AddLWPolylineByPolyCurve(civilDocument._document, polyCurve, "0"));

            Utils.Log(string.Format("Polyline 2D added: {0}", pl.Handle));

            var alignments = civilDocument._document.AlignmentsSiteless;

            AeccAlignmentStyle alignmentStyle = null;
            try
            {
                alignmentStyle = civilDocument._document.AlignmentStyles[0];
            }
            catch
            {
                alignmentStyle = civilDocument._document.AlignmentStyles.Add("CivilConnection_AlignmentStyle");
            }

            Utils.Log(string.Format("Alignment Style: {0}", alignmentStyle));

            AeccAlignmentLabelStyleSet alignmentLabelStyleSet = null;

            try
            {
                alignmentLabelStyleSet = civilDocument._document.AlignmentLabelStyleSets[0];
            }
            catch (Exception)
            {

                alignmentLabelStyleSet = civilDocument._document.AlignmentLabelStyleSets.Add("CivilConnection_AlignmentLabelStyle");
            }

            Utils.Log(string.Format("Alignment Label Style Set: {0}", alignmentLabelStyleSet));

            AeccAlignment al = alignments.AddFromPolylineEx(name, layer, pl, alignmentStyle, alignmentLabelStyleSet, true, false);

            Utils.Log(string.Format("Alignment Created: {0}, {1}", name, al.Handle));

            return new Alignment(al);
        }
        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Returns the list of vertical Profiles associated to the Alignment.
        /// </summary>
        /// <returns>The list of associated Profiles.</returns>
        public IList<Profile> GetProfiles()
        {
            Utils.Log("Alignment.GetProfiles Started...");

            IList<Profile> output = new List<Profile>();

            foreach (AeccProfile profile in this._alignment.Profiles)
            {
                output.Add(new Profile(profile));
            }

            Utils.Log("Alignment.GetProfiles Completed.");

            return output;
        }

        /// <summary>
        /// Returns the list of vertical ProfileViews associated to the Alignment.
        /// </summary>
        /// <returns>The list of assocaited ProfileViews.</returns>
        public IList<ProfileView> GetProfileViews()
        {
            Utils.Log("Alignment.GetProfileViews Started...");

            IList<ProfileView> output = new List<ProfileView>();

            foreach (AeccProfileView profile in this._alignment.ProfileViews)
            {
                output.Add(new ProfileView(profile));
            }

            Utils.Log("Alignment.GetProfileViews Completed.");

            return output;
        }


        /// <summary>
        /// Factorial function. Returns a double to allow for values bigger than 20!
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        private double Factorial(int f)
        {
            if (f < 0)
            {
                throw new Exception("Factorial is undefined for negative numbers.");
            }

            double output = 1;

            if (f == 0)
            {
                return output;
            }
            else
            {
                for (int i = 1; i <= f; i++)
                {
                    output *= i;
                }
            }

            return output;
        }

        /// <summary>
        /// Returns the list of Dynamo curves that defines the Alignment.
        /// </summary>
        /// <param name="tessellation">The length of the tessellation for spirals, by default is 1 unit.</param>
        /// <returns>A list of curves that represent the Alignment.</returns>
        /// <remarks>The tool returns only lines and arcs.</remarks>
        public IList<Curve> GetCurves(double tessellation = 1)
        {
            Utils.Log(string.Format("Alignment.GetCurves {0} Started...", this.Name));

            IList<Curve> output = new List<Curve>();

            if (this._entities == null)
            {
                Utils.Log(string.Format("ERROR: Alignment Entities are null", ""));

                var stations = this.GeometryStations.ToList();
                stations.AddRange(this.PIStations.ToList());
                stations.AddRange(this.SuperTransStations.ToList());

                stations.Sort();

                var pts = new List<Point>();

                foreach (var s in stations)
                {
                    pts.Add(this.PointByStationOffsetElevation(s));
                }

                pts = Point.PruneDuplicates(pts).ToList();

                output.Add(PolyCurve.ByPoints(pts));

                return output;
            }

            Utils.Log(string.Format("Total Entities: {0}", this._entities.Count));

            try
            {
                var entities = new List<AeccAlignmentEntity>();

                for (int c = 0; c < this._entities.Count; ++c)
                {
                    try
                    {
                        var ce = this._entities.Item(c);

                        Utils.Log(string.Format("Entity: {0}", ce.Type));

                        if (ce.Type != AeccAlignmentEntityType.aeccArc && ce.Type != AeccAlignmentEntityType.aeccTangent)
                        {
                            int count = ce.SubEntityCount;

                            if (count > 0)
                            {
                                for (int i = 0; i < ce.SubEntityCount; ++i)
                                {
                                    try
                                    {
                                        var se = ce.SubEntity(i);

                                        Utils.Log(string.Format("SubEntity: {0}", se.Type));

                                        entities.Add(se);
                                    }
                                    catch (Exception ex)
                                    {
                                        Utils.Log(string.Format("ERROR1: {0} {1}", ex.Message, ex.StackTrace));
                                    }
                                }
                            }
                            else
                            {
                                entities.Add(ce);
                            }
                        }
                        else
                        {
                            entities.Add(ce);
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR2: {0} {1}", ex.Message, ex.StackTrace));
                    }
                }

                Utils.Log(string.Format("Missing Entities: {0}", this._entities.Count - entities.Count));

                foreach (AeccAlignmentEntity e in entities)
                {
                    try
                    {
                        switch (e.Type)
                        {
                            case AeccAlignmentEntityType.aeccTangent:
                                {
                                    //Utils.Log(string.Format("Tangent..", ""));

                                    AeccAlignmentTangent a = e as AeccAlignmentTangent;

                                    var start = Point.ByCoordinates(a.StartEasting, a.StartNorthing);
                                    var end = Point.ByCoordinates(a.EndEasting, a.EndNorthing);

                                    output.Add(Line.ByStartPointEndPoint(start, end));

                                    start.Dispose();
                                    end.Dispose();

                                    //Utils.Log(string.Format("OK", ""));

                                    break;
                                }
                            case AeccAlignmentEntityType.aeccArc:
                                {
                                    //Utils.Log(string.Format("Arc..", ""));

                                    AeccAlignmentArc a = e as AeccAlignmentArc;

                                    Point center = Point.ByCoordinates(a.CenterEasting, a.CenterNorthing);
                                    Point start = Point.ByCoordinates(a.StartEasting, a.StartNorthing);
                                    Point end = Point.ByCoordinates(a.EndEasting, a.EndNorthing);

                                    Arc arc = null;
                                    if (!a.Clockwise)
                                    {
                                        arc = Arc.ByCenterPointStartPointEndPoint(center, start, end);
                                    }
                                    else
                                    {
                                        arc = Arc.ByCenterPointStartPointEndPoint(center, end, start);
                                    }

                                    output.Add(arc);

                                    center.Dispose();
                                    start.Dispose();
                                    end.Dispose();

                                    //Utils.Log(string.Format("OK", ""));

                                    break;
                                }
                            default:
                                {
                                    //Utils.Log(string.Format("Curve...", ""));
                                    try
                                    {
                                        AeccAlignmentCurve a = e as AeccAlignmentCurve;

                                        var pts = new List<Point>();

                                        double start = this.Start;

                                        try
                                        {
                                            start = a.StartingStation;
                                        }
                                        catch (Exception ex)
                                        {
                                            Utils.Log(string.Format("ERROR11: {0} {1}", ex.Message, ex.StackTrace));

                                            break;
                                        }

                                        //Utils.Log(string.Format("start: {0}", start));

                                        double length = a.Length;

                                        //Utils.Log(string.Format("length: {0}", length));

                                        int subs = Convert.ToInt32(Math.Ceiling(length / tessellation));

                                        if (subs < 10)
                                        {
                                            subs = 10;
                                        }

                                        double delta = length / subs;

                                        for (int i = 0; i < subs + 1; ++i)
                                        {
                                            try
                                            {
                                                double x = 0;
                                                double y = 0;

                                                this._alignment.PointLocation(start + i * delta, 0, out x, out y);

                                                pts.Add(Point.ByCoordinates(x, y));
                                            }
                                            catch (Exception ex)
                                            {
                                                Utils.Log(string.Format("ERROR21: {2} {0} {1}", ex.Message, ex.StackTrace, start + i * delta));
                                            }
                                        }

                                        //Utils.Log(string.Format("Points: {0}", pts.Count));

                                        if (pts.Count < 2)
                                        {
                                            Utils.Log(string.Format("ERROR211: not enough points to create a spiral", ""));
                                            break;
                                        }

                                        NurbsCurve spiral = NurbsCurve.ByPoints(pts);  // Degree by default is 3

                                        output.Add(spiral);

                                        foreach (var pt in pts)
                                        {
                                            if (pts != null)
                                            {
                                                pt.Dispose();
                                            }
                                        }

                                        pts.Clear();

                                        //Utils.Log(string.Format("OK", ""));

                                        //prevStation += length;
                                    }
                                    catch (Exception ex)
                                    {
                                        Utils.Log(string.Format("ERROR22: {0} {1}", ex.Message, ex.StackTrace));
                                    }

                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.Log(string.Format("ERROR3: {0} {1}", ex.Message, ex.StackTrace));
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR4: {0} {1}", ex.Message, ex.StackTrace));
            }

            output = SortCurves(output);

            Utils.Log("Alignment.GetCurves Completed.");

            return output;
        }


        /// <summary>
        /// Sorts a list of Dynamo Curves.
        /// </summary>
        /// <param name="curves">THe list of Curves.</param>
        /// <returns></returns>
        private IList<Curve> SortCurves(IList<Curve> curves)
        {
            int n = curves.Count;

            for (int i = 0; i < n; i++)
            {
                Curve c = curves[i];

                Point endPoint = c.EndPoint;

                bool found = i + 1 >= n;

                Point p = null;

                for (int j = i + 1; j < n; j++)
                {
                    p = curves[j].StartPoint;

                    if (p.DistanceTo(endPoint) < 0.00001)
                    {
                        if (i + 1 != j)
                        {
                            Curve temp = curves[i + 1];

                            curves[i + 1] = curves[j];

                            curves[j] = temp;
                        }

                        found = true;

                        break;
                    }

                    p = curves[j].EndPoint;

                    if (p.DistanceTo(endPoint) < 0.00001)
                    {
                        if (i + 1 == j)
                        {
                            curves[i + 1] = curves[j].Reverse();
                        }
                        else
                        {
                            Curve temp = curves[i + 1];

                            curves[i + 1] = curves[j].Reverse();

                            curves[j] = temp;
                        }

                        found = true;

                        break;
                    }
                }
            }

            return curves;
        }

        /// <summary>
        /// Returns the Sample Lines parameters associated with the alignment.
        /// </summary>
        /// <returns></returns>
        [MultiReturn(new string[] { "station", "lengthLeft", "lengthRight", "elevationMin", "elevationMax" })]
        public IList<Dictionary<string, object>> SampleLinesParameters()
        {
            return RevitUtils.AlignmentSampleLinesParameters(this._alignment);
        }

        /// <summary>
        /// Returns the station, offset and elevation of a point from the alignment.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Station", "Offset", "Elevation" })]
        public Dictionary<string, object> GetStationOffsetElevation(Point point)
        {
            double station = 0;
            double offset = 0;

            ((AeccAlignment)this.InternalElement).StationOffset(point.X, point.Y, out station, out offset);

            return new Dictionary<string, object>() { { "Station", station }, { "Offset", offset }, { "Elevation", point.Z } };
        }

        /// <summary>
        /// Returns a CoordinateSystem along the Alignment at the specified station.
        /// </summary>
        /// <param name="station">The station value.</param>
        /// <param name="offset">The offset value.</param>
        /// <param name="elevation">The elevation value.</param>
        /// <returns></returns>
        public CoordinateSystem CoordinateSystemByStation(double station, double offset = 0, double elevation = 0)
        {
            Utils.Log("Alignment.CoordinateSystemByStation Started...");

            double northing = 0;
            double easting = 0;
            double northingX = 0;
            double eastingX = 0;

            this._alignment.PointLocation(station, offset, out easting, out northing);

            Point point = Point.ByCoordinates(easting, northing, elevation);

            this._alignment.PointLocation(station, offset + 1, out eastingX, out northingX);

            Point pointX = Point.ByCoordinates(eastingX, northingX, elevation);

            Vector x = Vector.ByTwoPoints(point, pointX).Normalized();

            Vector y = Vector.ZAxis().Cross(x).Normalized();

            CoordinateSystem cs = CoordinateSystem.ByOriginVectors(point, x, y);

            point.Dispose();
            pointX.Dispose();
            x.Dispose();
            y.Dispose();

            Utils.Log("Alignment.CoordinateSystemByStation Completed.");

            return cs;
        }

        /// <summary>
        /// Returns a Point along the Alignment at the specified station.
        /// </summary>
        /// <param name="station">The station value.</param>
        /// <param name="offset">The offset value.</param>
        /// <param name="elevation">The elevation value.</param>
        /// <returns></returns>
        public Point PointByStationOffsetElevation(double station, double offset = 0, double elevation = 0)
        {
            Utils.Log("Alignment.PointByStationOffsetElevation Started...");

            double northing = 0;
            double easting = 0;

            this._alignment.PointLocation(station, offset, out easting, out northing);

            Point point = Point.ByCoordinates(easting, northing, elevation);

            Utils.Log("Alignment.PointByStationOffsetElevation Completed.");

            return point;
        }

        /// <summary>
        /// Public textual representation in the Dynamo node preview.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("Alignment(Name = {0}, Length = {1}, Start = {2}, End = {3})",
                this.Name,
                Math.Round(this.Length, 2).ToString(),
                Math.Round(this.Start, 2).ToString(),
                Math.Round(this.End, 2).ToString());
        }
        #endregion
    }
}
