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
using Autodesk.AECC.Interop.Roadway;
using Autodesk.AECC.Interop.UiRoadway;
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CivilConnection
{
    /// <summary>
    /// Corridor obejct type.
    /// </summary>
    public class Corridor
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The corridor
        /// </summary>
        internal AeccCorridor _corridor;
        /// <summary>
        /// The baselines
        /// </summary>
        internal IList<Baseline> _baselines;
        /// <summary>
        /// The document
        /// </summary>
        internal AeccRoadwayDocument _document;
        /// <summary>
        /// Gets the baselines.
        /// </summary>
        /// <value>
        /// The baselines.
        /// </value>
        public IList<Baseline> Baselines { get { return _baselines; } }
        /// <summary>
        /// Gets the Corridor name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return _corridor.DisplayName; } }
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal AeccCorridor InternalElement { get { return this._corridor; } }

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="Corridor"/> class.
        /// </summary>
        /// <param name="corridor">The corridor.</param>
        /// <param name="doc">The document.</param>
        internal Corridor(AeccCorridor corridor, AeccRoadwayDocument doc)
        {
            this._corridor = corridor;
            this._document = doc;
            IList<Baseline> bls = new List<Baseline>();

            int index = 0;
            foreach (AeccBaseline b in corridor.Baselines)
            {
                bls.Add(new Baseline(b, index));
                ++index;
            }

            this._baselines = bls;
        }

        /// <summary>
        /// Rebuilds the Corridor in Civil 3D.
        /// </summary>
        /// <returns></returns>
        public Corridor Rebuild()
        {
            this.InternalElement.Rebuild();
            return this;
        }

        #endregion

        #region PRIVATE METHODS

        /// <summary>
        /// Returns the points that define the subassemblies in a corridor organized by:
        /// Corridor &gt; Baseline &gt; Region &gt; Assembly &gt; Subassembly
        /// </summary>
        /// <returns>
        /// The list of points that define each subassembly in the corridor
        /// </returns>
        /// <search> Subassembly, section</search>
        private IList<IList<IList<IList<IList<IList<Point>>>>>> GetPointsBySubassembly()
        {
            Utils.Log(string.Format("Corridor.GetPointsBySubassembly started...", ""));

            IList<IList<IList<IList<IList<IList<Point>>>>>> output = new List<IList<IList<IList<IList<IList<Point>>>>>>();

            foreach (AeccBaseline b in this._corridor.Baselines)
            {
                IList<IList<IList<IList<IList<Point>>>>> baselineColl = new List<IList<IList<IList<IList<Point>>>>>();

                foreach (AeccBaselineRegion blr in b.BaselineRegions)
                {
                    IList<IList<IList<IList<Point>>>> regionColl = new List<IList<IList<IList<Point>>>>();

                    foreach (AeccAppliedAssembly assembly in blr.AppliedAssemblies)
                    {
                        IList<IList<IList<Point>>> assemblyColl = new List<IList<IList<Point>>>();

                        foreach (AeccAppliedSubassembly sub in assembly.AppliedSubassemblies)
                        {
                            IList<IList<Point>> subColl = new List<IList<Point>>();

                            foreach (AeccCalculatedShape shape in sub.CalculatedShapes)
                            {
                                IList<Point> shapeColl = new List<Point>();

                                foreach (IAeccCalculatedLink link in shape.CalculatedLinks)
                                {
                                    foreach (AeccCalculatedPoint p in link.CalculatedPoints)
                                    {
                                        dynamic xyz = b.StationOffsetElevationToXYZ(p.GetStationOffsetElevationToBaseline());

                                        Point point = Point.ByCoordinates(xyz[0], xyz[1], xyz[2]);

                                        shapeColl.Add(point);
                                    }
                                }

                                subColl.Add(Point.PruneDuplicates(shapeColl));
                            }

                            assemblyColl.Add(subColl);
                        }

                        regionColl.Add(assemblyColl);
                    }

                    baselineColl.Add(regionColl);
                }

                output.Add(baselineColl);
            }

            Utils.Log(string.Format("Corridor.GetPointsBySubassembly completed.", ""));

            return output;
        }

        /// <summary>
        /// Gets the corridor surfaces.
        /// </summary>
        /// <returns></returns>
        private IList<IList<Surface>> GetCorridorSurfaces()
        {
            Utils.Log(string.Format("Corridor.GetCorridorSurfaces started...", ""));

            IList<IList<Surface>> output = new List<IList<Surface>>();

            if (null != this._corridor.CorridorSurfaces)
            {
                foreach (AeccCorridorSurface s in this._corridor.CorridorSurfaces)
                {
                    IList<Surface> surfaces = new List<Surface>();

                    foreach (AeccCorridorSurfaceMask b in s.Masks)
                    {
                        IList<Point> dSpoints = new List<Point>();

                        foreach (double[] point in b.GetPolygonPoints())
                        {
                            dSpoints.Add(Point.ByCoordinates(point[0], point[1], point[2]));
                        }

                        surfaces.Add(Surface.ByPerimeterPoints(Point.PruneDuplicates(dSpoints)));  // [20181009]
                    }

                    foreach (AeccCorridorSurfaceBoundary b in s.Boundaries)
                    {
                        IList<Point> dSpoints = new List<Point>();

                        foreach (double[] point in b.GetPolygonPoints())
                        {
                            dSpoints.Add(Point.ByCoordinates(point[0], point[1], point[2]));
                        }

                        surfaces.Add(Surface.ByPerimeterPoints(Point.PruneDuplicates(dSpoints)));  // [20181009]
                    }

                    if (surfaces.Count > 0)
                    {
                        output.Add(surfaces);
                    }
                }
            }

            Utils.Log(string.Format("Corridor.GetCorridorSurfaces completed.", ""));

            //TODO raise exception for no corridor surfaces
            return output;
        }


        #region Old Code
        /// <summary>
        /// Gets the points by code1.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        private IList<IList<IList<IList<Point>>>> GetPointsByCode1(string code)
        {
            IList<IList<IList<IList<Point>>>> output = new List<IList<IList<IList<Point>>>>();

            foreach (AeccBaseline b in this._corridor.Baselines)
            {
                IList<IList<IList<Point>>> baseline = new List<IList<IList<Point>>>();

                foreach (AeccBaselineRegion reg in b.BaselineRegions)
                {
                    IList<IList<Point>> region = new List<IList<Point>>();

                    foreach (AeccAppliedAssembly assembly in reg.AppliedAssemblies)
                    {
                        IList<Point> temp = new List<Point>();

                        foreach (AeccCalculatedPoint p in assembly.GetPointsByCode(code))
                        {
                            dynamic soe = p.GetStationOffsetElevationToBaseline();

                            if (soe[0] >= reg.StartStation && soe[0] <= reg.EndStation)
                            {
                                dynamic xyz = b.StationOffsetElevationToXYZ(soe);

                                Point point = Point.ByCoordinates(xyz[0], xyz[1], xyz[2]);

                                temp.Add(point);
                            }
                        }

                        region.Add(temp);
                    }

                    baseline.Add(region);
                }

                output.Add(baseline);
            }

            return output;
        }

        /// <summary>
        /// Gets the feature line points.
        /// </summary>
        /// <returns></returns>
        private IList<IList<IList<IList<Point>>>> GetFeatureLinePoints()
        {
            IList<IList<IList<IList<Point>>>> output = new List<IList<IList<IList<Point>>>>();

            foreach (AeccBaseline b in this._corridor.Baselines)
            {
                IList<IList<IList<Point>>> baseline = new List<IList<IList<Point>>>();

                foreach (AeccBaselineRegion blr in b.BaselineRegions)
                {
                    IList<IList<Point>> region = new List<IList<Point>>();

                    foreach (AeccFeatureLines coll in b.MainBaselineFeatureLines.FeatureLinesCol)
                    {
                        foreach (AeccFeatureLine f in coll)
                        {
                            IList<Point> featureline = new List<Point>();

                            foreach (AeccFeatureLinePoint p in f.FeatureLinePoints)
                            {
                                if (p.Station >= blr.StartStation && p.Station <= blr.EndStation)
                                {
                                    Point point = Point.ByCoordinates(p.XYZ[0], p.XYZ[1], p.XYZ[2]);

                                    featureline.Add(point);
                                }
                            }

                            region.Add(featureline);
                        }
                    }

                    baseline.Add(region);
                }

                output.Add(baseline);
            }

            return output;
        }

        [MultiReturn(new string[] { "Featurelines" })]
        private Dictionary<string, object> TestCorridorInfo_Old(string code)
        {
            IList<string[]> corridorCodes = new List<string[]>();
            IList<IList<Featureline>> corridorFeaturelines = new List<IList<Featureline>>();

            foreach (Baseline bl in this.Baselines)
            {
                IList<Featureline> blFeaturelines = new List<Featureline>();

                var b = bl._baseline;

                foreach (AeccFeatureLines coll in b.MainBaselineFeatureLines.FeatureLinesCol)
                {
                    foreach (AeccFeatureLine f in coll)
                    {
                        if (f.CodeName == code)
                        {
                            IList<Point> featureline = new List<Point>();

                            foreach (AeccFeatureLinePoint p in f.FeatureLinePoints)
                            {
                                Point point = Point.ByCoordinates(p.XYZ[0], p.XYZ[1], p.XYZ[2]);

                                featureline.Add(point);
                            }

                            featureline = Point.PruneDuplicates(featureline);

                            PolyCurve pc = PolyCurve.ByPoints(featureline);

                            var offset = bl.GetArrayStationOffsetElevationByPoint(pc.PointAtParameter(0.5))[1];

                            Featureline.SideType side = Featureline.SideType.Right;

                            if (offset < 0)
                            {
                                side = Featureline.SideType.Left;
                            }

                            blFeaturelines.Add(new Featureline(bl, pc, f.CodeName, side));
                        }
                    }
                }

                corridorFeaturelines.Add(blFeaturelines);
            }

            return new Dictionary<string, object>() { { "Featurelines", corridorFeaturelines } };
        }

        /// <summary>
        /// Gets the featurelines by Code &gt; Baseline &gt; Region.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        private IList<IList<IList<Featureline>>> GetFeaturelinesByCode_Old1(string code)
        {
            IList<string[]> corridorCodes = new List<string[]>();
            IList<IList<IList<Featureline>>> corridorFeaturelines = new List<IList<IList<Featureline>>>();

            foreach (Baseline bl in this.Baselines)
            {
                IList<IList<Featureline>> blFeaturelines = new List<IList<Featureline>>();

                var b = bl._baseline;

                int regionIndex = 0;

                foreach (AeccBaselineRegion region in b.BaselineRegions)
                {
                    IList<Featureline> regFeaturelines = new List<Featureline>();

                    foreach (AeccFeatureLines coll in b.MainBaselineFeatureLines.FeatureLinesCol)
                    {
                        foreach (AeccFeatureLine f in coll)
                        {
                            if (f.CodeName == code)
                            {
                                IList<Point> points = new List<Point>();

                                foreach (AeccFeatureLinePoint p in f.FeatureLinePoints)
                                {
                                    Point point = Point.ByCoordinates(p.XYZ[0], p.XYZ[1], p.XYZ[2]);

                                    double s = Math.Round(bl.GetArrayStationOffsetElevationByPoint(point)[0], 5);

                                    if (s >= region.StartStation || Math.Abs(s - region.StartStation) < 0.001)
                                    {
                                        if (s <= region.EndStation || Math.Abs(s - region.EndStation) < 0.001)
                                        {
                                            points.Add(point);
                                        }
                                    }
                                }

                                points = Point.PruneDuplicates(points);

                                if (points.Count > 1)
                                {
                                    PolyCurve pc = PolyCurve.ByPoints(points);

                                    var soeStart = bl.GetArrayStationOffsetElevationByPoint(pc.PointAtParameter(0));
                                    var soeEnd = bl.GetArrayStationOffsetElevationByPoint(pc.PointAtParameter(1));
                                    double offset = soeStart[1];

                                    if (soeStart[0] > soeEnd[0])
                                    {
                                        pc = pc.Reverse() as PolyCurve;
                                        offset = bl.GetArrayStationOffsetElevationByPoint(pc.PointAtParameter(0))[1];
                                    }

                                    Featureline.SideType side = Featureline.SideType.Right;

                                    if (offset < 0)
                                    {
                                        side = Featureline.SideType.Left;
                                    }

                                    regFeaturelines.Add(new Featureline(bl, pc, f.CodeName, side, regionIndex));
                                }
                            }
                        }
                    }

                    blFeaturelines.Add(regFeaturelines);

                    regionIndex++;
                }

                corridorFeaturelines.Add(blFeaturelines);
            }

            return corridorFeaturelines;
        }

        #endregion
        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Returns a Point by station offset elevation.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="station">The station.</param>
        /// <param name="offset">The offset.</param>
        /// <param name="elevation">The elevation.</param>
        /// <returns></returns>
        public Point PointByStationOffsetElevation(Baseline baseline, double station = 0, double offset = 0, double elevation = 0)
        {
            return baseline.PointByStationOffsetElevation(station, offset, elevation);
        }

        /// <summary>
        /// Returns a CoordinateSystem by station.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="station">The station.</param>
        /// <returns></returns>
        public CoordinateSystem CoordinateSystemByStation(Baseline baseline, double station = 0)
        {
            return baseline.CoordinateSystemByStation(station);
        }

        /// <summary>
        /// Returns a CoordinateSystem by point.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="point">The point.</param>
        /// <returns></returns>
        public CoordinateSystem CoordinateSystemByPoint(Baseline baseline, Point point)
        {
            return baseline.GetCoordinateSystemByPoint(point);
        }

        /// <summary>
        /// Gets the PointCodes.
        /// </summary>
        /// <returns></returns>
        public IList<string> GetCodes()
        {
            Utils.Log(string.Format("Corridor.GetCodes started...", ""));

            IList<string> output = new List<string>();

            IList<AeccCorridorCodes> codeList = new List<AeccCorridorCodes>();

            try
            {
                foreach (AeccBaseline b in this._corridor.Baselines)
                {
                    foreach (string code in b.MainBaselineFeatureLines.CodeNames)
                    {
                        if (!output.Contains(code))
                        {
                            output.Add(code);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.Log(string.Format("ERROR: {0}", ex.Message));
            }

            Utils.Log(string.Format("Corridor.GetCodes completed.", ""));

            return output.OrderBy(x => x).ToList();
        }

        /// <summary>
        /// Gets the corridor Featurelies organized by Corridor-Baseline-Code-Region
        /// </summary>
        /// <returns></returns>
        public IList<IList<IList<Featureline>>> GetFeaturelines()
        {
            return Utils.GetFeaturelines(this);
        }

        /// <summary>
        /// Gets the subassembly points organized by: Corridor &gt; Baseline &gt; Region &gt; Assembly &gt; Subassembly.
        /// </summary>
        /// <param name="dumpXML">If true dumps a LandXML in the Temp folder.</param>
        /// <returns></returns>
        public IList<IList<IList<IList<IList<Point>>>>> GetSubassemblyPoints(bool dumpXML=false)
        {
            return Utils.GetCorridorSubAssembliesFromLandXML(this, dumpXML);
        }

        /// <summary>
        ///  Gets the subassembly points organized by: Corridor &gt; Baseline &gt; Region &gt; Code.
        ///  It requires to export a LandXML to the %Temp% folder, named like the Civil 3D Document, containing only the corridor data.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        public IList<IList<IList<IList<IList<Point>>>>> GetPointsByCode(string code)
        {
            return Utils.GetCorridorPointsByCodeFromLandXML(this, code);
        }

       
        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("Corridor(Name = {0})", this.Name);
        }

        /// <summary>
        /// Gets the closest featureline by point code side.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <param name="baselineIndex">Index of the baseline.</param>
        /// <param name="code">The code.</param>
        /// <param name="side">The side.</param>
        /// <returns></returns>
        [MultiReturn(new string[] { "Featureline" })]
        public Dictionary<string, object> GetFeaturelineByPointCodeSide(Point point, int baselineIndex, string code, string side)
        {
            return new Dictionary<string, object>() { { "Featureline", UtilsObjectsLocation.ClosestFeaturelineByPoint(point, this, baselineIndex, code, side) } };
        }

        /// <summary>
        /// Gets the featurelines by Code &gt; Baseline &gt; Region.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        public IList<IList<IList<Featureline>>> GetFeaturelinesByCode(string code)  // 1.1.0
        {
            Utils.Log(string.Format("Corridor.GetFeaturelinesByCode started...", ""));

            IList<IList<IList<Featureline>>> corridorFeaturelines = new List<IList<IList<Featureline>>>();

            foreach (Baseline bl in this.Baselines)
            {
                IList<IList<Featureline>> blFeaturelines = bl.GetFeaturelinesByCode(code);

                corridorFeaturelines.Add(blFeaturelines);
            }

            Utils.Log(string.Format("Corridor.GetFeaturelinesByCode completed.", ""));

            return corridorFeaturelines;
        }

        /// <summary>
        /// Gets the featurelines by Code &gt; Baseline &gt; Region.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <param name="station">The station.</param>
        /// <returns></returns>
        public IList<IList<Featureline>> GetFeaturelinesByCodeStation(string code, double station)  // 1.1.0
        {
            Utils.Log(string.Format("Corridor.GetFeaturelinesByCodeStation started...", ""));

            IList<IList<Featureline>> corridorFeaturelines = new List<IList<Featureline>>();

            foreach (Baseline bl in this.Baselines)
            {
                IList<Featureline> blFeaturelines = bl.GetFeaturelinesByCodeStation(code, station);

                corridorFeaturelines.Add(blFeaturelines);
            }

            Utils.Log(string.Format("Corridor.GetFeaturelinesByCodeStation completed.", ""));

            return corridorFeaturelines;
        }

        #endregion
    }
}
