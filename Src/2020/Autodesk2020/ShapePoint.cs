// Copyright (c) 2017 Autodesk, Inc. All rights reserved.
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
using Newtonsoft.Json;
using System;


namespace CivilConnection
{
    /// <summary>
    /// Shape Point Class
    /// </summary>
    //[IsVisibleInDynamoLibrary(false)]

    [JsonConverter(typeof(ShapePointConverter))]
    public class ShapePoint
    {
        #region PRIVATE PROPERTIES
        [JsonIgnore]
        private int _id = 0;

        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the unique identifier.
        /// </summary>
        /// <value>
        /// The unique identifier.
        /// </value>
        [JsonProperty("UniqueId")]
        public string UniqueId { get; set; }
        /// <summary>
        /// Gets the identifier.
        /// </summary>
        /// <value>
        /// The identifier.
        /// </value>
        /// 
        [JsonProperty("Id")]
        public int Id { get { return this._id; } set { this._id = value; } }
        /// <summary>
        /// Gets the corridor.
        /// </summary>
        /// <value>
        /// The corridor.
        /// </value>
        /// 
        [JsonProperty("Corridor")]
        public string Corridor { get; set; }
        /// <summary>
        /// Gets the index of the baseline.
        /// </summary>
        /// <value>
        /// The index of the baseline.
        /// </value>
        /// 
        [JsonProperty("BaselineIndex")]
        public int BaselineIndex { get; set; }
        /// <summary>
        /// Gets the side.
        /// </summary>
        /// <value>
        /// The side.
        /// </value>
        /// 
        [JsonProperty("Side")]
        public Featureline.SideType Side { get; set; }
        /// <summary>
        /// Gets the code.
        /// </summary>
        /// <value>
        /// The code.
        /// </value>
        /// 
        [JsonProperty("Code")]
        public string Code { get; set; }
        /// <summary>
        /// Gets the station.
        /// </summary>
        /// <value>
        /// The station.
        /// </value>
        /// 
        [JsonProperty("Station")]
        public double Station { get; set; }
        /// <summary>
        /// Gets the offset.
        /// </summary>
        /// <value>
        /// The offset.
        /// </value>
        /// 
        [JsonProperty("Offset")]
        public double Offset { get; set; }
        /// <summary>
        /// Gets the elevation.
        /// </summary>
        /// <value>
        /// The elevation.
        /// </value>
        [JsonProperty("Elevation")]
        public double Elevation { get; set; }
        /// <summary>
        /// Gets the point.
        /// </summary>
        /// <value>
        /// The point.
        /// </value>
        [JsonIgnore]
        public Point Point { get; private set; }
        /// <summary>
        /// Gets the featureline.
        /// </summary>
        /// <value>
        /// The featureline.
        /// </value>
        [JsonIgnore]
        public Featureline Featureline { get; private set; }

        /// <summary>
        /// Gets the point in Revit local coordinate system.
        /// </summary>
        /// <value>
        /// The revit point.
        /// </value>
        [JsonIgnore]
        public Point RevitPoint { get; private set; }

        /// <summary>
        /// Gets the Featureline region index
        /// </summary>
        [JsonProperty("RegionIndex")]
        public int RegionIndex { get; set; }

        /// <summary>
        /// Gets the Featureline region index
        /// </summary>
        [JsonProperty("RegionRelative")]
        public double RegionRelative { get; set; }  // 1.1.0

        /// <summary>
        /// Gets the Featureline region index
        /// </summary>
        [JsonProperty("RegionNormalized")]
        public double RegionNormalized { get; set; }  // 1.1.0

        #endregion

        #region CONSTRUCTOR
        internal ShapePoint() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapePoint"/> class.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <param name="featureline">The featureline.</param>
        /// <param name="id">The initial ID of the ShapePoint.</param>
        internal ShapePoint(Point point, Featureline featureline, int id = 0)
        {
            this.UniqueId = Guid.NewGuid().ToString();
            this._id = id;

            this.Corridor = featureline.Baseline.CorridorName;
            this.BaselineIndex = featureline.Baseline.Index;
            this.RegionIndex = featureline.BaselineRegionIndex;

            this.Side = featureline.Side;
            this.Code = featureline.Code;

            var soe = featureline.GetStationOffsetElevationByPoint(point);
            this.Station = (double)soe["Station"];
            this.Offset = (double)soe["Offset"];
            this.Elevation = (double)soe["Elevation"];

            this.RegionRelative = this.Station - featureline.Start;  // 1.1.0
            this.RegionNormalized = this.RegionRelative / (featureline.End - featureline.Start);  // 1.1.0

            this.Point = point;
            this.Featureline = featureline;

            this.RevitPoint = point.Transform(RevitUtils.DocumentTotalTransform()) as Point;
        }

        [JsonConstructor]
        internal ShapePoint(string guid, int id, string corridor, int baselineIndex, int regionIndex, double regionRelative,
            double regionNormalized, string code, int side, double station, double offset, double elevation)  // 1.1.0
        {
            this.UniqueId = guid;
            this.Id = id;

            this.Corridor = corridor;
            this.BaselineIndex = baselineIndex;
            this.RegionIndex = regionIndex;
            this.RegionRelative = regionRelative;  // 1.1.0
            this.RegionNormalized = regionNormalized;  // 1.1.0
            this.Code = code;
            this.Station = station;
            this.Offset = offset;
            this.Elevation = elevation;

            if (side == 0)
            {
                this.Side = CivilConnection.Featureline.SideType.None;
            }
            else if (side == 1)
            {
                this.Side = CivilConnection.Featureline.SideType.Left;
            }
            else if (side == 2)
            {
                this.Side = CivilConnection.Featureline.SideType.Right;
            }
        }

        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Bies the point featureline.
        /// </summary>
        /// <param name="point">The point.</param>
        /// <param name="featureline">The featureline.</param>
        /// <returns></returns>
        public static ShapePoint ByPointFeatureline(Point point, Featureline featureline)
        {
            ShapePoint sp = new ShapePoint(point, featureline);

            return sp;
        }

        /// <summary>
        /// Copies this instance.
        /// </summary>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public ShapePoint Copy(int id=0)
        {
            return new ShapePoint(this.Point, this.Featureline, id);
        }

        /// <summary>
        /// Sets the identifier.
        /// </summary>
        /// <param name="newId">The new identifier.</param>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public ShapePoint SetId(int newId)
        {
            this._id = newId;
            return this;
        }

        /// <summary>
        /// Sets the point.
        /// </summary>
        /// <returns></returns>
        [IsVisibleInDynamoLibrary(false)]
        public ShapePoint SetPoint(Point point)
        {
            this.Point = point;
            this.RevitPoint = this.Point.Transform(RevitUtils.DocumentTotalTransform()) as Point;
            return this;
        }
        
        /// <summary>
        /// Calculates the new ShapePoint location on the new Featureline using the Station, Offset and Elevation parameters.
        /// </summary>
        /// <param name="featureline">The featureline used to update the ShapePoint</param>
        /// <returns></returns>
        public ShapePoint UpdateByFeatureline(Featureline featureline)
        {
            Utils.Log(string.Format("ShapePoint.UpdateByFeatureline started...", ""));

            var p = featureline.PointByStationOffsetElevation(this.Station, this.Offset, this.Elevation, false);
            this.Point = p;
            this.RevitPoint = this.Point.Transform(RevitUtils.DocumentTotalTransform()) as Point;
            this.Featureline = featureline;

            this.BaselineIndex = featureline.Baseline.Index;  // 1.1.0
            this.RegionIndex = featureline.BaselineRegionIndex;  // 1.1.0
            this.RegionRelative = this.Station - featureline.Start;  // 1.1.0
            this.RegionNormalized = this.RegionRelative / (featureline.End - featureline.Start);  // 1.1.0
            this.Code = featureline.Code;  // 1.1.0
            this.Corridor = featureline.Baseline.CorridorName;  // 1.1.0
            this.Side = featureline.Side;  // 1.1.0

            Utils.Log(string.Format("ShapePoint.UpdateByFeatureline completed.", ""));

            return this;
        }


        /// <summary>
        /// Updates the ShapePoint parameters Station, Offset and Elevation against the new Featureline.
        /// </summary>
        /// <param name="featureline">The featureline assigned to the ShapePoint</param>
        /// <returns></returns>
        public ShapePoint AssignFeatureline(Featureline featureline)  // 1.1.0
        {
            Utils.Log(string.Format("ShapePoint.AssignFeatureline started...", ""));

            this.Featureline = featureline;

            var soe = featureline.GetStationOffsetElevationByPoint(this.Point);
            this.Station = (double)soe["Station"];
            this.Offset = (double)soe["Offset"];
            this.Elevation = (double)soe["Elevation"];

            this.BaselineIndex = featureline.Baseline.Index;
            this.RegionIndex = featureline.BaselineRegionIndex;
            this.RegionRelative = this.Station - featureline.Start;
            this.RegionNormalized = this.RegionRelative / (featureline.End - featureline.Start);
            this.Code = featureline.Code;
            this.Corridor = featureline.Baseline.CorridorName;
            this.Side = featureline.Side;

            Utils.Log(string.Format("ShapePoint.AssignFeatureline completed.", ""));

            return this;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("ShapePoint({0}, Id={1}, Code={2}, Corridor={3}, Station={4}, Offset={5}, Elevation={6})",
                this.Point,
                this.Id,
                this.Code,
                this.Corridor,
                this.Station,
                this.Offset,
                this.Elevation);
        }

        #endregion
    }
}
