// Copyright (c) 2019 DIMCO, N.V All rights reserved.
// Author: de.maesschalck.bart@deme-group.com
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

namespace CivilConnection
{
    /// <summary>
    /// FeaturelinePoint obejct type.
    /// </summary>
    [DynamoServices.RegisterForTrace()]
    public class FeaturelinePoint
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The baseline
        /// </summary>
        Baseline _baseline;
        /// <summary>
        /// The code
        /// </summary>
        string _code;
        /// <summary>
        /// The position
        /// </summary>
        Point _position;
        /// <summary>
        /// The station
        /// </summary>
        double _station;
        /*// <summary>
        /// The offset
        /// </summary>
        //double _offset;
        /// <summary>
        */// The CoordinateSystem for the baseline on the point
        /// </summary>
        CoordinateSystem _cs;
     
        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the baseline.
        /// </summary>
        /// <value>
        /// The baseline.
        /// </value>
        public Baseline Baseline { get { return this._baseline; } }
        /// <summary>
        /// Gets the code.
        /// </summary>
        /// <value>
        /// The code.
        /// </value>
        public string Code { get { return this._code; } }
        /// <summary>
        /// Gets the position.
        /// </summary>
        /// <value>
        /// XYZ arraty of the position.
        /// </value>
        public Point Position { get { return this._position; } }
        /// <summary>
        /// Gets the station.
        /// </summary>
        /// <value>
        /// The station.
        /// </value>
        public double Station { get  { return this._station; } }
         
        ///<summary>
        /// Get the coordinate system at the point
        ///</summary>        
        public CoordinateSystem CoordinateSystem {  get { return this._cs; } }
        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="FeaturelinePoint"/> class.
        /// </summary>
        /// <param name="baseline">The baseline.</param>
        /// <param name="pt">The point.</param>
        /// <param name="code">The code.</param>
        /// <param name="station">The station on the baseline.</param>
        internal FeaturelinePoint(Baseline baseline, Point pt, string code, double station)
        {
            this._baseline = baseline;
            this._position = pt;
            this._code = code;
            this._station = station;
            this._cs = this._baseline. CoordinateSystemByStation(_station);
        }

        #endregion
    }
}
