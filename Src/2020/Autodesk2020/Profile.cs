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
using Autodesk.DesignScript.Geometry;
using System.Collections.Generic;

namespace CivilConnection
{
    /// <summary>
    /// Profile obejct type.
    /// </summary>
    public class Profile
    {
        #region PRIVATE PROPERTIES

        /// <summary>
        /// The profile
        /// </summary>
        private AeccProfile _profile;
        /// <summary>
        /// The entities
        /// </summary>
        private AeccProfileEntities _entities;

        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return _profile.DisplayName; } }
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this._profile; } }
        /// <summary>
        /// Gets the length.
        /// </summary>
        /// <value>
        /// The length.
        /// </value>
        public double Length { get { return _profile.Length; } }
        /// <summary>
        /// Gets the maximum elevation.
        /// </summary>
        /// <value>
        /// The maximum elevation.
        /// </value>
        public double MaxElevation { get { return _profile.ElevationMax; } }
        /// <summary>
        /// Gets the minimum elevation.
        /// </summary>
        /// <value>
        /// The minimum elevation.
        /// </value>
        public double MinElevation { get { return _profile.ElevationMin; } }
        /// <summary>
        /// Gets the start station.
        /// </summary>
        /// <value>
        /// The start station.
        /// </value>
        public double StartStation { get { return _profile.StartingStation; } }
        /// <summary>
        /// Gets the end station.
        /// </summary>
        /// <value>
        /// The end station.
        /// </value>
        public double EndStation { get { return _profile.EndingStation; } }
        /// <summary>
        /// Gets the weed grade factor.
        /// </summary>
        /// <value>
        /// The weed grade factor.
        /// </value>
        public double WeedGradeFactor { get { return _profile.WeedGradeFactor; } }

        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Initializes a new instance of the <see cref="Profile"/> class.
        /// </summary>
        /// <param name="profile">The profile.</param>
        internal Profile(AeccProfile profile)
        {
            this._profile = profile;
            this._entities = profile.Entities;
        }

        #endregion

        #region PRIVATE METHODS

        //TODO: Get profile curves
        /// <exclude />
        private void XX()
        {
            Utils.Log(string.Format("Profile.XX started...", ""));

            IList<Curve> output = new List<Curve>();

            for (int i = 0; i < this._entities.Count; ++i)
            {
                var e = this._entities.Item(i);

                if (e.Type == AeccProfileEntityType.aeccProfileEntityTangent)
                {
                    var tangent = e as aeccProfileTangent;

                    var start = Point.ByCoordinates(tangent.StartStation, tangent.StartElevation);
                    var end = Point.ByCoordinates(tangent.EndStation, tangent.EndElevation);

                    output.Add(Line.ByStartPointEndPoint(start, end));
                }
                else if (e.Type == AeccProfileEntityType.aeccProfileEntityCurveCircular)
                {
                    var arc = e as aeccProfileCurveCircular;

                    var start = Point.ByCoordinates(arc.StartStation, arc.StartElevation);
                    var end = Point.ByCoordinates(arc.EndStation, arc.EndElevation);
                    var pvi = Point.ByCoordinates(arc.PVIStation, arc.PVIElevation);
                    double radius = arc.Radius;
                }
            }

            Utils.Log(string.Format("Profile.XX completed.", ""));
        }

        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Gets the elevation at station.
        /// </summary>
        /// <param name="station">The station.</param>
        /// <returns></returns>
        public double GetElevationAtStation(double station)
        {
            return this._profile.ElevationAt(station);
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("Profile(Name = {0})", this.Name);
        }

        #endregion
    }
}
