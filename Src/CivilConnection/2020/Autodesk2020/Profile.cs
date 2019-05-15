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

        /// <summary>
        /// Gets the stations of the PVIs.
        /// </summary>
        /// <value>
        /// The PVIStations.
        /// </value>
        public double[] PVIStations { get { return _profile.PVIs.Cast<AeccProfilePVI>().Select(x => x.Station).ToArray(); } }

        /// <summary>
        /// Gets the elevation of the PVIs.
        /// </summary>
        /// <value>
        /// The PVIElevations.
        /// </value>
        public double[] PVIElevations { get { return _profile.PVIs.Cast<AeccProfilePVI>().Select(x => x.Elevation).ToArray(); } }

        /// <summary>
        /// Gets the grade in of the PVIs.
        /// </summary>
        /// <value>
        /// The PVIGradeIns.
        /// </value>
        private double[] PVIGradeIns { get { return _profile.PVIs.Cast<AeccProfilePVI>().Select(x => x.GradeIn).ToArray(); } }

        /// <summary>
        /// Gets the grade out of the PVIs.
        /// </summary>
        /// <value>
        /// The PVIGradeOuts.
        /// </value>
        private double[] PVIGradeOuts { get { return _profile.PVIs.Cast<AeccProfilePVI>().Select(x => x.GradeOut).ToArray(); } }
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
        /// Gets the elevations of the entities in the profile.
        /// </summary>
        /// <returns></returns>
        public IList<double> GetEntitiesElevations()
        {
            Utils.Log(string.Format("Profile.GetEntitiesElevations Started...", ""));

            IList<double> elevations = new List<double>();

            Dictionary<double, IAeccProfileEntity> entities = new Dictionary<double, IAeccProfileEntity>();

            foreach (IAeccProfileEntity ent in _profile.Entities)
            {
                double start = 0;

                if (ent.Type == AeccProfileEntityType.aeccProfileEntityTangent)
                {
                    var c = ent as aeccProfileTangent;
                    start = c.StartStation;
                    entities.Add(start, c);
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveSymmetricParabola)
                {
                    var c = ent as AeccProfileCurveParabolic;
                    start = c.StartStation;
                    entities.Add(start, c);
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveCircular)
                {
                    var c = ent as aeccProfileCurveCircular;
                    start = c.StartStation;
                    entities.Add(start, c);
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveAsymmetricParabola)
                {
                    var c = ent as AeccProfileCurveAsymmetric;
                    start = c.StartStation;
                    entities.Add(start, c);
                }
            }

            double[] stations = entities.Keys.OrderBy(x => x).ToArray();

            for (int i = 0; i < stations.Length; ++i )
            {
                double s = stations[i];

                IAeccProfileEntity ent = entities[s];

                if (ent.Type == AeccProfileEntityType.aeccProfileEntityTangent)
                {
                    var c = ent as aeccProfileTangent;
                    elevations.Add(c.StartElevation);
                    if (i == stations.Length - 1)
                    {
                        elevations.Add(c.EndElevation);
                    }
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveSymmetricParabola)
                {
                    var c = ent as AeccProfileCurveParabolic;
                    elevations.Add(c.StartElevation);
                    elevations.Add(c.HighLowPointElevation);
                    if (i == stations.Length - 1)
                    {
                        elevations.Add(c.EndElevation);
                    }
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveCircular)
                {
                    var c = ent as aeccProfileCurveCircular;
                    elevations.Add(c.StartElevation);
                    elevations.Add(c.HighLowPointElevation);
                    if (i == stations.Length - 1)
                    {
                        elevations.Add(c.EndElevation);
                    }
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveAsymmetricParabola)
                {
                    var c = ent as AeccProfileCurveAsymmetric;
                    elevations.Add(c.StartElevation);
                    elevations.Add(c.HighLowPointElevation);
                    if (i == stations.Length - 1)
                    {
                        elevations.Add(c.EndElevation);
                    }
                }
            }

            Utils.Log(string.Format("Elements: {0}", elevations.Count));

            foreach (double el in elevations)
            {
                Utils.Log(string.Format("{0}", el));
            }

            Utils.Log(string.Format("Profile.GetEntitiesElevations Completed.", ""));

            return elevations;
        }

        /// <summary>
        /// Gets the stations of the entities in the profile.
        /// </summary>
        /// <returns></returns>
        public IList<double> GetEntitiesStations()
        {
            Utils.Log(string.Format("Profile.GetEntitiesStations Started...", ""));

            IList<double> stations = new List<double>();

            foreach (IAeccProfileEntity ent in _profile.Entities)
            {
                if (ent.Type == AeccProfileEntityType.aeccProfileEntityTangent)
                {
                    var c = ent as aeccProfileTangent;
                    stations.Add(c.StartStation);
                    stations.Add(c.EndStation);                   
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveSymmetricParabola)
                {
                    var c = ent as AeccProfileCurveParabolic;
                    stations.Add(c.StartStation);
                    stations.Add(c.HighLowPointStation);
                    stations.Add(c.EndStation);
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveCircular)
                {
                    var c = ent as aeccProfileCurveCircular;
                    stations.Add(c.StartStation);
                    stations.Add(c.HighLowPointStation);
                    stations.Add(c.EndStation);
                }
                else if (ent.Type == AeccProfileEntityType.aeccProfileEntityCurveAsymmetricParabola)
                {
                    var c = ent as AeccProfileCurveAsymmetric;
                    stations.Add(c.StartStation);
                    stations.Add(c.HighLowPointStation);
                    stations.Add(c.EndStation);
                }
            }

            stations = stations.Distinct().OrderBy(x => x).ToList();

            Utils.Log(string.Format("Elements: {0}", stations.Count));

            foreach (double el in stations)
            {
                Utils.Log(string.Format("{0}", el));
            }

            Utils.Log(string.Format("Profile.GetEntitiesStations Completed.", ""));

            return stations;
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

