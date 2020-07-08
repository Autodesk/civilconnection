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
    /// ProfilePVI object type.
    /// </summary>
    [IsVisibleInDynamoLibrary(false)]
    public class ProfilePVI
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The pvi
        /// </summary>
        private AeccProfilePVI _pvi;

        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the station.
        /// </summary>
        /// <value>
        /// The station.
        /// </value>
        public double Station { get { return _pvi.Station; } }
        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this._pvi; } }
        /// <summary>
        /// Gets the elevation.
        /// </summary>
        /// <value>
        /// The elevation.
        /// </value>
        public double Elevation { get { return _pvi.Elevation; } }
        /// <summary>
        /// Gets the grade in.
        /// </summary>
        /// <value>
        /// The grade in.
        /// </value>
        public double GradeIn { get { return _pvi.GradeIn; } }
        /// <summary>
        /// Gets the grade out.
        /// </summary>
        /// <value>
        /// The grade out.
        /// </value>
        public double GradeOut { get { return _pvi.GradeOut; } }

        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Initializes a new instance of the <see cref="ProfilePVI"/> class.
        /// </summary>
        /// <param name="pvi">The pvi.</param>
        internal ProfilePVI(AeccProfilePVI pvi)
        {
            this._pvi = pvi;
        }

        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS


        //TODO: Get profile curves        

        /// <summary>
        /// Returns a text representation of the object.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return string.Format("ProfilePVI(Station = {0}, Elevation = {1}, GradeIn = {2}, GradeOut = {3})", this.Station, this.Elevation, this.GradeIn, this.GradeOut);
        }
        #endregion
    }
}
