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




namespace CivilConnection
{
    /// <summary>
    /// Session Variables utilities.
    /// </summary>
    class SessionVariables
    {
        #region PRIVATE PROPERTIES

        #endregion

        #region PUBLIC PROPERTIES

        /// <summary>
        /// Gets or sets the land XML path.
        /// </summary>
        /// <value>
        /// The land XML path.
        /// </value>
        public static string LandXMLPath { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is land XML exported.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is land XML exported; otherwise, <c>false</c>.
        /// </value>
        public static bool IsLandXMLExported { get; set; }

        /// <summary>
        /// Returns the CivilApplication object
        /// </summary>
        public static CivilApplication CivilApplication { get; set; }

        /// <summary>
        /// Returns true if the shared parameters have been created for this session.
        /// </summary>
        public static bool ParametersCreated { get; set; }

        /// <summary>
        /// Returns a Dynamo CoordinateSystem that represents the Revit Document Total Transform for the session.
        /// </summary>
        public static Autodesk.DesignScript.Geometry.CoordinateSystem DocumentTotalTransform {get;set;}

        /// <summary>
        /// Returns a Dynamo CoordinateSystem that represents the Revit Document Total Transform Inverse for the session.
        /// </summary>
        public static Autodesk.DesignScript.Geometry.CoordinateSystem DocumentTotalTransformInverse {get;set;}
        
        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Initializes a new instance of the <see cref="SessionVariables"/> class.
        /// </summary>
        internal SessionVariables() 
        {
            ParametersCreated = false;
            DocumentTotalTransform = null;
            DocumentTotalTransformInverse = null;
        }

        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS


        #endregion
    }
}
