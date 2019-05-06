// Copyright (c) 2019 Autodesk, Inc. All rights reserved.
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
using Autodesk.DesignScript.Runtime;
using RevitServices.Persistence;
using System;




namespace CivilConnection
{
    /// <summary>
    /// Static Obejct that returns the ProjectPosition of the Revti Document
    /// </summary>
    [IsVisibleInDynamoLibrary(false)]
    class ProjectPositionUtils
    {
        #region PRIVATE MEMBERS

        private static Autodesk.Revit.DB.ProjectPosition _position;
        private static double _angle;
        private static ProjectPositionUtils _instance;
        private static Autodesk.Revit.DB.ProjectLocation _location;

        #endregion

        #region PUBLIC MEMBERS
        /// <summary>
        /// The ProjectPosition
        /// </summary>
        public Autodesk.Revit.DB.ProjectPosition ProjectPosition 
        {
            get 
            {
                return _position;
            }

        }
        /// <summary>
        /// The angle of the Project Position
        /// </summary>
        public double Angle
        {
            get
            {
                return _angle;
            }

        }
        /// <summary>
        /// The ProjectLocation
        /// </summary>
        public Autodesk.Revit.DB.ProjectLocation ProjectLocation
        {
            get 
            {
                return _location;
            }
        }

        /// <summary>
        /// ProjectPositionUtils Instance
        /// </summary>
        public static ProjectPositionUtils Instance
        {
            get 
            {
                if (_instance == null)
                {
                    _instance = new ProjectPositionUtils();
                }

                return _instance;
            }
        }

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Constructor.
        /// </summary>
        protected ProjectPositionUtils() 
        {
            _location = DocumentManager.Instance.CurrentDBDocument.ActiveProjectLocation;

            if (Convert.ToInt32(DocumentManager.Instance.CurrentDBDocument.Application.VersionNumber) <= 2018)
            {
                //try
                //{
                //    _position = _location.get_ProjectPosition(Autodesk.Revit.DB.XYZ.Zero);
                //    _angle = _position.Angle;
                //}
                //catch (Exception ex)
                //{
                //    Utils.Log(string.Format("ERROR {0}: {1}", this, ex.Message));
                //}
            }
            else
            {
                _position = _location.GetProjectPosition(Autodesk.Revit.DB.XYZ.Zero);
                _angle = _position.Angle;
            }
        }

        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// String representation
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return base.ToString();
        }

        /// <summary>
        /// Set the new ProjectLocation of the Revit document origin.
        /// </summary>
        /// <param name="location">The ProjectLocation</param>
        /// <param name="newPosition">The new ProjectPosition of the document origin.</param>
        public void SetProjectPosition(Autodesk.Revit.DB.ProjectLocation location, Autodesk.Revit.DB.ProjectPosition newPosition)
        {
            Utils.Log(string.Format("ProjectPositionUtils.SetProjectPosition started...", ""));

            // location.set_ProjectPosition(Autodesk.Revit.DB.XYZ.Zero, newPosition);  // deprecated
            location.SetProjectPosition(Autodesk.Revit.DB.XYZ.Zero, newPosition);

            Utils.Log(string.Format("ProjectPositionUtils.SetProjectPosition completed.", ""));
        }

        #endregion

        #region PRIVATE METHODS


        #endregion

    }
}
