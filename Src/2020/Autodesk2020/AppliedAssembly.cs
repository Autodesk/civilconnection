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
using Autodesk.AECC.Interop.Roadway;
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using System.Collections.Generic;

namespace CivilConnection
{
    /// <summary>
    /// AppliedAssembly object type.
    /// </summary>
    [IsVisibleInDynamoLibrary(false)]
    public class AppliedAssembly
    {
        #region PRIVATE PROPERTIES

        internal AeccAppliedAssembly _assembly;
        internal AeccAppliedSubassemblies _appliedSubassemblies;
        internal AeccCorridor _corridor;
        internal BaselineRegion _region;
        
        internal object InternalElement { get { return this._assembly; } }

        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Internal Constructor.
        /// </summary>
        /// <param name="blr">The BaselineRegion</param>
        /// <param name="appliedAssembly">the AeccAppliedAssembly form Civil 3D.</param>
        /// <param name="corridor">the AeccCorridor from Civil 3D.</param>
        internal AppliedAssembly(BaselineRegion blr, AeccAppliedAssembly appliedAssembly, AeccCorridor corridor)
        {
            this._region = blr;
            this._assembly = appliedAssembly;
            this._appliedSubassemblies = appliedAssembly.AppliedSubassemblies;
            this._corridor = corridor;
        }

        #endregion

        #region PUBLIC METHODS
        /// <summary>
        /// Public textual representation in the Dynamo node preview.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("AppliedAssembly");
        }
        #endregion
    }


     /// <summary>
     /// Base class for applied subassemblies geometry objects.
     /// </summary>
    [IsVisibleInDynamoLibrary(false)]
    public abstract class AbstractAppliedSubassemblyGeometryObject
    {
        #region PRIVATE PROPERTIES
        /// <exclude/>
        protected string _name;
        /// <exclude/>
        protected Geometry _geometry;
        /// <exclude/>
        protected IList<string> _codes = new List<string>();
        /// <exclude/>
        protected double _station;
        #endregion

        #region PUBLIC PROPERTIES

        #endregion

        #region CONSTRUCTOR
        internal AbstractAppliedSubassemblyGeometryObject(string name, Geometry geometry, IList<string> codes, double station)
        {
            this._name = name;
            this._geometry = geometry;
            this._codes = codes;
            this._station = station;
        }
        #endregion

        #region PUBLIC METHODS

        #endregion
    }

    //[IsVisibleInDynamoLibrary(false)]
     ///<summary>
     ///The Applied Subassembly link object
     ///</summary>
    public class AppliedSubassemblyLink : AbstractAppliedSubassemblyGeometryObject
    {
        // It's necessary to add these public properties to have them shown in the Dynamo Library
        #region PUBLIC PROPERTIES 
         /// <summary>
        /// Returns the unique name of the object.
        /// </summary>
        public string Name { get { return this._name; } }

        /// <summary>
        /// Returns the Dynamo geometry associated to the object.
        /// </summary>
        public Geometry Geometry { get { return this._geometry; } }

        /// <summary>
        /// Returns the list of codes associated to the object.
        /// </summary>
        public IList<string> Codes { get { return this._codes; } }

        /// <summary>
        /// Returns the station associated to the object.
        /// </summary>
        public double Station { get { return this._station; } }

         #endregion

        #region CONSTRUCTOR
        internal AppliedSubassemblyLink(string name, Geometry geometry, IList<string> codes, double station)
            : base(name, geometry, codes, station)
        {
        }
        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Public textual representation of the Dynamo node preview
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("AppliedSubassemblyLink(Name={0})", this.Name);
        }
        #endregion
    }

    //[IsVisibleInDynamoLibrary(false)]
     ///<summary>
     ///The Applied Subassembly shape object
     ///</summary>
    public class AppliedSubassemblyShape : AbstractAppliedSubassemblyGeometryObject
    {
        // It's necessary to add these public properties to have them shown in the Dynamo Library
        #region PUBLIC PROPERTIES
        /// <summary>
        /// Returns the unique name of the object.
        /// </summary>
        public string Name { get { return this._name; } }

        /// <summary>
        /// Returns the Dynamo geometry associated to the object.
        /// </summary>
        public Geometry Geometry { get { return this._geometry; } }

        /// <summary>
        /// Returns the list of codes associated to the object.
        /// </summary>
        public IList<string> Codes { get { return this._codes; } }

        /// <summary>
        /// Returns the station associated to the object.
        /// </summary>
        public double Station { get { return this._station; } }

         #endregion

        #region CONSTRUCTOR
        internal AppliedSubassemblyShape(string name, Geometry geometry, IList<string> codes, double station)
            : base(name, geometry, codes, station)
        {
        }
        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Public textual representation of the Dynamo node preview
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("AppliedSubassemblyShape(Name={0})", this.Name);
        }
        #endregion
    }
}
