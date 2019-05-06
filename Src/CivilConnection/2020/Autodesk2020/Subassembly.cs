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
using Autodesk.AECC.Interop.Land;
using Autodesk.AECC.Interop.Roadway;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CivilConnection
{
    /// <summary>
    /// Subassembly object type.
    /// </summary>
    public class Subassembly
    {
        #region PRIVATE PROPERTIES

        /// <summary>
        /// The subassembly
        /// </summary>
        internal AeccSubassembly _subassembly;
        /// <summary>
        /// The corridor
        /// </summary>
        internal AeccCorridor _corridor;
        /// <summary>
        /// The parameters
        /// </summary>
        internal IList<SubassemblyParameter> _parameters = new List<SubassemblyParameter>();

        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this._subassembly; } }

        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return this._subassembly.Name; } }
        /// <summary>
        /// Gets the parameters.
        /// </summary>
        /// <value>
        /// The parameters.
        /// </value>
        public IList<SubassemblyParameter> Parameters { get { return this._parameters; } }

        #endregion

        #region CONSTRUCTOR

        /// <summary>
        /// Initializes a new instance of the <see cref="Subassembly"/> class.
        /// </summary>
        /// <param name="subassembly">The subassembly.</param>
        /// <param name="corridor">The corridor.</param>
        internal Subassembly(AeccSubassembly subassembly, AeccCorridor corridor)
        {
            this._subassembly = subassembly;
            this._corridor = corridor;

            foreach (var p in subassembly.ParamsBool)
            {
                this._parameters.Add(new SubassemblyParameter((IAeccParam)p));
            }
            foreach (var p in subassembly.ParamsDouble)
            {
                this._parameters.Add(new SubassemblyParameter((IAeccParam)p));
            }
            foreach (var p in subassembly.ParamsLong)
            {
                this._parameters.Add(new SubassemblyParameter((IAeccParam)p));
            }
            foreach (var p in subassembly.ParamsString)
            {
                this._parameters.Add(new SubassemblyParameter((IAeccParam)p));
            }
        }

        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Sets SubassemblyParameter value by name.
        /// </summary>
        /// <param name="name">The parameter name.</param>
        /// <param name="value">The value.</param>
        /// <param name="rebuild">if set to <c>true</c> [rebuild].</param>
        /// <returns></returns>
        /// <exception cref="Exception">
        /// The parameter name cannot be null
        /// or
        /// The value cannot be null
        /// or
        /// or
        /// </exception>
        public Subassembly SetParameterByName(string name, object value, bool rebuild = false)
        {
            Utils.Log(string.Format("Subassembly.SetParameterByName started...", ""));

            if (null == name)
            {
                throw new Exception("The parameter name cannot be null");
            }

            if (null == value)
            {
                throw new Exception("The value cannot be null");
            }

            SubassemblyParameter parameter = null;

            try
            {
                parameter = this.Parameters.First(x => x.Name == name);
            }
            catch (Exception ex)
            {
                var message = string.Format("No parameter {0} found on this Subassembly", name);

                Utils.Log(string.Format("ERROR: Subassembly.SetParameterByName {0} {1}", message, ex.Message));

                throw new Exception(message);
            }

            if (null != parameter)
            {
                if(parameter.Type == SubassemblyParameterType.Bool && value is bool)
                {
                    var p = parameter.InternalElement as AeccParamBool;
                    p.Value = Convert.ToBoolean(value);
                }

                 else if(parameter.Type == SubassemblyParameterType.Double && (value is double || value is int))
                {
                    var p = parameter.InternalElement as AeccParamDouble;
                    p.Value = Convert.ToDouble(value);
                }

                else if (parameter.Type == SubassemblyParameterType.Long && (value is long || value is double || value is int))
                {
                    var p = parameter.InternalElement as AeccParamLong;
                    p.Value = Convert.ToInt32(value);
                }
                else if (parameter.Type == SubassemblyParameterType.String && value is string)
                {
                    var p = parameter.InternalElement as AeccParamString;
                    p.Value = Convert.ToString(value);
                }
                else 
                {
                    var message = string.Format("The value does not match the parameter data type {0}", parameter.Type);

                    Utils.Log(string.Format("ERROR: Subassembly.SetParameterByName {0}", message));

                    throw new Exception(message);
                }
            }

            if (rebuild)
            {
                AeccSubassembly sa = this.InternalElement as AeccSubassembly;
                this._corridor.Rebuild();
            }

            Utils.Log(string.Format("Subassembly.SetParameterByName completed.", ""));

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
            return string.Format("Subassembly(Name={0}, Corridor={1})", this.Name, this._corridor.Name);
        }
        #endregion
    }
}
