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
using Autodesk.AECC.Interop.Land;

using Autodesk.DesignScript.Runtime;



namespace CivilConnection
{
    /// <summary>
    /// SubassemblyParameter obejct type.
    /// </summary>
    public class SubassemblyParameter
    {
        #region PRIVATE PROPERTIES
        /// <summary>
        /// The parameter
        /// </summary>
        internal IAeccParam _parameter;
        /// <summary>
        /// The value
        /// </summary>
        internal object _value;
        /// <summary>
        /// The type
        /// </summary>
        internal SubassemblyParameterType _type;
        //internal SubassemblyParameter[] _parameters;

        /// <summary>
        /// Gets the internal element.
        /// </summary>
        /// <value>
        /// The internal element.
        /// </value>
        internal object InternalElement { get { return this._parameter; } }

        #endregion

        #region PUBLIC PROPERTIES
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get { return this._parameter.DisplayName; } }
        /// <summary>
        /// Gets the Comment.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Comment { get { return this._parameter.Comment; } }
        /// <summary>
        /// Gets the Description.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Description { get { return this._parameter.Description; } }
        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public object Value { get { return this._value; } }
        /// <summary>
        /// Gets the type.
        /// </summary>
        /// <value>
        /// The type.
        /// </value>
        public SubassemblyParameterType Type { get { return this._type; } }
        //public Parameters[] Parameters {get {return this._parameters}}

        #endregion

        #region CONSTRUCTOR
        /// <summary>
        /// Initializes a new instance of the <see cref="SubassemblyParameter"/> class.
        /// </summary>
        /// <param name="parameter">The parameter.</param>
        internal SubassemblyParameter(IAeccParam parameter)
        {
            this._parameter = parameter;

            if (parameter is AeccParamBool)
            {
                var p = parameter as AeccParamBool;
                this._type = SubassemblyParameterType.Bool;
                this._value = p.Value;
            }
            else if (parameter is AeccParamDouble)
            {
                var p = parameter as AeccParamDouble;
                this._type = SubassemblyParameterType.Double;
                this._value = p.Value;
            }
            else if (parameter is AeccParamLong)
            {
                this._type = SubassemblyParameterType.Long;
                var p = parameter as AeccParamLong;
                this._value = p.Value;
            }
            else if (parameter is AeccParamString)
            {
                this._type = SubassemblyParameterType.String;
                var p = parameter as AeccParamString;
                this._value = p.Value;
            }
        }


        #endregion

        #region PRIVATE METHODS


        #endregion

        #region PUBLIC METHODS

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.Format("SubassemblyParameter(Name={0}, Value={1}, Type={2}, Comment={3}, Description={4})", this._parameter.DisplayName, this.Value, this._type, this._parameter.Comment, this._parameter.Description);
        }   
        #endregion
    }

    /// <summary>
    /// SubassemblyParameterType enumerator.
    /// </summary>
    [IsVisibleInDynamoLibrary(false)]
    public enum SubassemblyParameterType
    {
        /// <summary>
        /// Boolean Type
        /// </summary>
        Bool,
        /// <summary>
        /// Double Type
        /// </summary>
        Double,
        /// <summary>
        /// Long Type
        /// </summary>
        Long,
        /// <summary>
        /// Stirng Type
        /// </summary>
        String
    }

}
