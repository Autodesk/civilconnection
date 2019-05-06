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

using Autodesk.DesignScript.Runtime;
using Autodesk.Revit.DB;

namespace CivilConnection
{
    /// <summary>
    /// FamilyLoadOptions
    /// </summary>
    /// <seealso cref="Autodesk.Revit.DB.IFamilyLoadOptions" />
    [IsVisibleInDynamoLibrary(false)]
    public class RevitFamilyLoadOptions : IFamilyLoadOptions
    {
#pragma warning disable CS0169 // The field 'RevitFamilyLoadOptions.overwriteParameters' is never used
        /// <summary>
        /// The overwrite parameters
        /// </summary>
        bool overwriteParameters;
#pragma warning restore CS0169 // The field 'RevitFamilyLoadOptions.overwriteParameters' is never used
        
#pragma warning disable CS0169 // The field 'RevitFamilyLoadOptions.source' is never used
        /// <summary>
        /// The source
        /// </summary>
        FamilySource source;
#pragma warning restore CS0169 // The field 'RevitFamilyLoadOptions.source' is never used

        /// <summary>
        /// Initializes a new instance of the <see cref="RevitFamilyLoadOptions"/> class.
        /// </summary>
        public RevitFamilyLoadOptions()
        { }

        /// <summary>
        /// Called when [family found].
        /// </summary>
        /// <param name="familyInUse">if set to <c>true</c> [family in use].</param>
        /// <param name="overwriteParameters">if set to <c>true</c> [overwrite parameters].</param>
        /// <returns></returns>
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameters)
        {
            overwriteParameters = true;

            return true;
        }

        /// <summary>
        /// Called when [shared family found].
        /// </summary>
        /// <param name="sharedFamily">The shared family.</param>
        /// <param name="familyInUse">if set to <c>true</c> [family in use].</param>
        /// <param name="source">The source.</param>
        /// <param name="overwriteParameters">if set to <c>true</c> [overwrite parameters].</param>
        /// <returns></returns>
        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameters)
        {
            source = FamilySource.Family;
            overwriteParameters = true;
            return true;
        }
    }
}
