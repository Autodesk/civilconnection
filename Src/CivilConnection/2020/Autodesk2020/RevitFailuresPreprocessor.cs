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

using Autodesk.Revit.DB;
using Autodesk.DesignScript.Runtime;

namespace CivilConnection
{
    /// <summary>
    /// Revit Failure Preprocessor.
    /// </summary>
    /// <seealso cref="Autodesk.Revit.DB.IFailuresPreprocessor" />
    [IsVisibleInDynamoLibrary(false)]
    public class RevitFailuresPreprocessor : IFailuresPreprocessor
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RevitFailuresPreprocessor"/> class.
        /// </summary>
        public RevitFailuresPreprocessor()
        { }

        /// <summary>
        /// Preprocesses the failures.
        /// </summary>
        /// <param name="fa">The failure accessor.</param>
        /// <returns></returns>
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            fa.DeleteAllWarnings();
           
            return FailureProcessingResult.Continue;
        }
    }
}
