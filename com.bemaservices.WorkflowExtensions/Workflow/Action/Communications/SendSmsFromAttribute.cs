﻿// <copyright>
// Copyright by BEMA Information Technologies
//
// Licensed under the Rock Community License (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.rockrms.com/license
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.Linq;
using Rock;
using Rock.Attribute;
using Rock.Communication;
using Rock.Data;
using Rock.Model;
using Rock.Web.Cache;
using Rock.Workflow;

namespace com.bemaservices.WorkflowExtensions.Workflow.Action
{
    /// <summary>
    /// Sends email
    /// </summary>
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "THIS ACTION WILL BE REMOVED IN V18." )]
    [Export( typeof( ActionComponent ) )]
    [ExportMetadata( "ComponentName", "SMS Send From Attribute" )]

   public class SendSms : Rock.Workflow.Action.SendSms
    {
    }
}