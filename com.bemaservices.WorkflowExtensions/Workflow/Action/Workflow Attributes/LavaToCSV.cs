// <copyright>
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
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web;
using Rock;
using Rock.Attribute;
using Rock.Data;
using Rock.Model;
using Rock.Web.Cache;
using Rock.Web.UI.Controls;
using Rock.Workflow;


namespace com.bemaservices.WorkflowExtensions.Workflow.Action
{
    /// <summary>
    /// Sets an entity attribute.
    /// </summary>
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "Takes a Lava template and renders it as an CSV. The resulting CSV is placed into a provided workflow attribute of type file." )]
    [Export( typeof( ActionComponent ) )]
    [ExportMetadata( "ComponentName", "Lava to CSV" )]

    [WorkflowAttribute( "Result Attribute", "The Attribute to put the resulting file into.", true, "", "", 0, "ResultAttribute", new string[] { "Rock.Field.Types.BinaryFileFieldType" } )]
    [TextField( "File Name", "The file name used for the file <span class='tip tip-lava'></span>", true, "document.csv", "", 1, "FileName" )]
    [CodeEditorField( "Lava Template", "The Lava templated used to create the CSV. <span class='tip tip-lava'></span>", CodeEditorMode.Lava, CodeEditorTheme.Rock, 300, true, "Attribute Key,Attribute Value\n{{ Workflow | Attribute:'AttributeKey' }},{{ Workflow | Attribute:'AttributeValue' }}", "", 2, "LavaTemplate" )]
    [LavaCommandsField( "Enabled Lava Commands", "The Lava commands that should be enabled for this action.", false, order: 3 )]

    public class LavaToCSV : ActionComponent
    {

        /// <summary>
        /// Executes the specified workflow.
        /// </summary>
        /// <param name="rockContext">The rock context.</param>
        /// <param name="action">The action.</param>
        /// <param name="entity">The entity.</param>
        /// <param name="errorMessages">The error messages.</param>
        /// <returns></returns>
        public override bool Execute( RockContext rockContext, WorkflowAction action, Object entity, out List<string> errorMessages )
        {

            errorMessages = new List<string>();

            var attribute = AttributeCache.Get( GetAttributeValue( action, "ResultAttribute" ).AsGuid(), rockContext );
            if ( attribute != null )
            {

                var binaryFileTypeGuid = attribute.QualifierValues.Where( x => x.Key == "binaryFileType" ).FirstOrDefault().Value.Value.AsGuidOrNull();

                if ( binaryFileTypeGuid.HasValue )
                {

                    BinaryFileType binaryFileType = new BinaryFileTypeService( rockContext ).Get( binaryFileTypeGuid.Value );

                    char[] illegalCharacters = new char[] { '<', '>', ':', '"', '/', '\\', '|', '?', '*' };
                    string fileName = GetAttributeValue( action, "FileName" ).ResolveMergeFields( GetMergeFields( action ), GetAttributeValue( action, "EnabledLavaCommands" ) ).Trim();
                    if ( fileName.IndexOfAny( illegalCharacters ) >= 0 || fileName.EndsWith( "." ) )
                    {
                        errorMessages.Add( "Invalid Filename.  Please remove any special characters (" + string.Join( " ", illegalCharacters ) + ")." );
                        return false;
                    }

                    string value = GetAttributeValue( action, "LavaTemplate" ).ResolveMergeFields( GetMergeFields( action ), GetAttributeValue( action, "EnabledLavaCommands" ) ).Trim();

                    // concert string to stream
                    byte[] byteArrary = Encoding.ASCII.GetBytes( value );
                    MemoryStream contentStream = new MemoryStream( byteArrary );

                    // always create a new BinaryFile record of IsTemporary when a file is uploaded
                    var binaryFileService = new BinaryFileService( rockContext );
                    var binaryFile = new BinaryFile();
                    binaryFileService.Add( binaryFile );

                    binaryFile.IsTemporary = false;
                    binaryFile.BinaryFileTypeId = binaryFileType.Id;
                    binaryFile.MimeType = "text/csv";
                    binaryFile.FileSize = contentStream.Length;
                    binaryFile.ContentStream = contentStream;
                    binaryFile.FileName = fileName;

                    rockContext.SaveChanges();

                    SetWorkflowAttributeValue( action, attribute.Guid, binaryFile.Guid.ToString() );

                    action.AddLogEntry( string.Format( "Set '{0}' attribute to '{1}'.", attribute.Name, binaryFile.Guid.ToString() ) );

                }
                else
                {
                    errorMessages.Add( "Binary file type must be specified on the Binary File Attribte" );
                    return false;
                }

            }

            return true;
        }
    }
}
