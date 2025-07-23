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
using System.Linq;
using System.Reflection;
using Rock;
using Rock.Attribute;
using Rock.Data;
using Rock.Model;
using Rock.Web.Cache;
using Rock.Workflow;


namespace com.bemaservices.WorkflowExtensions.Workflow.Action
{
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "Adds an item to a specified Attribute Matrix with the specified values." )]
    [Export( typeof( ActionComponent ) )]
    [ExportMetadata( "ComponentName", "Attribute Matrix Item Add" )]

    [WorkflowAttribute( "Matrix", "The Matrix to add the item to", true, "", "", 0, "TargetMatrix", new string[] { "Rock.Field.Types.MatrixFieldType" } )]
    [MatrixField( "BF2EE993-AC09-4EDA-95D8-7E3757CB8A46", "Attribute Item Values", "", false, "", 1, "ItemMatrix" )]
    public class AddItemToMatrix : ActionComponent
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
            var attributeMatrixItemService = new AttributeMatrixItemService( rockContext );
            var attributeMatrixService = new AttributeMatrixService( rockContext );
            AttributeMatrix targetMatrix = null;

            // Get Target Matrix
            var targetMatrixAttributeGuid = GetAttributeValue( action, "TargetMatrix" ).AsGuidOrNull();
            if ( targetMatrixAttributeGuid.HasValue )
            {
                action.AddLogEntry( "Found Guid for the matrix attribute" );
                var targetMatrixAttribute = AttributeCache.Get( targetMatrixAttributeGuid.Value );
                if ( targetMatrixAttribute != null )
                {
                    action.AddLogEntry( "Found Attribute containing matrix" );

                    var targetMatrixGuid = action.GetWorkflowAttributeValue( targetMatrixAttributeGuid.Value ).AsGuidOrNull();
                    if ( !targetMatrixGuid.HasValue )
                    {
                        action.AddLogEntry( "No matrix found. Adding new one" );

                        var templateQualifier = targetMatrixAttribute.QualifierValues.Where( aq => aq.Key == "attributematrixtemplate" ).FirstOrDefault();
                        if ( targetMatrixAttribute.QualifierValues.ContainsKey( "attributematrixtemplate" )
                            && templateQualifier.Value != null
                            && templateQualifier.Value.Value != null
                            && templateQualifier.Value.Value.ToString().AsIntegerOrNull() != null )
                        {
                            // Create the AttributeMatrix now and save it even though they haven't hit save yet. We'll need the AttributeMatrix record to exist so that we can add AttributeMatrixItems to it
                            // If this ends up creating an orphan, we can clean up it up later
                            targetMatrix = new AttributeMatrix { Guid = Guid.NewGuid() };
                            targetMatrix.AttributeMatrixTemplateId = templateQualifier.Value.Value.ToString().AsIntegerOrNull().Value;
                            targetMatrix.AttributeMatrixItems = new List<AttributeMatrixItem>();
                            attributeMatrixService.Add( targetMatrix );
                            SetWorkflowAttributeValue( action, targetMatrixAttribute.Guid, targetMatrix.Guid.ToString() );
                            rockContext.SaveChanges();
                            action.AddLogEntry( String.Format( "Added Matrix with Guid {0}", targetMatrix.Guid ) );
                        }
                        else
                        {
                            errorMessages.Add( "Matrix specified in attribute does not exist, and no default settings exist to create one. Please select a default AttributeMatrixTemplate in the workflow settings." );
                            return false;
                        }
                    }
                    else
                    {
                        targetMatrix = attributeMatrixService.Get( targetMatrixGuid.Value );
                        targetMatrix.AttributeMatrixTemplate = new AttributeMatrixTemplateService( rockContext ).Get( targetMatrix.AttributeMatrixTemplateId );
                        action.AddLogEntry( String.Format( "Found matrix with Id {0}", targetMatrix.Id ) );
                    }

                    if ( targetMatrix != null )
                    {
                        action.AddLogEntry( "Building new Matrix Item" );
                        var newMatrixItem = new AttributeMatrixItem();
                        newMatrixItem.AttributeMatrix = targetMatrix;
                        newMatrixItem.AttributeMatrixId = targetMatrix.Id;

                        // if we check if the property exists and then set it, we can support both v16 and v17
                        PropertyInfo attributeMatrixTemplateIdProperty = newMatrixItem.GetType().GetProperty( "AttributeMatrixTemplateId" );
                        if ( attributeMatrixTemplateIdProperty != null && 
                            attributeMatrixTemplateIdProperty.PropertyType == typeof( int ) && 
                            attributeMatrixTemplateIdProperty.CanWrite )
                        {
                            attributeMatrixTemplateIdProperty.SetValue( newMatrixItem, targetMatrix.AttributeMatrixTemplateId );
                        }

                        newMatrixItem.LoadAttributes();
                        action.AddLogEntry( string.Format( "New Matrix Item has the following {0} columns available: {1}",
                            newMatrixItem.Attributes.Count,
                            newMatrixItem.Attributes.Select(a=> a.Key ).JoinStringsWithRepeatAndFinalDelimiterWithMaxLength(", ",", and ", null) ));

                        // Get Matrix with new Matrix Item
                        var attributeMatrixGuid = GetAttributeValue( action, "ItemMatrix" ).AsGuid();
                        var attributeMatrix = attributeMatrixService.Get( attributeMatrixGuid );
                        if ( attributeMatrix != null )
                        {
                            action.AddLogEntry( string.Format("Found Mapping Matrix ID {0} with {1} items", attributeMatrix.Id, attributeMatrix.AttributeMatrixItems.Count) );
                            foreach ( AttributeMatrixItem attributeMatrixItem in attributeMatrix.AttributeMatrixItems )
                            {
                                action.AddLogEntry( "Loading Mapping Item" );
                                attributeMatrixItem.LoadAttributes();

                                string columnKey = attributeMatrixItem.GetMatrixAttributeValue( action, "ColumnKey", true ).ResolveMergeFields( GetMergeFields( action ) );
                                action.AddLogEntry( String.Format("Found Mapping for |{0}|", columnKey ));

                                if ( newMatrixItem.Attributes.ContainsKey( columnKey ) )
                                {
                                    action.AddLogEntry( String.Format( "Found matching column in target matrix for {0}", columnKey ) );

                                    string columnValue = attributeMatrixItem.GetMatrixAttributeValue( action, "ColumnValue", true ).ResolveMergeFields( GetMergeFields( action ) );
                                    newMatrixItem.SetAttributeValue( columnKey, columnValue );
                                    action.AddLogEntry( String.Format( "Set Attribute with Key {0} to {1}", columnKey, columnValue ) );
                                }
                            }
                        }

                        if ( newMatrixItem.AttributeValues.Where( av => av.Value.Value.IsNotNullOrWhiteSpace() ).Any() )
                        {
                            action.AddLogEntry( "Matrix has at least one non whitespace column" );
                            attributeMatrixItemService.Add( newMatrixItem );
                            rockContext.SaveChanges();
                            newMatrixItem.SaveAttributeValues();
                            action.AddLogEntry( String.Format( "New Matrix Item added with Id {0}", newMatrixItem.Id ) );
                        }
                    }
                }
            }

            return true;
        }
    }
}