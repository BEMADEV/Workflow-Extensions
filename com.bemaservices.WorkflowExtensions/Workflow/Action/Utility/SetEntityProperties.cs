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
using System.Reflection;
using Rock;
using Rock.Attribute;
using Rock.Data;
using Rock.Model;
using Rock.Web.Cache;
using Rock.Workflow;

namespace com.bemaservices.WorkflowExtensions.Workflow.Action
{
    /// <summary>
    /// Sets an entity property.
    /// </summary>
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "Sets entity properties." )]
    [Export( typeof( ActionComponent ) )]
    [ExportMetadata( "ComponentName", "Entity Properties Set" )]

    [EntityTypeField( "Entity Type", false, "The type of Entity.", true, "", 0, "EntityType" )]
    [WorkflowTextOrAttribute( "Entity Id or Guid", "Entity Attribute", "The id or guid of the entity. <span class='tip tip-lava'></span>", true, "", "", 1, "EntityIdGuid" )]
    [MatrixField( "91339C27-CA24-4038-A382-7C59A1DE5906", "Set Entity Properties", "", false, "", 2, "Matrix" )]
    [CustomDropdownListField( "Empty Value Handling", "How to handle empty property values.", "IGNORE^Ignore empty values,EMPTY^Set to empty,NULL^Set to NULL", true, "", "", 4, "EmptyValueHandling" )]
    [WorkflowAttribute( "Entity Guid", "The attribute that the entity's guid is saved to. If the Entity Id is and this attribute is set, a new entity will be created.",
         false, "", "", 5 )]
    public class SetEntityProperty : ActionComponent
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
            var entityGuidAttribute = AttributeCache.Get( GetAttributeValue( action, "EntityGuid" ).AsGuid(), rockContext );

            errorMessages = new List<string>();

            // Get the entity type
            EntityTypeCache entityType = null;
            var entityTypeGuid = GetAttributeValue( action, "EntityType" ).AsGuidOrNull();
            if ( entityTypeGuid.HasValue )
            {
                entityType = EntityTypeCache.Get( entityTypeGuid.Value );
            }
            if ( entityType == null )
            {
                errorMessages.Add( string.Format( "Entity Type could not be found for selected value ('{0}')!", entityTypeGuid.ToString() ) );
                return false;
            }

            var mergeFields = GetMergeFields( action );
            RockContext _rockContext = new RockContext();

            // Get the entity
            EntityTypeService entityTypeService = new EntityTypeService( _rockContext );
            IEntity entityObject = null;
            string entityIdGuidString = GetAttributeValue( action, "EntityIdGuid", true ).ResolveMergeFields( mergeFields ).Trim();
            var entityGuid = entityIdGuidString.AsGuidOrNull();
            var newEntity = false;
            if ( entityGuid.HasValue )
            {
                entityObject = entityTypeService.GetEntity( entityType.Id, entityGuid.Value );
            }
            else
            {
                var entityId = entityIdGuidString.AsIntegerOrNull();
                if ( entityId.HasValue )
                {
                    entityObject = entityTypeService.GetEntity( entityType.Id, entityId.Value );
                }
            }

            if ( entityObject == null && entityType.Guid == Rock.SystemGuid.EntityType.PERSON.AsGuid() && entityGuid.HasValue )
            {
                var personAlias = new PersonAliasService( _rockContext ).Get( entityGuid.Value );
                if ( personAlias != null )
                {
                    entityObject = personAlias.Person;
                }
            }

            if ( entityObject == null )
            {
                var value = GetActionAttributeValue( action, "EntityIdGuid" );
                entityObject = action.GetEntityFromAttributeValue( value, rockContext );
            }

            if ( entityObject == null )
            {
                if ( entityGuidAttribute == null )
                {
                    errorMessages.Add( string.Format( "Entity could not be found for selected value ('{0}')!", entityIdGuidString ) );
                    return false;
                }
                else
                {
                    newEntity = true;
                    entityObject = Activator.CreateInstance( entityType.GetEntityType() ) as IEntity;
                    var entityService = Reflection.GetServiceForEntityType( entityType.GetEntityType(), _rockContext );
                    entityService.GetType().GetMethod( "Add" ).Invoke( entityService, new object[] { entityObject } );
                }
            }

            // Get the property settings
            string emptyValueHandling = GetAttributeValue( action, "EmptyValueHandling" );
            var attributeMatrixGuid = GetAttributeValue( action, "Matrix" ).AsGuid();
            var attributeMatrix = new AttributeMatrixService( _rockContext ).Get( attributeMatrixGuid );
            if ( attributeMatrix != null )
            {
                foreach ( AttributeMatrixItem attributeMatrixItem in attributeMatrix.AttributeMatrixItems )
                {
                    attributeMatrixItem.LoadAttributes();


                    string propertyName = attributeMatrixItem.GetMatrixAttributeValue( action, "PropertyName", true ).ResolveMergeFields( mergeFields );
                    string propertyValue = attributeMatrixItem.GetMatrixAttributeValue( action, "Value", true ).ResolveMergeFields( mergeFields );

                    if ( emptyValueHandling == "IGNORE" && String.IsNullOrWhiteSpace( propertyValue ) )
                    {
                        action.AddLogEntry( "Skipping empty value." );
                        return true;
                    }

                    PropertyInfo propInf = entityObject.GetType().GetProperty( propertyName, BindingFlags.Public | BindingFlags.Instance );

                    if ( propInf == null )
                    {
                        errorMessages.Add( string.Format( "Property does not exist ('{0}')!", propertyName ) );
                        return false;
                    }

                    if ( !propInf.CanWrite )
                    {
                        errorMessages.Add( string.Format( "Property is not writable ('{0}')!", entityIdGuidString ) );
                        return false;
                    }

                    try
                    {
                        propInf.SetValue( entityObject, ConvertObject( propertyValue, propInf.PropertyType, _rockContext, emptyValueHandling == "NULL" ), null );
                    }
                    catch ( Exception ex ) when ( ex is InvalidCastException || ex is FormatException || ex is OverflowException )
                    {
                        errorMessages.Add( string.Format( "Could not convert property value ('{0}')! {1}", propertyValue, ex.Message ) );
                        return false;
                    }

                    if ( newEntity == false )
                    {

                        if ( !entityObject.IsValid )
                        {
                            errorMessages.Add( string.Format( "Invalid property value ('{0}')! {1}", propertyValue, entityObject.ValidationResults.Select( r => r.ErrorMessage ).ToList().AsDelimited( " " ) ) );
                            return false;
                        }

                        try
                        {
                            _rockContext.SaveChanges();
                        }
                        catch ( Exception ex )
                        {
                            errorMessages.Add( string.Format( "Could not save value ('{0}')! {1}", propertyValue, ex.Message ) );
                            return false;
                        }
                    }

                    action.AddLogEntry( string.Format( "Set '{0}' property to '{1}'.", propertyName, propertyValue ) );
                }
            }

            if ( newEntity == true )
            {

                if ( !entityObject.IsValid )
                {
                    errorMessages.Add( entityObject.ValidationResults.Select( r => r.ErrorMessage ).ToList().AsDelimited( " " ) );
                    return false;
                }

                try
                {
                    _rockContext.SaveChanges();
                }
                catch ( Exception ex )
                {
                    errorMessages.Add( string.Format( "Could not create new entity. {1}", ex.Message ) );
                    return false;
                }
            }

            if ( entityObject != null )
            {
                if ( entityGuidAttribute != null )
                {
                    SetWorkflowAttributeValue( action, entityGuidAttribute.Guid, entityObject.Guid.ToString() );
                    action.AddLogEntry( string.Format( "Set '{0}' attribute to '{1}'.", entityGuidAttribute.Name, entityObject.Guid ) );
                    return true;
                }
            }
            else
            {
                errorMessages.Add( "Entity could not be Set!" );
            }

            return true;
        }

        /// <summary>
        /// Converts a string to the specified type of object.
        /// </summary>
        /// <param name="theObject">The string to convert.</param>
        /// <param name="objectType">The type of object desired.</param>
        /// <param name="tryToNull">If empty strings should return as null.</param>
        /// <returns></returns>
        private static object ConvertObject( string theObject, Type objectType, RockContext rockContext, bool tryToNull = true )
        {
            if ( objectType.IsEnum )
            {
                return string.IsNullOrWhiteSpace( theObject ) ? null : Enum.Parse( objectType, theObject, true );
            }

            var theObjectGuid = theObject.AsGuidOrNull();
            if ( objectType.IsClass == true && theObjectGuid.HasValue )
            {
                // Get the service type corresponding to the object type
                Type serviceType = typeof( Rock.Data.Service<> );
                Type[] modelType = { objectType };
                Type service = serviceType.MakeGenericType( modelType );

                // Create new service of the above type
                var serviceInstance = Activator.CreateInstance( service, new object[] { rockContext } );

                // Find the get method for the service
                var getMethod = service.GetMethod( "Get", new Type[] { typeof( Guid ) } );

                // Call the get method to get the object
                object entity = getMethod.Invoke( serviceInstance, new object[] { theObjectGuid.Value } ) as object;

                // Return object
                return entity;
            }

            Type underType = Nullable.GetUnderlyingType( objectType );
            if ( underType == null ) // not nullable
            {
                return Convert.ChangeType( theObject, objectType );
            }

            if ( tryToNull && string.IsNullOrWhiteSpace( theObject ) )
            {
                return null;
            }
            return Convert.ChangeType( theObject, underType );
        }

    }
}
