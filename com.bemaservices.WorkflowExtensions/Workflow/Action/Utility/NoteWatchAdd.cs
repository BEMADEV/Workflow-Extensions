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
using System.Data.Entity;
using System.Linq;
using System.Web;
using Rock;
using Rock.Attribute;
using Rock.Data;
using Rock.Model;
using Rock.Web.Cache;
using Rock.Workflow;

namespace com.bemaservices.WorkflowExtensions.Workflow.Action.Utility
{
    /// <summary>
    /// Adds a Note to an Entity
    /// </summary>
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "Adds a note watch." )]
    [Export( typeof( Rock.Workflow.ActionComponent ) )]
    [ExportMetadata( "ComponentName", "Note Watch Add" )]

    [EntityTypeField(
        "Entity Type",
        IncludeGlobalAttributeOption = false,
        Description = "The type of Entity.",
        IsRequired = true,
        Order = 0,
        Key = AttributeKeys.EntityType )]
    [WorkflowTextOrAttribute(
        "Entity Id or Guid",
        "Entity Attribute",
        Description = "The id or guid of the entity. <span class='tip tip-lava'></span>.",
        IsRequired = true,
        Order = 1,
        Key = AttributeKeys.EntityIdGuid )]
    [NoteTypeField(
        "Note Type",
        "The type of note to add.",
        AllowMultiple = false,
        Order = 2 )]
    [WorkflowAttribute(
        "Watcher Person",
        "Workflow attribute that contains the person to use as the watcher of the note.",
        false,
        "",
        "",
        3,
        null,
        new string[] { "Rock.Field.Types.PersonFieldType" } )]
    [WorkflowAttribute(
        "Watcher Group",
        "Workflow attribute that contains the person to use as the watcher of the note.",
        false,
        "",
        "",
        4,
        null,
        new string[] { "Rock.Field.Types.GroupFieldType" } )]
    [BooleanField(
        "Is Watching",
        "Set this to false to block notifications.",
        false,
        "",
        5 )]
    [BooleanField(
        "Allow Override",
        "Set this to false to prevent other note watches from blocking this note watch.",
        false,
        "",
        6 )]
    public class NoteWatchAdd : Rock.Workflow.ActionComponent
    {

        #region Attribute Keys

        private static class AttributeKeys
        {
            public const string EntityType = "EntityType";
            public const string EntityIdGuid = "EntityIdGuid";
        }

        #endregion

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

            // Get the entity type
            EntityTypeCache entityType = null;
            var entityTypeGuid = GetAttributeValue( action, AttributeKeys.EntityType ).AsGuidOrNull();
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
            string entityIdGuidString = GetAttributeValue( action, AttributeKeys.EntityIdGuid, true ).ResolveMergeFields( mergeFields ).Trim();
            var entityId = entityIdGuidString.AsIntegerOrNull();
            if ( entityId.HasValue )
            {
                entityObject = entityTypeService.GetEntity( entityType.Id, entityId.Value );
            }
            else
            {
                var entityGuid = entityIdGuidString.AsGuidOrNull();
                if ( entityGuid.HasValue )
                {
                    entityObject = entityTypeService.GetEntity( entityType.Id, entityGuid.Value );
                }
            }

            if ( entityObject == null )
            {
                var value = GetActionAttributeValue( action, AttributeKeys.EntityIdGuid );
                entityObject = action.GetEntityFromAttributeValue( value, rockContext );
            }

            var noteType = NoteTypeCache.Get( GetAttributeValue( action, "NoteType" ).AsGuid() );

            // Create the Note Watch

            NoteWatch noteWatch;

            var noteWatchService = new NoteWatchService( rockContext );
            noteWatch = new NoteWatch();
            noteWatchService.Add( noteWatch );

            if ( noteType != null )
            {
                noteWatch.NoteTypeId = noteType.Id;
            }

            if ( entityType != null )
            {
                noteWatch.EntityTypeId = entityType.Id;
            }

            if ( noteWatch.EntityTypeId.HasValue )
            {
                if ( entityObject != null )
                {
                    noteWatch.EntityId = entityObject.Id;
                }
            }

            // Get Watcher
            var watcherPerson = GetPersonAliasFromActionAttribute( "WatcherPerson", rockContext, action, errorMessages );
            if ( watcherPerson != null )
            {
                noteWatch.WatcherPersonAliasId = watcherPerson.PrimaryAlias.Id;
            }

            var watcherGroup = GetGroupFromActionAttribute( "WatcherPerson", rockContext, action, errorMessages );
            if ( watcherGroup != null )
            {
                noteWatch.WatcherGroupId = watcherGroup.Id;
            }


            noteWatch.IsWatching = GetAttributeValue( action, "IsWatching" ).AsBoolean();
            noteWatch.AllowOverride = GetAttributeValue( action, "AllowOverride" ).AsBoolean();

            // see if the Watcher parameters are valid
            if ( !noteWatch.IsValidWatcher )
            {
                errorMessages.Add( "A Person or Group must be specified as the watcher" );
                return false;
            }

            // see if the Watch filters parameters are valid
            if ( !noteWatch.IsValidWatchFilter )
            {
                errorMessages.Add( "A Watch Filter must be specified" );
                return false;
            }

            if ( !noteWatch.IsValid )
            {
                return false;
            }

            // See if there is a matching filter that doesn't allow overrides
            if ( noteWatch.IsWatching == false )
            {
                if ( !noteWatch.IsAbleToUnWatch( rockContext ) )
                {
                    var nonOverridableNoteWatch = noteWatch.GetNonOverridableNoteWatches( rockContext ).FirstOrDefault();
                    if ( nonOverridableNoteWatch != null )
                    {
                        errorMessages.Add( "Unable to set Watching to false. This would override another note watch that doesn't allow overrides" );
                        return false;
                    }
                }
            }

            // see if the NoteType allows following
            if ( noteWatch.NoteTypeId.HasValue )
            {
                var noteTypeCache = NoteTypeCache.Get( noteWatch.NoteTypeId.Value );
                if ( noteTypeCache != null )
                {
                    if ( noteTypeCache.AllowsWatching == false )
                    {
                        errorMessages.Add( "This note type doesn't allow note watches." );
                        return false;
                    }
                }
            }

            rockContext.SaveChanges();

            errorMessages.ForEach( m => action.AddLogEntry( m, true ) );

            return true;
        }

        private Person GetPersonAliasFromActionAttribute( string key, RockContext rockContext, WorkflowAction action, List<string> errorMessages )
        {
            string value = GetAttributeValue( action, key );
            Guid guidPersonAttribute = value.AsGuid();
            if ( !guidPersonAttribute.IsEmpty() )
            {
                var attributePerson = AttributeCache.Get( guidPersonAttribute, rockContext );
                if ( attributePerson != null )
                {
                    string attributePersonValue = action.GetWorkflowAttributeValue( guidPersonAttribute );
                    if ( !string.IsNullOrWhiteSpace( attributePersonValue ) )
                    {
                        if ( attributePerson.FieldType.Class == "Rock.Field.Types.PersonFieldType" )
                        {
                            Guid personAliasGuid = attributePersonValue.AsGuid();
                            if ( !personAliasGuid.IsEmpty() )
                            {
                                PersonAliasService personAliasService = new PersonAliasService( rockContext );
                                return personAliasService.Queryable().AsNoTracking()
                                    .Where( a => a.Guid.Equals( personAliasGuid ) )
                                    .Select( a => a.Person )
                                    .FirstOrDefault();
                            }
                            else
                            {
                                errorMessages.Add( string.Format( "Person could not be found for selected value ('{0}')!", guidPersonAttribute.ToString() ) );
                                return null;
                            }
                        }
                        else
                        {
                            errorMessages.Add( string.Format( "The attribute used for {0} to provide the person was not of type 'Person'.", key ) );
                            return null;
                        }
                    }
                }
            }

            return null;
        }

        private Group GetGroupFromActionAttribute( string key, RockContext rockContext, WorkflowAction action, List<string> errorMessages )
        {
            string value = GetAttributeValue( action, key );
            Guid guidGroupAttribute = value.AsGuid();
            if ( !guidGroupAttribute.IsEmpty() )
            {
                var attributeGroup = AttributeCache.Get( guidGroupAttribute, rockContext );
                if ( attributeGroup != null )
                {
                    string attributeGroupValue = action.GetWorkflowAttributeValue( guidGroupAttribute );
                    if ( !string.IsNullOrWhiteSpace( attributeGroupValue ) )
                    {
                        Guid groupGuid = attributeGroupValue.AsGuid();
                        if ( !groupGuid.IsEmpty() )
                        {
                            GroupService groupService = new GroupService( rockContext );
                            return groupService.Queryable().AsNoTracking()
                                .Where( a => a.Guid.Equals( groupGuid ) )
                                .FirstOrDefault();
                        }
                        else
                        {
                            errorMessages.Add( string.Format( "Group could not be found for selected value ('{0}')!", guidGroupAttribute.ToString() ) );
                            return null;
                        }
                    }
                }
            }

            return null;
        }

    }
}
