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
    [ActionCategory("BEMA Services > Workflow Extensions")]
    [Description("Adds a note the selected entity.")]
    [Export(typeof(Rock.Workflow.ActionComponent))]
    [ExportMetadata("ComponentName", "Entity Note Add")]

    [EntityTypeField(
        "Entity Type",
        IncludeGlobalAttributeOption = false,
        Description = "The type of Entity.",
        IsRequired = true,
        Order = 0,
        Key = AttributeKeys.EntityType)]
    [WorkflowTextOrAttribute(
        "Entity Id or Guid",
        "Entity Attribute",
        Description = "The id or guid of the entity. <span class='tip tip-lava'></span>.",
        IsRequired = true,
        Order = 1,
        Key = AttributeKeys.EntityIdGuid)]
    [NoteTypeField(
        "Note Type",
        "The type of note to add.",
        AllowMultiple = false,
        Order = 2)]
    [TextField(
        "Caption",
        "The title/caption of the note. If none is provided then the author's name will be displayed. <span class='tip tip-lava'></span>",
        false,
        "",
        "",
        Order = 3)]
    [MemoField(
        "Text",
        "The body of the note. <span class='tip tip-lava'></span>",
        true,
        "",
        "",
        Order = 4)]
    [WorkflowAttribute(
        "Author",
        "Workflow attribute that contains the person to use as the author of the note. While not required it is recommended.",
        false,
        "",
        "",
        5,
        null,
        new string[] { "Rock.Field.Types.PersonFieldType" })]
    [BooleanField(
        "Alert",
        "Determines if the note should be flagged as an alert.",
        false,
        "",
        6)]
    public class EntityNoteAdd : Rock.Workflow.ActionComponent
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

        public override bool Execute(RockContext rockContext, WorkflowAction action, Object entity, out List<string> errorMessages)
        {
            errorMessages = new List<string>();

            // Get the entity type
            EntityTypeCache entityType = null;
            var entityTypeGuid = GetAttributeValue(action, AttributeKeys.EntityType).AsGuidOrNull();
            if (entityTypeGuid.HasValue)
            {
                entityType = EntityTypeCache.Get(entityTypeGuid.Value);
            }
            if (entityType == null)
            {
                errorMessages.Add(string.Format("Entity Type could not be found for selected value ('{0}')!", entityTypeGuid.ToString()));
                return false;
            }

            var mergeFields = GetMergeFields(action);
            RockContext _rockContext = new RockContext();

            // Get the entity
            EntityTypeService entityTypeService = new EntityTypeService(_rockContext);
            IEntity entityObject = null;
            string entityIdGuidString = GetAttributeValue(action, AttributeKeys.EntityIdGuid, true).ResolveMergeFields(mergeFields).Trim();
            var entityId = entityIdGuidString.AsIntegerOrNull();
            if (entityId.HasValue)
            {
                entityObject = entityTypeService.GetEntity(entityType.Id, entityId.Value);
            }
            else
            {
                var entityGuid = entityIdGuidString.AsGuidOrNull();
                if (entityGuid.HasValue)
                {
                    entityObject = entityTypeService.GetEntity(entityType.Id, entityGuid.Value);
                }
            }

            if (entityObject == null)
            {
                var value = GetActionAttributeValue(action, AttributeKeys.EntityIdGuid);
                entityObject = action.GetEntityFromAttributeValue(value, rockContext);
            }

            if (entityObject == null)
            {
                errorMessages.Add(string.Format("Entity could not be found for selected value ('{0}')!", entityIdGuidString));
                return false;
            }


            // Create the Note

            NoteService noteService = new NoteService(rockContext);

            Note note = new Note
            {
                EntityId = entityObject.Id,
                Caption = GetAttributeValue(action, "Caption").ResolveMergeFields(mergeFields),
                IsAlert = GetAttributeValue(action, "Alert").AsBoolean(),
                IsPrivateNote = false,
                Text = GetAttributeValue(action, "Text").ResolveMergeFields(mergeFields)
            };

            var noteType = NoteTypeCache.Get(GetAttributeValue(action, "NoteType").AsGuid());
            if (noteType != null)
            {
                note.NoteTypeId = noteType.Id;
            }

            // Get Author
            var author = GetPersonAliasFromActionAttribute("Author", rockContext, action, errorMessages);
            if (author != null)
            {
                note.CreatedByPersonAliasId = author.PrimaryAlias.Id;
            }

            // Add the Note
            noteService.Add(note);
            rockContext.SaveChanges();

            errorMessages.ForEach(m => action.AddLogEntry(m, true));

            return true;
        }

        private Person GetPersonAliasFromActionAttribute(string key, RockContext rockContext, WorkflowAction action, List<string> errorMessages)
        {
            string value = GetAttributeValue(action, key);
            Guid guidPersonAttribute = value.AsGuid();
            if (!guidPersonAttribute.IsEmpty())
            {
                var attributePerson = AttributeCache.Get(guidPersonAttribute, rockContext);
                if (attributePerson != null)
                {
                    string attributePersonValue = action.GetWorkflowAttributeValue(guidPersonAttribute);
                    if (!string.IsNullOrWhiteSpace(attributePersonValue))
                    {
                        if (attributePerson.FieldType.Class == "Rock.Field.Types.PersonFieldType")
                        {
                            Guid personAliasGuid = attributePersonValue.AsGuid();
                            if (!personAliasGuid.IsEmpty())
                            {
                                PersonAliasService personAliasService = new PersonAliasService(rockContext);
                                return personAliasService.Queryable().AsNoTracking()
                                    .Where(a => a.Guid.Equals(personAliasGuid))
                                    .Select(a => a.Person)
                                    .FirstOrDefault();
                            }
                            else
                            {
                                errorMessages.Add(string.Format("Person could not be found for selected value ('{0}')!", guidPersonAttribute.ToString()));
                                return null;
                            }
                        }
                        else
                        {
                            errorMessages.Add(string.Format("The attribute used for {0} to provide the person was not of type 'Person'.", key));
                            return null;
                        }
                    }
                }
            }

            return null;
        }

    }
}
