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
using Rock.Data;
using Rock.Model;
using Rock.Web.Cache;
using Rock.Workflow;

namespace com.bemaservices.WorkflowExtensions.Workflow.Action
{
    /// <summary>
    /// Adds person to a group using a workflow attribute.
    /// </summary>
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "Adds person to a group using a workflow attribute." )]
    [Export( typeof( ActionComponent ) )]
    [ExportMetadata( "ComponentName", "Group Member Add From Attribute" )]

    [WorkflowAttribute( "Person",
        Description = "Workflow attribute that contains the person to add to the group.",
        Key = AttributeKey.PersonKey,
        IsRequired = true,
        FieldTypeClassNames = new string[] { "Rock.Field.Types.PersonFieldType" },
        Order = 0 )]

    [WorkflowAttribute( "Group",
        Description = "Workflow Attribute that contains the group to add the person to.",
        Key = AttributeKey.GroupKey,
        IsRequired = true,
        FieldTypeClassNames = new string[] { "Rock.Field.Types.GroupFieldType" },
        Order = 1 )]

    [WorkflowAttribute( "Group Role Attribute",
        Description = "The group role the person will have.",
        IsRequired = true,
        Order = 2,
        Key = AttributeKey.GROUP_ROLE_ATTRIBUTE_KEY,
        FieldTypeClassNames = new string[] { "Rock.Field.Types.GroupRoleFieldType" } )]

    [BooleanField( "Disable Security Groups",
        Description = "When set to 'Yes', if the group given is a security type group the action will not be marked as a success and an error will be logged.",
        Key = AttributeKey.DisableSecurityGroups,
        DefaultBooleanValue = false,
        IsRequired = false,
        Order = 3 )]

    [GroupTypeField( "Limit To Groups Of Type",
        Description = "When set, if the group given does not match this group type the action will not be marked as a success and an error will be logged.",
        Key = AttributeKey.LimitToGroupsOfType,
        IsRequired = false,
        Order = 4 )]

    [GroupField( "Limit To Groups Under Specific Parent Group",
        Description = "When set, if the group given is not found under this parent group the action will not be marked as a success and an error will be logged.",
        Key = AttributeKey.LimitToGroupsUnderSpecificParentGroup,
        IsRequired = false,
        Order = 5 )]

    [WorkflowAttribute( "Group Member", 
        Description = "An optional GroupMember attribute to store the group member that is added.", 
        Key = AttributeKey.GroupMember,
        IsRequired = false,
        FieldTypeClassNames = new string[] { "Rock.Field.Types.GroupMemberFieldType" },
        Order =  6
         )]

    public class AddPersonToGroupWFAttribute : ActionComponent
    {
        private class AttributeKey
        {
            public const string PersonKey = "Person";
            public const string GroupKey = "Group";
            public const string GROUP_ROLE_ATTRIBUTE_KEY = "GroupRoleAttribute";
            public const string DisableSecurityGroups = "DisableSecurityGroups";
            public const string LimitToGroupsOfType = "LimitToGroupsOfType";
            public const string LimitToGroupsUnderSpecificParentGroup = "LimitToGroupsUnderSpecificParentGroup";
            public const string GroupMember = "GroupMember";
        }

        /// <summary>
        /// Executes the specified workflow.
        /// </summary>
        /// <param name="rockContext">The rock context.</param>
        /// <param name="action">The action.</param>
        /// <param name="entity">The entity.</param>
        /// <param name="errorMessages">The error messages.</param>
        /// <returns></returns>
        public override bool Execute( RockContext rockContext, WorkflowAction action, object entity, out List<string> errorMessages )
        {
            errorMessages = new List<string>();

            // Determine which group to add the person to
            Group group = null;
            int? groupRoleId = null;

            var guidGroupAttribute = GetAttributeValue( action, AttributeKey.GroupKey ).AsGuidOrNull();

            if ( guidGroupAttribute.HasValue )
            {
                var attributeGroup = AttributeCache.Get( guidGroupAttribute.Value, rockContext );
                if ( attributeGroup != null )
                {
                    var groupGuid = action.GetWorkflowAttributeValue( guidGroupAttribute.Value ).AsGuidOrNull();

                    if ( groupGuid.HasValue )
                    {
                        group = new GroupService( rockContext ).Get( groupGuid.Value );

                        if ( group != null )
                        {
                            var groupRoleGuid = GetAttributeValue( action, AttributeKey.GROUP_ROLE_ATTRIBUTE_KEY, true ).AsGuid();

                            var groupRole = group.GroupType.Roles.Where( gr => gr.Guid == groupRoleGuid ).FirstOrDefault();
                            if ( groupRole == null )
                            {
                                // use the group's grouptype's default group role if a group role wasn't specified
                                groupRole = group.GroupType.DefaultGroupRole;
                            }

                            if ( groupRole == null )
                            {
                                // use the group's grouptype's first found group role if a group role wasn't specified
                                groupRole = group.GroupType.Roles.FirstOrDefault();                                
                            }

                            groupRoleId = groupRole.Id;

                            if ( groupRoleId == null )
                            {
                                errorMessages.Add( "Invalid or no Group Role provided." );
                            }
                        }
                    }
                }
            }

            if ( group == null )
            {
                errorMessages.Add( "No group was provided" );
            }
            else
            {
                // Check if this is a security group and show an error if that functionality has been disabled.
                var disableSecurityGroups = GetAttributeValue( action, AttributeKey.DisableSecurityGroups ).AsBooleanOrNull() ?? false;
                if ( group.IsSecurityRoleOrSecurityGroupType() && disableSecurityGroups )
                {
                    errorMessages.Add( $"\"{group.Name}\" is a Security group. The settings for this workflow action do not allow it to add a person to a security group." );
                }

                // If LimitToGroupsOfType has any values check if this Group's GroupType is in that list
                var limitToGroupsOfTypeGuid = GetAttributeValue( action, AttributeKey.LimitToGroupsOfType ).AsGuidOrNull();
                if ( limitToGroupsOfTypeGuid.HasValue )
                {
                    var limitToGroupType = GroupTypeCache.Get( limitToGroupsOfTypeGuid.Value );
                    var limitToGroupTypeIds = new GroupTypeService( rockContext ).GetChildGroupTypes( limitToGroupType.Id ).Select( a => a.Id ).ToList();
                    limitToGroupTypeIds.Add( limitToGroupType.Id );

                    if ( !limitToGroupTypeIds.Contains( group.GroupTypeId ) )
                    {
                        errorMessages.Add( $"The group type for group \"{group.Name} is \"{group.GroupType.Name}\". This action is configured to only add persons to groups of type \"{limitToGroupType.Name}\" and its child types." );
                    }
                }

                // If LimitToGroupsUnderSpecificParentGroup has any values check if this Group is a child of that group
                var limitToChildGroupsOfGroupGuid = GetAttributeValue( action, AttributeKey.LimitToGroupsUnderSpecificParentGroup ).AsGuidOrNull();
                if ( limitToChildGroupsOfGroupGuid.HasValue )
                {
                    var groupService = new GroupService( rockContext );
                    var limitToChildGroupsOfGroup = groupService.Get( limitToChildGroupsOfGroupGuid.Value );
                    var limitToGroupIds = groupService.GetAllDescendentGroupIds( limitToChildGroupsOfGroup.Id, true );
                    limitToGroupIds.Add( limitToChildGroupsOfGroup.Id );

                    if ( !limitToGroupIds.Contains( group.Id ) )
                    {
                        errorMessages.Add( $"Cannot add the person to group \"{group.Name}\". This workflow action is configured to only add persons to groups that are a descendant of Group {limitToChildGroupsOfGroup.Name}." );
                    }
                }
            }

            if ( !groupRoleId.HasValue )
            {
                errorMessages.Add( "Provided group doesn't have a default group role" );
            }

            // determine the person that will be added to the group
            Person person = null;

            // get the Attribute.Guid for this workflow's Person Attribute so that we can lookup the value
            var guidPersonAttribute = GetAttributeValue( action, AttributeKey.PersonKey ).AsGuidOrNull();

            if ( guidPersonAttribute.HasValue )
            {
                var attributePerson = AttributeCache.Get( guidPersonAttribute.Value, rockContext );
                if ( attributePerson != null )
                {
                    string attributePersonValue = action.GetWorkflowAttributeValue( guidPersonAttribute.Value );
                    if ( !string.IsNullOrWhiteSpace( attributePersonValue ) )
                    {
                        if ( attributePerson.FieldType.Class == typeof( Rock.Field.Types.PersonFieldType ).FullName )
                        {
                            Guid personAliasGuid = attributePersonValue.AsGuid();
                            if ( !personAliasGuid.IsEmpty() )
                            {
                                person = new PersonAliasService( rockContext ).Queryable()
                                    .Where( a => a.Guid.Equals( personAliasGuid ) )
                                    .Select( a => a.Person )
                                    .FirstOrDefault();
                            }
                        }
                        else
                        {
                            errorMessages.Add( "The attribute used to provide the person was not of type 'Person'." );
                        }
                    }
                }
            }

            if ( person == null )
            {
                errorMessages.Add( string.Format( "Person could not be found for selected value ('{0}')!", guidPersonAttribute.ToString() ) );
            }

            // Add Person to Group
            if ( !errorMessages.Any() )
            {
                var groupMemberService = new GroupMemberService( rockContext );
                var groupMember = groupMemberService.GetByGroupIdAndPersonIdAndPreferredGroupRoleId( group.Id, person.Id, groupRoleId.Value );
                bool isNew = false;
                if ( groupMember == null )
                {
                    groupMember = new GroupMember();
                    groupMember.PersonId = person.Id;
                    groupMember.GroupId = group.Id;
                    isNew = true;
                }
                else
                {
                    action.AddLogEntry( $"{person.FullName} was already a member of the selected group.", true );
                }

                groupMember.GroupRoleId = groupRoleId.Value;
                groupMember.GroupMemberStatus = GroupMemberStatus.Active;

                if ( groupMember.IsValidGroupMember( rockContext ) )
                {
                    if (isNew)
                    {
                        groupMemberService.Add( groupMember );
                    }

                    rockContext.SaveChanges();
                }
                else
                {
                    // if the group member couldn't be added (for example, one of the group membership rules didn't pass), add the validation messages to the errormessages
                    errorMessages.AddRange( groupMember.ValidationResults.Select( a => a.ErrorMessage ) );
                }

                // If group member attribute was specified, requery the request and set the attribute's value
                Guid? groupMemberAttributeGuid = GetAttributeValue( action, AttributeKey.GroupMember ).AsGuidOrNull();
                if ( groupMemberAttributeGuid.HasValue )
                {
                    groupMember = groupMemberService.Get( groupMember.Id );
                    if ( groupMember != null )
                    {
                        SetWorkflowAttributeValue( action, groupMemberAttributeGuid.Value, groupMember.Guid.ToString() );
                    }
                }
            }

            errorMessages.ForEach( m => action.AddLogEntry( m, true ) );

            return true;
        }
    }
}