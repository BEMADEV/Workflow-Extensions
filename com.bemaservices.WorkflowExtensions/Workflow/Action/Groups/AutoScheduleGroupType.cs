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

namespace com.bemaservices.WorkflowExtensions.Workflow.Action
{
    /// <summary>
    /// Adds person to a group using a workflow attribute.
    /// </summary>
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "Adds person to a group using a workflow attribute." )]
    [Export( typeof( ActionComponent ) )]
    [ExportMetadata( "ComponentName", "Auto Schedule Group Type" )]

    [WorkflowAttribute( "Group Type",
        Description = "When set, if the group given does not match this group type the action will not be marked as a success and an error will be logged.",
        Key = AttributeKey.GroupType,
        IsRequired = false,
        FieldTypeClassNames = new string[] { "Rock.Field.Types.GroupTypeFieldType" },
        Order = 0 )]

    [IntegerField( "Number of Weeks",
        Description = "How many weeks out should be auto scheduled?",
        Key = AttributeKey.WeeksOut,
        IsRequired = true,
        DefaultIntegerValue = 7,
        Order = 1
        )]

    [WorkflowAttribute( "Auto Scheduler",
        Description = "Workflow attribute that contains the person doing the auto scheduling.",
        Key = AttributeKey.AutoScheduler,
        IsRequired = true,
        FieldTypeClassNames = new string[] { "Rock.Field.Types.PersonFieldType" },
        Order = 2 )]

    [WorkflowTextOrAttribute(
        "Auto-Schedule Attribute Key",
        "Auto-Schedule Attribute Key Attribute",
        Description = "The key of an optional boolean group attribute to check before auto-scheduling. <span class='tip tip-lava'></span>",
        IsRequired = true,
        Order = 2,
        Key = AttributeKey.AutoScheduleAttributeKey )]

    public class AutoScheduleGroupType : ActionComponent
    {
        private class AttributeKey
        {
            public const string GroupType = "GroupType";
            public const string WeeksOut = "WeeksOut";
            public const string AutoScheduler = "AutoScheduler";
            public const string AutoScheduleAttributeKey = "AutoScheduleAttributeKey";
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

            GroupType groupType = null;
            var guidGroupTypeAttribute = GetAttributeValue( action, AttributeKey.GroupType ).AsGuidOrNull();

            if ( guidGroupTypeAttribute.HasValue )
            {
                var attributeGroupType = AttributeCache.Get( guidGroupTypeAttribute.Value, rockContext );
                if ( attributeGroupType != null )
                {
                    var groupTypeGuid = action.GetWorkflowAttributeValue( guidGroupTypeAttribute.Value ).AsGuidOrNull();

                    if ( groupTypeGuid.HasValue )
                    {
                        groupType = new GroupTypeService( rockContext ).Get( groupTypeGuid.Value );
                    }
                }
            }

            if ( groupType == null )
            {
                errorMessages.Add( "No group type was provided" );
            }

            // determine the person that will serve as the auto scheduler
            Person autoSchedulerPerson = null;

            // get the Attribute.Guid for this workflow's Person Attribute so that we can lookup the value
            var guidPersonAttribute = GetAttributeValue( action, AttributeKey.AutoScheduler ).AsGuidOrNull();

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
                                autoSchedulerPerson = new PersonAliasService( rockContext ).Queryable()
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

            if ( autoSchedulerPerson == null )
            {
                errorMessages.Add( string.Format( "Person could not be found for selected value ('{0}')!", guidPersonAttribute.ToString() ) );
            }

            var mergeFields = GetMergeFields( action );
            int weeksOut = GetAttributeValue( action, AttributeKey.WeeksOut ).AsInteger();
            string attributeKey = GetAttributeValue( action, AttributeKey.AutoScheduleAttributeKey, true ).ResolveMergeFields( mergeFields );

            if ( !errorMessages.Any() )
            {
                var attendanceService = new AttendanceService( rockContext );
                var groupService = new GroupService( rockContext );
                int occurrenceScheduledCount = 0;
                var loopCount = 0;

                // Find next Sunday Dates
                List<DateTime> sundayDates = new List<DateTime>();

                var sundayDate = RockDateTime.Now.SundayDate();
                int weekNum = 0;
                while ( weekNum < weeksOut )
                {
                    sundayDates.Add( sundayDate );
                    weekNum++;
                    sundayDate = sundayDate.AddDays( 7 );
                }

                //Find Groups
                var groupList = groupService.Queryable().AsNoTracking()
                    .Include( a => a.GroupType )
                    .Where( a =>
                        a.IsActive
                    && !a.IsArchived
                        && ( a.GroupTypeId == groupType.Id )
                        && a.ParentGroupId.HasValue
                        && a.GroupType.IsSchedulingEnabled
                        && !a.DisableScheduling )
                    .ToList();

                List<int> filteredGroupIds = new List<int>();

                if ( attributeKey.IsNotNullOrWhiteSpace() )
                {
                    foreach ( Group group in groupList )
                    {
                        group.LoadAttributes();
                        if ( group.GetAttributeValue( attributeKey ).AsBoolean() )
                        {
                            filteredGroupIds.Add( group.Id );
                        }
                    }
                }
                else
                {
                    filteredGroupIds = groupList.Select( g => g.Id ).ToList();
                }

                var groupLocationQry = new GroupLocationService( rockContext ).Queryable().Where( a => filteredGroupIds.Contains( a.GroupId ) );
                var scheduleList = groupLocationQry.SelectMany( a => a.Schedules ).Where( s => s.IsActive ).Distinct().AsNoTracking().ToList();

                var attendanceOccurrenceIdList = new List<int>();

                foreach ( var date in sundayDates )
                {
                    var startDate = date.Date.AddDays( -6 );
                    var endDate = date.Date;

                    var scheduleOccurrenceDateList = scheduleList
                    .Select( s => new
                    {
                        OccurrenceDate = s.GetNextStartDateTime( startDate ).HasValue ? ( DateTime? ) s.GetNextStartDateTime( startDate ).Value.Date : ( DateTime? ) null,
                        ScheduleId = s.Id
                    } )
                    .Where( a => a.OccurrenceDate.HasValue )
                    .OrderBy( a => a.OccurrenceDate )
                    .ToList();

                    var groupLocationOrdered = groupLocationQry.OrderBy( a => a.Order ).ThenBy( a => a.Location.Name ).Select( s => new { s.LocationId, s.GroupId } ).ToList();

                    foreach ( var scheduleOccurrenceDate in scheduleOccurrenceDateList )
                    {
                        foreach ( var glOrdered in groupLocationOrdered )
                        {
                            var attendanceOccurrence = new AttendanceOccurrenceService( rockContext ).GetOrAdd( scheduleOccurrenceDate.OccurrenceDate.Value, glOrdered.GroupId, glOrdered.LocationId, scheduleOccurrenceDate.ScheduleId );
                            attendanceOccurrenceIdList.Add( attendanceOccurrence.Id );
                        }
                    }
                }

                try
                {
                    //@#@# BD, BEMA 7/15/2024: Chunk occurrences into batches of 10K to avoid SQL compilation error with large number of occurrences
                    var index = 0;
                    var chunkSize = 10000;

                    while ( index < attendanceOccurrenceIdList.Count )
                    {
                        var chunk = attendanceOccurrenceIdList.Skip( index ).Take( chunkSize ).ToList();
                        index += chunkSize;
                        loopCount++;

                        //@#@# BD, BEMA 2/1/2023: Add scheduler alias id so the confirmation block will work
                        attendanceService.SchedulePersonsAutomaticallyForAttendanceOccurrences( chunk, autoSchedulerPerson.PrimaryAlias );

                        occurrenceScheduledCount += chunk.Count;

                        rockContext.SaveChanges();
                    }

                }
                catch ( Exception ex )
                {
                    errorMessages.Add( ex.Message );
                }

                try
                {
                    // Mark all occurrences with schedules as confirmed
                    foreach ( int occurrenceId in attendanceOccurrenceIdList )
                    {
                        List<int> attendanceIds = attendanceService.Queryable().AsNoTracking()
                            //BD, BEMA 11/16/2022: Add RequestedToAttend check, only auto-confirm if true
                            .Where( a => a.OccurrenceId == occurrenceId && a.DidAttend != true && a.RequestedToAttend == true ) 
                            .Where( a => a.RSVP == RSVP.Maybe || a.RSVP == RSVP.Unknown )
                            .Select( a => a.Id )
                            .ToList();
                        foreach ( int attendanceId in attendanceIds )
                        {
                            attendanceService.ScheduledPersonConfirm( attendanceId );
                        }
                    }
                    rockContext.SaveChanges();
                }
                catch ( Exception ex )
                {
                    errorMessages.Add( ex.Message );
                }

                action.AddLogEntry(
                    String.Format(
                             "{0} occurrences identified and {1} occurrences scheduled in {2} scheduling loops."
                            , attendanceOccurrenceIdList.Count.ToString()
                            , occurrenceScheduledCount.ToString()
                            , loopCount.ToString()
                            )
                    , true );
            }

            errorMessages.ForEach( m => action.AddLogEntry( m, true ) );
            return true;
        }
    }
}