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
using Rock.Plugin;

namespace com.bemaservices.WorkflowExtensions.Migrations
{
    [MigrationNumber( 4, "1.18.0" )]
    public class RenameScheduleBuilder : Migration
    {
        public override void Up()
        {
            RockMigrationHelper.UpdateFieldType(
                "BEMA Schedule Builder",
                "",
                "com.bemaservices.WorkflowExtensions",
                "com.bemaservices.WorkflowExtensions.Field.Types.ScheduleBuilderFieldType",
                "9AAF1F39-E485-4CCB-92CF-5F5BA0CD9822" );
        }

        public override void Down()
        {
            RockMigrationHelper.UpdateFieldType(
                "Schedule Builder",
                "",
                "com.bemaservices.WorkflowExtensions",
                "com.bemaservices.WorkflowExtensions.Field.Types.ScheduleBuilderFieldType",
                "9AAF1F39-E485-4CCB-92CF-5F5BA0CD9822" );
        }
    }
}