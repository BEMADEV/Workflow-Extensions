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
    public class UpdateFieldTypes : Migration
    {
        public override void Up()
        {
            RenameScheduleBuilder();
            ReplaceReCaptchaFieldType();
        }

        public override void Down()
        {
            RestoreScheduleBuilderName();

            // The CAPTCHA consolidation cannot be safely reversed without changing CAPTCHA references that originated with Rock core.
        }

        private void RenameScheduleBuilder()
        {
            RockMigrationHelper.UpdateFieldType(
                "BEMA Schedule Builder",
                "",
                "com.bemaservices.WorkflowExtensions",
                "com.bemaservices.WorkflowExtensions.Field.Types.ScheduleBuilderFieldType",
                "9AAF1F39-E485-4CCB-92CF-5F5BA0CD9822" );
        }

        private void RestoreScheduleBuilderName()
        {
            RockMigrationHelper.UpdateFieldType(
                "Schedule Builder",
                "",
                "com.bemaservices.WorkflowExtensions",
                "com.bemaservices.WorkflowExtensions.Field.Types.ScheduleBuilderFieldType",
                "9AAF1F39-E485-4CCB-92CF-5F5BA0CD9822" );
        }

        private void ReplaceReCaptchaFieldType()
        {
            Sql( @"
        DECLARE @BemaFieldTypeId INT = (Select Top 1 Id From FieldType Where [Class] = 'com.bemaservices.WorkflowExtensions.Field.Types.ReCaptchaFieldType');
        DECLARE @CoreFieldTypeId INT = (Select Top 1 Id From FieldType Where [Guid] = '22F43337-7177-4064-9D4B-841EAD671678');

        If @BemaFieldTypeId Is Not Null And @CoreFieldTypeId Is Not Null
        Begin
        Update EntityType
        Set SingleValueFieldTypeId = @CoreFieldTypeId
        Where SingleValueFieldTypeId = @BemaFieldTypeId

        Update EntityType
        Set MultiValueFieldTypeId = @CoreFieldTypeId
        Where MultiValueFieldTypeId = @BemaFieldTypeId

        Update Attribute
        Set FieldTypeId = @CoreFieldTypeId
        Where FieldTypeId = @BemaFieldTypeId
        End
" );
        }
    }
}