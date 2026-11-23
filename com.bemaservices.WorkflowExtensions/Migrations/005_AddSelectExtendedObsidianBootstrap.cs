using Rock.Plugin;

namespace com.bemaservices.WorkflowExtensions.Migrations
{
    [MigrationNumber( 5, "1.0.12" )]
    public class AddSelectExtendedObsidianBootstrap : Migration
    {
        private const string BootstrapStartMarker = "<!-- BEMA Workflow Extensions: Select Extended Obsidian START -->";

        private const string Bootstrap = @"<!-- BEMA Workflow Extensions: Select Extended Obsidian START -->
<script>
    Obsidian.onReady(() => {
        System.import('/Plugins/com_bemaservices/WorkflowExtensions/fieldTypes.js')
            .catch(error => {
                console.error('Unable to load WorkflowExtensions field types.', error);
            });
    });
</script>
<!-- BEMA Workflow Extensions: Select Extended Obsidian END -->";

        public override void Up()
        {
            var bootstrap = Bootstrap.Replace( "'", "''" );

            Sql( $@"
DECLARE @Bootstrap NVARCHAR(MAX) = N'{bootstrap}';

UPDATE
    [Site]
SET
    [PageHeaderContent] = CASE
        WHEN NULLIF([PageHeaderContent], N'') IS NULL THEN @Bootstrap
        ELSE [PageHeaderContent] + CHAR(13) + CHAR(10) + @Bootstrap
    END
WHERE
    CHARINDEX(N'{BootstrapStartMarker}', ISNULL([PageHeaderContent], N'')) = 0
;" );
        }

        public override void Down()
        {
            var bootstrap = Bootstrap.Replace( "'", "''" );

            Sql( $@"
DECLARE @Bootstrap NVARCHAR(MAX) = N'{bootstrap}';

UPDATE
    [Site]
SET
    [PageHeaderContent] = REPLACE(
        REPLACE([PageHeaderContent], CHAR(13) + CHAR(10) + @Bootstrap, N''),
        @Bootstrap,
        N'')
WHERE
    CHARINDEX(N'{BootstrapStartMarker}', ISNULL([PageHeaderContent], N'')) > 0
;" );
        }
    }
}