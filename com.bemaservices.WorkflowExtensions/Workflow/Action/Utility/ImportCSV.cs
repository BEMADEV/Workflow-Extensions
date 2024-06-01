using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;

using CsvHelper;
using Newtonsoft.Json;

using Rock;
using Rock.Attribute;
using Rock.Data;
using Rock.Model;
using Rock.Workflow;

namespace com.bemaservices.WorkflowExtensions.Workflow.Action.Utility
{

    /// <summary>
    /// This workflow action imports a CSV to a workflow attribute, where the contents of the CSV are stored as JSON.
    /// </summary>
    [ActionCategory( "BEMA Services > Workflow Extensions" )]
    [Description( "Imports a CSV file to text workflow attribute, where the contents of the CSV are stored as JSON. The first row of your CSV should contain field headers. This should be used only for files of moderate size, as the whole file will be loaded into memory." )]
    [Export( typeof( ActionComponent ) )]
    [ExportMetadata( "ComponentName", "Import CSV" )]

    // The workflow attribute containing the input CSV file
    [WorkflowAttribute( "Input CSV File", "The CSV file to be parsed.", true, "", "", 0, CSV_ATTRIBUTE_KEY, new string[] { "Rock.Field.Types.FileFieldType" } )]
    [WorkflowAttribute( "Output JSON Attribute", "The attribute to store the parsed JSON output.", true, "", "", 1, JSON_ATTRIBUTE_KEY, new string[] { "Rock.Field.Types.TextFieldType" } )]

    public class ImportCSV : ActionComponent
    {

        private const string CSV_ATTRIBUTE_KEY = "CSVAttribute";
        private const string JSON_ATTRIBUTE_KEY = "JSONAttribute";

        public override bool Execute(RockContext rockContext, WorkflowAction action, object entity, out List<string> errorMessages)
        {
            
            errorMessages = new List<string>();

            var binaryFileService = new BinaryFileService( rockContext );

            var bfAttrGuid = GetAttributeValue(action, CSV_ATTRIBUTE_KEY).AsGuidOrNull();
            var binaryFileGuid = action.GetWorkflowAttributeValue(bfAttrGuid.Value).AsGuidOrNull();

            if (binaryFileGuid == null)
            {
                errorMessages.Add("The provided file attribute is empty or invalid.");
                return false;
            }

            var binaryFile = binaryFileService.Get(binaryFileGuid.Value);

            if (binaryFile == null)
            {
                errorMessages.Add("The specified binary file could not be found.");
                return false;
            }

            var contentStream = binaryFile.ContentStream;

            try 
            {

                using (var csv = new CsvReader(new StreamReader(contentStream)))
                {
                    var records = csv.GetRecords<dynamic>().ToList();
                    var json = JsonConvert.SerializeObject(records);

                    SetWorkflowAttributeValue(action, JSON_ATTRIBUTE_KEY, json);
                }

            }
            catch (Exception e)
            {
                errorMessages.Add("An error occurred while parsing the CSV file: " + e.Message);
                return false;
            }

            return true;

        }

    }
}
