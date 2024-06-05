using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;

using CsvHelper;
using CsvHelper.Configuration;
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
    [CustomDropdownListField( 
        "Output Format",
        Description = "How to format the records in the JSON output. Dictionary format REQUIRES the CSV to have a header row.",
        ListSource = OUTPUT_FORMAT_DICTIONARY + "^Dictionary (key-value pairs)," + OUTPUT_FORMAT_ARRAY + "^Array",
        IsRequired = true, 
        Order = 3,
        Key = OUTPUT_FORMAT_KEY )]
    [WorkflowTextOrAttribute("CSV Has Header Row", "CSV Has Header Row Attribute", "Does the CSV file have a header row? Specify a boolean attribute or a string value that converts to a boolean.", true, "True", "", 3, HAS_HEADER_ROW_KEY, new string[] { "Rock.Field.Types.BooleanFieldType" } )]

    public class ImportCSV : ActionComponent
    {

        private const string CSV_ATTRIBUTE_KEY = "CSVAttribute";
        private const string JSON_ATTRIBUTE_KEY = "JSONAttribute";
        private const string HAS_HEADER_ROW_KEY = "HasHeaderRow";
        private const string OUTPUT_FORMAT_KEY = "OutputFormat";

        private const string OUTPUT_FORMAT_DICTIONARY = "DICTIONARY";
        private const string OUTPUT_FORMAT_ARRAY = "ARRAY";

        public override bool Execute(RockContext rockContext, WorkflowAction action, object entity, out List<string> errorMessages)
        {
            
            errorMessages = new List<string>();

            // Get output format value
            var outputFormatValue = GetAttributeValue(action, OUTPUT_FORMAT_KEY);
            var outputFormat = string.Empty;

            if ( outputFormatValue == OUTPUT_FORMAT_DICTIONARY )
            {
                outputFormat = OUTPUT_FORMAT_DICTIONARY;
            } else if ( outputFormatValue == OUTPUT_FORMAT_ARRAY )
            {
                outputFormat = OUTPUT_FORMAT_ARRAY;
            }

            // Get and validate the Has Header Row value
            var hasHeaderRow = GetAttributeValue(action, HAS_HEADER_ROW_KEY, true).AsBooleanOrNull();

            if (hasHeaderRow == null)
            {
                errorMessages.Add("The Has Header Row value is invalid. Please provide a valid Boolean value.");
                return false;
            }

            if (!(bool)hasHeaderRow && outputFormat == OUTPUT_FORMAT_DICTIONARY)
            {
                errorMessages.Add("The output format is set to Dictionary, but the CSV does not have a header row. Cannot continue.");
                return false;
            }

            // Get and validate the binary file and content stream
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

            // Parse and output
            try 
            {
                var json = string.Empty;

                // We have a header row and the output format is set to dictionary
                if ((bool)hasHeaderRow && outputFormat == OUTPUT_FORMAT_DICTIONARY)
                {
                    json = ParseToDictionary(contentStream);
                }

                // Output format is set to array
                else if (outputFormat == OUTPUT_FORMAT_ARRAY)
                {

                    json = ParseToJsonArray(contentStream, (bool)hasHeaderRow);

                }

                // Shouldn't be possible to get here, but bail if we do.
                else
                {
                    errorMessages.Add("Unable to parse CSV. Specified parsing options are invalid.");
                    return false;
                }

                SetWorkflowAttributeValue(action, JSON_ATTRIBUTE_KEY, json);

            }
            catch (Exception e)
            {
                errorMessages.Add("An error occurred while parsing the CSV file: " + e.Message);
                return false;
            }

            return true;

        }

        /// <summary>
        /// Parse to a dictionary.
        /// </summary>
        /// <param name="contentStream">The stream object representing the file content.</param>
        /// <returns></returns>
        private string ParseToDictionary(Stream contentStream)
        {

            using (var streamReader = new StreamReader(contentStream))
            using (var csv = new CsvReader(streamReader))
            {
                var records = csv.GetRecords<dynamic>().ToList();
                return JsonConvert.SerializeObject(records);
            }

        }

        /// <summary>
        /// Parse to a JSON array.
        /// </summary>
        /// <param name="contentStream">The stream object representing the file content.</param>
        /// <param name="hasHeader">Indicates whether or not the file has a header. Header values are ignored when outputting an array.</param>
        /// <returns>JSON representation of the input file.</returns>
        private string ParseToJsonArray(Stream contentStream, bool hasHeader)
        {

            // Disable CsvHelper's header handling in all cases if we're outputting an array
            var config = new CsvConfiguration()
            {
                HasHeaderRecord = false,
            };

            using (var streamReader = new StreamReader(contentStream))
            using (var csv = new CsvReader(streamReader, config))
            {
                var records = csv.GetRecords<dynamic>();

                if ( hasHeader )
                {
                    records.First(); // Skip the header row if present
                }

                var output = new List<List<string>>();

                var fieldCount = -1;
                var fieldList = new List<string>();

                foreach (var row in records)
                {
                    var fieldDict = ((IDictionary<string, object>)row);
                            
                    if (fieldCount == -1)
                    {
                        fieldCount = fieldDict.Count;
                        for (int i = 1; i <= fieldCount; i++)
                        {
                            fieldList.Add("Field" + i);
                        }
                    }

                    var outRow = new List<string>();

                    foreach(var field in fieldList)
                    {
                        outRow.Add(fieldDict.GetPropertyValue(field).ToString());
                    }

                    output.Add(outRow);
                }

                return JsonConvert.SerializeObject(output);

            }
        }

    }

}
