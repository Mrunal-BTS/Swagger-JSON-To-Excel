using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Swagger_json_to_excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonFilePath = $"C:\\Users\\{Environment.UserName}\\input.json"; // Replace with your JSON file path

            try
            {
                string jsonContent = File.ReadAllText(jsonFilePath);
                var swagger = JsonConvert.DeserializeObject<JObject>(jsonContent);

                var paths = swagger?["paths"].ToObject<Dictionary<string, JObject>>();
                var components = swagger?["components"] as JObject;
                var schemas = components?["schemas"]?.ToObject<Dictionary<string, JObject>>();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Ensure EPPlus is set to non-commercial license mode

                using (var excelPackage = new ExcelPackage())
                {
                    var worksheet = excelPackage.Workbook.Worksheets.Add("API Endpoints");

                    int row = 1;
                    worksheet.Cells[row, 1].Value = "Api Endpoint";
                    worksheet.Cells[row, 2].Value = "HTTP Verb";
                    worksheet.Cells[row, 3].Value = "Request Parameters in Query String";
                    worksheet.Cells[row, 4].Value = "Request JSON Body Schema";
                    worksheet.Cells[row, 5].Value = "Response JSON Body Schema";
                    worksheet.Cells[row, 6].Value = "Response Description";

                    // Apply formatting to header row
                    using (var headerCells = worksheet.Cells[row, 1, row, 6])
                    {
                        headerCells.Style.Font.Bold = true;
                        headerCells.Style.Font.Size = 12;
                        headerCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        headerCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        headerCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        headerCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        headerCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        headerCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        headerCells.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Center vertically
                        headerCells.Style.WrapText = true; // Enable word wrapping
                    }

                    foreach (var pathEntry in paths)
                    {
                        var path = pathEntry.Key;
                        var verbs = pathEntry.Value as JObject;

                        foreach (var verbEntry in verbs)
                        {
                            var verb = verbEntry.Key;
                            var details = verbEntry.Value as JObject;

                            var requestParameters = "-";
                            var requestJsonBodySchema = GetRequestJsonBodySchema(details["requestBody"]?["content"] as JObject, schemas) ?? "-";
                            var responseJsonBodySchema = GetResponseSchemaReference(details["responses"] as JObject, schemas) ?? "-";
                            var responseDescription = details["responses"]?["200"]?["description"]?.ToString() ?? "-";

                            // Extracting request parameters
                            var parameters = details["parameters"] as JArray;
                            if (parameters != null && parameters.Count > 0)
                            {
                                requestParameters = string.Join(", ", parameters.Select(p => $"{p["name"]} ({p["in"]})"));
                            }

                            row++;
                            worksheet.Cells[row, 1].Value = path;
                            worksheet.Cells[row, 2].Value = verb.ToUpper();
                            worksheet.Cells[row, 3].Value = requestParameters;
                            worksheet.Cells[row, 4].Value = requestJsonBodySchema;
                            worksheet.Cells[row, 5].Value = responseJsonBodySchema;
                            worksheet.Cells[row, 6].Value = responseDescription;

                            // Apply formatting to data rows
                            using (var dataCells = worksheet.Cells[row, 1, row, 6])
                            {
                                dataCells.Style.WrapText = true; // Enable word wrapping
                            }
                        }
                    }

                    // AutoFit columns based on content
                    worksheet.Cells.AutoFitColumns();

                    // Set minimum column widths to ensure readability
                    worksheet.Column(1).Width = Math.Max(worksheet.Column(1).Width, 50); // Path
                    worksheet.Column(2).Width = Math.Max(worksheet.Column(2).Width, 10); // Verb
                    worksheet.Column(3).Width = Math.Max(worksheet.Column(3).Width, 25); // Request Parameters
                    worksheet.Column(4).Width = Math.Max(worksheet.Column(4).Width, 30); // Request JSON Body Schema
                    worksheet.Column(5).Width = Math.Max(worksheet.Column(5).Width, 30); // Response JSON Body Schema
                    worksheet.Column(6).Width = Math.Max(worksheet.Column(6).Width, 30); // Response Description
                    if(File.Exists($"C:\\Users\\{Environment.UserName}\\output.xlsx"))
                    {
                        File.Copy($"C:\\Users\\{Environment.UserName}\\output.xlsx", $"C:\\Users\\{Environment.UserName}\\output_{DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")}.xlsx");
                    }
                    excelPackage.SaveAs(new FileInfo($"C:\\Users\\{Environment.UserName}\\output.xlsx"));
                }

                Console.WriteLine($"Excel file generated successfully and saved at below location: \nC:\\Users\\{Environment.UserName}\\output.xlsx");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Error: {ex.Message}\n"+ex);
                Console.WriteLine($"Error: {ex.Message}\nPlease Check if the excel file is still open.");
                Console.ReadKey();
            }
        }

        static string GetRequestJsonBodySchema(JObject content, Dictionary<string, JObject> schemas)
        {
            if (content == null || schemas == null)
                return "No schema";

            // Try to get schema reference from multiple content types
            var contentTypes = new[] { "application/json", "text/json", "application/*+json" };
            foreach (var contentType in contentTypes)
            {
                if (content.TryGetValue(contentType, out var contentDetail))
                {
                    var schemaRef = contentDetail["schema"]?["$ref"]?.ToString();
                    if (!string.IsNullOrEmpty(schemaRef))
                    {
                        return GetSchemaReference(schemaRef, schemas);
                    }
                }
            }

            return "No schema";
        }

        static string GetResponseSchemaReference(JObject responses, Dictionary<string, JObject> schemas)
        {
            if (responses == null || schemas == null)
                return "No schema";

            // Try to get schema reference from multiple content types
            var contentTypes = new[] { "application/json", "text/plain", "text/json" };
            if (responses.TryGetValue("200", out var response))
            {
                var content = response["content"] as JObject;
                foreach (var contentType in contentTypes)
                {
                    if (content != null && content.TryGetValue(contentType, out var contentDetail))
                    {
                        var schemaRef = contentDetail["schema"]?["$ref"]?.ToString();
                        if (!string.IsNullOrEmpty(schemaRef))
                        {
                            return GetSchemaReference(schemaRef, schemas);
                        }
                    }
                }
            }

            return "No schema";
        }

        static string GetSchemaReference(string refPath, Dictionary<string, JObject> schemas)
        {
            if (string.IsNullOrWhiteSpace(refPath))
                return null;

            // Strip the leading "#/" and split by "/"
            var refPathSegments = refPath.TrimStart('#').Split('/');
            var schemaKey = refPathSegments.Last();

            if (schemas != null && schemas.ContainsKey(schemaKey))
            {
                return ResolveSchema(schemas[schemaKey], schemas);
            }

            return null;
        }

        static string ResolveSchema(JObject schema, Dictionary<string, JObject> schemas)
        {
            if (schema == null)
                return null;

            var schemaCopy = schema.DeepClone() as JObject;

            // Resolve $ref references in schema
            foreach (var property in schemaCopy.Properties().ToList())
            {
                if (property.Value is JObject propertyObject)
                {
                    var refPath = propertyObject["$ref"]?.ToString();
                    if (!string.IsNullOrEmpty(refPath))
                    {
                        var resolvedSchema = GetSchemaReference(refPath, schemas);
                        if (resolvedSchema != null)
                        {
                            property.Value = JObject.Parse(resolvedSchema);
                        }
                    }
                    else
                    {
                        // Recursively resolve nested properties
                        ResolveSchema(propertyObject, schemas);
                    }
                }
                else if (property.Value is JArray array)
                {
                    // Handle arrays
                    foreach (var item in array.Children<JObject>())
                    {
                        var itemSchemaRef = item["$ref"]?.ToString();
                        if (!string.IsNullOrEmpty(itemSchemaRef))
                        {
                            var resolvedItemSchema = GetSchemaReference(itemSchemaRef, schemas);
                            if (resolvedItemSchema != null)
                            {
                                property.Value = JObject.Parse(resolvedItemSchema);
                            }
                        }
                    }
                }
            }
            // Format the schema for better readability
            return schemaCopy.ToString(Formatting.Indented);
        }
    }
}