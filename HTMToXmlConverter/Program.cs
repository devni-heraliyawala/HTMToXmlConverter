using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using HtmlAgilityPack;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Drawing;
using Microsoft.Extensions.Configuration;
using ClosedXML.Excel;

namespace Magic5XmlGenerator
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfiguration configuration = builder.Build();

            Console.Write("Please enter the base path (e.g) W:\\Boris\\LandSec\\BoarLane\\LS Fire Stopping - Boar Lane: ");
            string basePath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(basePath) || !Directory.Exists(basePath))
            {
                Console.WriteLine("Invalid base path.");
                return;
            }

            string htmlDirectory = Path.Combine(basePath, "Records");
            string imageDirectory = Path.Combine(basePath, "Images");
            string drawingDirectory = Path.Combine(basePath, "Plans");
            string xmlDirectory = Path.Combine(basePath, "XmlDirectory");
            Directory.CreateDirectory(xmlDirectory);

            GenerateMagic5Xml(htmlDirectory, imageDirectory, drawingDirectory, xmlDirectory, basePath, configuration);

            Console.WriteLine("Processing complete!");
            Console.ReadKey();
        }

        static void GenerateMagic5Xml(string htmlDirectory, string imageDirectory, string drawingDirectory, string xmlDirectory, string basePath, IConfiguration configuration)
        {
            string defaultConnection = configuration.GetConnectionString("DefaultConnection");

            string fullFolderName = new DirectoryInfo(basePath).Name;
            string locationName = fullFolderName.Contains('-')
                ? fullFolderName.Substring(fullFolderName.LastIndexOf('-') + 1).Trim()
                : fullFolderName;

            string customerName = GetCustomerName(defaultConnection, locationName);
            string customerExternalPK = customerName.Replace(" ", "");
            // Create the <Customer> element with desired attributes
            var customerElement = new XElement("Customer",
                new XAttribute("name", customerName),
                new XAttribute("externalPK", customerExternalPK)
            );

            
            //string locationExternalPK = locationName.Replace(" ", "");
            string locationExternalPK = fullFolderName.Replace(" ","");
            string drawingListExternalPK = $"DRAWINGS_{locationExternalPK}";

            var locationElement = new XElement("Location",
                new XAttribute("name", locationName),
                new XAttribute("externalPK", locationExternalPK),
                //new XAttribute("drawingListIdPK", drawingListExternalPK)
                new XAttribute("AttListPK_drawingListId", drawingListExternalPK)
            );

            // Add existing elements to the <Customer> element
            customerElement.Add(locationElement);

            var drawingList = new XElement("List",
                new XAttribute("externalPK", drawingListExternalPK),
                new XAttribute("name", $"Drawing list - {locationName}")
            );

            Dictionary<string, XElement> assetLists = new();

            var htmlFiles = Directory.GetFiles(htmlDirectory, "*.htm");
            Dictionary<string, (int width, int height)> drawingDimensions = new Dictionary<string, (int width, int height)>();
            List<Dictionary<string, string>> allRecords = new();

            foreach (var htmlFile in htmlFiles)
            {
                var record = ExtractRecord(htmlFile);
                var imagePaths = ExtractImagePaths(htmlFile);
                var drawingPaths = ExtractDrawingPaths(htmlFile);
                var markers = ExtractMarkers(htmlFile);

                if (htmlFile.Contains("template", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                allRecords.Add(record);

                string drawingName = record.TryGetValue("Level", out var dname) && !string.IsNullOrWhiteSpace(dname) ? dname : "Drawing 1";
                string drawingNameId = drawingName.Replace(" ", "");
                string drawingId = $"DRAWINGS_{locationExternalPK}_{drawingNameId}";
                string assetListId = $"ASSETS_{locationExternalPK}_{drawingNameId}";

                if (!drawingList.Elements("ListEntry").Any(e => (string)e.Attribute("externalPK") == drawingId))
                {
                    var drawingEntry = new XElement("ListEntry",
                        new XAttribute("externalPK", drawingId),
                        new XAttribute("name", drawingName),
                        //new XAttribute("assetListIdPK", assetListId)
                        //new XAttribute("AttListPK_assetListId", assetListId),
                        new XAttribute("AttListPK_drawingFSAssetsListId", assetListId)
                    );

                    // Attach drawings here only
                    if (drawingPaths.Any())
                    {
                        var attachments = new XElement("Attachments",
                            new XAttribute("ignoreAttachmentIfSameNameAlreadyExists", "true"),
                            new XAttribute("ignoreAttachmentIfFileDoesntExist", "true")
                        );

                        foreach (var drawingPath in drawingPaths)
                        {
                            string fullDrawingFile = FindFileWithExtension(drawingDirectory, Path.GetFileName(drawingPath));
                            var localPath = Path.Combine(drawingDirectory, Path.GetFileName(fullDrawingFile));

                            if (File.Exists(localPath))
                            {
                                using (var image = Image.FromFile(localPath))
                                {
                                    int width = image.Width;
                                    int height = image.Height;

                                    drawingDimensions[drawingName] = (width, height);
                                }
                                attachments.Add(new XElement("fileName",
                                new XAttribute("friendlyName", Path.GetFileName(fullDrawingFile)),
                                new XAttribute("text", localPath)
                            ));
                            }
                        }

                        drawingEntry.Add(attachments);
                    }

                    drawingList.Add(drawingEntry);
                }

                string assetName = record.TryGetValue("New Fire Stopping Number", out var aname) && !string.IsNullOrWhiteSpace(aname) ? aname : "Asset 1";

                var assetEntry = new XElement("ListEntry",
                    new XAttribute("externalPK", Path.GetFileNameWithoutExtension(htmlFile)),
                    new XAttribute("name", assetName)
                );

                foreach (var kv in record)
                {
                    if (!string.IsNullOrWhiteSpace(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value))
                    {
                        string key = Regex.Replace(kv.Key, "[^\\w]", "");
                        key = char.ToLowerInvariant(key[0]) + key.Substring(1);
                        assetEntry.SetAttributeValue(key, kv.Value);
                    }
                }

                string assetId = Path.GetFileNameWithoutExtension(htmlFile);
                var marker = markers.FirstOrDefault(m => m.id == Convert.ToInt32(assetId)); // Replace 'assetId' with the appropriate identifier
                
                if (marker != null && drawingPaths.Any())
                {
                    // find the drawing from drawing list and load and get the image values
                    var drawDims = drawingDimensions.TryGetValue(drawingName, out var dims);

                    int scaledX = (int)(marker.x * dims.width);
                    int scaledY = (int)(marker.y * dims.height);

                    assetEntry.SetAttributeValue("markerId", marker.id);
                    assetEntry.SetAttributeValue("x", scaledX);
                    assetEntry.SetAttributeValue("y", scaledY);
                }

                if (imagePaths.Any())
                {
                    var attachments = new XElement("Attachments",
                        new XAttribute("ignoreAttachmentIfSameNameAlreadyExists", "true"),
                        new XAttribute("ignoreAttachmentIfFileDoesntExist", "true")
                    );

                    foreach (var imgPath in imagePaths)
                    {
                        var localPath = Path.Combine(imageDirectory, Path.GetFileName(imgPath));
                        attachments.Add(new XElement("fileName",
                            new XAttribute("friendlyName", Path.GetFileName(imgPath)),
                            new XAttribute("text", localPath)
                        ));
                    }
                    assetEntry.Add(attachments);
                }

                if (!assetLists.ContainsKey(assetListId))
                {
                    assetLists[assetListId] = new XElement("List",
                        new XAttribute("externalPK", assetListId),
                        new XAttribute("name", $"Asset List - Drawing list - {locationName} - {drawingName}")
                    );
                }

                assetLists[assetListId].Add(assetEntry);
            }

            // Export to Excel file
            ExportRecordsToExcel(allRecords, "output.xlsx");

            var magic5 = new XElement("magic5In",
                new XAttribute("version", "2.1.0"),
                customerElement,
                drawingList
            );

            foreach (var list in assetLists.Values)
            {
                magic5.Add(list);
            }

            var xmlDoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), magic5);
            string xmlPath = Path.Combine(xmlDirectory, locationExternalPK + ".xml");
            xmlDoc.Save(xmlPath);

            Console.WriteLine($"XML saved to: {xmlPath}");
        }

        private static void ExportRecordsToExcel(List<Dictionary<string, string>> records, string outputPath)
        {
            var columnMappings = new Dictionary<string, string>
            {
                { "name", "New Fire Stopping Number" },
                { "buildingName", "Building Name" },
                { "level", "Level" },
                { "area", "Area" },
                { "existingInstallType", "Existing Install Type" },
                { "existingCondition", "Existing Condition" },
                { "repairReplace", "Repair / Replace" },
                { "status", "Status" },
                { "lDform1Fitter", "Created By" },
                { "lDpenServ", "Type of Penetration" },
                { "repairType", "Repair Type" },
                { "access", "Is the Area Accessible?" },
                { "cwDate", "Date of Inspection" },
                { "dateNext", "Date of Next Inspection" },
                { "role", "role" },
                { "other", "If Other, please specify" },
                { "completeBy", "Works to be completed by" },
                { "detailsRepairReplacement", "Details of Repair/Replacement Required (If Required)" },
                { "comments", "General Comments" },
                { "detailsRepairsReplacementsCarriedOut", "Details of Repairs/Replacements Carried Out" },
                { "existingFireStoppingNumber", "Existing Fire Stopping Number" },
                { "newFireStoppingNumber", "New Fire Stopping Number" },
                { "productManufacturersDetails", "Product Manufacturers Detail" },
                { "otherRelevantInformation", "Other Relevant information" }
            };


            var headers = new[]
            {
                "name", "buildingName", "level", "area", "existingInstallType", "existingCondition", "repairReplace",
                "status", "lDform1Fitter", "lDpenServ", "repairType", "access", "cwDate", "dateNext", "role",
                "other", "completeBy", "detailsRepairReplacement", "comments", "detailsRepairsReplacementsCarriedOut",
                "existingFireStoppingNumber", "newFireStoppingNumber", "productManufacturersDetails", "otherRelevantInformation"
            };

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("HTM Records");

            // Write header row
            int colIndex = 1;
            foreach (var header in columnMappings.Keys)
            {
                worksheet.Cell(1, colIndex++).Value = header;
            }

            // Write record data rows
            for (int r = 0; r < records.Count; r++)
            {
                var record = records[r];
                colIndex = 1;

                foreach (var header in columnMappings.Keys)
                {
                    string mappedKey = columnMappings[header];
                    record.TryGetValue(mappedKey, out var value); // safe if key missing
                    worksheet.Cell(r + 2, colIndex++).Value = value ?? "";
                }
            }

            workbook.SaveAs(outputPath);
            Console.WriteLine($"Excel saved: {outputPath}");
        }

        static Dictionary<string, string> ExtractRecord(string htmlFile)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlFile);

            var record = new Dictionary<string, string>();
            var rows = doc.DocumentNode.SelectNodes("//table[contains(@class, 'table-details')]//tr");

            if (rows != null)
            {
                foreach (var row in rows)
                {
                    var tds = row.SelectNodes("td");
                    if (tds != null && tds.Count == 2)
                    {
                        var key = HtmlEntity.DeEntitize(tds[0].InnerText.Trim());
                        var value = HtmlEntity.DeEntitize(tds[1].InnerText.Trim());
                        record[key] = value;
                    }
                }
            }

            return record;
        }

        static List<string> ExtractImagePaths(string htmlFile)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlFile);

            var imageNodes = doc.DocumentNode
                .SelectNodes("//h5[text()='Images']/following::div[contains(@class, 'carousel')]//img");

            var imagePaths = new List<string>();

            if (imageNodes != null)
            {
                foreach (var node in imageNodes)
                {
                    var src = node.GetAttributeValue("src", null);
                    if (!string.IsNullOrEmpty(src))
                    {
                        var cleanPath = src.Replace("../images/", "").Trim();
                        if (!string.IsNullOrWhiteSpace(cleanPath))
                            imagePaths.Add(cleanPath);
                    }
                }
            }

            return imagePaths.Distinct().ToList();
        }

        static List<string> ExtractDrawingPaths(string htmlFile)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlFile);

            var drawingPaths = new List<string>();
            var scriptNodes = doc.DocumentNode.SelectNodes("//script[contains(text(), 'AddPlan')]");

            if (scriptNodes != null)
            {
                foreach (var node in scriptNodes)
                {
                    var match = Regex.Match(node.InnerText, "AddPlan\\('planId',\\s*'',\\s*\\d+,\\s*\\d+,\\s*'(?<url>[^']+)'");
                    if (match.Success)
                    {
                        var path = match.Groups["url"].Value.Replace("../plans/", "");
                        drawingPaths.Add(path);
                    }
                }
            }

            return drawingPaths.Distinct().ToList();
        }

        public static List<Marker> ExtractMarkers(string htmlFilePath)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlFilePath);

            // Find all script tags
            var scriptNodes = doc.DocumentNode.SelectNodes("//script");
            if (scriptNodes == null)
                return new List<Marker>();

            foreach (var script in scriptNodes)
            {
                var scriptContent = script.InnerText;

                // Look for the markersdata variable
                if (scriptContent.Contains("var markersdata"))
                {
                    // Extract the JSON array assigned to markersdata
                    var match = Regex.Match(scriptContent, @"var\s+markersdata\s*=\s*(\[\{.*?\}\]);", RegexOptions.Singleline);
                    if (match.Success)
                    {
                        var jsonArray = match.Groups[1].Value;

                        // Deserialize the JSON array into a list of Marker objects
                        var markers = JsonConvert.DeserializeObject<List<Marker>>(jsonArray);
                        return markers;
                    }
                }
            }

            return new List<Marker>();
        }

        static string GetCustomerName(string connectionString, string locationName)
        {
            string query = @"
                SELECT c.Description
                FROM dbo.tblCustomers c
                INNER JOIN dbo.tblLocations l ON c.Id = l.CustomerId
                WHERE l.Description LIKE @LocationName";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand(query, connection))
                {

                    command.Parameters.Add("@LocationName", SqlDbType.NVarChar).Value = $"%{locationName}%";

                    connection.Open();
                    object result = command.ExecuteScalar();

                    // Check if result is not null and not DBNull
                    if (result != null && result != DBNull.Value)
                    {
                        return result.ToString();
                    }
                    else
                    {
                        // Handle case where no matching customer is found
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log exception or handle accordingly
                Console.WriteLine($"Error retrieving customer name: {ex.Message}");
                return null;
            }
        }

        static string FindFileWithExtension(string directory, string fileStem)
        {
            string cleanName = Path.GetFileName(fileStem);
            try
            {
                var files = Directory.GetFiles(directory, cleanName + ".*");
                return files.Length > 0 ? Path.GetFileName(files[0]) : null;
            }
            catch
            {
                return null;
            }
        }
    }
}
