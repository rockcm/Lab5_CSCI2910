using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Text.Json;
using SteamProject;
using OfficeOpenXml;
using System.Formats.Asn1;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;

class Program
{
    static async Task Main()
    {
        string apiUrl = $"https://bymykel.github.io/CSGO-API/api/en/skins.json";

        List<Skin> skinList = new List<Skin>();


        // Creates an instance of HttpClient for sending HTTP requests
        using (HttpClient httpClient = new HttpClient())
        {
            //try catch to ensure no errors crash program 
            try
            {
                HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    // makes the JSON response into a JArray, Jarray. works like deserializer 
                    JArray skinsArray = JArray.Parse(jsonResponse);

                    // iterates through the Jarray and adds data to the list, uses the token like a key to pull out the information relating to each token. 
                    foreach (JToken obj in skinsArray)
                    {
                        string name = obj["name"].ToString();
                        string description = obj["description"].ToString();
                        string weaponName = obj["weapon"]["name"].ToString(); // gets Weapons "name"
                        string rarityName = obj["rarity"]["name"].ToString(); // gets Rarity's Name 

                        // Create a SkinData object and add it to the list
                        Skin skinData = new Skin
                        {
                            Name = name,
                            Description = description,
                            WeaponName = weaponName,
                            RarityName = rarityName
                        };

                        skinList.Add(skinData);

                        // Display data in the console so you can view whats going in sheets 
                        Console.WriteLine($"Name: {name}");
                        Console.WriteLine($"Description: {description}");
                        Console.WriteLine($"Weapon: {weaponName}");
                        Console.WriteLine($"Rarity: {rarityName}");
                        Console.WriteLine();
                      
                    }
                }
                else
                {
                    Console.WriteLine("Error fetching CS:GO inventory data. Error Code: " + response.StatusCode);
                }
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine("HTTP Request Error: " + e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
        }

        using (ExcelPackage package = new ExcelPackage()) // downloaded packages to make excel sheets 
        {
            // Add a new worksheet to the Excel package with the name "CSGO Skins"
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("CSGO Skins");

            // Set column headers in the first row of the Excel worksheet
            worksheet.Cells[1, 1].Value = "Name";
            worksheet.Cells[1, 2].Value = "Description";
            worksheet.Cells[1, 3].Value = "Weapon";
            worksheet.Cells[1, 4].Value = "Rarity";

            // Initialize the row counter to start from the second row
            int row = 2;

            // Iterate through the list of SkinData objects to fill the Excel worksheet
            foreach (Skin skinData in skinList)
            {
                // Set the values for each column in the current row
                worksheet.Cells[row, 1].Value = skinData.Name;
                worksheet.Cells[row, 2].Value = skinData.Description;
                worksheet.Cells[row, 3].Value = skinData.WeaponName;
                worksheet.Cells[row, 4].Value = skinData.RarityName;

                // Move to the next row
                row++;
            }

            // Defines the file path and name for the Excel file
            FileInfo fileInfo = new FileInfo("CSGOSkins.xlsx");

            // Save the Excel package to the specified file
            package.SaveAs(fileInfo);

            // Prints message that the Excel file has been created
            Console.WriteLine("Excel file has been created: " + fileInfo.FullName);
        }

        // Creates a CSV file
        using (var writer = new StreamWriter("CSGOSkins.csv"))
        using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
        {
            // Write the records from the skinList to the CSV file
            csv.WriteRecords(skinList);
        }

        // Prints message that the CSV file has been created
        Console.WriteLine("CSV file has been created: CSGOSkins.csv");

    }
}


