using System;
using System.IO;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;
using System.Linq;

namespace ExcelToJson
{
    class Program
    {
        public static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var inFilePath = "worldairports.xlsx";

            List<Location> unsortedAllLocations = new List<Location>();      

            //open the file for reading
            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))

            {
                List<string> countryNames = new List<string>();     //create something to match against

                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                { FallbackEncoding = Encoding.GetEncoding(1252) }))
                                        
                    //try to read the whole document and make classes from it
                    while (reader.Read())
                    {
                        try
                        {
                            Location location = new Location()
                            {
                                country = reader.GetString(6).ToString(),
                                state = reader.GetString(0).ToString(),
                                city = reader.GetString(1).ToString(),
                                name = reader.GetString(2).ToString(),
                                icao = reader.GetString(3).ToString(),
                                lat = reader.GetDouble(4),
                                lon = reader.GetDouble(5)
                            };

                            unsortedAllLocations.Add(location);

                            if (!countryNames.Contains(location.country))
                            {
                                countryNames.Add(location.country);                 //adding to our list of comparable keys
                            }
                        }

                        catch (Exception ex)
                        {
                            Console.WriteLine("issue:{0}", ex.Message);
                        }
                    }
                SortedList<string, List<Location>> sortedLocations = new SortedList<string, List<Location>>();
                
                

                foreach (var countryName in countryNames)
                {
                    List<Location> separateLocationListsByCountry = new List<Location>();
                    separateLocationListsByCountry = unsortedAllLocations.TakeWhile(x => x.country == countryName).ToList();
                    sortedLocations.Add(countryName, separateLocationListsByCountry);
                }               
                             

                CreateCountryLocationFiles(sortedLocations);
            }

            void CreateCountryLocationFiles(SortedList<string, List<Location>> list)
            {
                //get the name of each country
                foreach (var countryAsString in list.Keys.Distinct())
                {
                    //and a list of locations within that country
                    List<Location> locations = list.Values as List<Location>;


                    //make a new file to hold them
                    using (var locationStreamWriter = new StreamWriter(string.Format(@".\Countries\{0}.txt", countryAsString)))
                    {
                        using (var writer = new JsonTextWriter(locationStreamWriter))
                        {
                            writer.Formatting = Formatting.Indented;
                            writer.WriteStartArray();

                            foreach (var location in locations)
                            {
                                try
                                {
                                    writer.WriteStartObject();

                                    writer.WritePropertyName("country");
                                    writer.WriteValue(location.country);

                                    writer.WritePropertyName("state");
                                    writer.WriteValue(location.state);

                                    writer.WritePropertyName("city");
                                    writer.WriteValue(location.city);

                                    writer.WritePropertyName("name");
                                    writer.WriteValue(location.name);

                                    writer.WritePropertyName("icao");
                                    writer.WriteValue(location.icao);

                                    writer.WritePropertyName("lat");
                                    writer.WriteValue(location.lat);

                                    writer.WritePropertyName("lon");
                                    writer.WriteValue(location.lon);

                                    writer.WriteEndObject();
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("failed");
                                }
                                writer.WriteEndArray();
                            }

                        }
                    }
                }
            }
        }
    }
}
