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
            List<string> countryNames = new List<string>();     //create something to match against
            SortedList<string, List<Location>> sortedLocations = new SortedList<string, List<Location>>();

            //open the file for reading
            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                { FallbackEncoding = Encoding.GetEncoding(1252) }))

                    //try to read the whole document and make classes from it
                    while (reader.Read())
                    {
                        try
                        {
                            unsortedAllLocations.Add(new Location()
                            {
                                country = reader.GetString(6).ToString(),
                                state = reader.GetString(0).ToString(),
                                city = reader.GetString(1).ToString(),
                                name = reader.GetString(2).ToString(),
                                icao = reader.GetString(3).ToString(),
                                lat = reader.GetDouble(4),
                                lon = reader.GetDouble(5)
                            });

                            if (!countryNames.Contains(reader.GetString(6).ToString()))
                            {
                                countryNames.Add(reader.GetString(6).ToString());                 //adding to our list of comparable keys
                            }

                            reader.Read();

                        }

                        catch (Exception ex)
                        {
                            Console.WriteLine("issue:{0}", ex.Message);
                        }
                    }

                foreach (var countryName in countryNames)
                {
                    List<Location> separateLocationListsByCountry = new List<Location>();
                    separateLocationListsByCountry = unsortedAllLocations.Where(x => x.country == countryName).ToList();
                    sortedLocations.Add(countryName, separateLocationListsByCountry);
                }


                CreateCountryLocationFiles(sortedLocations);
            }


            void CreateCountryLocationFiles(SortedList<string, List<Location>> list)
            {
                List<Location> fileWritingLocations = new List<Location>();

                //get the name of each country
                foreach (var countryAsString in list.Keys.Distinct())
                {
                    foreach (var listLocation in list.Values)
                    {

                        foreach (var location in listLocation)
                        {
                            fileWritingLocations.AddRange(listLocation.Where(x => x.country == countryAsString).ToList());

                            using (var sw = new StreamWriter("./Countries/" + countryAsString + ".txt"))
                                Write(sw, fileWritingLocations);
                           
                        }
                    }
                    fileWritingLocations.RemoveAll(x => x.country == countryAsString);
                }

            }
        }

        internal static void Write(StreamWriter sw, List<Location> locations)
        {
            using (var writer = new JsonTextWriter(sw))
            {
                writer.Formatting = Formatting.Indented;
                writer.WriteStartArray();

                foreach (var location in locations)
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

            }
            sw.Close();
            
            return;
        }

    }
}


