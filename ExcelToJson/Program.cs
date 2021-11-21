using System;
using System.IO;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;

namespace ExcelToJson
{
    class Program
    {
        public static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var inFilePath = "worldairports.xlsx";

            //create a list of countries and the locations in them in a sorted set
            SortedList<string, Location> list = new SortedList<string, Location>();

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

                            list.Add(reader.GetString(6).ToString(), location);

                        }
                        catch (Exception)
                        {
                            //if we can't read the record, skip it.
                            reader.Read();
                        }
                    }

                CreateCountryLocationFiles(list);
            }

           void CreateCountryLocationFiles(SortedList<string, Location> list)
            {
                //create a list of locations within country files
                foreach (var countryAsString in list.Keys.Distinct())
                {

                    using (var locationStreamWriter = new StreamWriter(string.Format(@".\Countries\{0}.txt", countryAsString)))
                    {
                        using (var writer = new JsonTextWriter(locationStreamWriter))
                        {
                            writer.Formatting = Formatting.Indented;
                            writer.WriteStartArray();
                            foreach (var location in list)
                            {
                                try
                                {
                                    writer.WriteStartObject();

                                    writer.WritePropertyName("country");
                                    writer.WriteValue(location.Value.country);

                                    writer.WritePropertyName("state");
                                    writer.WriteValue(location.Value.state);

                                    writer.WritePropertyName("city");
                                    writer.WriteValue(location.Value.city);

                                    writer.WritePropertyName("name");
                                    writer.WriteValue(location.Value.name);

                                    writer.WritePropertyName("icao");
                                    writer.WriteValue(location.Value.icao);

                                    writer.WritePropertyName("lat");
                                    writer.WriteValue(location.Value.lat);

                                    writer.WritePropertyName("lon");
                                    writer.WriteValue(location.Value.lon);

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
