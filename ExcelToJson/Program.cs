using System;
using System.IO;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;

namespace ExcelToJson
{
    class Program
    {
        static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var inFilePath = "worldairports.xlsx";


            SortedList<string, Location> list = new SortedList<string, Location>();

            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))

            {
                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                { FallbackEncoding = Encoding.GetEncoding(1252) }))

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

                            reader.Read();
                        }


                        foreach (var Country in list.Keys.Distinct())
                        {
                            var countryFileString = String.Format(@".\Positions\{0}.txt", Country);
                            using (var countryFile = (File.CreateText(countryFileString)))
                            {
                                using (var writer = new JsonTextWriter(countryFile))
                                {
                                    writer.Formatting = Formatting.Indented;
                                    writer.WriteStartArray();
                                    reader.Read(); //skip the headings
                                    do
                                    {
                                        while (reader.Read())
                                        {
                                            try
                                            {
                                                writer.WriteStartObject();
                                                writer.WritePropertyName("country");
                                                writer.WriteValue(reader.GetString(6).ToString());

                                                writer.WritePropertyName("state");
                                                writer.WriteValue(reader.GetString(0).ToString());

                                                writer.WritePropertyName("city");
                                                writer.WriteValue(reader.GetString(1).ToString());

                                                writer.WritePropertyName("name");
                                                writer.WriteValue(reader.GetString(2).ToString());

                                                writer.WritePropertyName("icao");
                                                writer.WriteValue(reader.GetString(3).ToString());

                                                writer.WritePropertyName("lat");
                                                writer.WriteValue(reader.GetDouble(4));

                                                writer.WritePropertyName("lon");
                                                writer.WriteValue(reader.GetDouble(5));

                                                writer.WriteEndObject();
                                            }
                                            catch (Exception)
                                            {
                                                reader.NextResult();
                                            }
                                        }
                                    } while (reader.NextResult());
                                    writer.WriteEndArray();
                                }
                            }


                            foreach (var countryName in list.Keys.Distinct())
                            {
                                using (var countryFile = (File.CreateText(@".\Positions\Countries.txt")))
                                {
                                    using (var writer = new JsonTextWriter(countryFile))
                                    {
                                        try
                                        {
                                            writer.WriteStartObject();
                                            writer.WritePropertyName("country");
                                            writer.WriteValue(reader.GetString(6).ToString());
                                        }
                                        catch (Exception)
                                        {
                                            reader.NextResult();
                                        }

                                    }
                                };
                            }
                        }
                    }
            }
        }
    }
}