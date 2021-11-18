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

                        if (!File.Exists(@".\Positions\Countries.txt"))
                        {
                            File.Create(@".\Positions\Countries.txt");
                        }
                        using (var countryWriterStream = (File.CreateText(@".\Positions\Countries.txt")))
                        {
                            //create a list of the countries in a file
                            foreach (var countryName in list.Keys.Distinct())
                            {

                                using (var writer = new JsonTextWriter(countryWriterStream))
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
                            }
                        };

                        //create a list of locations within country files
                        foreach (var countryAsString in list.Keys.Distinct())
                        {
                            //here I'm trying to create a file for each country that's distinct in the list of keys.

                            var countryFileNameString = String.Format(@".\Positions\{0}.txt", countryAsString);
                            if (!File.Exists(countryFileNameString))
                            {
                                File.Create(countryFileNameString);
                            }

                            File.CreateText(countryFileNameString);
                            using (var locationStreamWriter = (File.CreateText(countryFileNameString)))
                            {
                                using (var writer = new JsonTextWriter(locationStreamWriter))
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
                        }
                    }
            }
        }
    }
}