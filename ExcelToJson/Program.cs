using System;
using System.IO;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {            
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var inFilePath = args[0];
            var outFilePath = args[1];

            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
            using (var outFile = File.CreateText(outFilePath))
            {
                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                { FallbackEncoding = Encoding.GetEncoding(1252) }))
                using (var writer = new JsonTextWriter(outFile))
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

                                writer.WritePropertyName("country");
                                writer.WriteValue(reader.GetString(6).ToString());

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