using Newtonsoft.Json;

var inFilePath = "Countries.txt";

using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
{

    var reader = new StreamReader(inFile);
    List<string> files = new List<string>();
    var text = reader.ReadToEnd();

    string[] data = text.Replace("txt", "").Split('.');
    files.Add(data.ToString());

    
        StreamWriter sw = new StreamWriter("Countries.json");
    using (var writer = new JsonTextWriter(sw))
    {
        writer.Formatting = Formatting.Indented;
        writer.WriteStartArray();

        foreach (var item in files)
        {
            writer.WriteStartObject();


            writer.WritePropertyName("country");
            writer.WriteValue(item);
        }

        writer.WriteEndObject();
        
    }
}
    