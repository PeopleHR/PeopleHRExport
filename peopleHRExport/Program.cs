using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace peopleHRExport
{
   public static class Program
    {
        static void Main(string[] args)
        {
            if (args.Count()==1 && args[0]!="")
            {
                // Specify a name for your top-level folder. 
                string destinationPath = @args[0];

               // Console.WriteLine(destinationPath);

                if (System.IO.Directory.Exists(destinationPath))
                {

                    // To create a string that specifies the path to a subfolder under your  
                    // top-level folder, add a name for the subfolder to folderName. 
                    string destination_folderPath = destinationPath;

                    string APIKey = "";
                    string Queries = "";

                    Console.WriteLine("Program started V3.0");
                    //try
                    //{
                    // // Or specify a specific name in the current dir
                    //var MyIni = new IniFile("Settings.ini");

                    //var DefaultVolume = MyIni.Read("DefaultVolume");
                    //Console.WriteLine(DefaultVolume);
                    //Console.ReadKey();
                    //}
                    //catch(Exception ex)
                    //{
                    //    Console.WriteLine(ex.Message);
                    //    Console.ReadKey();
                    //}


                    string[] lines = System.IO.File.ReadAllLines(@"settings.txt");

                    // Display the file contents by using a foreach loop.
                    // System.Console.WriteLine("Contents of WriteLines2.txt = ");
                    if (lines.Count() == 2)
                    {

                        // Get API key 
                        // Use a tab to indent each line of the file.

                        // explode "=" to check it has valid data or not

                        string[] lineitems1 = lines[0].Split('=');

                        if (lineitems1.Count() == 2)
                        {
                            if (lineitems1[0] == "API_KEY")
                            {
                                APIKey = lineitems1[1];
                            }
                            else
                            {
                                // wrong format 
                            }
                        }
                        else
                        {
                            // wrong format
                        }



                        // get Queries list 

                        string[] lineitems2 = lines[1].Split('=');

                        if (lineitems2.Count() == 2)
                        {
                            if (lineitems2[0] == "Query")
                            {
                                Queries = lineitems2[1];
                            }
                            else
                            {
                                // wrong format 
                            }
                        }
                        else
                        {
                            // wrong format
                        }


                    }

                    if (APIKey != "" && Queries != "")
                    {
                        //Console.WriteLine(APIKey);
                        //Console.WriteLine(Queries);

                        string[] queries_list = Queries.Split(',');
                        foreach (var query_single in queries_list)
                        {
                            // Call PeopleHR to get query result 

                            try
                            {
                                // Limit  API Call 
                                System.Threading.Thread.Sleep(1000);

                                WebRequest req_peopleHR = WebRequest.Create("https://api.peoplehr.net/Query");
                                HttpWebRequest httpreq_peopleHR = (HttpWebRequest)req_peopleHR;
                                httpreq_peopleHR.Method = "POST";
                                httpreq_peopleHR.Timeout = 100000;

                                httpreq_peopleHR.ContentType = "application/json";
                                // httpreq_mandrill.Headers.Add("Authorization", "Basic " + asana_APIKey.Text);
                                Stream str_peopleHR = httpreq_peopleHR.GetRequestStream();
                                StreamWriter strwriter_peoleHR = new StreamWriter(str_peopleHR, Encoding.ASCII);


                                string soaprequest_peopleHR = "{\"APIKey\": \"" + APIKey + "\",\"Action\": \"GetQueryResult\",\"QueryName\":\"" + query_single + "\"}";
                                //MessageBox.Show(soaprequest_mandrill);
                               // Console.WriteLine(soaprequest_peopleHR);

                                strwriter_peoleHR.Write(soaprequest_peopleHR.ToString());
                                strwriter_peoleHR.Close();
                                HttpWebResponse res_peopleHR = (HttpWebResponse)httpreq_peopleHR.GetResponse();
                                if (res_peopleHR.StatusCode == HttpStatusCode.OK || res_peopleHR.StatusCode == HttpStatusCode.Accepted)
                                {
                                    StreamReader rdr_peopleHR = new StreamReader(res_peopleHR.GetResponseStream());
                                    string result_QueryPeopleHR = rdr_peopleHR.ReadToEnd();

                                    JObject json_export = JObject.Parse(result_QueryPeopleHR);

                                    JToken token = JObject.Parse(result_QueryPeopleHR);

                                    bool isError = (bool)token.SelectToken("isError");
                                    string API_message = (string)token.SelectToken("Message");

                                    if (!isError)
                                    {

                                        string csvString = "";
                                        int cnt = 0;



                                        foreach (JToken child in json_export["Result"].Children())
                                        {
                                            if (cnt == 0)
                                            {
                                                foreach (JToken grandGrandChild in child)
                                                {
                                                    var property = grandGrandChild as JProperty;

                                                    csvString += "\"" + property.Name + "\"" + ",";

                                                }
                                            }
                                            csvString += "\n";

                                            foreach (JToken grandGrandChild in child)
                                            {
                                                var property = grandGrandChild as JProperty;

                                                csvString += "\"" + property.Value + "\"" + ",";

                                            }
                                            cnt++;


                                        }



                                        string destination_queryFile = System.IO.Path.Combine(destination_folderPath, query_single + ".csv");
                                        System.IO.File.WriteAllText(destination_queryFile, csvString);
                                        Console.WriteLine(query_single + " - Exported");
                                    }

                                    else
                                    {
                                        Console.WriteLine(query_single + " - Error - " + API_message);

                                    }
                                }

                                else
                                {
                                    Console.WriteLine("API Error");

                                }
                                
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }
                    }
                    else
                    {
                        Console.WriteLine("Wrong settings");
                    }

                    // Keep the console window open in debug mode.
                    Console.WriteLine("Program Successfully Executed . Press any key to exit.");
                    System.Console.ReadKey();

                }
                else
                {
                    Console.WriteLine("Destination Path Not found . Please provide correct destination path ");
                    System.Console.ReadKey();
                }
        }
       else
   {
       Console.WriteLine("Incomplete command . Destination path is not specified");
       System.Console.ReadKey();
   }
   }



        public static string ToCSV(this DataTable table, string delimator)
        {
            var result = new StringBuilder();
            for (int i = 0; i < table.Columns.Count; i++)
            {
                result.Append(table.Columns[i].ColumnName);
                result.Append(i == table.Columns.Count - 1 ? "\n" : delimator);
            }
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    result.Append(row[i].ToString());
                    result.Append(i == table.Columns.Count - 1 ? "\n" : delimator);
                }
            }
            return result.ToString().TrimEnd(new char[] { '\r', '\n' });
            //return result.ToString();
        }
    }
}
