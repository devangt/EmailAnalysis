using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using EmailAnalysis;

namespace TestConsole
{
    public class ScoreData
    {
        public Dictionary<string, string> FeatureVector { get; set; }
        public Dictionary<string, string> GlobalParameters { get; set; }
    }

    public class ScoreRequest
    {
        public string Id { get; set; }
        public ScoreData Instance { get; set; }
    }

    public class WebserviceClient
    {
        public async Task<double?> InvokeRequestResponseService(Dictionary<string, string> featureVector)
        {
            double? result = null;
            using (var client = new HttpClient())
            {
                var scoreData = new ScoreData
                {
                    FeatureVector = featureVector,
                    GlobalParameters = new Dictionary<string, string>()
                };

                var scoreRequest = new ScoreRequest
                {
                    Id = "score00001",
                    Instance = scoreData
                };

                string apiKey = ConfigurationManager.AppSettings["apiKey"];
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                client.BaseAddress = new Uri(ConfigurationManager.AppSettings["BaseAddress"]);
                HttpResponseMessage response = await client.PostAsJsonAsync(string.Empty, scoreRequest);

                if (response.IsSuccessStatusCode)
                {
                    var responseArray = await response.Content.ReadAsAsync<string[]>();
                    result = double.Parse(responseArray[9]);
                }
                else
                {
                    Console.WriteLine("Failed with status code: {0}", response.StatusCode);
                }
            }

            return result;
        }
    }

    public class ValidatePredictionRunner
    {
        public async void Run(Func<Dictionary<string, object>, bool> filter)
        {
            var client = new WebserviceClient();

            var analyser = new EmailAnalyser();

            var emailData = analyser.GetEmailData();

            var testOutput = new List<Dictionary<string, object>>();

            foreach (var data in emailData.Where(filter))
            {
                var featureVector = new Dictionary<string, string>
                {
                    {"TestFolder", (bool)data["TestFolder"] ? "1" : "0"},
                    {"HasAttachments", (bool)data["HasAttachments"] ? "1" : "0"},
                    {"SentDirect", (bool)data["SentDirect"] ? "1" : "0"},
                    {"MayContainATime", (bool) data["MayContainATime"] ? "1" : "0"},
                    {"ReceivedHour", data["ReceivedHour"].ToString()},
                    {"SubjectWordCount", data["SubjectWordCount"].ToString()},
                    {"SenderDomain", data["SenderDomain"].ToString()},
                    {"HasCC", (bool)data["HasCC"] ? "1" : "0"},
                    {"SpecialCharacterCount", data["SpecialCharacterCount"].ToString()}
                };

                var result = await client.InvokeRequestResponseService(featureVector);

                data.Add("Result", result);

                testOutput.Add(data);

                Console.WriteLine(result);
            }

            EmailAnalyser.OutputDataset(testOutput, "EmailDatasetPredictions.csv");
            Console.WriteLine("DONE");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Get 2 days of test data

            Func<Dictionary<string, object>, bool> filter = data =>
                ((string)data["FolderName"] == "Whereabouts" && (DateTime)data["Recieved"] >= DateTime.Now.Subtract(TimeSpan.FromDays(2))) ||
                ((string)data["FolderName"] == "Inbox" && (DateTime)data["Recieved"] >= DateTime.Now.Subtract(TimeSpan.FromDays(2)));

            new ValidatePredictionRunner().Run(filter);

            Console.ReadLine();
        }
    }
}