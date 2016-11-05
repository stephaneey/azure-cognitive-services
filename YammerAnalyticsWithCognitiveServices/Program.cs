using Microsoft.ProjectOxford.Vision;
using Microsoft.ProjectOxford.Vision.Contract;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Nito.AsyncEx;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Rest;
using System.Configuration;

namespace YammerAnalyticsWithCognitiveServices
{
    class YammerGroup
    {
        public List<YammerMessage> Messages
        {
            get;set;
        }
        public List<YammerImageFile> Files { get; set; }
        public List<Topic> Topics { get; set; }
        
    }
    class Topic
    {
        public string Label { get; set; }
        public double Score { get; set; }
    }
    class YammerMessage
    {
        public string Id { get; set; }
        public string Body { get; set; }
        public double SentimentScore { get; set; }
        public List<string> KeyPhrases { get; set; }
    }
    class YammerImageFile
    {
        public string Id { get; set; }
        public string Download_Url { get; set; }
        public string Name { get; set; }
        public List<string> Tags { get; set; }
        public bool IsAdultContent { get; set;}
        public bool IsRacyContent { get; set; }
    }
    class Program
    {
        static string ComputerVisionKey = ConfigurationManager.AppSettings["ComputerVisionKey"];
        static string TextAnalyticsKey = ConfigurationManager.AppSettings["TextAnalyticsKey"];
        static string YammerToken = ConfigurationManager.AppSettings["YammerToken"];

        static string YammerGroupId = ConfigurationManager.AppSettings["YammerGroupId"];
        static void Main(string[] args)
        {            
            AsyncContext.Run(() => MainAsync(args));
        }
        static async void MainAsync(string[] args)
        {            
            YammerGroup group = new YammerGroup();
            //retrieving the conversations of a given group
            List<YammerMessage> messages = GetYammerMessages(YammerGroupId);
            //preparing dataset for APIs
            byte[] data = GetAPIData(messages);
            //performing the keyPhrase analysis against the Yammer messages
            await KeyPhraseOrSentiment(data, messages);
            //performing the sentiment analysis against the Yammer messages
            await KeyPhraseOrSentiment(data, messages, true);
            group.Messages = messages;
            //retrieving the images of that group
            List<YammerImageFile> files = await GetImageFiles(YammerGroupId);
            await DonwloadAndAnalyzeYammerFile(files);
            group.Files = files;
            Console.WriteLine("---------------TOPIC DETECTION-----------------");
            if (await GetGroupTopics(data, group))
            {                
                var MostDiscussedTopics = group.Topics.OrderByDescending(t => t.Score).Take(5);
                foreach (var MostDiscussedTopic in MostDiscussedTopics)
                {
                    Console.WriteLine("{0} {1}", 
                        MostDiscussedTopic.Score, 
                        MostDiscussedTopic.Label);
                }
            }
            else
            {
                Console.WriteLine("Toppic Detection failed");
            }
            Console.WriteLine("------------------SENSITIVE CONTENT---------------");
            var TrickyImages = group.Files.Where(f => f.IsAdultContent == true || f.IsRacyContent == true);
            foreach (var TrickyImage in TrickyImages)
            {
                Console.WriteLine("{0} {1}",
                    TrickyImage.Name,
                    (TrickyImage.Tags != null) ? 
                        string.Join("/", TrickyImage.Tags) : String.Empty);
            }            
            Console.WriteLine("----------------NEGATIVE MESSAGES--------------");
                    var NegativeMessages = group.Messages.Where(m => m.SentimentScore < 0.5).Take(5);
                    var PositiveMessages = group.Messages.Where(m => m.SentimentScore >= 0.5).Take(5);
            foreach (var NegativeMessage in NegativeMessages)
            {
                Console.WriteLine("ID : {0} Score : {1} KP : {2}",
                    NegativeMessage.Id,
                    NegativeMessage.SentimentScore,
                    (NegativeMessage.KeyPhrases!=null) ?
                      string.Join("/", NegativeMessage.KeyPhrases) : string.Empty);
            }
            Console.WriteLine("--------------POSITIVE MESSAGES--------------");            
            foreach (var PositiveMessage in PositiveMessages)
            {
                Console.WriteLine(
                    "ID : {0} Score : {1} KP : {2}",
                    PositiveMessage.Id,
                    PositiveMessage.SentimentScore,
                    (PositiveMessage.KeyPhrases!=null) ? 
                        string.Join("/", PositiveMessage.KeyPhrases):String.Empty);
            }
        }
        
        static async Task DonwloadAndAnalyzeYammerFile(List<YammerImageFile> files)
        {
            HttpClient cli = new HttpClient();
            cli.DefaultRequestHeaders.Add("Authorization", YammerToken);
            //watch out : with the free pricing tier, no more than 20 calls per minute.
            foreach(YammerImageFile file in files)
            {
                using (Stream image =
                    await cli.GetStreamAsync(
                        string.Format(
                            "https://www.yammer.com/api/v1/uploaded_files/{0}/download",
                            file.Id)))
                {
                    VisionServiceClient vscli = new VisionServiceClient(ComputerVisionKey);
                    VisualFeature[] features = new VisualFeature[] { VisualFeature.Adult, VisualFeature.Tags };
                    AnalysisResult result = await vscli.AnalyzeImageAsync(image, features);
                    file.IsAdultContent = result.Adult.IsAdultContent;
                    file.IsRacyContent = result.Adult.IsRacyContent;
                    if (result.Tags.Count() > 0)
                    {
                        List<string> tags = new List<string>();
                        var EligibleTags = result.Tags.Where(t => t.Confidence >= 0.5);
                        if (EligibleTags != null && EligibleTags.Count() > 0)
                        {
                            foreach (Tag EligibleTag in EligibleTags)
                            {
                                tags.Add(EligibleTag.Name);
                            }
                        }
                        file.Tags = tags;
                    }
                }
            }           
        }
        
        static async Task<List<YammerImageFile>> GetImageFiles(string groupid)
        {           
            List<YammerImageFile> YammerFiles = new List<YammerImageFile>();
            HttpClient cli = new HttpClient();
            cli.DefaultRequestHeaders.Add("Authorization", YammerToken);            
            HttpResponseMessage resp = 
                await cli.GetAsync(
                    string.Format(
                        "https://www.yammer.com/api/v1/uploaded_files/in_group/{0}.json?content_class=images",
                        groupid));
            
            JObject files = JObject.Parse(await resp.Content.ReadAsStringAsync());
            foreach(var file in files["files"])
            {
                YammerFiles.Add(new YammerImageFile
                {
                    Id = file["id"].ToString(),
                    Download_Url = file["download_url"].ToString(),
                    Name = file["name"].ToString()
                });
            }
            return YammerFiles;
        }
        static byte [] GetAPIData(List<YammerMessage> messages)
        {
            StringBuilder s = new StringBuilder();
            s.Append("{\"documents\":[");
            foreach (var message in messages)
            {
                var json = JsonConvert.SerializeObject(message.Body, new JsonSerializerSettings
                {
                    StringEscapeHandling = StringEscapeHandling.EscapeNonAscii
                });                        
                s.AppendFormat("{{\"id\":\"{0}\",\"text\":{1}}},", message.Id,  json);                        
            }      
                  
            return Encoding.UTF8.GetBytes(
                string.Concat(
                    s.ToString().Substring(0, s.ToString().Length - 1), "]}"));
        }
        static List<YammerMessage> GetYammerMessages(string groupid)
        {
            List<YammerMessage> ReturnedMessages = new List<YammerMessage>();
            bool FullyParsed = false;
            string LastMessageId = "";                

            HttpWebRequest req = HttpWebRequest.Create(
                "https://www.yammer.com/api/v1/messages/in_group/" +
                groupid + ".json?threaded=true")
                    as HttpWebRequest;
            req.Headers.Add("Authorization", YammerToken);
            req.Accept = "application/json; odata=verbose";
                    int MessageCount = 1;     
            while (!FullyParsed && MessageCount < 105)
            {
                using (StreamReader sr =
                        new StreamReader(req.GetResponse().GetResponseStream()))
                {                    
                    JObject resp = JObject.Parse(sr.ReadToEnd());
                    JArray messages = JArray.Parse(resp["messages"].ToString());                    
                    LastMessageId = "";
                    foreach (var message in messages)
                    {
                     
                        if(!string.IsNullOrEmpty(message["body"]["parsed"].ToString()))
                        {
                                    MessageCount++;
                            ReturnedMessages.Add(new YammerMessage
                            {
                                Id = message["id"].ToString(),
                                SentimentScore = -1,
                                Body = message["body"]["parsed"].ToString()
                            });
                        }                                            
                        LastMessageId = message["id"].ToString();                                         
                    }                    
                }
               
                if(string.IsNullOrEmpty(LastMessageId))
                {
                    FullyParsed = true;
                }
                else
                {                   
                    req = HttpWebRequest.Create(
                        "https://www.yammer.com/api/v1/messages/in_group/" +
                        groupid + ".json?threaded=true&older_than=" +
                        LastMessageId) as HttpWebRequest;
                    req.Headers.Add("Authorization", YammerToken);
                    req.Accept = "application/json; odata=verbose";
                }
            }
            return ReturnedMessages;
        }
        
        static async Task<bool> GetGroupTopics(byte[] data,YammerGroup group)
        {
            group.Topics = new List<Topic>();
            HttpClient cli = new HttpClient();
            cli.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", TextAnalyticsKey);
            var response = await cli.PostAsync(                
                "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/Topics/?minDocumentsPerWord=5",
                new ByteArrayContent(data));
            var OperationLocation = response.Headers.GetValues("Operation-Location").First();
            while(true)//should implement timeout
            {
                JObject documents = 
                        JObject.Parse(await GetTopicResult(cli, OperationLocation));
                string status = documents["status"].ToString().ToLowerInvariant();
                if (status == "succeeded")
                {
                    JArray topics = JArray.Parse(
                        documents["operationProcessingResult"]["topics"].ToString());
                    foreach (var topic in topics)
                    {
                        group.Topics.Add(new Topic
                        {
                            Label = topic["keyPhrase"].ToString(),
                            Score = Convert.ToDouble(topic["score"])
                        });
                    }
                    return true;
                }
                else if (status == "failed")
                    return false;
                else
                {
                    Thread.Sleep(60000);
                }                
            }
            
        }
        static async Task<string> GetTopicResult(HttpClient client, string uri)
        {
            var response = await client.GetAsync(uri);
            return await response.Content.ReadAsStringAsync();
        }
        static async Task KeyPhraseOrSentiment(byte[] data, List<YammerMessage> messages, bool sentiment=false)
        {            
            HttpClient cli = new HttpClient();
            cli.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", TextAnalyticsKey);
            var uri = (sentiment == true) ? 
                    "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment" : 
                    "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases";
            var response =
                await cli.PostAsync(
                    uri,
                    new ByteArrayContent(data));
            JObject analysis = JObject.Parse(response.Content.AsString());
            foreach (var document in analysis["documents"])
            {
                var TargetMessage = 
                        messages.Where(
                            m => m.Id.Equals(document["id"].ToString()))
                            .SingleOrDefault();
                if (TargetMessage != null)
                {
                    if(sentiment)
                        TargetMessage.SentimentScore =
                                Convert.ToDouble(document["score"]);
                    else
                    {
                        JArray KeyPhrases = 
                                JArray.Parse(document["keyPhrases"].ToString());
                        List<string> kp = new List<string>();
                        foreach(var KeyPhrase in KeyPhrases)
                        {
                            kp.Add(KeyPhrase.ToString());
                        }
                        TargetMessage.KeyPhrases = kp;
                    }                       
                }
            }
        }       
        
    }
}
