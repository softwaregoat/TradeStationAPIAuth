using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace TradeStationAPIAuth2
{
    class TradeStationWebApi
    {
        private string Key { get; set; }
        private string Secret { get; set; }
        private string Host { get; set; }
        private string RedirectUri { get; set; }
        private AccessToken Token { get; set; }
        string path = Directory.GetParent(Application.ExecutablePath).ToString();
        public TradeStationWebApi(string key, string secret, string host, string redirecturi)
        {
            this.Key = key;
            this.Secret = secret;
            this.RedirectUri = redirecturi;

            this.Host = host;

            // Disable Tls 1.0 and use Tls 1.2
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
            // these two lines are only needed if .net 4.5 is not installed
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            
            string refresh_token = "";
            try
            {
                refresh_token = System.IO.File.ReadAllText(path + "\\RefreshToken.txt");
                this.Token = GetAccessToken(refresh_token, "");
            }
            catch(Exception ex)
            {
                ErrorLog(ex);
            }
            if (refresh_token == "")
            {
                this.Token = GetAccessToken(GetAuthorizationCode());
                System.IO.File.WriteAllText(path + "\\RefreshToken.txt", this.Token.refresh_token);
            }
            try
            {
                if (this.Token != null)
                {
                    System.IO.File.WriteAllText(path + "\\Token.txt", this.Token.access_token);
                }
                else
                {
                    MessageBox.Show("Please check API credential");
                    return;
                }
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
            }
        }

        private string GetAuthorizationCode()
        {
            var url = string.Format("{0}/{1}", this.Host,
                    string.Format(
                        "authorize?client_id={0}&response_type=code&redirect_uri={1}",
                        this.Key,
                        "http://localhost:1125/"));
            System.IO.File.WriteAllText(path + "\\AuthUrl.txt", url);
            MessageBox.Show("Open Auth Url in your browser from AuthUrl.txt file");
            using (var listener = new HttpListener())
            {
                listener.Prefixes.Add(this.RedirectUri);
                listener.Start();

                var context = listener.GetContext();
                var req = context.Request;
                var res = context.Response;

                var responseString = "<html><body><script>window.open('','_self').close();</script></body></html>";
                var buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
                res.ContentLength64 = buffer.Length;
                var output = res.OutputStream;
                output.Write(buffer, 0, buffer.Length);
                output.Close();

                listener.Stop();
                return req.QueryString.Get("code");
            }
        }

        private AccessToken GetAccessToken(string authcode)
        {

            var request = WebRequest.Create(string.Format("{0}/security/authorize", this.Host)) as HttpWebRequest;
            request.Method = "POST";
            var postData =
                string.Format(
                    "grant_type=authorization_code&code={0}&client_id={1}&redirect_uri={2}&client_secret={3}",
                    authcode,
                    this.Key,
                    this.RedirectUri,
                    this.Secret);
            var byteArray = Encoding.UTF8.GetBytes(postData);

            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = byteArray.Length;

            var dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            try
            {
                return GetDeserializedResponse<AccessToken>(request);
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        private AccessToken GetAccessToken(string refresh_token, string refresh)
        {

            var request = WebRequest.Create(string.Format("{0}/security/authorize", this.Host)) as HttpWebRequest;
            request.Method = "POST";
            var postData =
                string.Format(
                    "grant_type=refresh_token&reponse_type=token&refresh_token={0}&client_id={1}&redirect_uri={2}&client_secret={3}",
                    refresh_token,
                    this.Key,
                    this.RedirectUri,
                    this.Secret);
            var byteArray = Encoding.UTF8.GetBytes(postData);

            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = byteArray.Length;

            var dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            try
            {
                return GetDeserializedResponse<AccessToken>(request);
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }

        private static T GetDeserializedResponse<T>(WebRequest request)
        {
            var response = request.GetResponse() as HttpWebResponse;
            var receiveStream = response.GetResponseStream();
            var readStream = new StreamReader(receiveStream, Encoding.UTF8);
            var ser = new JavaScriptSerializer();
            var json = readStream.ReadToEnd();
            var scrubbedJson =
                json.Replace(
                    "\"__type\":\"EquitiesOptionsOrderConfirmation:#TradeStation.Web.Services.DataContracts\",", ""); // hack
            var deserializaed = ser.Deserialize<T>(scrubbedJson);
            response.Close();
            readStream.Close();
            return deserializaed;
        }
        internal IEnumerable<Quote> GetQuote(string symbols)
        {
            var resourceUri = new Uri(string.Format("{0}/data/quote/{1}?access_token={2}", this.Host,
                symbols, this.Token.access_token));

            var request = WebRequest.Create(resourceUri) as HttpWebRequest;
            request.Method = "GET";

            try
            {
                return GetDeserializedResponse<IEnumerable<Quote>>(request);
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        internal IEnumerable<Symbol> SymbolSuggest(string suggestText)
        {
            var resourceUri = new Uri(string.Format("{0}/{1}/{2}?access_token={3}", this.Host, "data/symbols/suggest", suggestText, this.Token.access_token));

            var request = WebRequest.Create(resourceUri) as HttpWebRequest;
            request.Method = "GET";

            try
            {
                return GetDeserializedResponse<IEnumerable<Symbol>>(request);
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        internal IEnumerable<BarChart> GetBarChart(string symbol, string interval, string unit, string barsBack, string endDate, string SessionTemplate)
        {
            var resourceUri =
                new Uri(string.Format("{0}/stream/barchart/{1}/{2}/{3}/{4}/{5}?SessionTemplate={6}&access_token={7}"
                , this.Host, symbol, interval, unit, barsBack, endDate, SessionTemplate, this.Token.access_token));

            var request = WebRequest.Create(resourceUri) as HttpWebRequest;
            request.Accept = "application/vnd.tradestation.streams+json";
            request.Method = "GET";

            try
            {
                var response = request.GetResponse() as HttpWebResponse;
                var receiveStream = response.GetResponseStream();
                var readStream = new StreamReader(receiveStream, Encoding.UTF8);
                var ser = new JavaScriptSerializer();
                var json = readStream.ReadToEnd();
                var scrubbedJson = json.Replace("END", "").Trim(); // hack
                scrubbedJson = scrubbedJson.Replace("{\"Close\"", ",{\"Close\"") + "]";
                scrubbedJson = '[' + scrubbedJson.Remove(0, 1);
                var deserializaed = ser.Deserialize<IEnumerable<BarChart>>(scrubbedJson);
                response.Close();
                readStream.Close();
                return deserializaed;
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        internal IEnumerable<BarChart> GetBarChartStartingOnDate(string symbol, string interval, string unit, string StartDate, string SessionTemplate)
        {
            var resourceUri =
                new Uri(string.Format("{0}/stream/barchart/{1}/{2}/{3}/{4}?SessionTemplate={5}&access_token={6}"
                , this.Host, symbol, interval, unit, StartDate, SessionTemplate, this.Token.access_token));
            var request = WebRequest.Create(resourceUri) as HttpWebRequest;
            request.Accept = "application/vnd.tradestation.streams+json";
            request.Method = "GET";

            try
            {
                var response = request.GetResponse() as HttpWebResponse;
                var receiveStream = response.GetResponseStream();
                var readStream = new StreamReader(receiveStream, Encoding.UTF8);
                var ser = new JavaScriptSerializer();
                var json = readStream.ReadToEnd();
                var scrubbedJson = json.Replace("END", "").Trim(); // hack
                scrubbedJson = scrubbedJson.Replace("{\"Close\"", ",{\"Close\"") + "]";
                scrubbedJson = '[' + scrubbedJson.Remove(0, 1);
                var deserializaed = ser.Deserialize<IEnumerable<BarChart>>(scrubbedJson);
                response.Close();
                readStream.Close();
                return deserializaed;
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        internal IEnumerable<BarChart> GetBarChartDateRange(string symbol, string interval, string unit, string StartDate, string EndDate, string SessionTemplate)
        {
            var resourceUri =
                new Uri(string.Format("{0}/stream/barchart/{1}/{2}/{3}/{4}/{5}?SessionTemplate={6}&access_token={7}"
                , this.Host, symbol, interval, unit, StartDate, EndDate, SessionTemplate, this.Token.access_token));
            var request = WebRequest.Create(resourceUri) as HttpWebRequest;
            request.Accept = "application/vnd.tradestation.streams+json";
            request.Method = "GET";

            try
            {
                var response = request.GetResponse() as HttpWebResponse;
                var receiveStream = response.GetResponseStream();
                var readStream = new StreamReader(receiveStream, Encoding.UTF8);
                var ser = new JavaScriptSerializer();
                var json = readStream.ReadToEnd();
                var scrubbedJson = json.Replace("END", "").Trim(); // hack
                scrubbedJson = scrubbedJson.Replace("{\"Close\"", ",{\"Close\"") + "]";
                scrubbedJson = '[' + scrubbedJson.Remove(0, 1);
                var deserializaed = ser.Deserialize<IEnumerable<BarChart>>(scrubbedJson);
                response.Close();
                readStream.Close();
                return deserializaed;
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        internal IEnumerable<BarChart> GetBarChartDaysBack(string symbol, string interval, string unit, string SessionTemplate, string daysBack, string lastDate)
        {
            var resourceUri =
                new Uri(string.Format("{0}/stream/barchart/{1}/{2}/{3}?SessionTemplate={4}&daysBack={5}&lastDate={6}&access_token={7}"
                , this.Host, symbol, interval, unit, SessionTemplate, daysBack, lastDate, this.Token.access_token));
            var request = WebRequest.Create(resourceUri) as HttpWebRequest;
            request.Accept = "application/vnd.tradestation.streams+json";
            request.Method = "GET";

            try
            {
                var response = request.GetResponse() as HttpWebResponse;
                var receiveStream = response.GetResponseStream();
                var readStream = new StreamReader(receiveStream, Encoding.UTF8);
                var ser = new JavaScriptSerializer();
                var json = readStream.ReadToEnd();
                var scrubbedJson = json.Replace("END", "").Trim(); // hack
                scrubbedJson = scrubbedJson.Replace("{\"Close\"", ",{\"Close\"") + "]";
                scrubbedJson = '[' + scrubbedJson.Remove(0, 1);
                var deserializaed = ser.Deserialize<IEnumerable<BarChart>>(scrubbedJson);
                response.Close();
                readStream.Close();
                return deserializaed;
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        internal IEnumerable<Symbol> SearchSymbols(string criteria)
        {
            var resourceUri = new Uri(string.Format("{0}/{1}/{2}?access_token={3}",
                this.Host, "data/symbols/search", criteria, this.Token.access_token));

            var request = WebRequest.Create(resourceUri) as HttpWebRequest;
            request.Method = "GET";

            try
            {
                return GetDeserializedResponse<IEnumerable<Symbol>>(request);
            }
            catch (Exception ex)
            {
                ErrorLog(ex);
                return null;
            }
        }
        private void ErrorLog(Exception ex)
        {
            string logfile = path + "\\errorlog.txt";
            if (!File.Exists(logfile))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(logfile))
                {
                    sw.WriteLine(ex.ToString());
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(logfile))
                {
                    sw.WriteLine(ex.ToString());
                }
            }
        }
    }
}
