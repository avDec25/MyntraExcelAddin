using System;
using System.Collections.Generic;
using MyntraExcelAddin.Constant;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using MyntraExcelAddin.Entity;
using System.Text;
using System.Diagnostics;

namespace MyntraExcelAddin.SystemObjects
{
    public class ExternalServiceMessenger : IDisposable
    {
        private const string MediaType = "application/json";
        public HttpClient httpClient;
        public ExternalServiceMessenger()
        {
            httpClient = new HttpClient();            
        }

        public void Dispose()
        {
            httpClient.Dispose();
        }

        public List<Handover> GetHandovers(string hids)
        {
            string responseString = "";
            try
            {
                var response = httpClient.GetAsync(Addin.ServiceBaseURL + "addin/cthandover?hids=" + hids).Result;

                if (response.IsSuccessStatusCode)
                {
                    var responseContent = response.Content;
                    responseString = responseContent.ReadAsStringAsync().Result;
                }
                return JsonConvert.DeserializeObject<List<Handover>>(responseString);
            }
            catch (AggregateException ae)
            {
                ae.Handle(ex => {
                    if (ex.InnerException.InnerException is SocketException)
                        MessageBox.Show(ex.InnerException.InnerException.Message + "\n\nFailed to Download Handovers", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (ex.InnerException is WebException)
                        MessageBox.Show(ex.InnerException.Message + "\n\nFailed to Download Handovers", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (ex is HttpRequestException)
                        MessageBox.Show(ex.Message + "\n\nFailed to Download Handovers", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return ex is HttpRequestException;
                });

                return null;
            }
        }

        public Tuple<List<HandoverTableView>,uint> GetFilteredHandovers(Dictionary<string, List<string>> query, uint page, uint size)
        {
            List<HandoverTableView> handoverlist = new List<HandoverTableView>();

            JObject jResponse = new JObject();

            string payload = JsonConvert.SerializeObject(query);
            System.Diagnostics.Debug.WriteLine(payload);
            using (StringContent content = new StringContent(payload, Encoding.UTF8, MediaType))
            {
                Uri uri = new Uri(Addin.ServiceBaseURL + "handover/filter?page=" + page + "&size="+ size);

                using (var resp = httpClient.PostAsync(uri, content).Result)
                {
                    if (resp.StatusCode == HttpStatusCode.OK)
                    {
                        jResponse = JObject.Parse(resp.Content.ReadAsStringAsync().Result);
                    }
                }
            }

            handoverlist = JsonConvert.DeserializeObject<List<HandoverTableView>>(jResponse.GetValue("content").ToString());
            uint totalpages = jResponse.GetValue("totalPages").ToObject<uint>();
            return Tuple.Create(handoverlist,totalpages);
        }

        public void UpdateHandovers(List<Handover> handoverlist)
        {
            string payload = JsonConvert.SerializeObject(handoverlist);
            System.Diagnostics.Debug.WriteLine(payload);
            using (StringContent content = new StringContent(payload, Encoding.UTF8, MediaType))
            {
                Uri uri = new Uri(Addin.ServiceBaseURL + "addin/handover/update");
                using (var resp = httpClient.PostAsync(uri, content).Result)
                {
                    resp.EnsureSuccessStatusCode();                    
                }
            }
        }

        public List<long> SubmitHandovers(List<Handover> handoverlist)
        {
            List<long> ans = new List<long>();
            HandoverCreateRequestToService request = new HandoverCreateRequestToService();
            request.uploadType = "JSON";
            request.allCtData = handoverlist;

            string payload = JsonConvert.SerializeObject(request);
            System.Diagnostics.Debug.WriteLine(payload);
            using (StringContent content = new StringContent(payload, Encoding.UTF8, MediaType))
            {
                Uri uri = new Uri(Addin.ServiceBaseURL + "cthandover");
                using (var resp = httpClient.PutAsync(uri, content).Result)
                {
                    //resp.EnsureSuccessStatusCode();         
                    if(resp.StatusCode == HttpStatusCode.Created)
                    { 
                        ans = JsonConvert.DeserializeObject<List<long>>(resp.Content.ReadAsStringAsync().Result);
                    }
                }
            }
            return ans;
        }

        internal double RetrieveBMTargetValue(string brand, string articletype, string gender, string repeated)
        {
            double bmtval = 0.0;
            var bmtreq = new BMTargetRequestToService(brand, articletype, gender, repeated);
            string payload = JsonConvert.SerializeObject(bmtreq);
            System.Diagnostics.Debug.WriteLine(payload);
            using (StringContent content = new StringContent(payload, Encoding.UTF8, MediaType))
            {
                Uri uri = new Uri(Addin.ServiceBaseURL + "addin/determine/bmtarget");
                using (var resp = httpClient.PostAsync(uri, content).Result)
                {
                    resp.EnsureSuccessStatusCode();
                    bmtval = double.Parse(resp.Content.ReadAsStringAsync().Result);
                }
            }
            return bmtval;
        }

        public List<ValidatorResult> GetValidationInfo(List<Handover> handoverlist)
        {
            string validationResult;

            string payload = JsonConvert.SerializeObject(handoverlist);
            using (StringContent content = new StringContent(payload, Encoding.UTF8, MediaType))
            {
                Uri uri = new Uri(Addin.ServiceBaseURL + "validator");
                using (var resp = httpClient.PostAsync(uri, content).Result)
                {
                    resp.EnsureSuccessStatusCode();
                    validationResult = resp.Content.ReadAsStringAsync().Result;
                }
            }
            List<ValidatorResult> reportcard = JsonConvert.DeserializeObject<List<ValidatorResult>>(validationResult);
            return reportcard;
        }

        public DropDownData GetDropDownData(string dropdowns)
        {
            string responseString = "";
            try 
            {
                var response = httpClient.GetAsync(Addin.ServiceBaseURL + "dropdown?names=" + dropdowns).Result;
                
                if (response.IsSuccessStatusCode)
                {
                    var responseContent = response.Content;

                    // by calling .Result we synchronously read the result
                    responseString = responseContent.ReadAsStringAsync().Result;
                }
                return JsonConvert.DeserializeObject<DropDownData>(responseString);
            }
            catch (AggregateException ae)
            {
                ae.Handle(ex => {
                    if (ex.InnerException.InnerException is SocketException)
                        MessageBox.Show(ex.InnerException.InnerException.Message + "\n\nFailed to Set Drop Downs", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (ex.InnerException is WebException)
                        MessageBox.Show(ex.InnerException.Message + "\n\nFailed to Set Drop Downs", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (ex is HttpRequestException)
                        MessageBox.Show(ex.Message + "\n\nFailed to Set Drop Downs", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return ex is HttpRequestException;
                });

                return null;
            }
        }
    }
}
