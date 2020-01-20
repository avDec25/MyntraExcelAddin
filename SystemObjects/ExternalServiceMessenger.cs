using System;
using System.Collections.Generic;
using MyntraExcelAddin.Constant;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using MyntraExcelAddin.Entity;
using System.Text;

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

        public void UpdateHandovers(List<Handover> handoverlist)
        {
            string payload = JsonConvert.SerializeObject(handoverlist);
            System.Diagnostics.Debug.WriteLine(payload);
            using (StringContent content = new StringContent(payload, Encoding.UTF8, MediaType))
            {
                Uri uri = new Uri(Addin.ServiceBaseURL + "handover/update");
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
                Uri uri = new Uri(Addin.ServiceBaseURL + "determine/bmtarget");
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

                    // by calling .Result you are synchronously reading the result
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
