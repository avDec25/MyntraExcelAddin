﻿using System;
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

        internal double RetrieveBMTargetValue(string brand, string articletype, string gender, bool repeated)
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

        public DropDownData GetDropDownData()
        {
            string responseString = "";
            try 
            {
                var response = httpClient.GetAsync(Addin.ServiceBaseURL + "dropdown?names=" +
                "brand,impression,articletype,gender,bodycode,cluster,color,subcategory,fpt,sizetype,datasource,source").Result;
                
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
