using System;
using MyntraExcelAddin.Constant;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;

namespace MyntraExcelAddin.SystemObjects
{
    public class ExternalServiceMessenger
    {
        public HttpClient httpClient;
        public ExternalServiceMessenger()
        {
            httpClient = new HttpClient();
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
                        System.Windows.Forms.MessageBox.Show(ex.InnerException.InnerException.Message + "\n\nFailed to Set Drop Downs", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (ex.InnerException is WebException)
                        System.Windows.Forms.MessageBox.Show(ex.InnerException.Message + "\n\nFailed to Set Drop Downs", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (ex is HttpRequestException)
                        System.Windows.Forms.MessageBox.Show(ex.Message + "\n\nFailed to Set Drop Downs", "External Service Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return ex is HttpRequestException;
                });

                return null;
            }
        }
    }
}
