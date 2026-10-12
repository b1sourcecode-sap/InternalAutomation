//using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Routing;


namespace IRDSync
{
    [System.Web.Http.Route("api/[controller]")]
    [ApiController]
    public class CustomerMasterController : ControllerBase
    {
        private readonly string serviceLayerBaseURL = "https://172.16.1.201:50000/b1s/v1/";
        private readonly string companyDB = "SUJALDAIRYLIVE";

        // Modify this method to handle Single Sign-On (SSO)
        private async Task<string> SsoLogin()
        {
            // In SSO, you may not need to pass credentials explicitly if SAP B1 handles it via user session.
            string loginURL = serviceLayerBaseURL + "Login";
            var client = new HttpClient();

            var loginBody = new
            {
                CompanyDB = companyDB,
                UserName = "manager",
                Password = "sap@9921", // Replace with actual SSO details if required
                // You can skip username and password in some SSO setups
            };

            StringContent content = new StringContent(JObject.FromObject(loginBody).ToString(), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await client.PostAsync(loginURL, content);

            if (response.IsSuccessStatusCode)
            {
                string responseBody = await response.Content.ReadAsStringAsync();
                var json = JObject.Parse(responseBody);
                return json["SessionId"].ToString();  // Return the session ID for further requests
            }
            else
            {
                throw new Exception("Failed to log in to SAP B1 Service Layer.");
            }
        }

        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("ExportCustomerData")]
        public async Task<IActionResult> ExportCustomerData()
        {
            try
            {
                // Get Session ID from SSO login
                string sessionId = await SsoLogin();

                // Create HttpClient for SAP B1 Service Layer request
                var client = new HttpClient();
                client.DefaultRequestHeaders.Add("Cookie", $"B1SESSION={sessionId}");

                // Construct the URL to get the customer master data
                string queryURL = serviceLayerBaseURL + "BusinessPartners?$filter=CardType eq 'C'"; // Modify filter as needed

                HttpResponseMessage response = await client.GetAsync(queryURL);

                if (response.IsSuccessStatusCode)
                {
                    string jsonData = await response.Content.ReadAsStringAsync();
                    var customerData = JObject.Parse(jsonData);

                    // Return customer data as JSON response
                    return Ok(customerData);
                }
                else
                {
                    return StatusCode((int)response.StatusCode, "Failed to fetch customer data from SAP B1.");
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
        }
    }
}