using System.Text;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Data;

namespace OpenXml_SpreadSheet
{
    public class Api
    {
        public static async Task<Dictionary<string, string>?> GetData(string internationlcode)
        {
            using var client = new HttpClient();

            var userName = AuthorizationParameters.Username;
            var passwd = AuthorizationParameters.Password;
            var url = AuthorizationParameters.Uri + internationlcode;

            var authToken = Encoding.ASCII.GetBytes($"{userName}:{passwd}");

            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Basic", Convert.ToBase64String(authToken));

            HttpResponseMessage response = await client.GetAsync(url);

            string content = await response.Content.ReadAsStringAsync();
            // string myResult = $"{{\"operationResult\":1,\"operation\":\"MemberStatusInquery\",\"responseDate\":\"2023-11-07 08:46:22\",\"personCode\":\"12586933\",\"lastName\":\"ايمانيان مزجين\",\"firstName\":\"صديقه\",\"areaCode\":\"\",\"areaName\":\"زنجان ناحيه - 2\",\"fatherName\":\"عربعلي\",\"nationalCode\":\"1639095349\",\"idNumber\":\"153\",\"issuePlace\":\"خلخال\",\"birthDate\":\"580401\",\"sex\":\"زن\",\"marriage\":\"متاهل\",\"employmentType\":\"رسمي\",\"employmentDate\":\"13800701\",\"totalState\":\"7\",\"settlemntAccountNumber\":\"0105397482009\",\"settlemntBank\":\"ملي\",\"membershipDate\":\"13830630\",\"countTransaction\":\"224\",\"lastPercent\":\"3\"}}";

            Dictionary<string, string>? dictionary =
                JsonConvert.DeserializeObject<Dictionary<string, string>?>(content);

            if (dictionary != null)
            {

                bool result = dictionary.TryGetValue("areaCode", out string? valueFromRegions);
                if (result)
                {
                    if (valueFromRegions != null)
                    {
                        try
                        {
                            DataRow dtrow = await DataExtraction.GetRegionRow(valueFromRegions);
                            if (dtrow != null)
                            {
                                dictionary?.Add("centerName", dtrow["CenterName"].ToString() ?? "");
                                dictionary?.Add("cneterId", dtrow["CenterId"].ToString() ?? "");
                            }
                            else
                            {
                                dictionary?.Add("centerName", "");
                                dictionary?.Add("cneterId", "");
                            }
                        }
                        catch (System.IO.FileNotFoundException ex)
                        {
                            Console.WriteLine("File not found: " + ex.Message);
                            dictionary?.Add("centerName", "");
                            dictionary?.Add("cneterId", "");
                        }
                    }
                    else
                    {
                        dictionary?.Add("centerName", "");
                        dictionary?.Add("cneterId", "");
                    }
                }
                else
                {
                    dictionary?.Add("centerName", "");
                    dictionary?.Add("cneterId", "");
                }
                return dictionary;

            }
            else { return null; }
        }
    }
}
