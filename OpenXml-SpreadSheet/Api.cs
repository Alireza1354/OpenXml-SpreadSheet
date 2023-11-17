﻿using System.Text;
using Newtonsoft.Json;
using System.Net.Http.Headers;

namespace OpenXml_SpreadSheet
{
    public class Api
    {
        public static async Task<Dictionary<string, string>> GetData(string internationlcode)
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
            //string myResult = $"{{\"operationResult\":1,\"operation\":\"MemberStatusInquery\",\"responseDate\":\"2023-11-07 08:46:22\",\"personCode\":\"12586933\",\"lastName\":\"ايمانيان مزجين\",\"firstName\":\"صديقه\",\"areaCode\":\"5702\",\"areaName\":\"زنجان ناحيه - 2\",\"fatherName\":\"عربعلي\",\"nationalCode\":\"1639095349\",\"idNumber\":\"153\",\"issuePlace\":\"خلخال\",\"birthDate\":\"580401\",\"sex\":\"زن\",\"marriage\":\"متاهل\",\"employmentType\":\"رسمي\",\"employmentDate\":\"13800701\",\"totalState\":\"7\",\"settlemntAccountNumber\":\"0105397482009\",\"settlemntBank\":\"ملي\",\"membershipDate\":\"13830630\",\"countTransaction\":\"224\",\"lastPercent\":\"3\"}}";


            //var res = new Dictionary<string, string>()
            //                            {
            //                                {".cbfmd", "bfres" },
            //                                {".cbfa",  "bfres" },
            //                                {".cbfsa", "bfres" },
            //                                {".cbntx", "bntx" },
            //                            };


            Dictionary<string, string> dictionary = new Dictionary<string, string>();

            var result = JsonConvert.DeserializeObject<Dictionary<string, string>>(content);

            if (result != null)
            {
                dictionary = result;
            }
            else
            {
                dictionary.Add("", "");
            }
            return dictionary;

        }
    }
}
