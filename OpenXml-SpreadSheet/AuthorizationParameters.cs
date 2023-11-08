using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml_SpreadSheet
{
    public static class AuthorizationParameters
    {
        public static string Username { get; set; } = $"leasing";
        public static string Password { get; set; } = $"api_1396*09*14_ls";
        public static string Uri { get; set; } = $"http://192.168.22.75:8080/tifcoRestApi/api/v1/getMemberStatusByNationalCode?nationalCode=";
    }
}
