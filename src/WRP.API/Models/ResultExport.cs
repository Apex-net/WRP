using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WRP.API.Models
{
    using Newtonsoft.Json;

    public class ResultExport
    {
        //[JsonProperty(PropertyName = "result_code")]
        public int ResultCode { get; set; }

        //[JsonProperty(PropertyName = "result_message")]
        public string ResultMessage { get; set; }
    }
}