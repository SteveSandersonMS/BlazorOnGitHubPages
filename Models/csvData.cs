using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorOnGitHubPages.Models
{
    public class csvData
    {
        public string Name { get; set; }
        public string Case_Status__c { get; set; }
        public string Case_Sub_Status__c { get; set; }
        public string Confirmed_Hearing_Date__c { get; set; }
        public string Lead_Advocate1__r_Name { get; set; }
        public string Advocate_Status__c { get; set; }
    }
}
