using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorOnGitHubPages.Models
{
    public class xlsData
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string Account { get; set; }
        public string Last4SSN { get; set; }
        public string HearingOfficeWithJurisdiction { get; set; }
        public DateTime HearingScheduledDate { get; set; }
        public string HearingTime { get; set; }
        public string ALJLastName { get; set; }
        public string MedicalExpert { get; set; }
        public string VocationalExpert { get; set; }

        public string Id { get; set; }
    }
}
