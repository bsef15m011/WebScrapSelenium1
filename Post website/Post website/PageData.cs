using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Post_website
{
    class PageData
    {
        public string LeadId, Make, Year, Model, InsuranceCompany, FirstName, LastName, Gender, ResidenceType, BirthDay, BirthMonth, BirthYear, MaritalStatus,
            creditRetain, Address, Phone, ZipCode, Email, Vertical, AnnualMiles,SourceId;

        public object Datetime { get; private set; }

        public void setData(List<string> list)
        {
            LeadId = list.ElementAt(0);
            Year = list.ElementAt(1);
            Make = list.ElementAt(2);

            Model = list.ElementAt(3);
            //Model = Model.Replace("-", ".");

            InsuranceCompany = list.ElementAt(4);
            FirstName = list.ElementAt(5);
            LastName = list.ElementAt(6);
            Gender = list.ElementAt(7);
            ResidenceType = list.ElementAt(8);

            string BirthDate = list.ElementAt(9);
            double d = double.Parse(BirthDate);
            BirthDate = DateTime.FromOADate(d).ToString();
            string[] dates = BirthDate.Split('/');
            BirthMonth = dates.ElementAt(0);
            BirthDay = dates.ElementAt(1);
            BirthYear = dates.ElementAt(2).Substring(0,4);
            BirthMonth = getMonth(BirthMonth);

            MaritalStatus = list.ElementAt(10);
            creditRetain = list.ElementAt(11);
            Address = list.ElementAt(12);
            ZipCode = list.ElementAt(13);

            Phone = list.ElementAt(14);
            Phone = Phone.Insert(0, "(");
            Phone = Phone.Insert(4, ") ");
            Phone = Phone.Insert(9, "-");

            
            Email = list.ElementAt(15);
            Vertical = list.ElementAt(16);
            AnnualMiles = list.ElementAt(17);
            SourceId = list.ElementAt(18); ;
        }
        public string getMonth(string mon)
        {
            if (mon == "1")
                return "Jan";
            else if(mon == "2")
                return "Feb";
            else if(mon == "3")
                return "Mar";
            else if(mon == "4")
                return "Apr";
            else if(mon == "5")
                return "May";
            else if(mon == "6")
                return "Jun";
            else if(mon == "7")
                return "Jul";
            else if(mon == "8")
                return "Aug";
            else if(mon == "9")
                return "Sep";
            else if(mon == "10")
                return "Oct";
            else if (mon == "11")
                return "Nov";
            else
                return "Dec";
        }
    }
}
