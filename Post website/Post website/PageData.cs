using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Post_website
{
    class PageData
    {
        public string LeadId, Make, Year, Model, InsuranceCompany, FirstName, LastName, Gender, ResidenceType, BirthDay, BirthMonth, BirthYear, MaritalStatus,
            creditRetain, Address, Phone, ZipCode, Email, Vertical, AnnualMiles,SourceId,BirthDate;
        Random rand = new Random();
        //public static int GetRandomNumber(int min, int max)
        //{
        //    using (RNGCryptoServiceProvider rg = new RNGCryptoServiceProvider())
        //    {
        //        byte[] rno = new byte[5];
        //        rg.GetBytes(rno);
        //        int randomvalue = BitConverter.ToInt32(rno, 0);
        //        randomvalue = (randomvalue % (max - min+1)) + min;
        //        return randomvalue;
        //    }
        //}

        public int GetRandomNumber(int min, int max)
        {
            return rand.Next(min,max);
        }
        public string getRandomGender()
        {
            int x = GetRandomNumber(1, 117)%2;
            return (x == 1 ? "Male" : "Female");
        }
        public string getRandomResidenceType()
        {
            int x = GetRandomNumber(1, 100)%2;
            return (x == 1 ? "My own house" : "I am renting");
        }
        public string getRandomMatrtalStatus()
        {
            
            int x = GetRandomNumber(1, 5);
            if (x == 1)
                return "Divorced";
            else if (x == 2)
                return "Married";
            else if (x == 3)
                return "Separated";
            else if (x == 4)
                return "Single";
            else
                return "Widowed";
        }
        public string getRandomCredit()
        {
            int x = GetRandomNumber(1, 4);
            if (x == 1)
                return "Excellent";
            else if (x == 2)
                return "Good";
            else if (x == 3)
                return "Some Problems";
            else
                return "Major Problems";
        }
        public string getRandomMiles()
        {
            int x = GetRandomNumber(1, 4);
            if (x == 1)
                return "2500";
            else if (x == 2)
                return "7500";
            else if (x == 3)
                return "12500";
            else
                return "15000";
        }
        public static string managePioint(string str)
        {
            
            int ind, strtind = 0;
            Double output;
            while (str.IndexOf('-', strtind) != -1)
            {
                ind = str.IndexOf('-', strtind);
                StringBuilder tempString = new StringBuilder(str.Substring(ind - 1, 3));
                tempString[1] = '.';
                if (Double.TryParse(tempString.ToString(), out output))
                {
                    StringBuilder someString = new StringBuilder(str);
                    someString[ind] = '.';
                    str = someString.ToString();
                }

                strtind = ind + 1;
            }
            return str;
        }
        public void setData(List<string> list)
        {
            LeadId = list.ElementAt(0);
            Year = list.ElementAt(1);
            Make = list.ElementAt(2);

            Model = managePioint( list.ElementAt(3));
            

            InsuranceCompany = list.ElementAt(4);
            if(InsuranceCompany=="Geico")
            {
                InsuranceCompany = "GEICO";
            }

            FirstName = list.ElementAt(5);
            LastName = list.ElementAt(6);

            Gender = list.ElementAt(7);
            if(Gender== "Seleted Randomly")
            {
                Gender = getRandomGender();
            }

            ResidenceType = list.ElementAt(8);
            if(ResidenceType== "Seleted Randomly")
            {
                ResidenceType = getRandomResidenceType();
            }

            BirthDate = list.ElementAt(9);
            double d = double.Parse(BirthDate);
            BirthDate = DateTime.FromOADate(d).ToString();
            string[] dates = BirthDate.Split('/');
            BirthMonth = dates.ElementAt(0);
            BirthDay = dates.ElementAt(1);
            BirthYear = dates.ElementAt(2).Substring(0,4);
            BirthMonth = getMonth(BirthMonth);

            MaritalStatus = list.ElementAt(10);
            if(MaritalStatus== "Seleted Randomly")
            {
                MaritalStatus = getRandomMatrtalStatus();
            }

            creditRetain = list.ElementAt(11);
            if(creditRetain== "Seleted Randomly")
            {
                creditRetain = getRandomCredit();
            }

            Address = list.ElementAt(12);
            ZipCode = list.ElementAt(13);

            Phone = list.ElementAt(14);
            Phone = Phone.Insert(0, "(");
            Phone = Phone.Insert(4, ") ");
            Phone = Phone.Insert(9, "-");

            
            Email = list.ElementAt(15);
            Vertical = list.ElementAt(16);
            AnnualMiles = list.ElementAt(17);
            if(AnnualMiles== "Seleted Randomly")
            {
                AnnualMiles = getRandomMiles();
            }
            AnnualMiles=AnnualMiles.Insert((AnnualMiles.Length - 3),",");
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
