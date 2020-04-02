using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace XMLDocs.Models
{
    public class Student
    {
         public static string HeaderRow = $"{nameof(Student.StudentId)},{nameof(Student.FirstName)},{nameof(Student.LastName)},{nameof(Student.DateOfBirth)},{nameof(Student.ImageData)}";

        public string StudentId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }

        private string _DateOfBirth;
        public string DateOfBirth
        {
            get { return _DateOfBirth; }
            set
            {
                _DateOfBirth = value;

                //Convert DateOfBirth to DateTime
                DateTime dtOut;
                DateTime.TryParse(_DateOfBirth, out dtOut);
                DateOfBirthDT = dtOut;
            }
        }

        public DateTime DateOfBirthDT { get; set; }
        public string ImageData { get; set; }
        public string MyRecord { get; set; }
        public string AbsoluteUrl { get; set; }
        public string Directory { get; set; }
      //public string InfoCSVPath { get { return (Constants.FTP.BaseUrl + "/" + Directory + "/" + Constants.Locations.InfoFile); } }
      //public string MyImagePath { get { return (Constants.FTP.BaseUrl + "/" + Directory + "/" + Constants.Locations.ImageFile); } }

        public string FullPathUrl
        {
            get
            {
                return AbsoluteUrl + "/" + Directory;
            }
        }

        public List<string> Exceptions { get; set; } = new List<string>();
        
        public void FromCSV(string csvdata)
        {
            string[] data = csvdata.Split(",", StringSplitOptions.None);
            try
            {
                StudentId = data[0];
                FirstName = data[1];
                LastName = data[2];
                DateOfBirth = data[3];
                ImageData = data[4];
            }
            catch (Exception e)
            {
                Exceptions.Add(e.Message);
            }
        }

        public bool IsValidImageFile(string imageFile)
        {
            try
            {
                // the using is important to avoid stressing the garbage collector
                using (var test = System.Drawing.Image.FromFile(imageFile))
                {
                    // image has loaded and so is fine
                    return true;
                }
            }
            catch
            {
                // technically some exceptions may not indicate a corrupt image, but this is unlikely to be an issue
                return false;
            }
        }

        public void FromDirectory(string directory)
        {
            Directory = directory;

            if (String.IsNullOrEmpty(directory.Trim()))
            {
                return;
            }

            string[] data = directory.Trim().Split(" ", StringSplitOptions.None);

            StudentId = data[0];
            FirstName = data[1];
            LastName = data[2];
        }



        public string ToCSV()
        {
            string result = $"{StudentId},{FirstName},{LastName},{DateOfBirthDT.ToShortDateString()},{ImageData}";
            return result;
        }

        public override string ToString()
        {
            string result = $"{StudentId} {FirstName} {LastName}";
            return result;
        }

        public virtual int age
        {
            get
            {
                DateTime Now = DateTime.Now;
                int Years = new DateTime(DateTime.Now.Subtract(DateOfBirthDT).Ticks).Year - 1;
                DateTime PastYearDate = DateOfBirthDT.AddYears(Years);
                int Months = 0;
                for (int i = 1; i <= 12; i++)
                {
                    if (PastYearDate.AddMonths(i) == Now)
                    {
                        Months = i;
                        break;
                    }
                    else if (PastYearDate.AddMonths(i) >= Now)
                    {
                        Months = i - 1;
                        break;
                    }
                }
                int Days = Now.Subtract(PastYearDate.AddMonths(Months)).Days;
                int Hours = Now.Subtract(PastYearDate).Hours;
                int Minutes = Now.Subtract(PastYearDate).Minutes;
                int Seconds = Now.Subtract(PastYearDate).Seconds;
                return Years;
            }
        }






        //public int age
        //{
        //    get
        //    {
        //        return (DateTime.Now - this.DateOfBirthDT).Days / 365;
        //    }
        //}

        //public DateTime DateOfBirthDT { get; set; }
        //public int Age
        //{
        //    get
        //    {
        //        //DateTime now = DateTime.Now;
        //        //int age = now.Year - DateOfBirth.Year;
        //        //if (now < DateOfBirth.AddYears(age)) age--;
        //        //return age;

        //        DateTime today = DateTime.Today;
        //        //DateTime bday = Convert.ToDateTime(student.DateOfBirthDT);
        //        int age = today.Year - DateOfBirthDT.Year;
        //        if (bday > today.AddYears(-age)) age--;
        //        return age;
        //    }
        //{

        //       // Console.WriteLine("Average age", )
        //}

    }
}
