using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bogus;

namespace EPPlusExporterDemo
{
    public class Employee
    {
        public Employee()
        {
            var person = new Person();
            UserName = person.UserName;
            FirstName = person.FirstName;
            LastName = person.LastName;
            Email = person.Email;
            Phone = person.Phone;
            DateOfBirth = person.DateOfBirth;

            var bogusRandom = new Bogus.Randomizer();
            var bogusDate = new Bogus.DataSets.Date();
            DateHired = bogusDate.Between(DateTime.Today.AddYears(-5), DateTime.Today);
            DateContractEnd = bogusDate.Between(DateTime.Today, DateTime.Today.AddYears(-5));
            LuckyNumber = bogusRandom.Int(0, 100);
            ChildrenCount = bogusRandom.Int(0, 6);
            ChangeInPocket = bogusRandom.Double(0, 5);
            CarValue = bogusRandom.Decimal(500, 40000);
        }

        public string UserName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }

        public int LuckyNumber { get; set; }

        [Display(Name = "Nb Children")]
        public int ChildrenCount { get; set; }

        public double ChangeInPocket { get; set; }
        
        [DisplayName("Car resale value")]
        public decimal CarValue { get; set; }

        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime DateOfBirth { get; set; }

        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime DateHired { get; set; }

        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime DateContractEnd { get; set; }
    }
}
