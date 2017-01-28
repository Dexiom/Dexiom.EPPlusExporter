using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExporterDemo
{
    public class Employee
    {
        public Employee()
        {
            
        }

        public Employee(Bogus.Person person)
        {
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
            ChangeInPocket = bogusRandom.Double(0, 5);
        }

        public string UserName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }

        public int LuckyNumber { get; set; }
        public double ChangeInPocket { get; set; }

        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime DateOfBirth { get; set; }

        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime DateHired { get; set; }

        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime DateContractEnd { get; set; }
    }
}
