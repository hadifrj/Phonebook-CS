using System.Collections.Generic;

namespace Phonebook
{
    public class Customer
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string PostalCode { get; set; }
        public List<Phone> Phones { get; set; } = new List<Phone>();
        public List<Person> Persons { get; set; } = new List<Person>();
    }

    public class Phone
    {
        public int Id { get; set; }
        public string Label { get; set; }
        public string Number { get; set; }
    }

    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Role { get; set; }
        public string Phone { get; set; }
    }
}
