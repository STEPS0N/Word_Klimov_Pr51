using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_Klimov.Models
{
    public class Owner
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string SurName { get; set; }
        public int NumberRoom { get; set; }
        public Owner(string FirstName, string LastName, string SurName, int NumberRoom)
        {
            this.FirstName = FirstName;
            this.LastName = LastName;
            this.SurName = SurName;
            this.NumberRoom = NumberRoom;
        }
    }
}
