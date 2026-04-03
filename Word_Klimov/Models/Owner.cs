using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_Klimov.Models
{
    public class Owner
    {
        public string img { get; set; } = "C:\\Users\\user\\Desktop\\флэшка\\3 курс\\Ощепков\\Практические работы\\Практическая работа №51\\Word_Klimov\\Word_Klimov\\Images\\owner.png";
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string SurName { get; set; }
        public int NumberRoom { get; set; }
        public Owner(string img, string FirstName, string LastName, string SurName, int NumberRoom)
        {
            this.img = img;
            this.FirstName = FirstName;
            this.LastName = LastName;
            this.SurName = SurName;
            this.NumberRoom = NumberRoom;
        }
    }
}
