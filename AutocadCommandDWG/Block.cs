using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutocadCommandDWG
{
    internal class Block
    {
        string id;
        double weight;
        int diameter;
        string fullname;

        public Block(string id, double weight, int diameter, string fullname)
        {
            this.Id = id;
            this.Weight = weight;
            this.Diameter = diameter;
            this.Fullname = fullname;
        }

        public Block(){}

        public string Id { get => id; set => id = value; }
        public double Weight { get => weight; set => weight = value; }
        public int Diameter { get => diameter; set => diameter = value; }
        public string Fullname { get => fullname; set => fullname = value; }

    }
}
