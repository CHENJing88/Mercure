using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercure
{
    class Program
    {   
        /// <summary>
        /// la fonction principale
        /// </summary>
        /// <param name="args">Args[0] est le fichier .xls et Args[1] est le fichier .sdf</param>
        static void Main(string[] Args)
        {
            if (Args.Length != 2)
                throw new FormatException("Le nombre d'argument n'est pas satisfaire.");
            if (!File.Exists(Args[0]))
                throw new FileNotFoundException("Fichier non trouve : " + Args[0]);
            if (!File.Exists(Args[1]))
                throw new FileNotFoundException("Fichier non trouve : " + Args[1]);

            GestionBDD BDD=new GestionBDD();
            BDD.LectureExcel(Args[0], Args[1]);
        }
    }
}
