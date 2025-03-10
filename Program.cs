using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Lexico_3 {
    class Program {
        static void Main(string[] args) {
            bool internalMatrix = true;
            if (internalMatrix) {
                Console.WriteLine("Usando matriz interna");
            } else {
                Console.WriteLine("Usando matriz exel");
            }
            try {
                using (Lexico l = new Lexico("Prueba.cpp")) {
                    while (!l.finArchivo()) {
                        l.nexToken(internalMatrix);
                    }
                }
            } catch (Exception e) {
                Console.WriteLine("Error: " + e.Message);
            }
        }
    }
}