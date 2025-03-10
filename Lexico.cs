using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Data.Common;
using System.IO.Compression;
using Microsoft.VisualBasic;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace Lexico_3 {
    public class Lexico : Token,  IDisposable {
        const int F = -1;
        const int E = -2;
        int line = 1;
        StreamReader archivo;
        StreamWriter log;
        StreamWriter asm;
        XLWorkbook workbook;
        IXLWorksheet hoja;
        int[,] TRAND = {
            {  0,  1,  2, 33,  1, 12, 14,  8,  9, 10, 11, 23, 16, 16, 18, 20, 21, 26, 25, 27, 29, 32, 34,  0,  F, 33  },
            {  F,  1,  1,  F,  1,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  2,  3,  5,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  E,  E,  4,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E  },
            {  F,  F,  4,  F,  5,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  E,  E,  7,  E,  E,  6,  6,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E  },
            {  E,  E,  7,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E  },
            {  F,  F,  7,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F, 13,  F,  F,  F,  F,  F, 13,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F, 13,  F,  F,  F,  F, 13,  F,  F,  F,  F,  F,  F, 15,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 17,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 19,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 19,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 22,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 24,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 24,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 24,  F,  F,  F,  F,  F,  F, 24,  F,  F,  F,  F,  F,  F,  F  },
            { 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 28, 27, 27, 27, 27,  E, 27  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            { 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30  },
            {  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E,  E, 31,  E,  E,  E,  E,  E  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F, 32,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F  },
            {  F,  F,  F,  F,  F,  F,  F,  F,  F,  F,  F, 17, 36,  F,  F,  F,  F,  F,  F,  F,  F,  F, 35,  F,  F,  F  },
            { 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35, 35,  0, 35, 35  },
            { 36, 36, 36, 36, 36, 36, 36, 36, 36, 36, 36, 36, 37, 36, 36, 36, 36, 36, 36, 36, 36, 36, 36, 36, 36, 36  },
            { 36, 36, 36, 36, 36, 36, 35, 36, 36, 36, 36, 36, 37, 36, 36, 36, 36, 36, 36, 36, 36, 36,  0, 36, 36, 36  }
        };

        //Constructor de la clase lexico
        public Lexico() {
            log = new StreamWriter("prueba.log");
            asm = new StreamWriter("prueba.asm");
            workbook = new XLWorkbook("TRAND.xlsx");
            hoja = workbook.Worksheet(2);

            log.AutoFlush = true;
            asm.AutoFlush = true;

            if (File.Exists("prueba.cpp")) {
                archivo = new StreamReader("prueba.cpp");
            } else {
                throw new Error("El archivo prueba.cpp no existe",  log);
            }
        }
        //Constructor sobrecargado
        public Lexico(string nombreArchivo) {
            log = new StreamWriter(Path.ChangeExtension(nombreArchivo,  ".log"));
            log.AutoFlush = true;

            workbook = new XLWorkbook("TRAND.xlsx");
            hoja = workbook.Worksheet(2);

            if (File.Exists(Path.ChangeExtension(nombreArchivo,  ".cpp"))) {
                archivo = new StreamReader(nombreArchivo);
            } else {
                throw new Error("El archivo " + nombreArchivo + " no existe",  log);
            }

            if (Path.GetExtension(nombreArchivo) == ".cpp") {
                asm = new StreamWriter(Path.ChangeExtension(nombreArchivo,  ".asm"));
                asm.AutoFlush = true;
            } else {
                throw new Error("El archivo tiene extension invalida",  log);
            }
        }
        //Destructor de la clase lexico
        public void Dispose() {
            log.WriteLine("Total de lineas {0}",  line);

            log.Close();
            archivo.Close();
            asm.Close();
            //workbook.Dispose();
            //hoja.Delete();
        }
        private int columna(char c)
        {
            if (c == '\n') {
                return 23;
            } else if (char.IsWhiteSpace(c)) {
                return 0; 
            } else if (char.ToLower(c) == 'e') {
                return 4;
            } else if (char.IsLetter(c)) {
                return 1;
            }  else if (char.IsDigit(c)) {
                return 2;
            } else if (c == '.') {
                return 3;
            } else if (c == '+') {
                return 5;
            } else if (c == '-') {
                return 6;
            } else if (c == ';') {
                return 7;
            } else if (c == '{') {
                return 8;
            } else if (c == '}') {
                return 9;
            } else if (c == '?') {
                return 10;
            } else if (c == '=') {
                return 11;
            } else if (c == '*') {
                return 12;
            } else if (c == '%') {
                return 13;
            } else if (c == '&') {
                return 14;
            } else if (c == '|') {
                return 15;
            } else if (c == '!') {
                return 16;
            } else if (c == '<') {
                return 17;
            } else if (c == '>') {
                return 18;
            } else if (c == '"') {
                return 19;
            } else if (c == '\'') {
                return 20;
            } else if (c == '#') {
                return 21;
            } else if (c == '/') {
                return 22;
            } else if (finArchivo()) {
                return 24;
            }
            return 25;
        }
        private void clasificacion(int estado) {
            switch (estado) {
                case 1: 
                    setClasificacion(Tipos.Identificador); 
                    break;
                case 2: 
                    setClasificacion(Tipos.Numero); 
                    break;
                case 8: 
                    setClasificacion(Tipos.FinSentencia); 
                    break;
                case 9: 
                    setClasificacion(Tipos.InicioBloque); 
                    break;
                case 10:
                    setClasificacion(Tipos.FinBloque); 
                    break;
                case 11:
                    setClasificacion(Tipos.OperadorTernario); 
                    break;
                case 12:
                case 14:
                    setClasificacion(Tipos.OperadorTermino); 
                    break;
                case 13:
                    setClasificacion(Tipos.IncrementoTermino); 
                    break;
                case 15:
                    setClasificacion(Tipos.Puntero); 
                    break;
                case 16:
                case 34: 
                    setClasificacion(Tipos.OperadorFactor); 
                    break;
                case 17: 
                    setClasificacion(Tipos.IncrementoFactor); 
                    break;
                case 18:
                case 20:
                case 29:
                case 32:
                case 33: 
                    setClasificacion(Tipos.Caracter); 
                    break;
                case 19:
                case 21: 
                    setClasificacion(Tipos.OperadorLogico); 
                    break;
                case 22:
                case 24:
                case 25:
                case 26: 
                    setClasificacion(Tipos.OperadorRelacional); 
                    break;
                case 23: 
                    setClasificacion(Tipos.Asignacion); 
                    break;
                case 27: 
                    setClasificacion(Tipos.Cadena); 
                    break;
            }
        }
        public void nexToken(bool internalMatrix) {
            char c;
            string Buffer = "";
            int estado = 0;

            while (estado >= 0) {

                if (estado == 0)
                {
                    Buffer = "";
                }

                c = (char)archivo.Peek();
                if (internalMatrix) {
                    estado = TRAND[estado,  columna(c)];
                } else {
                    estado = usarExel(estado, c);
                }
                clasificacion(estado);

                if (estado >= 0) {
                    archivo.Read();
                    if (c == '\n') {
                        line++;
                    }
                    if (estado > 0) {
                        Buffer += c;
                    } else  {
                        Buffer = "";
                    }
                }
            }
            if (estado == E) {
                String mensaje;
                    if (getClasificacion() == Tipos.Numero) {
                        mensaje ="Lexico, Se espera un digito";
                    } else if (getClasificacion() == Tipos.Cadena) {
                        mensaje = "Lexico, Se esperaban comillas";
                    } else if (getClasificacion() == Tipos.Caracter) {
                        mensaje = "Lexico, Se esperaba una comilla";
                    } else {
                        mensaje = "Lexico, Se esperaba cierre de comentario";
                    }
                    throw new Error(mensaje, log, line);
                }

                setContenido(Buffer);
                log.WriteLine("{0}  °°°°  {1}",  getContenido(),  getClasificacion());   
        }
        public bool finArchivo() {
            return archivo.EndOfStream;
        }

        public int usarExel(int estado, char c) {
            string valor;
            IXLCell celda;

            celda = hoja.Cell(estado + 1, columna(c) + 1);
                    valor = celda.GetValue<String>();
                    if (valor == "F") {
                        estado = F;
                    } else if (valor == "E") {
                        estado = E;
                    } else {
                        estado = int.Parse(valor);
                    }

            return estado;
        }
    }
}
/*
    Expresion Regular: Metodo formal que através de una secuencia de caracteres 
    que define un PATRON de busqueda
    
        a) Reglas BNF
        b) Reglas BNF extendidas
        c) Operaciones aplicadas al lenguaje

    OAL

        1. Concatenación simple (·)
        2. Concatenación exponencial (^)
        3. Cerradura de Kleene (*)
        4. Cerradura positiva (+)
        5. Cerradura Epsilon (?)
        6. Operador OR (|)
        7. Parentesis()

        L = {A,  B,  C,  D,  E,  ...,  Z,  a,  b,  c,  d,  e,  ...,  z}
        D = {0,  1,  2,  3,  4,  5,  6,  7,  8,  9}

        1. L·D
            LD

        2. L^3 = LLL
           L^3D^2 = LLLDD

           D^5 = DDDDD
           =^2 = ==

        3. L* = Cero o más letras
           D* = Cero o más digítos

        4. L+ = Una o mas letras
           D+ = Uno o más digitos 

        5. L? = Cero o una letra (La letra es optativa-opcional)

        6. L | D = Letra o digito
           + | - = + ó - 

        7. (L D) L? = (Letra seguido de Digito y al final letra opcional)


    Producción Gramatical

        Clasificación del Token -> Expresión regular.

        Identificador -> L (L | D)* 

        Numero -> D+ (.D+)? (E(+ | -)? D+)?
        
        FinSentencia -> ;
        InicioBloque -> {
        FinBloque -> }
        OperadorTernario -> ?

        Puntero -> ->

        OperadorTermino -> + | -
        IncrementoTermino -> + (+ | =) | - (- | =)

        Termino + -> + (+ | =)?
        Termino - -> - (- | = | >)?

        OperadorFactor -> * | / | %
        IncrementoFactor -> *= | /= | %=

        Factor -> * | / | % (=)?
 
        OperadorLogico -> && | || | !

        NotOpRel -> ! (=)?
        
        Asignacion -> =

        AsigOpRel -> = (=)?

        OperadorRelacional -> > (=) ? | < (> | =)? | == | !=

        Cadena -> "c*"
        Caracter -> 'c' | #D* | lamda

    Automata: Modelo matematico que representa una expresion regular a travez de 
    un GRAFO,  para una maquina de estado finito que consiste en un conjunto de 
    estados bien definidos: 
        - Un estado inicial
        - Un alfabeto de entrada
        - Una funcion de transición
*/